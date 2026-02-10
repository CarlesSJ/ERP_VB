VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixes Muntadora"
   ClientHeight    =   11280
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   11850
   Icon            =   "muntadora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11280
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command23 
      BackColor       =   &H00F8FDB5&
      Height          =   495
      Left            =   75
      Picture         =   "muntadora.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Calendari"
      Top             =   9405
      Width           =   630
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H005C31DD&
      Caption         =   "Escullir COLOR que s'està muntant."
      Height          =   450
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5430
      Width           =   4680
   End
   Begin VB.Data bobines 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11025
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1785
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Data bobinesent 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   11715
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1365
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton botoensenyarpacking 
      Caption         =   "Command7"
      Height          =   195
      Left            =   11730
      TabIndex        =   57
      Top             =   975
      Visible         =   0   'False
      Width           =   75
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
      Left            =   15
      MaxLength       =   40
      TabIndex        =   54
      Tag             =   "888"
      Top             =   11085
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.TextBox compantone 
      DataField       =   "lot8"
      DataSource      =   "imppantones"
      Height          =   285
      Index           =   7
      Left            =   3210
      MaxLength       =   30
      TabIndex        =   53
      Tag             =   "888"
      Top             =   11085
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox kbpantone 
      DataField       =   "kg8"
      DataSource      =   "imppantones"
      Height          =   285
      Index           =   7
      Left            =   4560
      MaxLength       =   8
      TabIndex        =   52
      Tag             =   "1"
      Top             =   11085
      Visible         =   0   'False
      Width           =   550
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
      Left            =   15
      MaxLength       =   40
      TabIndex        =   51
      Tag             =   "888"
      Top             =   11340
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.TextBox compantone 
      DataField       =   "lot9"
      DataSource      =   "imppantones"
      Height          =   285
      Index           =   8
      Left            =   3210
      MaxLength       =   30
      TabIndex        =   50
      Tag             =   "888"
      Top             =   11340
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox kbpantone 
      DataField       =   "kg9"
      DataSource      =   "imppantones"
      Height          =   285
      Index           =   8
      Left            =   4560
      MaxLength       =   8
      TabIndex        =   49
      Tag             =   "1"
      Top             =   11340
      Visible         =   0   'False
      Width           =   550
   End
   Begin VB.CommandButton btinterfet 
      DisabledPicture =   "muntadora.frx":06E3
      DownPicture     =   "muntadora.frx":0C6D
      Height          =   330
      Index           =   0
      Left            =   11295
      Picture         =   "muntadora.frx":11F7
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Marcar/Desmarcar clixé muntat o no."
      Top             =   6240
      Width           =   525
   End
   Begin VB.CommandButton btinterfet 
      DisabledPicture =   "muntadora.frx":1781
      DownPicture     =   "muntadora.frx":1D0B
      Height          =   330
      Index           =   1
      Left            =   11295
      Picture         =   "muntadora.frx":2295
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Marcar/Desmarcar clixé muntat o no."
      Top             =   6570
      Width           =   525
   End
   Begin VB.CommandButton btinterfet 
      DisabledPicture =   "muntadora.frx":281F
      DownPicture     =   "muntadora.frx":2DA9
      Height          =   330
      Index           =   2
      Left            =   11295
      Picture         =   "muntadora.frx":3333
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Marcar/Desmarcar clixé muntat o no."
      Top             =   6900
      Width           =   525
   End
   Begin VB.CommandButton btinterfet 
      DisabledPicture =   "muntadora.frx":38BD
      DownPicture     =   "muntadora.frx":3E47
      Height          =   330
      Index           =   3
      Left            =   11295
      Picture         =   "muntadora.frx":43D1
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Marcar/Desmarcar clixé muntat o no."
      Top             =   7230
      Width           =   525
   End
   Begin VB.CommandButton btinterfet 
      DisabledPicture =   "muntadora.frx":495B
      DownPicture     =   "muntadora.frx":4EE5
      Height          =   330
      Index           =   4
      Left            =   11295
      Picture         =   "muntadora.frx":546F
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Marcar/Desmarcar clixé muntat o no."
      Top             =   7560
      Width           =   525
   End
   Begin VB.CommandButton btinterfet 
      DisabledPicture =   "muntadora.frx":59F9
      DownPicture     =   "muntadora.frx":5F83
      Height          =   330
      Index           =   5
      Left            =   11295
      Picture         =   "muntadora.frx":650D
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Marcar/Desmarcar clixé muntat o no."
      Top             =   7890
      Width           =   525
   End
   Begin VB.CommandButton btinterfet 
      DisabledPicture =   "muntadora.frx":6A97
      DownPicture     =   "muntadora.frx":7021
      Height          =   330
      Index           =   6
      Left            =   11295
      Picture         =   "muntadora.frx":75AB
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Marcar/Desmarcar clixé muntat o no."
      Top             =   8220
      Width           =   525
   End
   Begin VB.CommandButton btinterfet 
      DisabledPicture =   "muntadora.frx":7B35
      DownPicture     =   "muntadora.frx":80BF
      Height          =   330
      Index           =   7
      Left            =   11295
      Picture         =   "muntadora.frx":8649
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Marcar/Desmarcar clixé muntat o no."
      Top             =   8550
      Width           =   525
   End
   Begin VB.TextBox observacionstreball 
      BackColor       =   &H00F3B378&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   5325
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   9975
      Width           =   6285
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      Caption         =   "Ja està muntat"
      Height          =   360
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2730
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton bcomprafoam 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   -15
      Picture         =   "muntadora.frx":8BD3
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1170
      Width           =   570
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   825
      Top             =   1425
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton ordredelescomandes 
      Height          =   405
      Left            =   -15
      Picture         =   "muntadora.frx":915D
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   375
      Width           =   405
   End
   Begin VB.CommandButton botohistorial 
      Caption         =   "Historial"
      Height          =   360
      Left            =   8910
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   495
      Width           =   960
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Comanda"
      Height          =   375
      Left            =   10785
      Picture         =   "muntadora.frx":96E7
      TabIndex        =   30
      ToolTipText     =   "Imprimir Packing-List"
      Top             =   480
      Width           =   900
   End
   Begin VB.CommandButton exportarapdf 
      Caption         =   "Arxiu Imp"
      Height          =   375
      Left            =   9870
      Picture         =   "muntadora.frx":9C71
      TabIndex        =   29
      ToolTipText     =   "Imprimir Packing-List"
      Top             =   480
      Width           =   915
   End
   Begin VB.ComboBox provadhesiu 
      Height          =   315
      Left            =   10125
      TabIndex        =   26
      Top             =   2805
      Width           =   1470
   End
   Begin VB.ComboBox numcomanda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Height          =   360
      Left            =   5265
      Picture         =   "muntadora.frx":9F7B
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Eliminar el registre horari."
      Top             =   2730
      Width           =   405
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Afegir Horari "
      Height          =   360
      Left            =   4035
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2730
      Width           =   1185
   End
   Begin VB.TextBox firmat 
      Height          =   285
      Left            =   5475
      TabIndex        =   13
      Top             =   75
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2220
      TabIndex        =   12
      Top             =   345
      Width           =   615
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
      Height          =   390
      Left            =   315
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Escull Operari"
      Top             =   2715
      Width           =   3675
   End
   Begin VB.Timer rellotge 
      Left            =   90
      Top             =   2145
   End
   Begin VB.Data datalinies 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "W:\progcomandes\dades\baixes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8235
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "muntadorescilindres"
      Top             =   2295
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Data datamuntadora 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "W:\progcomandes\dades\baixes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -465
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "muntadores"
      Top             =   885
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Firma ---->"
      Height          =   270
      Left            =   4485
      TabIndex        =   8
      Top             =   15
      Width           =   1215
   End
   Begin VB.CommandButton comandanoacabada 
      BackColor       =   &H008080FF&
      Caption         =   "No Acabada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9975
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   30
      Width           =   1725
   End
   Begin VB.CommandButton comandaacabada 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Muntats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      Width           =   1830
   End
   Begin VB.TextBox observacionsgenerals 
      BackColor       =   &H00EAD9CE&
      Height          =   240
      Left            =   5295
      MaxLength       =   100
      TabIndex        =   3
      Top             =   9765
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Totals"
      Height          =   1050
      Left            =   300
      TabIndex        =   2
      Top             =   9930
      Width           =   4920
      Begin VB.CheckBox checkpeuimprenta 
         Caption         =   "Verif. peu imprenta"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   315
         TabIndex        =   48
         Top             =   795
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.TextBox gruixpolimer 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   435
         Width           =   990
      End
      Begin VB.TextBox cilindre 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   435
         Width           =   990
      End
      Begin VB.TextBox totalhores 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   435
         Width           =   990
      End
      Begin VB.TextBox totalpolimers 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   435
         Width           =   990
      End
      Begin VB.CheckBox clixesmuntats 
         Caption         =   "Clixes Muntats"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3210
         TabIndex        =   10
         Top             =   810
         Width           =   1590
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Gruix Polimer"
         Height          =   285
         Left            =   3810
         TabIndex        =   23
         Top             =   225
         Width           =   1425
      End
      Begin VB.Label ecilindre 
         Caption         =   "Cilindre"
         Height          =   285
         Left            =   2865
         TabIndex        =   21
         Top             =   225
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Total Hores"
         Height          =   285
         Left            =   315
         TabIndex        =   19
         Top             =   225
         Width           =   930
      End
      Begin VB.Label Label4 
         Caption         =   "Total Polimers"
         Height          =   285
         Left            =   1425
         TabIndex        =   18
         Top             =   225
         Width           =   1425
      End
   End
   Begin MSDBGrid.DBGrid reixalinies 
      Bindings        =   "muntadora.frx":A505
      Height          =   3435
      Left            =   315
      OleObjectBlob   =   "muntadora.frx":A51A
      TabIndex        =   1
      Top             =   5925
      Width           =   10950
   End
   Begin MSDBGrid.DBGrid reixamuntadora 
      Bindings        =   "muntadora.frx":B8F9
      Height          =   2085
      Left            =   315
      OleObjectBlob   =   "muntadora.frx":B911
      TabIndex        =   0
      Top             =   3315
      Width           =   11355
   End
   Begin VB.Label etobsmuntadoravella 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   5355
      TabIndex        =   59
      ToolTipText     =   "Dos Clics per editar les observacions de la muntadora vella."
      Top             =   10815
      Width           =   6270
   End
   Begin VB.Label etampladesmaterialanonim 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005C31DD&
      Height          =   330
      Left            =   585
      TabIndex        =   56
      Top             =   1950
      Width           =   11190
   End
   Begin VB.Label etaviscomandaprogramada 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Comanda programada per les 6:30 i encara no està muntada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   55
      Top             =   2235
      Visible         =   0   'False
      Width           =   11760
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Observacions Generals (TREBALL):"
      Height          =   180
      Left            =   5445
      TabIndex        =   39
      Top             =   9795
      Width           =   2985
   End
   Begin VB.Label comanda 
      Height          =   270
      Left            =   3060
      TabIndex        =   36
      Top             =   555
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label avispeu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NO POSAR PEU NI DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1365
      TabIndex        =   34
      Top             =   45
      Width           =   3255
   End
   Begin VB.Label novaorepetida 
      BackStyle       =   0  'Transparent
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
      Left            =   5775
      TabIndex        =   31
      Top             =   2745
      Width           =   2910
   End
   Begin VB.Label infocomanda2 
      BackStyle       =   0  'Transparent
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
      Height          =   870
      Left            =   600
      TabIndex        =   28
      Top             =   1140
      Width           =   10980
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveïdor Adhesiu"
      Height          =   210
      Left            =   10140
      TabIndex        =   27
      Top             =   2610
      Width           =   1470
   End
   Begin VB.Label descripciocomanda 
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
      ForeColor       =   &H00FF0000&
      Height          =   750
      Left            =   2955
      TabIndex        =   15
      Top             =   450
      Width           =   7185
   End
   Begin VB.Label nomfirma 
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
      Height          =   285
      Left            =   5760
      TabIndex        =   9
      Top             =   15
      Width           =   2145
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Comanda:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   375
      TabIndex        =   5
      Top             =   105
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Observacions Generals (Comanda):"
      Height          =   180
      Left            =   5490
      TabIndex        =   4
      Top             =   9570
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Menu mimpimiri 
      Caption         =   "Imprimir"
      Begin VB.Menu metiquetabossaclixe 
         Caption         =   "Etiqueta Bossa del Clixé"
      End
      Begin VB.Menu impetxlperlalleixa 
         Caption         =   "Etiqueta XL per la lleixa"
      End
      Begin VB.Menu llitatXLs 
         Caption         =   "Llistat de bosses dins dels XL"
      End
   End
   Begin VB.Menu mestocadhesius 
      Caption         =   "Estoc Adhesius"
   End
   Begin VB.Menu mcingularreal2 
      Caption         =   "Pujar CingularReal²"
   End
   Begin VB.Menu mllistatdecamises 
      Caption         =   "Llistat de camises"
      Begin VB.Menu llistatexcelonline 
         Caption         =   "Llistat excel online"
      End
      Begin VB.Menu mantcamises 
         Caption         =   "Manteniment camises de muntadora"
      End
   End
   Begin VB.Menu mimportardadesmuntadoravella 
      Caption         =   "Exportar treball muntadora vella a nova"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbsql As Database
Function comandavalida(numc As Double, msg As String, Optional nocomprovarllista As Boolean) As Boolean

End Function
Sub possarbotonstintersfets()
  Dim rst As Recordset
  Dim i As Byte
  
  Set rst = dbbaixes.OpenRecordset("Select * from muntadorescilindres where numcomanda=" + atrim(comanda) + " order by numcilindre")
  For i = 1 To 8
    btinterfet(i - 1).visible = False
    btinterfet(i - 1).tag = ""
    rst.FindFirst "numcilindre=" + atrim(i)
    If Not rst.NoMatch Then
       If atrim(rst!descripcio) <> "" Then
         btinterfet(i - 1).visible = True
         If IsNull(rst!datamuntatge) Then
             btinterfet(i - 1).Picture = btinterfet(i - 1).DownPicture
              btinterfet(i - 1).tag = "0"
              Else: btinterfet(i - 1).Picture = btinterfet(i - 1).DisabledPicture: btinterfet(i - 1).tag = "1"
         End If
       End If
    End If
  Next i
  Set rst = Nothing
End Sub
Sub passarlotsaprincipal()

End Sub
Function buscartreball(numc As String) As String
  Dim rst As Recordset
  buscartreball = 0
  Set rst = dbcomandes.OpenRecordset("select numtreball from comandes where comanda=" + atrim(cadbl(numc)))
  If Not rst.EOF Then buscartreball = atrim(rst!numtreball)
End Function
Sub possarvalorscomanda(vcomanda As String)
  Dim i As Byte
  For i = 1 To 6
     If llegir_ini("Valors", "comanda" + atrim(i), fitxerini) = "1" Then
       vcomanda = vcomanda + llegir_ini("Valors", "adhesiu" + atrim(i), fitxerini) + Chr(10)
     End If
  Next i
  If vcomanda <> "" Then vcomanda = "Comandes ja fetes de:" + Chr(10) + vcomanda
End Sub
Sub comprovarestatenvio()
  Dim vcomanda As String
  possarvalorscomanda vcomanda
  If llegir_ini("Valors", "enviat", fitxerini) = "esperantcomanda" Then
        bcomprafoam.visible = True
        bcomprafoam.tag = "Esperant que es faci la comanda de foam." + Chr(10) + vcomanda
        bcomprafoam.BackColor = &H8080FF
   End If
   If llegir_ini("Valors", "enviat", fitxerini) = "si" Then
        bcomprafoam.visible = True
        bcomprafoam.tag = "Comanda de foam feta." + Chr(10) + vcomanda
        bcomprafoam.BackColor = &HFF8080
   End If
   If llegir_ini("Valors", "enviat", fitxerini) = "no" Then
        bcomprafoam.visible = False
        bcomprafoam.tag = ""
   End If
       
End Sub

Private Sub bcomprafoam_Click()
 MsgBox bcomprafoam.tag, vbInformation, "Atenció"
End Sub

Private Sub botohistorial_Click()
  Dim numc As String
  Dim rst As Recordset
  Dim numtreball As Double
  If botohistorial.tag <> "" Then
     botohistorial.BackColor = exportarapdf.BackColor
     numcomanda = botohistorial.tag
     botohistorial.tag = ""
     Command4_Click
     Exit Sub
  End If
  'ratoli "espera"
  
  
  numc = InputBox("Entra la comanda que vols buscar historial." + Chr(10) + "PD. AQUESTA BUSQUEDA POT TRIGAR UNA MICA", "Historial comanda", numcomanda)
  If cadbl(numc) = 0 Then Exit Sub
  
  numtreball = cadbl(buscartreball(numc))
  If numtreball = 0 Then ratoli "normal": MsgBox "Aquesta comanda no te numero de treball.", vbCritical, "Error": Exit Sub
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  sql = "SELECT impressores.comanda, First(impressores.numeromaquina) AS Imp, format(First(impressores.datainici),'dd/mm/yy') AS Data, First(comandes.numtreball) AS Treball, Last(comandes.numordremodificacio) AS Ordre FROM impressores RIGHT JOIN comandes ON impressores.comanda = comandes.comanda GROUP BY impressores.comanda, impressores.tipus Having (((First(comandes.numtreball)) = " + atrim(numtreball) + ") And ((impressores.tipus) = 'f')) ORDER BY impressores.comanda DESC , Last(comandes.numordremodificacio) DESC;"

  'formseleccio.Data1.RecordSource = "select comanda from comandes where numtreball=" + atrim(cadbl(idtreball)) + " order by comanda Desc"
  formseleccio.Data1.RecordSource = sql
  formseleccio.caption = "Historial"
  
  'formseleccio.Width = 7000
  ratoli "espera"
  formseleccio.refrescar
  ratoli "normal"
  formseleccio.DBGrid2.Columns(0).width = 2220
  formseleccio.DBGrid2.Columns(1).width = 720
  formseleccio.DBGrid2.Columns(2).width = 2370
  formseleccio.DBGrid2.Columns(3).width = 1500
  formseleccio.DBGrid2.Columns(4).width = 800
'  formseleccio.DBGrid2.Columns(5).width = 800
  formseleccio.Show 1
  If seleccioret = 1 Then
   botohistorial.tag = numcomanda
   botohistorial.BackColor = QBColor(12)
   numcomanda = formseleccio.Data1.Recordset!comanda
   Command4_Click
  End If
  Unload formseleccio
   ratoli "normal"
End Sub

Public Sub btinterfet_Click(Index As Integer)
  Dim v As String
  'si s'ha apretat CANCELAR deixar el clixe com a no preparat
  If btinterfet(Index).tag = "1" Then
     datalinies.Recordset.FindFirst "numcilindre=" + atrim(Index + 1)
     If Not datalinies.Recordset.NoMatch Then
         datalinies.Recordset.Edit
         datalinies.Recordset!op = 0
         datalinies.Recordset!datamuntatge = Null
         datalinies.Recordset.Update
     End If
     GoTo fi
  End If
  'si s'ha apretat OK deixar el clixe com a no preparat
  If btinterfet(Index).tag = "0" Then
     datalinies.Recordset.FindFirst "numcilindre=" + atrim(Index + 1)
     If Not datalinies.Recordset.NoMatch Then
         reixalinies.col = 4
         reixalinies.row = Index
         If atrim(reixalinies.Text) = "" Then
            reixalinies_ButtonClick reixalinies.col
         End If
         If cadbl(datalinies.Recordset!numpolimers) = 0 Then v = InputBox("Quants polimers hi ha en aquest cilindre?", "Quants", 1)
        
         datalinies.Recordset.Edit
         datalinies.Recordset!op = numop
         datalinies.Recordset!datamuntatge = Now
         If cadbl(v) > 0 Then datalinies.Recordset!numpolimers = cadbl(v)
         datalinies.Recordset.Update
     End If
     GoTo fi
  End If
fi:
  possarbotonstintersfets
End Sub

Private Sub cilindre_LostFocus()
gravartotals
End Sub

Sub possardatafi()
 Dim hores As Double
 If datamuntadora.Recordset.EOF Or datamuntadora.Recordset.BOF Then Exit Sub
 guardarcanvisreixa
 datamuntadora.Recordset.MoveLast
 If Not datamuntadora.Recordset.EOF Then
    If Not IsDate(datamuntadora.Recordset!datafi) And Not IsDate(datamuntadora.Recordset!horafi) Then
        reixamuntadora.Columns("datafi") = Date
        reixamuntadora.Columns("horafi") = Time
    End If
    possarhorestreballades
 End If
 If datamuntadora.Recordset.EditMode = 0 Then datamuntadora.Recordset.Edit
 datamuntadora.Recordset.Update
End Sub
Sub guardarcanvisreixa()
  On Error GoTo errorguardant
 
  reixamuntadora.CurrentCellModified = True
  reixamuntadora.EditActive = False
  
  If Not sidatacontrolcontedades(datamuntadora) Then Exit Sub
   datamuntadora.Recordset.Move 0
  Exit Sub
errorguardant:
   MsgBox "Hi ha hagut un error guardant la baixa", vbCritical, "Error..."
   If datamuntadora.Recordset.EditMode > 0 Then
      datamuntadora.Recordset.CancelUpdate
      datamuntadora.Refresh
   End If
End Sub
Function sidatacontrolcontedades(data As Control) As Boolean
   On Error GoTo err
   sidatacontrolcontedades = False
   If Not datamuntadora.Recordset.EOF Or Not datamuntadora.Recordset.BOF Then
     If datamuntadora.Recordset.EOF Then datamuntadora.Recordset.MoveLast
     If datamuntadora.Recordset.BOF Then datamuntadora.Recordset.MoveFirst
     sidatacontrolcontedades = True
   End If
   Exit Function
err:
End Function
Sub imprimir_fulla()
  'Command4_Click
  calculartotals
  crear_taula_muntadora_baixa
  Set rsttmp = datamuntadora.Database.OpenRecordset("tmp_muntadora_baixa")
  rsttmp.AddNew
  imp_possardadesgenerals
  imp_possardadesfuncionament
  imp_possardadesliniespolimers
  rsttmp.Update
  wait 2
  imp_tirarelllistat
  Set rsttmp = Nothing
End Sub

Sub imp_tirarelllistat()
    Dim resp As String
     Dim oapp As CRAXDDRT.Application
     Dim oreport As CRAXDDRT.Report
     Dim vnomfitxerpdf As String
     Dim vcarpetadesti As String
     
     Set oapp = New CRAXDDRT.Application
     Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "baixesmuntadora_pdf.rpt", 1)
     oreport.Database.Tables.Item(1).Location = cami
     oreport.FormulaFields.GetItemByName("numeromuntadora").Text = "'" + atrim(nummaq) + "-" + atrim(llegir_ini("Baixes", "nommaquina", "comandes.ini")) + "'"
     'oreport.RecordSelectionFormula = "mid({Clixes.ubicacio},1,5)<>'Palet' and {Clixes.arxiu}<>'' and isnull({Clixes.databaixaclixe}) and {clixes.estatclixe}<>'RETORNEM CLIXES'"
     'oreport.RecordSelectionFormula = "{@arxiusenseXL}>0 and (trim({Clixes.arxiu})<>'' and isnull({Clixes.databaixaclixe}) and {clixes.estatclixe}<>'RETORNEM CLIXES')"
     oreport.DiscardSavedData
     
     escriure_ini "General", "exportantpdfs", "si", llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\organitzar.ini"
     crearlacarpetaperexportar cadbl(numcomanda.Text), vcarpetadesti
  
     vnomfitxerpdf = vcarpetadesti + "\" + atrim(numcomanda.Text) + "_BaixaMuntadora.pdf"
     
     oreport.ExportOptions.DestinationType = crEDTDiskFile
     oreport.ExportOptions.FormatType = crEFTPortableDocFormat
     oreport.ExportOptions.DiskFileName = vnomfitxerpdf
     oreport.ExportOptions.PDFExportAllPages = True
     oreport.Export False
     escriure_ini "General", "exportantpdfs", "no", llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\organitzar.ini"
     
   'trec imprimir perque ara ja no volen imprimir
'     oreport.PrintOut False, 1
'     wait 1
'     oreport.PrintOut False, 1
   
   

End Sub
Function generar_baixapdf() As String

End Function
Sub crearlacarpetaperexportar(numc As Double, carpetadesti As String)
   Dim carpetaprincipal As String
   Dim vcarpetatemporal As String
   Dim vubicaciocarpetadesti As String
   Dim vnomfitxer As String
   vcarpetatemporal = rutadelfitxer(llegir_ini("General", "cami", "comandes.ini"))
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

Sub imp_tirarelllistat_8_5()
llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "baixesmuntadora.rpt"
' llistat.Destination = crptToWindow
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 2
 llistat.DataFiles(0) = cami
 llistat.DiscardSavedData = True
' llistat.PrinterName = llegir_ini("Impressores", "nomfulla", "baixesimpressora.ini")
' llistat.PrinterPort = llegir_ini("Impressores", "portfulla", "baixesimpressora.ini")
' llistat.PrinterDriver = llegir_ini("Impressores", "driverfulla", "baixesimpressora.ini")
  DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 llistat.Action = 1
End Sub
Sub imp_possardadesliniespolimers()
   Dim rst As Recordset
   Dim rstl As Recordset
   Set rst = datamuntadora.Database.OpenRecordset("select * from muntadorescilindres where numcomanda=" + atrim(cadbl(numcomanda)))
   Set rstl = datamuntadora.Database.OpenRecordset("select * from tmp_muntadora_baixa_bob ")
   While Not rst.EOF
    If cadbl(rst!gruixadhesiu) > 0 Then
     rstl.AddNew
     rstl!comanda = cadbl(numcomanda)
     rstl!operari = cadbl(rst!op)
     rstl!descripcio = atrim(rst!descripcio) + " "
     rstl!gruix = cadbl(rst!gruixadhesiu)
     rstl!nomadhesiu = atrim(rst!nomadhesiu) + " "
     rstl!numeropolimers = cadbl(rst!numpolimers)
     rstl!numeroremuntats = cadbl(rst!remuntats)
     rstl!observacions = atrim(rst!observacio) + " "
     rstl.Update
    End If
    rst.MoveNext
   Wend
End Sub
Sub imp_possardadesgenerals()
  Dim rst As Recordset
  Dim rstc As Recordset
  Set rst = datamuntadora.Database.OpenRecordset("select * from muntadoratot where comanda=" + atrim(cadbl(numcomanda)))
  If rst.EOF Then Exit Sub
  Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(numcomanda)))
  If rstc.EOF Then Exit Sub
  With rsttmp
  !tipusimpresio = "Comanda " + atrim(rst!tipusimpresio)
  !nomfirmat = atrim(rst!nomfirma)
  !firmat = atrim(rst!firma)
  !comanda = atrim(rst!comanda)
  !client = atrim(rst!nomclient)
  !texteimp = IIf(atrim(rstc!marcailinia) = "", atrim(rstc!texteimpressio), atrim(rstc!marcailinia))
  !codibarres = atrim(rstc!codibarras)
  !comandaacavada = clixesmuntats.Value
  !hfunc = cadbl(rst!totalhores)
  !tpolimers = cadbl(rst!totalpolimers)
  !cilindre = cadbl(rst!cilindre)
  !gruixpolimer = cadbl(rst!gruixpolimer)
  !observacionsgenerals = atrim(rst!observacions)
  End With
End Sub
Sub imp_possardadesfuncionament()
  Dim rst As Recordset
 'temps funcionament
  Set rst = datamuntadora.Database.OpenRecordset("select id,comanda,operari1,datainici,horainici,datafi,horafi,observacio,totalhores from muntadores where comanda=" + atrim(cadbl(numcomanda.Text)))
  If Not rst.EOF Then
    rst.MoveLast
   If Not rst.BOF Then rst.MovePrevious:
   If rst.BOF Then
      rst.MoveNext
    Else: rst.MovePrevious: If rst.BOF Then rst.MoveNext Else rst.MovePrevious: If rst.BOF Then rst.MoveNext
   End If
  End If
  i = 1
  With rsttmp
  While Not rst.EOF
    .Fields("tempsmunt_data" + Trim(i)) = Format(atrim(rst!datainici), "dd/mm/yy")
    .Fields("tempsmunt_op" + Trim(i)) = cadbl(rst!operari1)
    .Fields("tempsmunt_de" + Trim(i)) = Format(atrim(rst!horainici), "hh:nn")
    .Fields("tempsmunt_fins" + Trim(i)) = Format(atrim(rst!horafi), "hh:nn")
    .Fields("tempsmunt_observacio" + Trim(i)) = atrim(rst!observacio)
    '.Fields("tempsmunt_thores" + Trim(i)) = atrim(rst!totalhores)
    i = i + 1
    rst.MoveNext
  Wend
  Set rst = Nothing
  End With
End Sub
Function existeixlataula(vnomtaula As String) As Boolean
  Dim rstp As Recordset
  On Error GoTo errortaula
  existeixlataula = True
  Set rstp = dbbaixes.OpenRecordset("select * from " + vnomtaula)
  Set rstp = Nothing
  On Error GoTo 0
  Exit Function
errortaula:
  existeixlataula = False
  On Error GoTo 0
End Function

Sub crear_taula_muntadora_baixa()
  Dim camps As String
  Dim campscapcalera As String
  Dim camps2 As String
  Dim campscapcalera2 As String
  Dim campspantone As String
  Dim campstotal As String
  Set rsttmp = Nothing
  
  campsextra = " nomfirmat text,firmat text,"
  campscapcalera = " comanda double, client string,comandaacavada byte, observacionsgenerals string,texteimp string,codibarres string,"
  
  camps3 = camps3 + "tempsmunt_data1 string,tempsmunt_op1 byte,tempsmunt_de1 string, tempsmunt_fins1 string,tempsmunt_mtrsmin1 double,tempsmunt_mtrscola1 double,tempsmunt_mtrslaminats1 double,tempsmunt_observacio1 string,"
  camps3 = camps3 + "tempsmunt_data2 string,tempsmunt_op2 byte,tempsmunt_de2 string, tempsmunt_fins2 string,tempsmunt_mtrsmin2 double,tempsmunt_mtrscola2 double, tempsmunt_mtrslaminats2 double,tempsmunt_observacio2 string,"
  camps2 = "tempsmunt_data3 string,tempsmunt_op3 byte,tempsmunt_de3 string, tempsmunt_fins3 string,tempsmunt_mtrsmin3 double,tempsmunt_mtrscola3 double, tempsmunt_mtrslaminats3 double,tempsmunt_observacio3 string,"
  camps2 = camps2 + "tempsmunt_data4 string,tempsmunt_op4 byte,tempsmunt_de4 string, tempsmunt_fins4 string,tempsmunt_mtrsmin4 double,tempsmunt_mtrscola4 double, tempsmunt_mtrslaminats4 double,tempsmunt_observacio4 string"
  campstotal = ",hfunc double, tpolimers double, cilindre double,gruixpolimer double,tipusimpresio string"
  
  'On Error Resume Next
  ' datamuntadora.Database.Execute "drop table tmp_muntadora_baixa"
  ' datamuntadora.Database.Execute "drop table tmp_muntadora_baixa_bob"
  'On Error GoTo 0
  If Not existeixlataula("tmp_muntadora_baixa") Then
      datamuntadora.Database.Execute ("create table tmp_muntadora_baixa (" + campsextra + campscapcalera + campscapcalera2 + camps + camps3 + camps2 + campspantone + campstotal + ")")
        Else: datamuntadora.Database.Execute "delete * from tmp_muntadora_baixa"
  End If
  If Not existeixlataula("tmp_muntadora_baixa_bob") Then
    datamuntadora.Database.Execute ("create table tmp_muntadora_baixa_bob (comanda double,idbob integer,operari byte,descripcio string,gruix double,nomadhesiu string,numeropolimers integer,numeroremuntats integer,observacions string)")
       Else: datamuntadora.Database.Execute "delete * from tmp_muntadora_baixa_bob"
  End If
  wait 2
End Sub
Sub mirarsidemanardeguardarcomabona()
   Dim rstc As Recordset
   Dim numtreball As Double
   If comprovarsiguardaronocomabona(numtreball) Then
       If MsgBox("Aquesta comanda ha tingut canvis respecta a la ultima vegada que es va muntar." + Chr(10) + "Vols guardar aquests parametres com a correctes per la próxima vegada?", vbInformation + vbYesNo, "Atenció") = vbYes Then
            Set rstc = dbbaixes.OpenRecordset("SELECT comandes.comanda, comandes.numtreball, muntadoratot.valorsbons FROM comandes INNER JOIN muntadoratot ON comandes.comanda = muntadoratot.comanda Where (((comandes.numtreball) = " + atrim(numtreball) + ")) ORDER BY muntadoratot.valorsbons,comandes.comanda;")
            While Not rstc.EOF
              If cadbl(rstc!comanda) <> cadbl(numcomanda) And rstc!valorsbons = True Then
                 rstc.Edit
                 rstc!valorsbons = False
                 rstc.Update
              End If
              If cadbl(rstc!comanda) = cadbl(numcomanda) Then
                 rstc.Edit
                 rstc!valorsbons = True
                 rstc.Update
              End If
              rstc.MoveNext
            Wend
   
       End If
   End If
End Sub

Private Sub comandaacabada_Click()
  Dim jaestaacavada As Boolean
  Dim vverificaciopeu As Boolean
  If Not sidatacontrolcontedades(datamuntadora) Then MsgBox "No hi ha linia de treball.", vbCritical, "Error": Exit Sub
  If firmat = "" Then MsgBox "Primer has de firmar la fulla.", vbCritical, "Error": Exit Sub
  vverificaciopeu = verificacio_peuimprenta
  If clixesmuntats = 1 Then jaestaacavada = True
  
  possardatafinal
  passarcomandaacabada
  calculartotals
  mirarsidemanardeguardarcomabona
  'Command4_Click
  verificacio_netejaidespeje
  If jaestaacavada Then
     If MsgBox("Vols tornar a imprimir la baixa?", vbInformation + vbYesNo + vbDefaultButton2, "Imprimir") = vbYes Then
        imprimir_fulla
     End If
    Else: imprimir_fulla
  End If
  'imprimirfullnetejaiendreçar id_treball, ordremodificacio
  numcomanda_DropDown
End Sub
Function verificacio_peuimprenta() As Boolean
   Dim resp As String
   verificacio_peuimprenta = False
   If avispeu.visible Then
     While resp <> ""
       resp = InputBox("Aquesta comanda porta una descripcio de PEU D'IMPRENTA." + Chr(10) + "REVISA-LA I ESCRIU [PEU IMPRENTA] PER CONTINUAR")
       If StrPtr(resp) = 0 Then Exit Function
     Wend
   End If
   verificacio_peuimprenta = True
End Function

Sub verificacio_netejaidespeje()
  'Dim v As String
  'Dim vcont As Byte
  'vcont = 9
  'While UCase(v) <> "NETEJA" And vcont > 0
    MsgBox "Verificació de Neteja i despeje de línia." + Chr(10) + "Fes ACCEPTAR per continuar.", vbExclamation + vbOKOnly, "Neteja i despeje (" + atrim(vcont) + ")"
   ' vcont = vcont - 1
  'Wend
End Sub
Sub possardatafinal()
'datamuntadora.UpdateControls
'datamuntadora.Recordset.MoveLast
Dim hores As Double
'On Error Resume Next
 guardarcanvisreixa
 wait 1
 If Not sidatacontrolcontedades(datamuntadora) Then Exit Sub
 datamuntadora.Recordset.MoveLast
If Not datamuntadora.Recordset.EOF Then
'    datamuntadora.UpdateControls
 '   datamuntadora.Recordset.MoveLast
    If Not IsDate(datamuntadora.Recordset!datafi) And Not IsDate(datamuntadora.Recordset!horafi) Then
       ' datamuntadora.Recordset.Edit
       ' datamuntadora.Recordset!datafi = Date
       ' datamuntadora.Recordset!horafi = Time
       ' datamuntadora.Recordset.Update
        reixamuntadora.EditActive = True
        reixamuntadora.Columns("datafi") = Date
        reixamuntadora.Columns("horafi") = Time
        hores = DateDiff("n", CVDate(atrim(reixamuntadora.Columns("datainici")) + " " + atrim(reixamuntadora.Columns("horainici"))), CVDate(atrim(reixamuntadora.Columns("datafi")) + " " + atrim(reixamuntadora.Columns("horafi"))))
        hores = Redondejar(hores / 60, 2)
        reixamuntadora.Columns("totalhores") = hores
        reixamuntadora.EditActive = False
        guardarcanvisreixa
    End If
End If
End Sub

Sub passarcomandaacabada()
  clixesmuntats.Value = 1
  gravartotals
End Sub

Private Sub comandanoacabada_Click()
  If Not sidatacontrolcontedades(datamuntadora) Then MsgBox "No hi ha linia de treball.", vbCritical, "Error": Exit Sub
  clixesmuntats = 0
  possardatafinal
  gravartotals
  imprimir_fulla
  numcomanda_DropDown
End Sub


Sub agegircomandaperbaixar(numc As Double, data As Date, urgent As Boolean, observacio As String)
  'dbstocks.Execute "insert into bobinesperbaixar (comanda,dataavis,urgent,observacio) values (" + atrim(numc) + ",#" + Format(data, "mm/dd/yy hh:nn") + "#," + IIf(urgent, "True", "False") + ",'" + treure_apostruf(observacio) + "')"
End Sub

Private Sub Command1_Click()
' If datamuntadora.Recordset.BOF And datamuntadora.Recordset.EOF Then
'    agegircomandaperbaixar cadbl(numcomanda), Now, False, ""
' End If
 possardatafinal
 crear_seccio
 reixamuntadora.SetFocus
End Sub
Sub crear_seccio()
  datamuntadora.Recordset.AddNew
  datamuntadora.Recordset!comanda = numcomanda
  datamuntadora.Recordset!numeromaquina = nummaq
  datamuntadora.Recordset!operari1 = numop
  datamuntadora.Recordset!datainici = Format(Date, "dd/mm/yy")
  datamuntadora.Recordset!horainici = Format(Time, "hh:nn")
  datamuntadora.Recordset.Update
  datamuntadora.Recordset.MoveLast

End Sub

Private Sub Command2_Click()
   If datamuntadora.Recordset.EOF And datamuntadora.Recordset.BOF Then MsgBox "No hi ha cap linia d'horaris sel.leccionada per borrar.", vbCritical, "Atenció": Exit Sub
   If datamuntadora.Recordset.EOF Then datamuntadora.Refresh
   If UCase(InputBox("Per eliminar aquest registre escriu [ELIMINAR].", "Eliminacio del registre horari.")) = "ELIMINAR" Then
       datamuntadora.Recordset.Delete
       calculartotals
   End If
End Sub
Function datainicialitzat(datac As Control) As Boolean
   datainicialitzat = True
   On Error GoTo fi
   If datac.Recordset.EOF Then Exit Function
   Exit Function
fi:
   datainicialitzat = False
End Function

Sub calculartotals()
  Dim hores As Double
  Dim thores As Double
  Dim rstp As Recordset
  If cadbl(numcomanda.tag) = 0 Then Exit Sub
  If Not datainicialitzat(datamuntadora) Then Exit Sub
  If datamuntadora.Recordset.EOF Or datamuntadora.Recordset.BOF Then Exit Sub
  datamuntadora.Recordset.MoveFirst
  While Not datamuntadora.Recordset.EOF
   On Error Resume Next
    hores = DateDiff("n", CVDate(atrim(reixamuntadora.Columns("datainici")) + " " + atrim(reixamuntadora.Columns("horainici"))), CVDate(atrim(reixamuntadora.Columns("datafi")) + " " + atrim(reixamuntadora.Columns("horafi"))))
    hores = Redondejar(hores / 60, 2)
    If hores < 0.01 Then hores = 0
    datamuntadora.Recordset.Edit
    datamuntadora.Recordset!totalhores = hores
    datamuntadora.Recordset.Update
    thores = thores + hores
    datamuntadora.Recordset.MoveNext
  Wend
  On Error GoTo 0
  Set rstp = datalinies.Database.OpenRecordset("select sum(numpolimers) as polimers from muntadorescilindres where numcomanda=" + atrim(cadbl(numcomanda)))
  totalpolimers = "0"
  If Not rstp.EOF Then totalpolimers = cadbl(rstp!polimers)
  totalhores = atrim(thores)
  gravartotals
End Sub

Private Sub Command23_Click()
  ratoli "espera"
  Shell "\\serverprodu\Dades\progcomandes\aplicacio\CalendariThunderbird\ThunderbirdPortable\ThunderbirdPortable.exe", vbMaximizedFocus
  wait 2
  ratoli "normal"
End Sub

Private Sub Command3_Click()
   firmar_fulla
   gravartotals
End Sub
Sub firmar_fulla()
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
    If cadbl(firmat) > 0 Then
      Set rsttmp = dbcomandes.OpenRecordset("select codi,descripcio from operaris where actiu=1 and maquina='M' and codi=" + atrim(cadbl(firmat)))
      If rsttmp.EOF Then nomfirma = "": firmat = "": MsgBox "Aquest operari no existeix": Exit Sub
      nomfirma = rsttmp!descripcio
    End If
    'guarda_totals
End Sub

Private Sub Command4_Click()
   'hola
     Dim numtreball As Double
    ' Dim odremodificacio As Double
     Dim novacomanda As String
     Dim vmuntadora As Double
     Dim rst As Recordset
     Dim vjafeta As Boolean
     
     If cadbl(numcomanda) = 0 Then Exit Sub
     If Not comprovarsicomandacorrecte(cadbl(numcomanda)) Then MsgBox "Aquesta comanda no està dins a la llista de pendents o muntades", vbCritical, "Atenció": Exit Sub
     If Not clixesentratsafabrica(cadbl(numcomanda)) Then MsgBox "Aquesta comanda no te els CLIXES ENTRATS a disseny. No es poden utilitzar.", vbCritical, "Atenció": Exit Sub
     Set rst = dbtmpb.OpenRecordset("select * from muntadoratot where comanda=" + atrim(numcomanda))
     If Not rst.EOF Then If rst!acabada Then vjafeta = True
     If Not vjafeta Then
      If Not ordremuntatge.comandavalida(cadbl(numcomanda), True, numtreball, ordremodificacio) Then
         MsgBox "Aquesta comanda ESTÀ PARADA O NO ESTÀ APUNT PER SER IMPRESA", vbCritical, "Atenció"
         Exit Sub
         'If InputBox("Entra el numero de comanda de nou per entrar-hi sabent que no està apunt per imprimir.", "Atenció") <> numcomanda Then Exit Sub
      End If
     End If
     'comprovar si hi ha alguna altra comanda impresa amb la mateixa referencia
     id_treball = numtreball
     comprovarsihihaunaltrereferenciaperimprimir cadbl(numcomanda), numtreball
     comprovarsihihaentratlarxiuimuntadora cadbl(numcomanda)
     imprimiretbossatreball numtreball, True
     dbclixes.Execute "update clixes set ubicacio='' where id_treball=" + atrim(numtreball)
     novacomanda = numcomanda
     numcomanda = numcomanda.tag
     calculartotals
     numcomanda.tag = novacomanda
     numcomanda = novacomanda
     formannex.carregarcomanda cadbl(comanda)
     carregarcomanda
     
     'Command6.visible = False
     possarbotonstintersfets
     If datamuntadora.Recordset.EOF And datamuntadora.Recordset.BOF Then Command1_Click ':Command6.visible = True
     'ara no es fa servir cingular real
     'If Label2.tag = "F2" Then mirarsihihaCingularReal_MUNT numtreball, cadbl(ordremodificacio), cadbl(numcomanda)
     possarbotonstintersfets
End Sub
Function mirarsihihaCingularReal(vnumtreball As Double, vordremodificacio As Double) As Boolean
   Dim vurl As String
   Dim generarfitxer_pdf As String
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(vnumtreball, "00000") + "\pdf" + Format(vnumtreball, "00000") + "-" + Format(vordremodificacio, "000") + "_CR.pdf"
   If existeix(generarfitxer_pdf) Then
      mirarsihihaCingularReal = True
   End If
   
   
End Function

Sub mirarsihihaCingularReal_MUNT(vnumtreball As Double, vordremodificacio As Double, Optional vnumcomanda As Double)
   Dim vurl As String
   Dim generarfitxer_pdf As String
   vurl = "http://192.168.10.242/user/remote/pdfx/upload"
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(vnumtreball, "00000") + "\pdf" + Format(vnumtreball, "00000") + "-" + Format(vordremodificacio, "000") + "_CR.pdf"
   If existeix(generarfitxer_pdf) Then
     If MsgBox("Aquesta comanda te Cingular Real2" + Chr(10) + "Vols pujar el fitxer ara a la impresora?", vbInformation + vbYesNo, "Atenció") = vbYes Then
         crearinetejardirectoritemporal
         Copiar_Fitxer generarfitxer_pdf, "c:\temp\cingularreal\" + atrim(vnumtreball) + "-" + atrim(vordremodificacio) + ".pdf"
         obrir_document vurl
     End If
   End If
   
   
End Sub
Sub crearinetejardirectoritemporal()
   On Error Resume Next
   MkDir "c:\temp\"
   MkDir "c:\temp\cingularreal"
   Kill "c:\temp\cingularreal\*.*"
End Sub
Function demanararxiu(Optional vsuggerirXL As Boolean) As String
  Dim vControl As Control
  Dim vnumXLsuggerit As String
  vnumXLsuggerit = suggerirXL
  Unload formescullirlleixa
  Load formescullirlleixa
  formescullirlleixa.cxl = vnumXLsuggerit
  formescullirlleixa.cxl.tag = "1"
  formescullirlleixa.Top = Form1.Top + ((Form1.Height - formescullirlleixa.Height) / 2)
  formescullirlleixa.Left = Form1.Left + ((Form1.width - formescullirlleixa.width) / 2)
  'formescullirlleixa.Left = formscrooll.Left + formcomandes.Left
  formescullirlleixa.Show 1
  If seleccioret = 1 Then demanararxiu = formescullirlleixa.valorescullit
  
End Function
Function suggerirXL() As String
  Dim rst As Recordset
  Dim vamplada As Double
  Dim vXL As String
  Dim vQtreballsXRbossa As Integer
  Dim vcriteri As String
  vQtreballsXRbossa = 7
  While vamplada <> 55 And vamplada <> 64 And vamplada <> 80
    vamplada = cadbl(InputBox("Entra la emplada de la bossa amb CMs (55cm, 64cm o 80cm)", "Ampla de la bossa XL"))
  Wend
  vXL = " instr(1,trim(Clixes.arxiu),'XL')>0 and isnull(Clixes.databaixaclixe) and instr(1,trim(Clixes.ubicacio),'P')=0 "
  
  'BOSSA PETITA
  If vamplada = 55 Then vcriteri = "(NumXL>=55 and NumXL<=72) or (NumXL>=136 and NumXL<=262)"
  'BOSSA GRAN
  If vamplada = 80 Then vcriteri = "(NumXL>=1 and NumXL<=54) or (NumXL>=73 and NumXL<=90) or (NumXL>=484 and NumXL<=492)"
  'BOSSA NORMAL
  If vamplada = 64 Then vcriteri = "(not (NumXL>=55 and NumXL<=72) and not ((NumXL>=1 and NumXL<=54) or (NumXL>=73 and NumXL<=90) or (NumXL>=484 and NumXL<=492)))"
  vcriteri = " HAVING " + substituirtot(vcriteri, "NumXL", "(CDbl(Mid(Trim(First(Clixes.arxiu)),4)))")
  Set rst = dbclixes.OpenRecordset("SELECT count(Clixes.arxiu) as Q,first(clixes.arxiu) as Parxiu, cdbl(mid(trim(first(Clixes.arxiu)),4)) as NumXL FROM Clixes WHERE " + vXL + " GROUP BY Clixes.arxiu " + vcriteri + " ORDER BY Count(trim(arxiu)) asc;")
  Clipboard.Clear
  Clipboard.SetText "SELECT count(Clixes.arxiu) as Q,first(clixes.arxiu) as Parxiu, cdbl(mid(trim(first(Clixes.arxiu)),4)) as NumXL FROM Clixes WHERE " + vXL + " GROUP BY Clixes.arxiu " + vcriteri + " ORDER BY Count(Clixes.arxiu) asc;"
  
  If Not rst.EOF Then suggerirXL = rst!NumXL
fi:
  Set rst = Nothing
End Function
Sub comprovarsihihaentratlarxiuimuntadora(comanda As Double)
    Dim rst As Recordset
    Dim rstc As Recordset
    Dim vmuntadora As String
    Dim varxiu As String
    Set rstc = dbcomandes.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(comanda))
    If rstc.EOF Then Exit Sub
    Set rst = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(rstc!numtreball))
    If Not rst.EOF Then
     If atrim(rst!arxiu) = "" Then
      rst.Edit
      'While atrim(rst!arxiu) = ""
         varxiu = demanararxiu(True)
         rst!arxiu = varxiu
      'Wend
      rst.Update
     End If
    End If
    Set rst = dbclixes.OpenRecordset("select codimuntadora from clientsvinculats where id_treball=" + atrim(rstc!numtreball) + " and ordremodificacio=" + atrim(rstc!numordremodificacio))
    If Not rst.EOF Then
     If atrim(rst!codimuntadora) = "" Then
     ' While vmuntadora = ""
         vmuntadora = atrim(InputBox("Entra el numero de muntadora." + Chr(10) + "SENSE NUMERO DE MUNTADORA NO POTS CONTINUAR", "Entra la muntadora corresponent"))
      'Wend
      dbclixes.Execute "update clientsvinculats set codimuntadora='" + atrim(vmuntadora) + "' where id_Treball=" + atrim(rstc!numtreball) + " and ordremodificacio=" + atrim(rstc!numordremodificacio)
      wait 3
     End If
    End If
    'actualitzo larxiu de muntadora i larxiu de la comanada
    If varxiu <> "" Or vmuntadora <> "" Then
        Set rst = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(comanda))
        If Not rst.EOF Then
            rst.Edit
            If varxiu <> "" Then rst!arxiu = atrim(varxiu)
            If vmuntadora <> "" Then rst!arxiumontadora = atrim(vmuntadora)
            rst.Update
        End If
    End If
    'dbcomandes.Execute "update comandes set  arxiu='" + varxiu + "' and arxiumontadora='" + vmuntadora + "' where comanda=" + atrim(comanda)
    Set rst = Nothing
    Set rstc = Nothing
End Sub
Sub comprovarsihihaunaltrereferenciaperimprimir(comanda As Double, numtreball As Double)
   Dim rstc As Recordset
   Dim c As String
   If numtreball = 0 Then Exit Sub
   Set rstc = dbcomandes.OpenRecordset("SELECT comandes.comanda, comandes_extres.passaraimpresores fROM comandes LEFT JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda  where comandes.proximaseccio='E' and comandes.numtreball=" + atrim(numtreball))
   c = ""
   While Not rstc.EOF
     If cadbl(rstc!comanda) <> comanda And passaraimpresores = 1 Then c = c + atrim(rstc!comanda) + " "
     rstc.MoveNext
   Wend
   If c <> "" Then MsgBox "Hi ha mes comandes pendents d'imprimir amb aquest treball." + Chr(10) + c, vbInformation, "Atenció"
End Sub
Function clixesentratsafabrica(numc As Double) As Boolean
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim rutaclixes As String
   
   Dim ordrem As Integer
   clixesentratsafabrica = False
   rutaclixes = rutadelfitxer(cami) + "clixesnous.mdb"
   'rutaclixes = rutadelfitxer(cami) + "clixes.mdb"
   'Set dbclixes = OpenDatabase(rutaclixes)
   Set rstc = dbcomandes.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(numc))
   If rstc.EOF Then Exit Function
   ordrem = cadbl(rstc!numordremodificacio)
   If ordrem = 0 Then ordrem = 1
   Set rst = dbclixes.OpenRecordset("select id_estatclixe from clixes_modifi where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordremodificacio=" + atrim(ordrem) + " order by ordre DESC")
   If rst.EOF Then Exit Function
   If rst!id_estatclixe = 8 Then clixesentratsafabrica = True
   Set rst = Nothing
   Set rstc = Nothing
   'Set dbclixes = Nothing
End Function
Function comprovarsicomandacorrecte(numc As Double) As Boolean
   Dim rst As Recordset
   Label2.tag = ""
   Label2 = "Comanda: "
   comprovarsicomandacorrecte = True
   Set rst = dbbaixes.OpenRecordset("select * from muntadora_ordremuntatge where comanda=" + atrim(numc))
   If rst.EOF Then
       Set rst = dbbaixes.OpenRecordset("select * from muntadoratot where comanda=" + atrim(numc))
       If rst.EOF Then
          comprovarsicomandacorrecte = False
       End If
       Else:
         Label2.caption = "Comanda: " + atrim(rst!nummaquina)
         Label2.tag = atrim(rst!nummaquina)
   End If

   
End Function
Sub possar_peu_imprenta(denvio As Long)
  Dim rstd As Recordset
  avispeu = ""
  If denvio > 0 Then
      Set rstd = dbcomandes.OpenRecordset("SELECT Clients_envios.codi, peuimprenta.descripcio FROM Clients_envios LEFT JOIN peuimprenta ON Clients_envios.peuimprenta = peuimprenta.codi where clients_envios.id=" + atrim(denvio))
      If Not rstd.EOF Then avispeu = atrim(rstd!descripcio)
      If avispeu <> "" Then MsgBox """Oju"" amb el peu d'imprenta. " + Chr(10) + avispeu, vbExclamation + vbOKOnly, "Atenció"
  End If
  Set rstd = Nothing
End Sub
Sub carregarcomanda()
   Dim rstc As Recordset
   Dim rstt As Recordset
   Dim rstobs As Recordset
   Dim ampleutil As Double
   Dim amplebobina As Double
   ampleutil = 0
   etobsmuntadoravella = ""
   observacionstreball = ""
   etobsmuntadoravella.tag = ""
   Set rstc = dbcomandes.OpenRecordset("SELECT comandes.marcailinia,comandes.comanda,comandes.numtreball, comandes.ampleutil,comandes.impressio,comandes.amplereb,comandes.arxiu,comandes.cantitatex, comandes.ampleesq,comandes.arxiumontadora,comandes.refclient,comandes.codibarras,comandes.texteimpressio, clients.nom,comandes.client, productes.ruta,comandes.direnvio FROM (comandes INNER JOIN clients ON comandes.client = clients.codi) INNER JOIN productes ON comandes.producte = productes.codi Where comanda = " + atrim(cadbl(numcomanda)))
   If rstc.EOF Then MsgBox "Aquesta comanda no existeix", vbCritical: Exit Sub
   Set rstt = dbclixes.OpenRecordset("SELECT Clixes.*, Clientsvinculats.codimuntadora, Clixes.id_treball, Clientsvinculats.codimuntadora FROM Clixes INNER JOIN Clientsvinculats ON Clixes.id_treball = Clientsvinculats.id_treball WHERE (((Clixes.id_treball)=" + atrim(rstc!numtreball) + ") AND ((Clientsvinculats.codimuntadora)<>''));")
   If rstt.EOF Then
      MsgBox "No s'ha trobat el treball " + atrim(rstc!numtreball)
      datamuntadora.RecordSource = "SELECT * from muntadores where comanda=-999999"
      datamuntadora.Refresh
      Exit Sub
   End If
   If InStr(1, rstc!ruta, "I") = 0 Then MsgBox "Aquest comanda no te seccio de Impresores no hi ha res a Muntar.", vbCritical, "Atenció": Exit Sub
   If InStr(1, rstc!ruta, "R") > 0 Then ampleutil = cadbl(rstc!amplereb)
   If InStr(1, rstc!ruta, "L") > 0 Then ampleutil = cadbl(rstc!ampleutil)
   amplebobina = amplemaxcomanda(rstc!comanda, 0, 0)
   possar_peu_imprenta cadbl(rstc!direnvio)
   descripciocomanda.caption = atrim(rstc!nom) + Chr(10) + IIf(atrim(rstc!marcailinia) = "", atrim(rstc!texteimpressio), atrim(rstc!marcailinia))
   descripciocomanda.tag = atrim(rstc!client) + " - " + atrim(rstc!nom)
   etampladesmaterialanonim = llistaampladesbobinesespackinglist(rstc!comanda)
   infocomanda2 = "Ref: " + atrim(rstc!refclient) + "    CodiBarres:" + atrim(rstt!codidebarres) + Chr(10) + "Ample:" + atrim(cadbl(rstc!ampleesq)) + "  Ample Util: " + atrim(ampleutil) + "  Ample Bob: " + atrim(amplebobina) + " Quantitat: " + Format(cadbl(rstc!cantitatex), "#,##0") + "mts " + Chr(10) + "NºMun:" + atrim(rstt!codimuntadora) + "     Arxiu:" + atrim(rstt!arxiu)
   novaorepetida = tipusimpresio(rstc!impressio)
   idtreball = cadbl(rstc!numtreball)
   observacionsgenerals = ""
   observacionstreball = ""
   datamuntadora.RecordSource = "select * from muntadores where comanda=" + atrim(cadbl(numcomanda)) + " order by datainici,horainici"
   datamuntadora.Refresh
  
   carregartotals
   calculartotals
   possarbotonstintersfets
   Set rstobs = dbtmpb.OpenRecordset("select * from muntadora_obsmuntadoravella where numtreball=" + atrim(rstc!numtreball))
   If Not rstobs.EOF Then etobsmuntadoravella = treurecanvisdelinia(atrim(rstobs!observacions)): etobsmuntadoravella.tag = atrim(rstc!numtreball)
   If observacionstreball = "" And etobsmuntadoravella <> "" Then
       observacionstreball = etobsmuntadoravella: observacionstreball_LostFocus
   End If
   Set rstc = Nothing
   Set rstt = Nothing
   Set rstobs = Nothing
End Sub
Function treurecanvisdelinia(v As String) As String
   v = substituir(v, Chr(10), " ")
   treurecanvisdelinia = v
End Function

Function llistaampladesbobinesespackinglist(vnumc As Double) As String
   Dim rst As Recordset
   Dim rstopcions As Recordset
   Dim v As String
   Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb", , True)
'   Clipboard.Clear
'   Clipboard.SetText "SELECT distinct palet.ample FROM materials RIGHT JOIN (Palets LEFT JOIN (Bobines LEFT JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)) ON Palets.Idpalet = Bobines.Idpalet) ON materials.codi = Palets.codimatprognou WHERE (((Parcials.comanda)='" + atrim(vnumc) + "'));"
   Set rst = dbstocks.OpenRecordset("SELECT distinct palets.ample FROM materials RIGHT JOIN (Palets LEFT JOIN (Bobines LEFT JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)) ON Palets.Idpalet = Bobines.Idpalet) ON materials.codi = Palets.codimatprognou WHERE (((Parcials.comanda)='" + atrim(vnumc) + "'));")
   If Not rst.EOF Then rst.MoveLast: rst.MoveFirst
   While Not rst.EOF
      v = v + "Ample " + atrim(rst.AbsolutePosition + 1) + " : " + atrim(rst!ample) + "cm "
      rst.MoveNext
   Wend
   Set rst = Nothing
   If v <> "" Then
      llistaampladesbobinesespackinglist = "Packinglist " + v
       Else
        Set rstopcions = dbstocks.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(numcomanda))
        If Not rstopcions.EOF Then llistaampladesbobinesespackinglist = "Packinglist d'estoc " + atrim(rstopcions!grupdestoc)
   End If
   Set rstopcions = Nothing
   Set rst = Nothing
End Function
Function amplemaxcomanda(numc, numc2, numc3) As Double
  Dim rstc As Recordset
  If numc = 0 Then numc = -1
  If numc2 = 0 Then numc2 = -1
  If numc3 = 0 Then numc3 = -1
  Set rstc = dbstocks.OpenRecordset("SELECT Max(Palets.Ample) AS amplemax FROM Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet GROUP BY CDbl([comanda]) HAVING (CDbl([comanda])=" + atrim(numc) + " or cdbl([comanda])=" + atrim(numc3) + " or cdbl([comanda])=" + atrim(numc2) + ");")
  If Not rstc.EOF Then
    amplemaxcomanda = rstc!amplemax
  End If
  Set rstc = dbstocks.OpenRecordset("SELECT opcionsdajust.comanda, opcionsdajust.grupdestoc, grupsdepalets.ample FROM (grupsdepalets RIGHT JOIN opcionsdajust ON grupsdepalets.numerogrup = opcionsdajust.grupdestoc) LEFT JOIN Palets ON grupsdepalets.paletexemple = Palets.Idpalet WHERE (((opcionsdajust.comanda)=" + atrim(CDbl(numc)) + "));")
  If cadbl(rstc!ample) > 0 Then amplemaxcomanda = cadbl(rstc!ample)
  Set rstc = Nothing
End Function
Function tipusimpresio(tipus As String) As String
  Select Case tipus
     Case "M"
       tipusimpresio = "Modificada"
     Case "R"
       tipusimpresio = "Repetida"
     Case "C"
       tipusimpresio = "Canvi de direccio"
        Case "N"
       tipusimpresio = "Nova"
  End Select
End Function
Function ultimacomandafetabona(numtreball As Double, numcomanda As Double) As String
   Dim rstc As Recordset
   ultimacomandafetabona = "-9999"
   Set rstc = dbbaixes.OpenRecordset("SELECT comandes.comanda, comandes.numtreball, muntadoratot.valorsbons FROM comandes INNER JOIN muntadoratot ON comandes.comanda = muntadoratot.comanda Where (((comandes.numtreball) = " + atrim(numtreball) + ") and comandes.comanda<>" + atrim(numcomanda) + ") ORDER BY muntadoratot.valorsbons,comandes.comanda;")
   If rstc.EOF Then Exit Function
   ultimacomandafetabona = rstc!comanda
   
End Function
Sub crearliniescilindres()
  Dim i As Byte
  Dim rstc As Recordset
  Dim rstmtultima As Recordset
  Dim nomtinta As String
  Dim vnomcomplerttinta As String
  Dim vdetall As String
  Dim rstclixes As Recordset
  Dim rstclixes2 As Recordset
  Dim rstrepas As Recordset
  Dim rstcl As Recordset
  Dim vnomlinkat As String
  Dim rstlinkatbossa As Recordset
  Dim vbossaclixes As String
  Dim vidtreballbossa As String
  Dim vnumbandes As Long
  Dim vnomadhesiu As String
  Dim vidadhesiu As Long
  
  If datamuntadora.Recordset.EOF Then GoTo fi
  
  Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(numcomanda)))
  If rstc.EOF Then Exit Sub
  Set rstclixes = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(rstc!numtreball) + " and ordremodificacio=" + atrim(rstc!numordremodificacio), , ReadOnly)
  If rstclixes.EOF Then MsgBox "Error treball no trobat": GoTo fi
  Set rstcl = dbclixes.OpenRecordset("Select arxiu from clixes where id_treball=" + atrim(rstclixes!id_treball))
  Set rstrepas = dbclixes.OpenRecordset("select * from repasclixes where comanda=" + atrim(numcomanda))
  vnumbandes = 1:  If Not rstrepas.EOF Then vnumbandes = cadbl(rstrepas!numbandes)
  If Not rstrepas.EOF Then Set rstrepas = dbclixes.OpenRecordset("select * from repasdadestintes where id_repas=" + atrim(rstrepas!id_repas))
  Set rstmtultima = dbbaixes.OpenRecordset("select * from muntadorescilindres where numcomanda=" + ultimacomandafetabona(cadbl(rstc!numtreball), cadbl(numcomanda)) + " order by numcilindre")
  For i = 1 To 8
     vnomlinkat = ""
     rstclixes.FindFirst "ordretinter=" + atrim(i)
     If Not rstclixes.NoMatch Then
        If cadbl(rstclixes!tinterlinkambid_treball) > 0 Then
           Set rstclixes2 = dbclixes.OpenRecordset("SELECT Tintes.id_tinter, Clixes.id_treball, Clixes.arxiu FROM Tintes LEFT JOIN Clixes ON Tintes.id_treball = Clixes.id_treball WHERE (((Tintes.id_tinter)=" + atrim(rstclixes!tinterlinkambid_treball) + "))")
           If Not rstclixes2.EOF Then
              vnomlinkat = "T:" + atrim(rstclixes2!id_treball) + " " + atrim(rstclixes2!arxiu)
           End If
             Else:
               vbossaclixes = atrim(rstcl!arxiu)
               vidtreballbossa = atrim(rstclixes!id_treball)
               If cadbl(rstclixes!tinterlinkambid_treball) < 0 Then
                  Set rstlinkatbossa = dbclixes.OpenRecordset("SELECT Tintes.id_tinter, Clixes.id_treball, Clixes.arxiu FROM Tintes LEFT JOIN Clixes ON Tintes.id_treball = Clixes.id_treball WHERE (((Tintes.id_tinter)=" + atrim(rstclixes!tinterlinkambid_treball * -1) + "))")
                  If Not rstlinkatbossa.EOF Then
                     vbossaclixes = atrim(rstlinkatbossa!arxiu)
                     vidtreballbossa = atrim(rstlinkatbossa!id_treball)
                  End If
               End If
               vnomlinkat = "T:" + vidtreballbossa + " " + atrim(vbossaclixes)
        End If
          
     End If
    
     vdetall = atrim(rstc.Fields("detalltinter" + atrim(i)).Value)
     nomtinta = atrim(rstc.Fields("tinta" + atrim(i) + "a").Value)
     vnomcomplerttinta = nomtinta + IIf(vdetall <> "", "(" + vdetall + ")", "")
'     If nomtinta <> "" Then
       datalinies.Recordset.AddNew
       If Len(vnomcomplerttinta) > datalinies.Recordset.Fields("descripcio").Size Then vnomcomplerttinta = Mid(vnomcomplerttinta, 1, datalinies.Recordset.Fields("descripcio").Size)
       
       datalinies.Recordset!descripcio = nomtinta + IIf(vdetall <> "", "(" + vdetall + ")", "")
       datalinies.Recordset!numcilindre = atrim(i)
       datalinies.Recordset!numcomanda = numcomanda
       datalinies.Recordset!id_tinter = cadbl(rstclixes!id_tinter)
       'valor per defecte del gruixdahesiu i tipus adhesiu
       If nomtinta <> "" Then
            vidadhesiu = 0
            vnomadhesiu = ""
            If Not rstrepas.EOF Then
               rstrepas.FindFirst "ordretinter=" + atrim(i)
               If Not rstrepas.NoMatch Then
                    buscar_adhesiuescullit rstrepas, vidadhesiu, vnomadhesiu
               End If
            End If
            datalinies.Recordset!gruixadhesiu = 0.5
            datalinies.Recordset!numpolimers = vnumbandes
            datalinies.Recordset!idadhesiu = vidadhesiu
            datalinies.Recordset!nomadhesiu = vnomadhesiu
       End If
       If Not rstmtultima.EOF Then
           rstmtultima.FindFirst "DESCRIPCIO='" + atrim(datalinies.Recordset!descripcio) + "'"
           If Not rstmtultima.NoMatch And nomtinta <> "" Then
                datalinies.Recordset!numpolimers = cadbl(rstmtultima!numpolimers)
                datalinies.Recordset!idadhesiu = cadbl(rstmtultima!idadhesiu)
                datalinies.Recordset!nomadhesiu = atrim(rstmtultima!nomadhesiu)
                datalinies.Recordset!gruixadhesiu = rstmtultima!gruixadhesiu
                datalinies.Recordset!observacio = rstmtultima!observacio
           End If
       End If
       If nomtinta <> "" And InStr(1, atrim(datalinies.Recordset!observacio), vnomlinkat) = 0 And vnomlinkat <> "" Then datalinies.Recordset!observacio = vnomlinkat
       datalinies.Recordset.Update
       'If Not rstmtultima.EOF Then rstmtultima.MoveNext
  Next i
fi:
  Set rstclixes = Nothing
  Set rstclixes2 = Nothing
  Set rstrepas = Nothing
  Set rstcl = Nothing
End Sub
Sub buscar_adhesiuescullit(vrst As Recordset, vidadhesiu As Long, vnomadhesiu As String)
   Dim vtipusfoam As String
   Dim rstadhesius As Recordset
   vtipusfoam = atrim(vrst!tipusdefoam)
   Set rstadhesius = dbbaixes.OpenRecordset("select * from muntadoratot where comanda=" + atrim(comanda))
   If Not rstadhesius.EOF Then
       Set rstadhesius = dbbaixes.OpenRecordset("select * from adhesiusmuntadora where codiproveidor=" + atrim(cadbl(rstadhesius!proveidoradhesiu)) + " and inicialsfoam='" + vtipusfoam + "'")
       If Not rstadhesius.EOF Then
           vidadhesiu = cadbl(rstadhesius!codiintern)
           vnomadhesiu = atrim(rstadhesius!descripcioinplacsa)
       End If
   End If
   Set rstadhesius = Nothing
End Sub
Function comprovarsiguardaronocomabona(numtreball As Double) As Boolean   'torna true si s'ha de preguntar per guardarla
  Dim i As Byte
  Dim rstc As Recordset
  Dim rstmtultima As Recordset
  Dim nomtinta As String
  Dim canvis As Boolean
  'If datamuntadora.Recordset.EOF Then Exit Function
  Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(numcomanda)))
  If rstc.EOF Then Exit Function
  Set rstmtultima = dbbaixes.OpenRecordset("select * from muntadorescilindres where numcomanda=" + ultimacomandafetabona(cadbl(rstc!numtreball), cadbl(numcomanda)) + " order by numcilindre")
  numtreball = rstc!numtreball
  canvis = False
  datalinies.Recordset.MoveFirst
  While Not rstmtultima.EOF And Not datalinies.Recordset.EOF
          'If rstmtultima!descripcio = nomtinta Then
           If datalinies.Recordset!numpolimers <> cadbl(rstmtultima!numpolimers) Then canvis = True
           If datalinies.Recordset!idadhesiu <> cadbl(rstmtultima!idadhesiu) Then canvis = True
           rstmtultima.MoveNext
           datalinies.Recordset.MoveNext
  Wend
  comprovarsiguardaronocomabona = canvis
  Set rstc = Nothing
  Set rstmtultima = Nothing
End Function

Sub carregartotals()
   Dim rstt As Recordset
   Dim rstp As Recordset
   Dim rstobs As Recordset
   Set rstt = dbbaixes.OpenRecordset("select * from muntadoratot where comanda=" + atrim(cadbl(numcomanda)))
   If rstt.EOF Then
       creartotalsnous
       Set rstt = dbbaixes.OpenRecordset("select * from muntadoratot where comanda=" + atrim(cadbl(numcomanda)))
   End If
   Set rstobs = dbbaixes.OpenRecordset("select * from muntadores_obstreballs where numtreball=" + atrim(id_treball))
   If rstobs.EOF Then
      dbbaixes.Execute "insert into muntadores_obstreballs (numtreball,observacions) values (" + atrim(id_treball) + ",'')"
      Set rstobs = dbbaixes.OpenRecordset("select * from muntadores_obstreballs where numtreball=" + atrim(id_treball))
   End If
   Set rstp = dbcomandes.OpenRecordset("select * from proveidors where codi=" + atrim(cadbl(rstt!proveidoradhesiu)))
   If rstp.EOF Or cadbl(rstt!proveidoradhesiu) = 0 Then
       'provadhesiu = "3M ESPAÑA S.L"
      ' provadhesiu.tag = "569"
         escullir_proveidor_adhesiu
        Else
          provadhesiu = atrim(rstp!nom)
          provadhesiu.tag = atrim(rstp!codi)
   End If
   cilindre = rstt!cilindre
   gruixpolimer = rstt!gruixpolimer
   totalpolimers = cadbl(rstt!totalpolimers)
   clixesmuntats = IIf(rstt!acabada, 1, 0)
   checkpeuimprenta = IIf(rstt!verificaciopeu, 1, 0)
   firmat = rstt!firma
   nomfirma = atrim(rstt!nomfirma)
   observacionsgenerals = atrim(rstt!observacions)
   observacionstreball = atrim(rstobs!observacions)
   checkpeuimprenta.visible = avispeu.visible
   Set rstt = Nothing
   Set rstobs = Nothing
End Sub
Sub creartotalsnous()
   Dim rstmt As Recordset
   Dim rstc As Recordset

   Set rstc = dbcomandes.OpenRecordset("select cilindres,gruixpol from comandes where comanda=" + atrim(numcomanda))
   Set rstmt = dbbaixes.OpenRecordset("muntadoratot")
   
   rstmt.AddNew
   rstmt!comanda = numcomanda
   rstmt!cilindre = cadbl(rstc!cilindres)
   rstmt!gruixpolimer = cadbl(rstc!gruixpol)
   rstmt.Update
   Set rstmt = Nothing
   
End Sub
Sub gravartotals()
   Dim rstmt As Recordset
   Set rstmt = dbbaixes.OpenRecordset("select * from muntadoratot where comanda=" + atrim(cadbl(numcomanda)))
   If rstmt.EOF Then Exit Sub
   
   rstmt.Edit
   rstmt!proveidoradhesiu = cadbl(provadhesiu.tag)
   rstmt!cilindre = cadbl(cilindre)
   rstmt!gruixpolimer = cadbl(gruixpolimer)
   rstmt!totalhores = cadbl(totalhores)
   rstmt!observacions = atrim(observacionsgenerals)
   rstmt!totalpolimers = cadbl(totalpolimers)
   rstmt!acabada = clixesmuntats
   rstmt!firma = firmat
   rstmt!nomfirma = nomfirma
   rstmt!nomclient = descripciocomanda.tag
   rstmt!tipusimpresio = novaorepetida
   rstmt!verificaciopeu = checkpeuimprenta
   rstmt.Update
   Set rstmt = Nothing
   
End Sub

Private Sub Command5_Click()
  Dim numc As String
  
  numc = InputBox("Entra la comanda que vols visualitzar.", "Visualitzar comanda", numcomanda)
  If cadbl(numc) = 0 Then Exit Sub
  escriure_ini "Baixes", "imprimircomanda", numc, "comandes.ini"
  Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", "comandes.ini")) + "comandes.exe - imprimir", vbHide
  missatgevist.Show 1

End Sub

Private Sub Command6_Click()
  Dim vnumc As Double
  vnumc = cadbl(InputBox("Entra el numero de comanda on s'han muntat aquests clixes.", "Ja estan muntats"))
  If vnumc > 0 Then
      dbbaixes.Execute "update muntadoratot set comandajamuntada=" + atrim(vnumc) + " where comanda=" + numcomanda
      datalinies.Recordset.MoveLast
      datalinies.Recordset.MoveFirst
      While Not datalinies.Recordset.EOF
        reixalinies_Change
        datalinies.Recordset.MoveNext
      Wend
      DoEvents
      comandaacabada_Click
  End If
  
End Sub
Sub borrarfetes()
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim rsttintes As Recordset
  Dim dbtintes As Database
  'proces anulat
  Exit Sub
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  
  Set rsttintes = dbtintes.OpenRecordset("select * from comandesactives where gestionat='M'")
  Set rst = dbbaixes.OpenRecordset("select * from muntadora_ordremuntatge")
  
  While Not rst.EOF
    Set rst2 = dbbaixes.OpenRecordset("SELECT muntadoratot.comanda, comandes.comanda, muntadoratot.acabada, comandes.proximaseccio FROM muntadoratot RIGHT JOIN comandes ON muntadoratot.comanda = comandes.comanda WHERE comandes.comanda=" + atrim(rst!comanda))
    If Not rst2.EOF Then
      rsttintes.FindFirst "comanda=" + atrim(rst!comanda)
      If (rsttintes.NoMatch And rst2!acabada) Or (rst2!proximaseccio = "T") Then
         rst.Delete
      End If
    End If
    rst.MoveNext
  Wend
  'BORRO TOTES LES RECLAMADES A OFICINES TAMBÉ
  dbbaixes.Execute "DELETE planificacio_reclamades.numcomanda, comandes.proximaseccio, planificacio_reclamades.* FROM planificacio_reclamades INNER JOIN comandes ON planificacio_reclamades.numcomanda = comandes.comanda WHERE (((comandes.proximaseccio)='T'));"

  Set rst = Nothing
  Set rst2 = Nothing
  Set rsttintes = Nothing
  Set dbtintes = Nothing
End Sub

Private Sub Command7_Click()
  Dim vnumcolor As Double
  Dim vnomcolor As String
  'Dim seleccioret As Integer
  seleccioret = 1
 While seleccioret = 1
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(camicomandes) + "baixes.mdb"
  formseleccio.Data1.RecordSource = "select numcilindre as [Nº],descripcio as COLOR,gruixadhesiu,nomadhesiu  from muntadorescilindres where datamuntatge=null and descripcio<>'' and numcomanda=" + atrim(cadbl(numcomanda)) + " order by numcilindre"
  formseleccio.caption = "ESCULLIR COLOR"
  formseleccio.refrescar
  If formseleccio.Data1.Recordset.EOF Then GoTo fi
  formseleccio.width = 17500
  formseleccio.DBGrid2.width = formseleccio.width - Command1.width - 150
  formseleccio.Command1.Left = formseleccio.width - Command1.width
  formseleccio.Command3.Left = formseleccio.width - Command3.width
  formseleccio.Left = (Screen.width / 2) - (formseleccio.width / 2)
  formseleccio.DBGrid2.Columns("gruixadhesiu").visible = False
  formseleccio.DBGrid2.Columns("nomadhesiu").visible = False
  formseleccio.Show 1
  If seleccioret = 1 Then
        vnumcolor = formseleccio.Data1.Recordset![Nº]
        vnomcolor = formseleccio.Data1.Recordset![color]
        Load formesperant
        formesperant.etcolor = atrim(vnumcolor) + "-" + atrim(vnomcolor)
        formesperant.bfet.tag = atrim(vnumcolor)
        formesperant.etadhesiu = atrim(formseleccio.Data1.Recordset!gruixadhesiu) + "-" + atrim(formseleccio.Data1.Recordset!nomadhesiu)
        Unload formseleccio
        formesperant.Show 1
  End If
 Wend
fi:
Unload formseleccio
End Sub

Private Sub datamuntadora_Reposition()
   datalinies.RecordSource = "select * from muntadorescilindres where numcomanda=" + atrim(cadbl(numcomanda)) + " order by numcilindre"
   datalinies.Refresh
   If datalinies.Recordset.EOF Then crearliniescilindres
End Sub

Private Sub etobsmuntadoravella_DblClick()
  Dim v As String
  If cadbl(etobsmuntadoravella.tag) = 0 Then Exit Sub
  v = InputBox("Modifica el texte:", "Modificacio", etobsmuntadoravella)
  If Len(v) > 1 Then
       dbbaixes.Execute "update muntadora_obsmuntadoravella set observacions='" + treure_apostruf(v) + "' where numtreball=" + atrim(cadbl(etobsmuntadoravella.tag))
       etobsmuntadoravella = v
  End If
End Sub

Private Sub exportarapdf_Click()
 
veureelimp

End Sub


Sub veureelimp()
   Dim rstc As Recordset
  Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(numcomanda)))
  obrir_imp_treball cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), cadbl(rstc!client), cadbl(rstc!direnvio)
End Sub
Sub obrir_imp_treball(treball As Double, modificacio As Double, codiclient As Double, direnvio As Double)
   Dim generarfitxer_imp As String
   If modificacio = 0 Then modificacio = 1
   generarfitxer_imp = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\IMP" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "-" + Format(codiclient, "000000") + "_" + atrim(direnvio) + ".doc"
   If existeix(generarfitxer_imp) Then
     obrir_document generarfitxer_imp
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + generarfitxer_imp, vbCritical, "Error"
  End If
End Sub


Sub imprimirfullnetejaiendreçar(vnumtreball As Double, vordre As Double)
   Dim rst As Recordset
   Dim oapp As CRAXDDRT.Application
   Dim oreport As CRAXDDRT.Report
 
   Set rst = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(vnumtreball))
   If rst.EOF Then GoTo fi
  
   Set oapp = New CRAXDDRT.Application
   Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "fullnetejaiendreçar.rpt", 1)
   'oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
   'oreport.RecordSelectionFormula = "{diferenciescomandaitreball.comanda}=" + atrim(numc)
   'oreport.DiscardSavedData
   oreport.FormulaFields.GetItemByName("numtreball").Text = "'" + atrim(vnumtreball) + "/" + atrim(vordre) + "'"
   oreport.FormulaFields.GetItemByName("arxiu").Text = "'" + atrim(rst!arxiu) + "'"
   oreport.FormulaFields.GetItemByName("texte").Text = "'" + atrim(rst!marca) + "-" + atrim(rst!linia) + "'"
   Set rst = dbclixes.OpenRecordset("select ordretinter,color from tintes where id_treball=" + atrim(vnumtreball) + " and ordremodificacio=" + atrim(vordre))
   If rst.EOF Then GoTo fi
   While Not rst.EOF
     oreport.FormulaFields.GetItemByName("color" + atrim(rst!ordretinter)).Text = "'" + atrim(rst!color) + "'"
     rst.MoveNext
   Wend
   If existeix("c:\ordprog.ini") Then
    Load veurereport
    veurereport.CRViewer.ReportSource = oreport
    veurereport.CRViewer.DisplayGroupTree = False
    veurereport.CRViewer.ViewReport
    veurereport.WindowState = 2
    veurereport.Show 1
     Else
       oreport.PrintOut False, 1
   End If
fi:
  Set rst = Nothing
   
End Sub
Sub comprovarsihihamissatgesCHAT()

End Sub

Private Sub Form_Activate()
   escriure_ini "Baixes", "imprimircomanda", "0", "comandes.ini"
   If Not existeix("c:\ordprog.ini") Then
    If cadbl(llegir_ini("Baixes", "programaamaquina", "comandes.ini")) = 1 Then assignardecimalipunt
   End If
   comprovarestatenvio
   
End Sub
Sub imprimir_packinglistTICKET(vnumc As Double, v As Boolean)

End Sub
Private Sub Form_Click()
'Command6_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
tempseditant = Now
End Sub

Private Sub Form_Load()
  Dim camistocks As String
  arguments = ObtenerLíneaComando
  camicomandes = llegir_ini("General", "cami", "comandes.ini")
  cami = llegir_ini("General", "camibaixes", "comandes.ini")
  fitxerini = rutadelfitxer(camicomandes) + "muntadora.ini"
  ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  Shell ("Runas /user:administrador net time \\serverprodu /set /y")
  'desactiva ctrlaltsupr
  If cadbl(llegir_ini("Baixes", "programaamaquina", "comandes.ini")) = 1 Then Shell "Runas /user:administrador c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\desactivarctrl.reg"
  If llegir_ini("Baixes", "programaamaquina", "comandes.ini") = "{[}]" Then escriure_ini "Baixes", "programaamaquina", "0", fitxerini
  If cami = "{[}]" Then
    escriure_ini "General", "camibaixes", InputBox("Entra la ruta de baixes", "Atenció", "y:\comandes\baixes.mdb"), "comandes.ini"
  End If
  comanda = cadbl(llegir_ini("Baixes", "ultimacomanda", "comandes.ini"))
  r = cadbl(llegir_ini("Baixes", "nummaq", "comandes.ini"))
  If r = "{[}]" Then r = 100
  nummaq = cadbl(r)
  
  If UCase(arguments(1)) = "LLISTATBAIXESMUNTADORA" Then formclixesmuntats.Show 1: End
  
  If Not existeix("c:\ordprog.ini") And nummaq <> 0 Then assignardecimalipunt
  If nummaq = 0 Then Me.caption = "Baixes Muntadora  -  Nº de màquina a zero"
  'centerscreen Me
  Load formannex
'  centerscreen Me
  Me.Top = 1
  Me.Left = 1
  formannex.Top = 80
  formannex.Left = Me.width
  formannex.Show
  
  'cami = "\\SERVERprodu\dades\progcomandes\dades\baixesprova.mdb"
  Set dbbaixes = OpenDatabase(cami)
  Set dbcomandes = OpenDatabase(camicomandes)
  
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  Set dbmissatges = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
  
  
  reixalinies.RowHeight = 330
  datamuntadora.DatabaseName = cami
  datalinies.DatabaseName = cami
  rellotge.Enabled = True
  rellotge.Interval = 900
  desabilitartotselscontrols
  etaviscomandaprogramada = ""
  
End Sub
Sub desabilitartotselscontrols()

  For Each objecte In Me
      If objecte.Name <> "mestocadhesius" And objecte.Name <> "ordredelescomandes" And objecte.Name <> "nomoperari" And objecte.Name <> "Line1" And objecte.Name <> "rellotge" And objecte.Name <> "llistat" And objecte.Name <> "llistatbob" Then
        objecte.Enabled = False
      End If
     Next objecte
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload formannex
End Sub

Private Sub gruixpolimer_LostFocus()
gravartotals
End Sub

Private Sub impetxlperlalleixa_Click()
   Dim v As String
   Dim vultima As String
   Dim i As Integer
   
   mirarsihihalafontTTFdecodidebarres
   v = InputBox("Escriu la etiqueta que vols: Ex: XL-123" + Chr(10) + "Escriu [Totes] per fer-les totes.", "Etiqueta")
   If UCase(v) = "TOTES" Then
        vultima = InputBox("Quin es l'ultim número de XL que vols." + Chr(10) + "Ex: 489", "Última Etiqueta")
        If cadbl(vultima) = 0 Then Exit Sub
   End If
   vposicio = InputBox("Quina posició d'etiqueta vols començar?" + Chr(10) + "Es compta de la esquerra cap a dreta.", "Última Etiqueta", 1)
   Open "c:\temp\llistaxl.csv" For Output As #1
   Print #1, "NUMEROXL"
   For i = 1 To cadbl(vposicio) - 1
     Print #1, " "
   Next i
   If UCase(v) = "TOTES" Then
        For i = 1 To cadbl(vultima)
            Print #1, "XL-" + atrim(i)
        Next i
         Else: Print #1, UCase(v)
   End If
   Close #1
   imprimirXLs
End Sub
Sub imprimirXLs()
     Dim oapp As CRAXDDRT.Application
     Dim oreport As CRAXDDRT.Report
     
     
     Set oapp = New CRAXDDRT.Application
     Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "Etiqueta_XL_A4.rpt", 1)
     oreport.Database.Tables.Item(1).Location = "c:\temp\llistaxl.csv"
     oreport.DiscardSavedData
     oreport.PrintOut False, 1
     
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

Private Sub llistatexcelonline_Click()
obrir_document "https://docs.google.com/spreadsheets/d/1iq79PY2PvEq9ZpXeXjHi6-Efvm1pBK9qYm0wybajQno/edit?usp=sharing"
End Sub

Private Sub llitatXLs_Click()
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
 
 
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistat de quantitat de XLs a cada lleixa.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
  
  
  oreport.DiscardSavedData
  'If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
     
End Sub

Sub feines_parar_engegar_maquina(a As String, b As String)

End Sub
Private Sub mantcamises_Click()
formcamises.Show 1
End Sub

Private Sub mcingularreal2_Click()
  Dim numtreball As Double
  Dim ordremodificacio As Double
  numtreball = cadbl(InputBox("Entra el numero de treball.", "Treball"))
  If numtreball > 0 Then
       ordremodificacio = cadbl(InputBox("Entra el numero de versió.", "Versió"))

  End If
  If numtreball = 0 Or ordremodificacio = 0 Then Exit Sub
  'ara no es fa servir cingular real
  'mirarsihihaCingularReal_MUNT numtreball, ordremodificacio, 0
End Sub

Private Sub mestocadhesius_Click()
   Dim rst As Recordset
   vtancarestoc = False
   Set rst = dbbaixes.OpenRecordset("select horaultimaentrada from estoccintaadhesiva")
   If rst.EOF Then Exit Sub
   If tancarfinestraremota(atrim(rst!horaultimaentrada)) = True Then Exit Sub
   Set rst = Nothing
   escriure_ini "Muntadora", "tancarfinestraadhesius", "no", rutadelfitxer(cami) + "valorsprograma.ini"
   Unload estocadhesiu
   estocadhesiu.Show 1
   Unload estocadhesiu
End Sub
Function tancarfinestraremota(vh As String) As Boolean
   Dim c As Byte
   If vh = "0:00:00" Or vh = "" Then Exit Function
   vh = CVDate(vh)
   If DateDiff("s", vh, Now) < 10 Then
       MsgBox "Hi ha algú editant l'estoc en un altra ordinador espera uns segons i torna-ho a provar", vbCritical, "Error"
       Me.tag = "noguardar"
       tancarfinestraremota = True
       vtancarestoc = True
       Exit Function
   End If
   escriure_ini "Muntadora", "tancarfinestraadhesius", "si", rutadelfitxer(cami) + "valorsprograma.ini"
   MsgBox "Hi ha la finestra de control cinta adhesiva oberta en un altra ordinador, la tancaré d'allà per poder editar els canvis des d'aquest ordinador.", vbCritical, "Atenció"
   c = 0
   While llegir_ini("Muntadora", "tancarfinestraadhesius", rutadelfitxer(cami) + "valorsprograma.ini") = "si" And c < 5
      wait 1
      c = c + 1
   Wend
   If c > 4 Then MsgBox "No he pogut tancar la finestra remota.", vbCritical, "Error"
End Function

Private Sub metiquetabossaclixe_Click()
   Dim numtreball As Double
   Dim rst As Recordset
   numtreball = cadbl(InputBox("Entra el numero de treball que vols imprimir l'etiqueta.", "Impressió etiqueta de la bossa del clixé"))
   imprimiretbossatreball numtreball, False
End Sub
Sub imprimiretbossatreball(numtreball As Double, vcontrolarrepeticio As Boolean)
   Dim rst As Recordset
   If numtreball > 0 Then
        Set rst = dbclixes.OpenRecordset("select ordre from modificacions where id_treball=" + atrim(numtreball) + " order by ordre desc")
        If Not rst.EOF Then
           imprimiretiquetabossaclixes numtreball, cadbl(rst!ordre), llistat, vcontrolarrepeticio
        End If
   End If
   Set rst = Nothing
End Sub

Private Sub mimportardadesmuntadoravella_Click()
  importardades_muntadora_vella
End Sub
Sub importardades_muntadora_vella()
   Dim vnomBDsqlmuntadoravella As String
   Dim rstsql As Recordset
   
   Dim rstsqllinies As Recordset
   Dim vnomfitxerxml As String
   Dim vrutafitxerxml As String
   Dim vnumtreballEXPORTAR As Double
   vrutafitxerxml = "\\pc-vision-02\XmlInput\"
   vnomBDsqlmuntadoravella = rutadelfitxer(cami) + "\connexió SQL muntadora vella"
   vnumtreballEXPORTAR = 0
   vnumtreballEXPORTAR = cadbl(InputBox("Escriu el numero de treball que vols exportar.", "Exportar", atrim(vnumtreballEXPORTAR)))
   If vnumtreballEXPORTAR = 0 Then GoTo fi
   Set dbsql = OpenDatabase(vnomBDsqlmuntadoravella)
   If vnumtreballEXPORTAR = 0 Then
          Set rstsql = dbsql.OpenRecordset("select * from AVuser_Jobs where year(jobdate)>2020 order by jobdate", dbOpenSnapshot, dbSeeChanges)
            Else
             Set rstsql = dbsql.OpenRecordset("select * from AVuser_Jobs where jobnumber='" + atrim(Format(vnumtreballEXPORTAR, "00000000")) + "'", dbOpenSnapshot, dbSeeChanges)
   End If
   If vnumtreballEXPORTAR = 0 Then GoTo fi
   If rstsql.EOF Then MsgBox "No he trobat aquest treball a la muntadora vella.", vbCritical, "Error": GoTo fi
   rstsql.MoveLast
   rstsql.MoveFirst
   While Not rstsql.EOF
      Set rstsqllinies = dbsql.OpenRecordset("select * from avuser_reports where jobid=" + atrim(rstsql!JobID) + " order by reportnr,xoffset", dbOpenSnapshot, dbSeeChanges)
      vnumtreballiversio = verificartreballiversio(cadbl(rstsql!jobnumber))
      If vnumtreballiversio <> "" Then
            vnomfitxerxml = vrutafitxerxml + atrim(vnumtreballiversio) + ".xml"
            generar_xml rstsql, rstsqllinies, vnomfitxerxml
      End If
      Me.caption = " Treball " + atrim(rstsql.AbsolutePosition) + "/" + atrim(rstsql.RecordCount): DoEvents
      rstsql.MoveNext
   Wend
fi:
   Set rstsql = Nothing
End Sub
Sub generar_xml(rstsql As Recordset, rstsqllinies As Recordset, vnomfitxerxml As String)
     Dim veliminarfitxer As Boolean
     veliminarfitxer = False
     Open vnomfitxerxml For Output As #1
     genera_xml_capçalera rstsql
     genera_xml_linies rstsql, rstsqllinies, veliminarfitxer
     Print #1, vbTab + "</XML_Montacliche>" + vbNewLine + "</DocumentElement>" + vbNewLine
     Close #1
     If veliminarfitxer Then MsgBox "No he progut crear aquest treball a la nova muntadora les posicions horitzontals es sobreposen.", vbCritical, "Error": If existeix(vnomfitxerxml) Then Kill vnomfitxerxml
End Sub
Sub comptar_steps_i_posicions(vsteps As Long, vposicions As Long, rstsqllinies As Recordset, Optional vstep As Long, Optional vposicio As Long)
   Dim vultimstep As Double
   Dim vultimaposicio As Double
   Dim vbookmark As Variant
   Dim vrst As Recordset
   
   vsteps = 1
   vposicions = 1
   If rstsqllinies.EOF And rstsqllinies.BOF Then Exit Sub
   rstsqllinies.MoveFirst
   vultimestep = rstsqllinies!yOffset
   vultimaposicio = rstsqllinies!xOffset
   While Not rstsqllinies.EOF
      If rstsqllinies!yOffset <> vultimestep Then
           If vsteps = vstep Then GoTo fi
           vsteps = vsteps + 1
           vposicions = 1
           vultimestep = rstsqllinies!yOffset
           vultimaposicio = rstsqllinies!xOffset
      End If
      
      If rstsqllinies!xOffset <> vultimaposicio Then
           vposicions = vposicions + 1
      End If
      If vsteps = vstep And vposicio = 0 Then
           Set vrst = dbsql.OpenRecordset("select * from avuser_reports where jobid=" + atrim(rstsqllinies!JobID) + " and yoffset=" + atrim(rstsqllinies!yOffset) + " order by yoffset,xoffset", dbOpenSnapshot, dbSeeChanges)
           If Not vrst.EOF Then vrst.MoveLast: vposicions = vrst.RecordCount
           GoTo fi
           'If vposicions = vposicio Then GoTo fi
      End If
      If vsteps = vstep And vposicio = vposicions Then vbookmark = rstsqllinies.Bookmark
      rstsqllinies.MoveNext
   Wend
   
   If vstep > 0 Then If vsteps = vstep And vposicions = vposicio Then vsteps = 0: vposicions = 0
fi:
If Not IsEmpty(vbookmark) Then rstsqllinies.Bookmark = vbookmark
End Sub
Function observacio_delJOB(vidjob As Long) As String
   Dim rst As Recordset
   Set rst = dbsql.OpenRecordset("select * from AVuser_Notes where jobid=" + atrim(vidjob))
   While Not rst.EOF
       observacio_delJOB = observacio_delJOB + IIf(observacio_delJOB <> "", " | ", "") + atrim(rst!info)
       rst.MoveNext
   Wend
   Set rst = Nothing
End Function
Sub genera_xml_linies(rstsql As Recordset, rstsqllinies As Recordset, veliminarfitxer As Boolean)
   Dim vlinia As String
   Dim vsteps As Long
   Dim vposicions As Long
   comptar_steps_i_posicions vsteps, vposicions, rstsqllinies
   vlinia = vbTab + vbTab + "<Colore_0>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<Colore>QUALSEVOL COLOR</Colore>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<Note>" + observacio_delJOB(rstsql!JobID) + " </Note>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<TipoBiadesivo></TipoBiadesivo>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<TipoAnilox> </TipoAnilox>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<NrManica> </NrManica>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<NumStep>" + atrim(vsteps) + "</NumStep>" + vbNewLine
   For i = 0 To vsteps - 1
        vlinia = vlinia + vbTab + vbTab + vbTab + "<Step_" + atrim(i) + ">" + vbNewLine
        comptar_steps_i_posicions 0, vposicions, rstsqllinies, i + 1
        If rstsqllinies.EOF Then GoTo fistep
        vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + "<W>" + atrim(Redondejar((rstsqllinies!yOffset / 1000) - ((rstsql!CylinderCircumference / 1000) / 2))) + "</W>" + vbNewLine
        vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + "<NumPosizioni>" + atrim(vposicions) + "</NumPosizioni>" + vbNewLine
        For j = 0 To vposicions - 1
            comptar_steps_i_posicions 0, 0, rstsqllinies, i + 1, j + 1
            If rstsqllinies.EOF Then veliminarfitxer = True: GoTo fi
            vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + "<Posizione_" + atrim(j) + ">" + vbNewLine
            vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + vbTab + "<X1>" + atrim(Redondejar(rstsqllinies!xOffset / 1000, 0)) + "</X1>" + vbNewLine
            vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + vbTab + "<X2>" + atrim(Redondejar((rstsqllinies!xOffset / 1000) + (rstsqllinies!reportwidth / 1000), 0)) + "</X2>" + vbNewLine
            vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + "</Posizione_" + atrim(j) + ">" + vbNewLine
        Next j
fistep:
        vlinia = vlinia + vbTab + vbTab + vbTab + "</Step_" + atrim(i) + ">" + vbNewLine
   Next i
   vlinia = vlinia + vbTab + vbTab + "</Colore_0>"
   Print #1, vlinia
fi:
End Sub
Sub genera_xml_capçalera(rstsql As Recordset)
   Dim vlinia As String
   vlinia = "<?xml version=""1.0""?>" + vbNewLine
   vlinia = vlinia + "<DocumentElement Version=""1.0"">" + vbNewLine
   vlinia = vlinia + vbTab + "<XML_Montacliche>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + "<PR>" + atrim(rstsql!CylinderCircumference / 1000) + "</PR>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + "<Descrizione>_" + treure_simbolsextranys(atrim(rstsql!jobname)) + "</Descrizione>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + "<SpessoreCliche>1.14 </SpessoreCliche>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + "<NumeroColori>1</NumeroColori>"
   Print #1, vlinia
End Sub
Function treure_simbolsextranys(vtxt As String) As String
  Dim vaccents As String
  Dim vnoaccents As String
  Dim i As Byte
  Dim v As String
  v = vtxt
  vaccents = "àèòáéíóú&'´"
  vnoaccents = "aeoaeiou/,,"
  For i = 1 To Len(vaccents)
     v = substituirtot(v, Mid(vaccents, i, 1), Mid(vnoaccents, i, 1))
  Next i
  treure_simbolsextranys = v
     
End Function
Function verificartreballiversio(vnumtreball As Double) As String
  Dim rst As Recordset
  Set rst = dbbaixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(vnumtreball) + " order by ordre desc")
  If Not rst.EOF Then
      verificartreballiversio = atrim(rst!id_treball) + "-" + atrim(rst!ordre)
  End If
  Set rst = Nothing
End Function
Private Sub nomoperari_Click()
 Dim numoptmp As Integer
 Dim nomoptmp As String
 
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='M' and actiu<>0"
  formseleccio.caption = "Selecció d'Operari"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   numoptmp = cadbl(formseleccio.Data1.Recordset!codi)
   nomoptmp = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  If numoptmp <> 0 Then
     numop = numoptmp
     nomoperari = atrim(numop) + "-" + nomoptmp
     
     For Each objecte In Me
      If objecte.Name <> "llistat" And objecte.Name <> "llistatbob" And objecte.Name <> "Line1" Then
        objecte.Enabled = True
      End If
     Next objecte
      Else: If cadbl(numop) = 0 Then MsgBox "Has d'escullir un operari per treballar": Exit Sub
  End If
   If cadbl(comanda) > 0 Then
      Command4_Click
     Else: numcomanda.SetFocus
   End If
End Sub

Private Sub numcomanda_Change()
  'comanda = numcomanda
End Sub

Private Sub numcomanda_DropDown()
   borrarfetes
'   ordremuntatge.Show 1
'   If cadbl(numcomanda) > 0 Then Command4_Click
   obrir_llistaordrecomandes
End Sub
Sub obrir_llistaordrecomandes()
  Dim vcomandaactual As Double
  Dim vcomandafingerprint As Double
  Set dbtmp = dbcomandes
  Set dbtmpb = dbbaixes
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "\tintes.mdb")
  vcomandaactual = cadbl(comanda)
  vcomandafingerprint = vcomandaactual
  Load formordreimpresio
  formordreimpresio.bbobinesamaquina.visible = False
  formordreimpresio.bimprimir.visible = False
  formordreimpresio.Framemodificacions.visible = False
  formordreimpresio.Show
  formordreimpresio.reixa.row = 1
  While Screen.ActiveForm.Name = "formordreimpresio"
     vnumc = formordreimpresio.reixa.TextMatrix(formordreimpresio.reixa.row, 0)
     If Not IsNumeric(Mid(vnumc, 1, 1)) Then vnumc = Mid(vnumc, 2)
     If cadbl(vnumc) > 0 Then
       If cadbl(vnumc) <> vcomandaactual Then
          'carrego l'annex
           formannex.carregarcomanda cadbl(vnumc)
           formordreimpresio.SetFocus
           vcomandaactual = vnumc
       End If
       
     End If
    DoEvents
  Wend
senseescullir:
  If seleccioret = 1 Or seleccioret = 5 Or seleccioret = 2 Then
   vnumc = formordreimpresio.reixa.TextMatrix(formordreimpresio.reixa.row, 0)
   If Not IsNumeric(Mid(vnumc, 1, 1)) Then vnumc = Mid(vnumc, 2)
   
   If seleccioret = 1 Then
        vnopreguntar = False
        Form1.BackColor = &H80000005
        vestemfentfingerprint = False
        Unload formseleccio
        'formannex.carregarcomanda cadbl(comanda)
        'vnumc = cadbl(InputBox("Entra la comanda manualment.", "Comanda"))
   End If
   If seleccioret = 2 Then
        If MsgBox("Vols fer la comanda " + atrim(vcomandaactual) + " com a copia de la " + atrim(vcomandafingerprint) + "?", vbExclamation + vbDefaultButton2 + vbYesNo, "Finger Print") = vbYes Then
          vestemfentfingerprint = True
          Form1.BackColor = &HFFFF&
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
  If vnumc = 0 Then formannex.carregarcomanda cadbl(comanda): Exit Sub
  If vnoobrirla Then GoTo fi
  numcomanda = vnumc
  comanda = numcomanda
  Command4_Click
fi:
End Sub
Sub ensenyarordreplanificacio()
  Dim nummaq As String
  
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select comanda from muntadora_ordremuntatge order by ordre"
  formseleccio.caption = "Selecció de comanda"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 2200
  'formseleccio.DBGrid2.Columns(1).Width = 5500
  formseleccio.Show 1
  If seleccioret = 1 Then
   numcomanda = atrim(formseleccio.Data1.Recordset!comanda)
   Command4_Click
  End If
  Unload formseleccio
     
End Sub

Private Sub observacionsgenerals_LostFocus()
gravartotals
End Sub

Private Sub observacionstreball_LostFocus()
   dbbaixes.Execute "update muntadores_obstreballs set observacions='" + treure_apostruf(atrim(observacionstreball)) + "' where numtreball=" + atrim(id_treball)
End Sub

Private Sub ordredelescomandes_Click()
   borrarfetes
   ordredelescomandes.tag = "1"
  ordremuntatge.Show 1
  ordredelescomandes.tag = ""
  mirarsicomandaprogramada True
End Sub

Private Sub provadhesiu_DropDown()
  escullir_proveidor_adhesiu
  SendKeys "{TAB}"
  gravartotals
End Sub
Sub escullir_proveidor_adhesiu()
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select distinct codiproveidor,nomproveidor from adhesiuSmuntadora group by codiproveidor,nomproveidor"
  formseleccio.caption = "Selecció de proveidor"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.Show 1
  If seleccioret = 1 Then
   provadhesiu.tag = cadbl(formseleccio.Data1.Recordset!codiproveidor)
   provadhesiu = atrim(formseleccio.Data1.Recordset!nomproveidor)
   dbbaixes.Execute "update muntadoratot set proveidoradhesiu=" + provadhesiu.tag + " where comanda=" + atrim(cadbl(comanda))
  End If
  Unload formseleccio
End Sub

Private Sub reixalinies_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
   tempseditant = 0
End Sub

Private Sub reixalinies_ButtonClick(ByVal ColIndex As Integer)
   Dim gruix As Double
   Dim idad As Long
   tempseditant = 0
   copiarampleanterior
   gruix = cadbl(reixalinies.Columns("Gruix"))
   
   If reixalinies.Columns(ColIndex).DataField = "nomadhesiu" Then
      If gruix > 0 Then
          idad = cadbl(triar_adhesiu(gruix))
          If idad > 0 Then
            If datalinies.Recordset.EditMode = 0 Then datalinies.Recordset.Edit
            datalinies.Recordset!idadhesiu = idad
            reixalinies.Columns("nomadhesiu") = r
            datalinies.Recordset.Update
          End If
          reixalinies.col = reixalinies.col + 1
         Else: MsgBox "Primer has de possar el gruix de l'adhesiu", vbInformation, "Atenció"
      End If
   End If
   tempseditant = Now
End Sub
      
Sub copiarampleanterior()
   Dim rstt As Recordset
   If reixalinies.Columns(reixalinies.col).caption = "Gruix" Then Exit Sub
   If cadbl(reixalinies.Columns("Gruix")) > 0 Then Exit Sub
   Set rstt = datalinies.Recordset.Clone
   If Not rstt.EOF Then
     rstt.MoveFirst
     If cadbl(rstt!gruixadhesiu) > 0 Then reixalinies.Columns("Gruix") = atrim(rstt!gruixadhesiu)
     If cadbl(rstt!numpolimers) > 0 Then reixalinies.Columns("Nº Polimers") = atrim(rstt!numpolimers)
   End If
   
End Sub
Function triar_adhesiu(gruix) As String
  r = ""
   Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select codiintern,descripcioinplacsa from adhesiusmuntadora where codiproveidor=" + atrim(cadbl(provadhesiu.tag)) + " and gruix=" + passaradecimalpunt(atrim(gruix))
  formseleccio.caption = "Selecció adhesiu"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   triar_adhesiu = atrim(cadbl(formseleccio.Data1.Recordset!codiintern))
   r = atrim(formseleccio.Data1.Recordset!descripcioinplacsa)
  End If
  Unload formseleccio
End Function

Private Sub reixalinies_Change()
  If reixalinies.Columns("op") = "0" Then
     reixalinies.Columns("op") = atrim(numop)
  End If
  copiarampleanterior
End Sub

Private Sub reixalinies_ColEdit(ByVal ColIndex As Integer)
tempseditant = 0
End Sub

Private Sub reixalinies_DblClick()
If reixalinies.col = 0 Then
  reixalinies.Text = escullir_operari
End If

End Sub

Sub possarhorestreballades()
   Dim horainici As String
   Dim horafi As String
   Dim hores As Double
   horainici = atrim(reixamuntadora.Columns("datainici")) + " " + atrim(reixamuntadora.Columns("horainici"))
   horafi = atrim(reixamuntadora.Columns("datafi")) + " " + atrim(reixamuntadora.Columns("horafi"))
   If Not IsDate(horainici) Or Not IsDate(horafi) Then
      hores = 0
       Else
            hores = DateDiff("n", CVDate(horainici), CVDate(horafi))
            If hores < 0 Then hores = 0
            hores = Redondejar(hores / 60, 2)
   End If
   reixamuntadora.Columns("totalhores") = hores
End Sub

Private Sub reixalinies_KeyDown(KeyCode As Integer, Shift As Integer)
tempseditant = Now
End Sub

Private Sub reixamuntadora_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
tempseditant = 0
End Sub

Private Sub reixamuntadora_ColEdit(ByVal ColIndex As Integer)
tempseditant = 0
End Sub

Private Sub reixamuntadora_DblClick()
If reixamuntadora.col = 0 Then
  reixamuntadora.Text = escullir_operari
'  numop = reixamuntadora.Text
End If
End Sub
Function escullir_operari() As String
Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='M' and actiu<>0"
  formseleccio.caption = "Selecció d'Operari"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   escullir_operari = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   r = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
End Function

Private Sub reixamuntadora_KeyDown(KeyCode As Integer, Shift As Integer)
tempseditant = Now
 Command6.visible = False
End Sub

Private Sub reixamuntadora_KeyUp(KeyCode As Integer, Shift As Integer)
 If reixamuntadora.col = 4 And KeyCode > 46 Then
     If (Len(reixamuntadora.Text)) >= 4 Then reixamuntadora.col = 5
  End If
  If reixamuntadora.col = 3 And KeyCode > 46 Then
     If (Len(reixamuntadora.Text)) >= 6 Then reixamuntadora.col = 4
  End If
  If reixamuntadora.col = 2 And KeyCode > 46 Then
     If (Len(reixamuntadora.Text)) >= 4 Then
       reixamuntadora.col = 3
     End If
  End If
  If reixamuntadora.col = 1 And KeyCode > 46 Then
     If (Len(reixamuntadora.Text)) >= 6 Then reixamuntadora.col = 2
  End If
  
  If reixamuntadora.col = 14 And KeyCode > 46 Then
      If (Len(reixamuntadora.Text)) > 99 Then reixamuntadora.Text = Mid(reixamuntadora.Text, 1, 99)
  End If
End Sub

Private Sub reixamuntadora_LostFocus()
  'guardarcanvisreixa
  reixamuntadora.col = IIf(reixamuntadora.col + 1 < reixamuntadora.Columns.Count, reixamuntadora.col + 1, 0)
  
End Sub

Private Sub reixamuntadora_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'   guardarcanvisreixa
  If reixamuntadora.Bookmark = LastRow Then
    possardataihora LastCol
    possarhorestreballades
  End If
   guardarcanvisreixa
   tempseditant = Now
End Sub
Sub possardataihora(LastCol As Integer)
 Dim valtmp As String
 
 If LastCol = 1 Or LastCol = 2 Then
  valtmp = reixamuntadora.Columns(LastCol).Text
  
  If LastCol = 1 Then
      
      If InStr(1, valtmp, "/") = 0 Then valtmp = Mid(valtmp, 1, 2) + "/" + Mid(valtmp, 3, 2) + "/" + Mid(valtmp, 5, 2)
      If Not IsDate(valtmp) Then valtmp = ""
  End If
  
  If LastCol = 2 Then
    If InStr(1, valtmp, ":") = 0 Then valtmp = Mid(valtmp, 1, 2) + ":" + Mid(valtmp, 3, 2)
    If Not IsDate(Format(valtmp, "hh:nn")) Then valtmp = ""
  End If
  reixamuntadora.Columns(LastCol) = IIf(valtmp = "", Null, valtmp)
  End If
  
  If LastCol = 3 Or LastCol = 4 Then
  valtmp = reixamuntadora.Columns(LastCol).Text
  If LastCol = 3 Then
      
      If InStr(1, valtmp, "/") = 0 Then valtmp = Mid(valtmp, 1, 2) + "/" + Mid(valtmp, 3, 2) + "/" + Mid(valtmp, 5, 2)
      If Not IsDate(valtmp) Then valtmp = ""
  End If
  
  If LastCol = 4 Then
    If InStr(1, valtmp, ":") = 0 Then valtmp = Mid(valtmp, 1, 2) + ":" + Mid(valtmp, 3, 2)
      If Not IsDate(Format(valtmp, "hh:nn")) Then valtmp = ""

  End If
  reixamuntadora.Columns(LastCol) = IIf(valtmp = "", Null, valtmp)
 End If
End Sub
Sub mirarsiparar()
 Static contar
  If llegir_ini("General", "parar", llegir_ini("General", "rutallistats", "comandes.ini") + "parar.ini") = "si" Then
    contar = contar + 1
     If contar = 1 Then MsgBox2 "El programa es pararà d'aqui a 1 minut. TANCA TOT I ESPERA CINC MINUTS.", 5, "Actualització", vbCritical
     If contar = 15 Then MsgBox2 "El programa es pararà d'aqui a 30 segons. TANCA TOT I ESPERA CINC MINUTS.", 5, "Actualització, vbCritical"
     If contar = 27 Then MsgBox2 "El programa es pararà d'aqui a 5 segons. TANCA TOT I ESPERA CINC MINUTS.", 3, "Actualització", vbCritical
     If contar > 30 Then End
   Else: contar = 0
  End If
  If llegir_ini("General", "parar", llegir_ini("General", "rutallistats", "comandes.ini") + "parar.ini") = "ja" Then End
End Sub
Private Sub rellotge_Timer()
  mirarsicomandaprogramada False
  mirarsiparar
  mirarsieditant
  blinkingavispeu
  mirarsiavisarpecopia
End Sub
Sub mirarsicomandaprogramada(vcomprovarara As Boolean)
  Static vcont
  Dim rst As Recordset
  vcont = vcont + 1
  If vcont > 70 Or vcomprovarara Then
     etaviscomandaprogramada = ""
     Set rst = dbbaixes.OpenRecordset("select * from impresores_ordreimpresio where not muntada and dataprogramada<>null order by dataprogramada ")
     If Not rst.EOF Then
        If DateDiff("n", Now, rst!dataprogramada) < 60 Then
           etaviscomandaprogramada = "Comanda " + atrim(rst!comanda) + " programada per les " + Format(rst!dataprogramada, "hh:nn") + " i encara no està muntada."
           etaviscomandaprogramada.visible = True
        End If
     End If
     vcont = 0
  End If
End Sub
Sub blinkingavispeu()
   If avispeu <> "" Then
      avispeu.visible = Not avispeu.visible
   End If
   If etaviscomandaprogramada <> "" Then
       etaviscomandaprogramada.visible = Not etaviscomandaprogramada.visible
   End If
End Sub
Sub mirarsiavisarpecopia()
    Static avisat As Boolean
    If Format(Now, "w") = 6 And Format(Now, "hhnn") > 2130 And Not avisat Then
       Form1.BackColor = &H8080FF
       MsgBox "PENSEU EN FER LA COPIA DE L'ORDINADOR DE MUNTADORA", vbCritical, "ATENCIÓ": avisat = True
    End If
End Sub


Sub mirarsieditant()
If DateDiff("s", tempseditant, Now) > 6 And tempseditant > 0 Then
   On Error Resume Next
   If Not datalinies.Recordset.EOF Then
    If datalinies.Recordset.EditMode = 0 Then datalinies.Recordset.Edit
    datalinies.Recordset.Update
   End If
   datalinies.UpdateControls
   If Not datamuntadora.Recordset.EOF Then
    If datamuntadora.Recordset.EditMode = 0 Then datamuntadora.Recordset.Edit
    
    datamuntadora.Recordset.Update
   End If
   datamuntadora.UpdateControls
   
  
   tempseditant = 0
 End If
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

Sub preparaelPDF(vnomfitxerpdf As String, vrotacio As Double, vMirall As String)
    
  vMirall = UCase(vMirall) 'vMirall si es V es vertical H es horitzotal
  If existeix("c:\temp\pdfimpresio.gif") Then Kill "c:\temp\pdfimpresio.gif"
  ConvertirFormats vnomfitxerpdf, "c:\temp\pdfimpresio.gif", 50
  If vMirall = "H" Then InvertirHVImatge "c:\temp\pdfimpresio.gif", "c:\temp\pdfimpresio.gif"
  If vMirall = "V" Then InvertirHVImatge "c:\temp\pdfimpresio.gif", "c:\temp\pdfimpresio.gif", True
  If vrotacio > 0 Then RotarImatge "c:\temp\pdfimpresio.gif", "c:\temp\pdfimpresio.gif", vrotacio
  
End Sub

