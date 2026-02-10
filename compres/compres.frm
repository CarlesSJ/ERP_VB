VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form comandescompra 
   Caption         =   "Comandes de compra."
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10410
   Icon            =   "compres.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Tag             =   "primera"
   Begin VB.TextBox linia 
      Height          =   285
      Left            =   15
      TabIndex        =   77
      Top             =   3495
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Data liniescompra 
      Caption         =   "linies"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2475
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "liniescompra"
      Top             =   3780
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Data capcalera 
      Caption         =   "Comandes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from capcalera order by data desc"
      Top             =   75
      Width           =   2910
   End
   Begin VB.Frame fdetallcompra 
      Caption         =   "Detall de la linia de compra."
      Height          =   3705
      Left            =   105
      TabIndex        =   2
      Top             =   4560
      Width           =   10230
      Begin VB.CommandButton Command3 
         Height          =   330
         Left            =   570
         Picture         =   "compres.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Modificació descripcio del material."
         Top             =   210
         Width           =   345
      End
      Begin VB.CommandButton Command5 
         Height          =   285
         Left            =   9915
         Picture         =   "compres.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Afegir Observacions"
         Top             =   1665
         Width           =   270
      End
      Begin VB.CommandButton Command4 
         Height          =   285
         Left            =   9915
         Picture         =   "compres.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Eliminar totes les linies"
         Top             =   1980
         Width           =   270
      End
      Begin VB.CommandButton borrarliniesdescripcio 
         Height          =   285
         Left            =   4395
         Picture         =   "compres.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Eliminar totes les linies"
         Top             =   1995
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Left            =   4395
         Picture         =   "compres.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Afegir Observacions"
         Top             =   1710
         Width           =   285
      End
      Begin VB.Data liniesdescripcio 
         Caption         =   "liniesdescripcio"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1950
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "liniesdescripcio"
         Top             =   2895
         Visible         =   0   'False
         Width           =   2625
      End
      Begin MSDBGrid.DBGrid reixadescripcio 
         Bindings        =   "compres.frx":213C
         Height          =   2160
         Left            =   45
         OleObjectBlob   =   "compres.frx":2157
         TabIndex        =   53
         Top             =   1485
         Width           =   4320
      End
      Begin VB.Data comandesxlinia 
         Caption         =   "comandesxlinia"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6945
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "comandesxlinia"
         Top             =   2850
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.CommandButton acceptarlinia 
         Height          =   330
         Left            =   1275
         Picture         =   "compres.frx":29BE
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Acceptar canvis"
         Top             =   210
         Width           =   345
      End
      Begin VB.CommandButton eliminarlinia 
         Height          =   330
         Left            =   915
         Picture         =   "compres.frx":2F48
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Eliminacio Registres"
         Top             =   210
         Width           =   345
      End
      Begin VB.CommandButton novalinia 
         Height          =   330
         Left            =   225
         Picture         =   "compres.frx":34D2
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Alta  Registres"
         Top             =   210
         Width           =   345
      End
      Begin MSDBGrid.DBGrid reixacomandes 
         Bindings        =   "compres.frx":3A5C
         Height          =   2160
         Left            =   4710
         OleObjectBlob   =   "compres.frx":3A75
         TabIndex        =   49
         Top             =   1485
         Width           =   5190
      End
      Begin VB.Frame fdescmat 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Descripció del material comprat."
         Enabled         =   0   'False
         Height          =   870
         Left            =   60
         TabIndex        =   25
         Top             =   570
         Width           =   10125
         Begin VB.TextBox linicodmat 
            Appearance      =   0  'Flat
            BackColor       =   &H00EBC5C5&
            BorderStyle     =   0  'None
            DataField       =   "codimaterial"
            DataSource      =   "liniescompra"
            Height          =   195
            Left            =   840
            TabIndex        =   71
            Top             =   225
            Width           =   630
         End
         Begin VB.TextBox preu 
            DataField       =   "preu"
            DataSource      =   "liniescompra"
            Height          =   285
            Left            =   8850
            TabIndex        =   37
            Top             =   435
            Width           =   540
         End
         Begin VB.TextBox mandril 
            DataField       =   "mandril"
            DataSource      =   "liniescompra"
            Height          =   285
            Left            =   8265
            TabIndex        =   36
            Top             =   435
            Width           =   525
         End
         Begin VB.TextBox diamext 
            DataField       =   "diametreext"
            DataSource      =   "liniescompra"
            Height          =   285
            Left            =   7770
            TabIndex        =   35
            Top             =   435
            Width           =   405
         End
         Begin VB.TextBox kilosxrcomprar 
            BackColor       =   &H00C0C0C0&
            DataField       =   "quantitatkg"
            DataSource      =   "liniescompra"
            Height          =   300
            Left            =   9435
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   435
            Width           =   645
         End
         Begin VB.CheckBox imicrop 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Microperforat"
            DataField       =   "microperforat"
            DataSource      =   "liniescompra"
            Height          =   300
            Left            =   7350
            TabIndex        =   34
            Top             =   450
            Width           =   270
         End
         Begin VB.ComboBox iobert 
            DataField       =   "obert"
            DataSource      =   "liniescompra"
            Height          =   315
            ItemData        =   "compres.frx":4B22
            Left            =   4950
            List            =   "compres.frx":4B2F
            TabIndex        =   29
            Top             =   450
            Width           =   495
         End
         Begin VB.ComboBox icares 
            DataField       =   "carestractat"
            DataSource      =   "liniescompra"
            Height          =   315
            ItemData        =   "compres.frx":4B3C
            Left            =   4395
            List            =   "compres.frx":4B49
            TabIndex        =   28
            Top             =   435
            Width           =   540
         End
         Begin VB.ComboBox itl 
            DataField       =   "semielaborat"
            DataSource      =   "liniescompra"
            Height          =   315
            ItemData        =   "compres.frx":4B56
            Left            =   3915
            List            =   "compres.frx":4B60
            TabIndex        =   27
            Top             =   435
            Width           =   480
         End
         Begin VB.TextBox iespesor 
            DataField       =   "micres"
            DataSource      =   "liniescompra"
            Height          =   285
            Left            =   6705
            TabIndex        =   33
            Top             =   435
            Width           =   555
         End
         Begin VB.TextBox iplegat 
            DataField       =   "Plegat"
            DataSource      =   "liniescompra"
            Height          =   285
            Left            =   5985
            TabIndex        =   31
            Top             =   420
            Width           =   360
         End
         Begin VB.TextBox iample 
            DataField       =   "Ample"
            DataSource      =   "liniescompra"
            Height          =   285
            Left            =   5430
            TabIndex        =   30
            Top             =   435
            Width           =   540
         End
         Begin VB.TextBox isolapa 
            DataField       =   "Solapa"
            DataSource      =   "liniescompra"
            Height          =   285
            Left            =   6375
            TabIndex        =   32
            Top             =   435
            Width           =   300
         End
         Begin VB.ComboBox combomaterial 
            BackColor       =   &H00808080&
            DataField       =   "nommaterial"
            DataSource      =   "liniescompra"
            Height          =   315
            Left            =   840
            TabIndex        =   26
            Top             =   435
            Width           =   3000
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Preu"
            Height          =   255
            Index           =   7
            Left            =   8910
            TabIndex        =   58
            Top             =   225
            Width           =   495
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Mandril"
            Height          =   255
            Index           =   6
            Left            =   8265
            TabIndex        =   55
            Top             =   225
            Width           =   615
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Ø Ext."
            Height          =   255
            Index           =   5
            Left            =   7770
            TabIndex        =   54
            Top             =   225
            Width           =   525
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Kilos/Quant"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   9360
            TabIndex        =   48
            Top             =   210
            Width           =   810
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Micro"
            Height          =   255
            Index           =   1
            Left            =   7290
            TabIndex        =   46
            Top             =   225
            Width           =   600
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Obert"
            Height          =   300
            Index           =   1
            Left            =   5010
            TabIndex        =   45
            Top             =   210
            Width           =   540
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "C.Tractat"
            Height          =   300
            Index           =   2
            Left            =   4305
            TabIndex        =   44
            Top             =   210
            Width           =   1020
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "T/L"
            Height          =   300
            Index           =   3
            Left            =   3975
            TabIndex        =   43
            Top             =   210
            Width           =   360
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Micres"
            Height          =   255
            Index           =   15
            Left            =   6735
            TabIndex        =   42
            Top             =   210
            Width           =   600
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Pleg"
            Height          =   255
            Index           =   3
            Left            =   5985
            TabIndex        =   41
            Top             =   210
            Width           =   480
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Ample"
            Height          =   255
            Index           =   2
            Left            =   5490
            TabIndex        =   40
            Top             =   210
            Width           =   630
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Solp"
            Height          =   255
            Index           =   0
            Left            =   6360
            TabIndex        =   39
            Top             =   210
            Width           =   330
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Material:"
            Height          =   345
            Left            =   135
            TabIndex        =   38
            Top             =   480
            Width           =   735
         End
      End
   End
   Begin VB.Frame fliniescompra 
      Caption         =   "Linies de compra."
      Height          =   1890
      Left            =   105
      TabIndex        =   1
      Top             =   2670
      Width           =   10215
      Begin Crystal.CrystalReport llistat 
         Left            =   1035
         Top             =   900
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileName   =   "c:\prova.doc"
         PrintFileType   =   17
         PrintFileLinesPerPage=   60
      End
      Begin MSDBGrid.DBGrid reixalinies 
         Bindings        =   "compres.frx":4B6A
         Height          =   1635
         Left            =   60
         OleObjectBlob   =   "compres.frx":4B81
         TabIndex        =   24
         Top             =   210
         Width           =   10065
      End
   End
   Begin VB.Frame fcapcalera 
      BackColor       =   &H008080FF&
      Caption         =   "Dades de la capçalera."
      Enabled         =   0   'False
      Height          =   2130
      Left            =   105
      TabIndex        =   0
      Top             =   555
      Width           =   10200
      Begin VB.TextBox precomandafins 
         BackColor       =   &H008080FF&
         DataField       =   "precomandafins"
         DataSource      =   "capcalera"
         Height          =   285
         Left            =   9165
         TabIndex        =   75
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox empresa 
         BackColor       =   &H008080FF&
         DataField       =   "empresa"
         DataSource      =   "capcalera"
         Height          =   315
         ItemData        =   "compres.frx":6950
         Left            =   945
         List            =   "compres.frx":695A
         TabIndex        =   72
         Text            =   "Inplacsa"
         Top             =   195
         Width           =   2010
      End
      Begin MSMAPI.MAPIMessages MiMAPIMessages 
         Left            =   4455
         Top             =   660
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   1
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   0   'False
      End
      Begin MSMAPI.MAPISession MiMAPISession 
         Left            =   3510
         Top             =   570
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   -1  'True
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FEC7C7&
         BorderStyle     =   0  'None
         DataField       =   "total"
         DataSource      =   "capcalera"
         Height          =   270
         Left            =   3945
         TabIndex        =   63
         Top             =   1800
         Width           =   1035
      End
      Begin VB.ComboBox proveidor 
         DataField       =   "nomprov"
         DataSource      =   "capcalera"
         Height          =   315
         Left            =   5835
         TabIndex        =   14
         Top             =   165
         Width           =   3480
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Dades del proveidor"
         Height          =   1575
         Left            =   5040
         TabIndex        =   13
         Top             =   510
         Width           =   4110
         Begin VB.Label Label9 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "provincia"
            DataSource      =   "capcalera"
            Height          =   330
            Left            =   75
            TabIndex        =   18
            Top             =   1245
            Width           =   4110
         End
         Begin VB.Label Label8 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "codipipoblacio"
            DataSource      =   "capcalera"
            Height          =   330
            Left            =   75
            TabIndex        =   17
            Top             =   915
            Width           =   4110
         End
         Begin VB.Label Label7 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "direccio"
            DataSource      =   "capcalera"
            Height          =   330
            Left            =   75
            TabIndex        =   16
            Top             =   585
            Width           =   4110
         End
         Begin VB.Label Label6 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "nomprovcomercial"
            DataSource      =   "capcalera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   75
            TabIndex        =   15
            Top             =   255
            Width           =   4110
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         Height          =   1635
         Left            =   75
         TabIndex        =   3
         Top             =   450
         Width           =   3075
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Data confirmada."
            Height          =   345
            Left            =   1215
            TabIndex        =   74
            Top             =   1200
            Width           =   1725
         End
         Begin VB.TextBox Text3 
            DataField       =   "magatzem"
            DataSource      =   "capcalera"
            Height          =   285
            Left            =   1230
            TabIndex        =   10
            Top             =   1275
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox Text2 
            DataField       =   "dataentrega"
            DataSource      =   "capcalera"
            Height          =   285
            Left            =   1230
            TabIndex        =   9
            Top             =   900
            Width           =   1335
         End
         Begin VB.TextBox cnumcomanda 
            BackColor       =   &H00C0C0C0&
            DataField       =   "numcomanda"
            DataSource      =   "capcalera"
            Height          =   285
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   510
            Width           =   1335
         End
         Begin VB.TextBox data 
            DataField       =   "data"
            DataSource      =   "capcalera"
            Height          =   285
            Left            =   1230
            TabIndex        =   5
            Top             =   135
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Magatzem:"
            Height          =   285
            Left            =   165
            TabIndex        =   11
            Top             =   1290
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Entrega:"
            Height          =   345
            Left            =   180
            TabIndex        =   8
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Comanda:"
            Height          =   345
            Left            =   180
            TabIndex        =   6
            Top             =   570
            Width           =   1155
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Data:"
            Height          =   345
            Index           =   0
            Left            =   165
            TabIndex        =   4
            Top             =   195
            Width           =   705
         End
      End
      Begin VB.Label titprecomandafins 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pendent d'enviar fins:"
         Height          =   420
         Left            =   9210
         TabIndex        =   76
         Top             =   555
         Width           =   915
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
         Height          =   270
         Left            =   180
         TabIndex        =   73
         Top             =   225
         Width           =   810
      End
      Begin VB.Label msgpendent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pendent..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   840
         Left            =   3225
         TabIndex        =   64
         Top             =   195
         Width           =   1800
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Total €:"
         Height          =   375
         Left            =   3345
         TabIndex        =   62
         Top             =   1800
         Width           =   930
      End
      Begin VB.Image imatgeimpres 
         Height          =   315
         Index           =   0
         Left            =   9420
         Picture         =   "compres.frx":6970
         Stretch         =   -1  'True
         ToolTipText     =   "Aquesta comanda s'ha imprès."
         Top             =   180
         Width           =   300
      End
      Begin VB.Image imatgeenviat 
         Height          =   315
         Index           =   0
         Left            =   9780
         Picture         =   "compres.frx":6EFA
         Stretch         =   -1  'True
         ToolTipText     =   "Aquesta comanda s'ha enviat"
         Top             =   180
         Width           =   300
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveidor:"
         Height          =   345
         Left            =   4995
         TabIndex        =   12
         Top             =   225
         Width           =   885
      End
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   90
      TabIndex        =   19
      Top             =   -60
      Width           =   10230
      Begin VB.CommandButton Command9 
         Height          =   375
         Left            =   8415
         Picture         =   "compres.frx":7484
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Generar el PDF de la compra."
         Top             =   180
         Width           =   585
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   2235
         Top             =   120
      End
      Begin VB.CommandButton consultar 
         Height          =   375
         Left            =   9015
         Picture         =   "compres.frx":7A0E
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "Buscar Registres"
         Top             =   180
         Width           =   585
      End
      Begin VB.CommandButton sortir 
         Height          =   375
         Left            =   9600
         Picture         =   "compres.frx":7F98
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Sortir"
         Top             =   180
         Width           =   585
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   7815
         Picture         =   "compres.frx":8522
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Imprimir Comanda de compra."
         Top             =   180
         Width           =   585
      End
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   7230
         Picture         =   "compres.frx":8AAC
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Enviar per mail amb pdf."
         Top             =   180
         Width           =   585
      End
      Begin VB.CommandButton Command8 
         Height          =   375
         Left            =   6645
         Picture         =   "compres.frx":9036
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Selecciona comandes pendents de comprar"
         Top             =   180
         Width           =   585
      End
      Begin VB.Timer Timer1 
         Interval        =   700
         Left            =   7725
         Top             =   180
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   105
         Picture         =   "compres.frx":95C0
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Alta  Registres"
         Top             =   165
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   960
         Picture         =   "compres.frx":9B4A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Eliminacio Registres"
         Top             =   165
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   525
         Picture         =   "compres.frx":A0D4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Edicio del  Registres"
         Top             =   165
         Width           =   420
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   1395
         Picture         =   "compres.frx":A65E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Acceptar canvis"
         Top             =   165
         Width           =   420
      End
      Begin VB.Label estat 
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
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1995
         TabIndex        =   61
         Top             =   165
         Width           =   1560
      End
   End
   Begin VB.Menu mllistas 
      Caption         =   "Llistats"
      Begin VB.Menu llistatcompresperarticle 
         Caption         =   "Llistat de compres per article"
         Begin VB.Menu llcpmateriaprimera 
            Caption         =   "Materia prima i Varis"
         End
         Begin VB.Menu llctintes 
            Caption         =   "Tintes"
         End
      End
      Begin VB.Menu m_llitstatdekgtotals 
         Caption         =   "Llistat de Kg Totals entre dates"
         Begin VB.Menu llmateriaprimera 
            Caption         =   "Materia primera"
         End
         Begin VB.Menu Lltintes 
            Caption         =   "Tintes"
         End
         Begin VB.Menu llvaris 
            Caption         =   "Varis"
         End
      End
      Begin VB.Menu mllistatrefcupu 
         Caption         =   "Llistat EXCEL de Kg Totals entre dates"
         Begin VB.Menu llsitatambcupu 
            Caption         =   "Amb referencia cupu"
         End
         Begin VB.Menu llistattoteslescompres 
            Caption         =   "Totes les compres"
         End
      End
      Begin VB.Menu mllistatpendentderebre 
         Caption         =   "Llistat comandes pendents de rebre"
      End
      Begin VB.Menu menullistatcompresdetot 
         Caption         =   "Llistat de totes les compres entre dues dates XLS"
      End
   End
   Begin VB.Menu mfiltre 
      Caption         =   "Filtre"
      Begin VB.Menu mpendentsentrega 
         Caption         =   "Pendents d'entregar."
      End
      Begin VB.Menu m_entregades 
         Caption         =   "Entregades."
      End
      Begin VB.Menu mpendentsdenviar 
         Caption         =   "Pendents d'enviar."
      End
      Begin VB.Menu mcomandesclientconcret 
         Caption         =   "Comandes d'un client concret."
      End
   End
End
Attribute VB_Name = "comandescompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).

Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM
Dim vllistatcomprespendents As Boolean
Dim vimprimint As Boolean
Sub enviarcompra()
  
Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = "Pedido Inplacsa Nº: 16615"
objMessage.From = "miquel.inplacsa@gmail.com"
objMessage.To = "miquel.inplacsa@gmail.com"
objMessage.TextBody = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\temp\cosmissatge.txt", 1).ReadAll
objMessage.AddAttachment "c:\temp\Pedido_16615.pdf"


'==This section provides the configuration information for the remote SMTP server.

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

'Name or IP of Remote SMTP Server
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"

'Type of authentication, NONE, Basic (Base64 encoded), NTLM
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

'Your UserID on the SMTP server
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = "miquel.inplacsa@gmail.com"

'Your password on the SMTP server
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "ipc990900ipc"

'Server port (typically 25)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465

'Use SSL for the connection (False or True)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

objMessage.Configuration.Fields.Update

'==End remote SMTP server configuration section==

objMessage.Send
End Sub
Private Sub acceptarlinia_Click()
  oklinia
  'For i = 0 To liniescompra.Recordset.Fields.Count - 1
  '   If valorcamp(liniescompra.Recordset.Fields(i).Name) <> Empty And valorcamp(liniescompra.Recordset.Fields(i).Name) <> liniescompra.Recordset.Fields(i).Value Then MsgBox "canvi"
  'Next i
End Sub
Sub oklinia()
  Dim bklinia As Long
  If liniescompra.Recordset.EditMode = 0 Then Exit Sub
  'If liniescompra.Recordset.EOF Then Exit Sub
  If Not comprovarcampsminims Then MsgBox "Falta emplenar camps necessaris", vbExclamation + vbOKOnly, "Atenció": Exit Sub
  If liniescompra.Recordset.EditMode > 0 Then
    bklinia = liniescompra.Recordset!idliniacompra
    liniescompra.Recordset.Update
    liniescompra.Recordset.FindFirst "idliniacompra=" + atrim(bklinia)
    borrarlesliniesdedescripcio False
    If liniescompra.Recordset!tipusmaterialcomprat = "M" Then
       generar_linies_descripcio
       generar_reserva_corresponent
    End If
    If liniescompra.Recordset!tipusmaterialcomprat = "T" Then
      generar_linies_descripcio_tintes
    End If
    If liniescompra.Recordset!tipusmaterialcomprat = "V" Then
      generar_linies_descripcio_varis
    End If
    
  End If
  If liniescompra.Recordset!tipusmaterialcomprat = "M" Then sumar_kilos
  actualitzar_valors_comanda
  fdescmat.Enabled = False
End Sub

Function comprovarcampsminims() As Boolean
  comprovarcampsminims = False

  If formselecciotipuscompra.tag = "M" Then
        If atrim(itl) = "" Or atrim(icares) = "" Or atrim(iobert) = "" Or cadbl(iample) = 0 Or cadbl(iespesor) = 0 Or cadbl(preu) = 0 Then Exit Function
        If Not IsNumeric(diamext) Then diamext = "0"
        If Not IsNumeric(mandril) Then mandril = "0"
        comprovarcampsminims = True
          Else: If cadbl(preu) > 0 And cadbl(kilosxrcomprar) > 0 Then comprovarcampsminims = True
  End If
End Function
Function valorcamp(camp As String) As Variant
   Dim Control As Control
   On Error Resume Next
   For Each Control In comandescompra
     If campbd(Control) Then
      If Control.DataField = camp And Control.Container.Name = "fdescmat" Then
          valorcamp = Control.Value
          If valorcamp = "" Then valorcamp = Control.Text: GoTo fi
      End If
     End If
   Next
fi:
End Function
Function campbd(c As Control) As Boolean
  On Error Resume Next
  If c.DataField = "" Then
     campdb = False
    Else: campbd = True
  End If
End Function
Sub generar_linies_descripcio_tintes()
   Dim rsttintes As Recordset
   With liniescompra.Recordset
   Set rsttintes = dbtintes.OpenRecordset("SELECT codi, tintesreferencies.referencia, tipusbidons.nombido FROM (tintes LEFT JOIN tintesreferencies ON tintes.idtinta = tintesreferencies.idtinta) LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id Where tintesreferencies.id = " + atrim(cadbl(!diametreext)) + "")
   If rsttintes.EOF Then Exit Sub
   afegir_linia_descripcio atrim(cadbl(!idliniacompra)), 1, "Ref: " + atrim(rsttintes!referencia)
   afegir_linia_descripcio atrim(cadbl(!idliniacompra)), 2, atrim(rsttintes!nombido)
   End With
End Sub
Sub generar_linies_descripcio_varis()
 Dim rstcol As Recordset
   Dim rstsubcol As Recordset
   Dim rstmat As Recordset
   Dim merror As String
   Dim solapa As Boolean
   Dim desc As String
   Dim espesor As String
   Dim fammat As String
   Dim famcol As String
   Dim famad As String
   merror = "Error"
   liniescompra.Database.Execute "delete * from liniesdescripcio where descripcio='' and idliniacompra=" + atrim(cadbl(liniescompra.Recordset!idliniacompra))
   liniesdescripcio.Refresh
   'assigno el material i la fam de colorant
   Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(liniescompra.Recordset!codimaterial)))
   'If rstmat.EOF And combomaterial.tag = "Unitats" Then
'    Set rstmat = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(cadbl(liniescompra.Recordset!codimaterial)) + "'")
   'End If
   If Not rstmat.EOF Then
     ' If combomaterial.tag = "Unitats" Then generar_unaliniadeproducte rstmat: GoTo ultima
      'fammat
      Set rstcol = dbtmp.OpenRecordset("select * from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)))
      If Not rstcol.EOF Then
           fammat = atrim(rstcol!descripcio)
           Set rstsubcol = dbtmp.OpenRecordset("select * from subfamiliesmaterials where codi=" + atrim(cadbl(rstmat!subfamilia)))
           If Not rstsubcol.EOF Then fammat = atrim(fammat) + " " + atrim(rstsubcol!descripcio)
      End If
      
      'famcol
      Set rstcol = dbtmp.OpenRecordset("select * from familiescolorants where codi=" + atrim(cadbl(rstmat!familiacol)))
      If Not rstcol.EOF Then
           famcol = atrim(rstcol!descripcio)
           Set rstsubcol = dbtmp.OpenRecordset("select * from subfamiliescolorants where codi=" + atrim(cadbl(rstmat!subfamiliacol)))
           If Not rstsubcol.EOF Then famcol = atrim(famcol) + " " + atrim(rstsubcol!descripcio)
      End If
      
      'famad
      Set rstcol = dbtmp.OpenRecordset("select * from familiesaditius where codi=" + atrim(cadbl(rstmat!familiaad)))
      If Not rstcol.EOF Then
           famad = atrim(rstcol!descripcio)
           Set rstsubcol = dbtmp.OpenRecordset("select * from subfamiliesaditius where codi=" + atrim(cadbl(rstmat!subfamiliaad)))
           If Not rstsubcol.EOF Then famad = atrim(famad) + " " + atrim(rstsubcol!descripcio)
      End If
      
       Else:
         
         merror = "Falta el material.": GoTo fi
   End If
   afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 1, atrim(rstmat!refproducte)
   'afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 2, atrim(desc) + " " + atrim(fammat)
   'afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 3, atrim(famcol)
   'If Len(famad) > 3 Then afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 4, atrim(famad)
ultima:
   liniesdescripcio.Refresh
  Exit Sub
fi:
  MsgBox merror, vbCritical, "Error"
End Sub
Sub generar_linies_descripcio()
   Dim rstcol As Recordset
   Dim rstsubcol As Recordset
   Dim rstmat As Recordset
   Dim merror As String
   Dim solapa As Boolean
   Dim desc As String
   Dim espesor As String
   Dim fammat As String
   Dim famcol As String
   Dim famad As String
   merror = "Error"
   liniescompra.Database.Execute "delete * from liniesdescripcio where descripcio='' and idliniacompra=" + atrim(cadbl(liniescompra.Recordset!idliniacompra))
   liniesdescripcio.Refresh
   'If liniesdescripcio.Recordset.RecordCount > 0 Then Exit Sub
'creo les linies
   
   'assigno el material i la fam de colorant
   Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(liniescompra.Recordset!codimaterial)))
   If rstmat.EOF And combomaterial.tag = "Unitats" Then
       Set rstmat = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(cadbl(liniescompra.Recordset!codimaterial)) + "'")
   End If
   If Not rstmat.EOF Then
      If combomaterial.tag = "Unitats" Then generar_unaliniadeproducte rstmat: GoTo ultima
      'fammat
      Set rstcol = dbtmp.OpenRecordset("select * from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)))
      If Not rstcol.EOF Then
           fammat = atrim(rstcol!descripcio)
           Set rstsubcol = dbtmp.OpenRecordset("select * from subfamiliesmaterials where codi=" + atrim(cadbl(rstmat!subfamilia)))
           If Not rstsubcol.EOF Then fammat = atrim(fammat) + " " + atrim(rstsubcol!descripcio)
      End If
      
      'famcol
      Set rstcol = dbtmp.OpenRecordset("select * from familiescolorants where codi=" + atrim(cadbl(rstmat!familiacol)))
      If Not rstcol.EOF Then
           famcol = atrim(rstcol!descripcio)
           Set rstsubcol = dbtmp.OpenRecordset("select * from subfamiliescolorants where codi=" + atrim(cadbl(rstmat!subfamiliacol)))
           If Not rstsubcol.EOF Then famcol = atrim(famcol) + " " + atrim(rstsubcol!descripcio)
      End If
      
      'famad
      Set rstcol = dbtmp.OpenRecordset("select * from familiesaditius where codi=" + atrim(cadbl(rstmat!familiaad)))
      If Not rstcol.EOF Then
           famad = atrim(rstcol!descripcio)
           Set rstsubcol = dbtmp.OpenRecordset("select * from subfamiliesaditius where codi=" + atrim(cadbl(rstmat!subfamiliaad)))
           If Not rstsubcol.EOF Then famad = atrim(famad) + " " + atrim(rstsubcol!descripcio)
      End If
      
       Else:
         
         merror = "Falta el material.": GoTo fi
   End If
   desc = IIf(atrim(liniescompra.Recordset!semielaborat) = "T", "TUBO ", "LAMINA ")
   If cadbl(liniescompra.Recordset!obert) > 0 Then desc = desc + "ABIERTO " + atrim(cadbl(liniescompra.Recordset!obert)) + " LADO/S"
   afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 1, atrim(rstmat!refproducte)
   afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 2, atrim(desc) + " " + atrim(fammat)
   afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 3, atrim(famcol)
   If Len(famad) > 3 Then afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 4, atrim(famad)
   If liniescompra.Recordset!microperforat Then afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 5, "MICROPERFORADO"
   
   
      
   desc = "ANCHO " + atrim(cadbl(liniescompra.Recordset!ample) * 10)
   If cadbl(liniescompra.Recordset!plegat) > 0 Then desc = desc + "/" + atrim(cadbl(liniescompra.Recordset!plegat) * 10)
   If cadbl(liniescompra.Recordset!solapa) > 0 Then desc = desc + "+" + atrim(cadbl(liniescompra.Recordset!solapa) * 10): solapa = True
   If cadbl(liniescompra.Recordset!grmm2) > 0 Then
      espesor = atrim(liniescompra.Recordset!grmm2) + " Grm/m2 "
     Else: espesor = atrim(cadbl(liniescompra.Recordset!micres)) + "µ "
   End If
   desc = desc + " MM " + IIf(solapa, " SOLAPA ", "") + espesor
   If rstmat!mesuarespcompra = "Galgues" Then desc = desc + " G/" + atrim(cadbl(liniescompra.Recordset!micres) * 4)
   afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 10, desc
   
   desc = ""
   If cadbl(liniescompra.Recordset!carestractat) > 0 Then
      desc = "TRATADO " + atrim(cadbl(liniescompra.Recordset!carestractat)) + " CARA/S"
      afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 15, desc
   End If
   
   
   desc = ""
   If cadbl(liniescompra.Recordset!diametreext) > 0 Then desc = "DIAMETRO EXTERIOR: " + atrim(cadbl(liniescompra.Recordset!diametreext) * 10) + " MM"
   If desc <> "" Then afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 50, desc
   If cadbl(liniescompra.Recordset!mandril) > 0 Then desc = "MANDRIL: " + atrim(cadbl(liniescompra.Recordset!mandril) * 10) + " MM"
   If desc <> "" Then afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 51, desc
   liniadedetalldecomandes
ultima:
   liniesdescripcio.Refresh
  Exit Sub
fi:
  MsgBox merror, vbCritical, "Error"
End Sub
Sub generar_unaliniadeproducte(rstmat As Recordset)
   Dim desc As String
   afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 10, rstmat!descripcio + "."
End Sub
Sub afegir_linia_descripcio(idlinia As Double, ordre As Byte, desc As String)
   If desc = "-" Or desc = "- -" Then Exit Sub
   liniesdescripcio.Database.Execute "delete * from liniesdescripcio where idliniacompra=" + atrim(idlinia) + " and ordre=" + atrim(ordre)
   liniesdescripcio.Recordset.AddNew
   liniesdescripcio.Recordset!idliniacompra = idlinia
   liniesdescripcio.Recordset!ordre = ordre
   liniesdescripcio.Recordset!descripcio = desc
   liniesdescripcio.Recordset.Update
End Sub
Sub sumar_kilos()
  Dim kilos As Double
  Dim rstk As Recordset
  kilos = 0
  Set rstk = comandesxlinia.Database.OpenRecordset("select sum(kgcompra) as total from comandesxlinia where idliniacompra=" + atrim(liniescompra.Recordset!idliniacompra))
  If Not rstk.EOF Then
     kilos = cadbl(rstk!total)
  End If
  If liniescompra.Recordset.EditMode = 0 Then liniescompra.Recordset.Edit
  kilosxrcomprar = atrim(kilos)
  liniescompra.Recordset.Update
End Sub
Private Sub alta_Click()
 novacomanda
End Sub
Sub novacomanda()
  Dim rs As Recordset
  Dim bk As Double
  capcalera.Refresh
  capcalera.Recordset.AddNew
  
  fcapcalera.Enabled = True
  Set rs = dbtmpb.OpenRecordset("select max(numcomanda) as gran from capcalera")
  If Not rs.EOF Then
    cnumcomanda = cadbl(rs!gran) + 1
    Else: cnumcomanda = "130000001187"
  End If
  bk = cnumcomanda
  empresa = "Inplacsa"
  Set rs = Nothing
  capcalera.Recordset.Update
  capcalera.Recordset.FindFirst "numcomanda=" + atrim(bk)
  capcalera.Recordset.Edit
  If Screen.ActiveForm.Name = "comandescompra" Then Text2.SetFocus
End Sub

Private Sub borrarliniesdescripcio_Click()
  borrarlesliniesdedescripcio True
End Sub
Sub borrarlesliniesdedescripcio(Optional demanarok As Boolean)
'borro totes les linies anteriors
   If Not demanarok Then GoTo segur
   If MsgBox("Segur que vols borrar totes les linies de descripció", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
segur:
    liniescompra.Database.Execute "delete * from liniesdescripcio where " + IIf(Not demanarok, "(ordre <30 or ordre>40) and ", "") + " idliniacompra=" + atrim(cadbl(liniescompra.Recordset!idliniacompra))
    
    liniesdescripcio.Refresh
   End If
End Sub
Private Sub capcalera_Reposition()
    refrescasubtaules
    If Not capcalera.Recordset.EOF Then
       capcalera.caption = "Comanda " + atrim(capcalera.Recordset.AbsolutePosition + 1) + "/" + atrim(capcalera.Recordset.RecordCount)
         Else: capcalera.caption = ""
    End If
End Sub
Function mirarsipendentoparcial(numc As Double) As String
   Dim rstm As Recordset
   Dim rstt As Recordset
   Set rstm = capcalera.Database.OpenRecordset("SELECT capcalera.numcomanda, liniescompra.totentregat FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((capcalera.numcomanda)=" + atrim(numc) + ") AND ((liniescompra.totentregat)=True));")
   Set rstt = capcalera.Database.OpenRecordset("SELECT capcalera.numcomanda, liniescompra.totentregat FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((capcalera.numcomanda)=" + atrim(numc) + ") );")
   If rstm.EOF Or rstt.EOF Then Exit Function
   rstm.MoveLast: rstm.MoveFirst
   rstt.MoveLast: rstt.MoveFirst
   
   If rstm.RecordCount = rstt.RecordCount Then mirarsipendentoparcial = "Material rebut."
   If rstm.RecordCount < rstt.RecordCount Then mirarsipendentoparcial = "Entrega Parcial..."
   If rstm.RecordCount = 0 Then mirarsipendentoparcial = "Pendent..."
   Set rstm = Nothing
   Set rstt = Nothing
End Function
Sub refrescasubtaules()
   
   If Not capcalera.Recordset.EOF Then
    imatgeenviat(0).visible = capcalera.Recordset!enviat
    imatgeimpres(0).visible = capcalera.Recordset!imprimit
    If Not capcalera.Recordset!enviat Then
     titprecomandafins.visible = True
     precomandafins.visible = True
      Else
       titprecomandafins.visible = False
       precomandafins.visible = False
    End If
    If liniescompra.RecordSource <> "" Then
      liniesdescripcio.RecordSource = "select * from liniesdescripcio where idliniacompra=" + atrim(cadbl(liniescompra.Recordset!idcompra))
        comandesxlinia.RecordSource = "select * from comandesxlinia where idliniacompra=" + atrim(cadbl(liniescompra.Recordset!idliniacompra))
    End If
    liniescompra.RecordSource = "select * from liniescompra where idcompra=" + atrim(cadbl(capcalera.Recordset!id))
    msgpendent = mirarsipendentoparcial(capcalera.Recordset!numcomanda)  '"Pendent..."
    If msgpendent = "Material rebut." Then
        fcapcalera.BackColor = &H9AA6FA
       Else
          fcapcalera.BackColor = &HFEC7C7
    End If
     Else
      liniescompra.RecordSource = ""
      liniesdescripcio.RecordSource = ""
      comandesxlinia.RecordSource = ""
      msgpendent = ""
      fcapcalera.BackColor = &HFEC7C7
  End If
  liniescompra.Refresh
  liniesdescripcio.Refresh
  comandesxlinia.Refresh
End Sub
Sub triar_tintes()
  Dim rstmat As Recordset
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
  'Set formseleccio.Data1.Recordset = dbtintes.OpenRecordset("SELECT codi, descripcio, referenciacolor, refproveidor, nomproveidor, DescripcioSerie , descripciofam, descripciosubfam,descripciofamcol, descripciosubfamcol FROM tintes_tot ")
  formseleccio.Data1.RecordSource = "SELECT tintes.codi, tintes.descripcio, tintes.referenciacolor, tintesreferencies.referencia, tipusbidons.nombido,tipusbidons.litrescompres,tintesreferencies.id FROM (tintesreferencies INNER JOIN tintes ON tintesreferencies.idtinta = tintes.idtinta) INNER JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id where tintesreferencies.predeterminada=true and tintesreferencies.codiproveidor=" + atrim(capcalera.Recordset!codiproveidor)
'  "SELECT tintes.codi, tintes.descripcio, tintes.referenciacolor, tintesreferencies.referencia FROM tintesreferencies INNER JOIN tintes ON tintesreferencies.idtinta = tintes.idtinta where tintesreferencies.codiproveidor=" + atrim(capcalera.Recordset!codiproveidor))
  '"select codi,descripcio,referenciacolor from tintes order by descripcio")
  'formseleccio.Data1.RecordSource = "select * from proveidors"
  formseleccio.refrescar
  formseleccio.width = 9200
  formseleccio.DBGrid2.Columns(0).width = 800
  formseleccio.DBGrid2.Columns(1).width = 3000
  formseleccio.DBGrid2.Columns(2).width = 2000
  formseleccio.DBGrid2.Columns(3).width = 1200
  formseleccio.DBGrid2.Columns(6).width = 0
  formseleccio.Show 1
  If seleccioret = 1 Then
   liniescompra.Recordset!codimaterial = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   liniescompra.Recordset!nommaterial = Mid(atrim(formseleccio.Data1.Recordset!descripcio), 1, 40)
   liniescompra.Recordset!diametreext = atrim(cadbl(formseleccio.Data1.Recordset!id))
   liniescompra.Recordset!quantitatkg = atrim(cadbl(formseleccio.Data1.Recordset!litrescompres))
   combomaterial.Text = atrim(liniescompra.Recordset!nommaterial)
   kilosxrcomprar = cadbl(liniescompra.Recordset!quantitatkg)
   combomaterial.tag = "Unitats"
   'Set rstmat = formseleccio.Data1.Recordset
   'possar_families liniescompra.Recordset!codimaterial, rstmat
  End If
  Unload formseleccio
End Sub

Sub triar_material()
  Dim rstmat As Recordset
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  'formseleccio.Data1.DatabaseName = cami
  Set formseleccio.Data1.Recordset = dbtmp.OpenRecordset("select * from materials where codi>499 and proveidor=" + atrim(cadbl(capcalera.Recordset!codiproveidor)) + " order by descripcio")
  'formseleccio.Data1.RecordSource = "select * from proveidors"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   liniescompra.Recordset!codimaterial = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   liniescompra.Recordset!nommaterial = atrim(formseleccio.Data1.Recordset!descripcio)
   liniescompra.Recordset!grmm2 = cadbl(formseleccio.Data1.Recordset!grmm2)
   possarmicresogrmm2
   combomaterial.Text = liniescompra.Recordset!nommaterial
   combomaterial.tag = atrim(formseleccio.Data1.Recordset!mesuarespcompra)
   Set rstmat = formseleccio.Data1.Recordset
   possar_families liniescompra.Recordset!codimaterial, rstmat
  End If
  Unload formseleccio
End Sub
Sub possarmicresogrmm2()
  If liniescompra.Recordset.EOF And liniescompra.Recordset.EditMode = 0 Then Exit Sub
  If liniescompra.Recordset!grmm2 > 0 Then
      iespesor.DataField = "grmm2"
      iespesor.Locked = True
      lblLabels(15).caption = "Grm/m2"
      lblLabels(15).ForeColor = QBColor(9)
        Else
          iespesor.DataField = "micres"
          lblLabels(15).caption = "Micres"
          lblLabels(15).ForeColor = QBColor(0)
          iespesor.Locked = False
          
   End If

End Sub
Sub possar_families(codimat As Long, rstmat As Recordset)
   With liniescompra.Recordset
   !familia = rstmat!familia
   !subfamilia = rstmat!subfamilia
   !familiacol = rstmat!familiacol
   !subfamiliacol = rstmat!subfamiliacol
   !familiaad = rstmat!familiaad
   !subfamiliaad = rstmat!subfamiliaad
   End With
End Sub

Private Sub Command10_Click()
 
End Sub

Private Sub Command2_Click()
  afegir_observacio
End Sub
Sub afegir_observacio()
   Dim txt As String
   Dim gran As Double
   Dim rstg As Recordset
   txt = "."
   While txt <> ""
    Load forminputbox
    forminputbox.etmissatge = "Entra la observació que vols entrar o res per acabar."
    forminputbox.Show 1
    If forminputbox.bacceptar.tag = "1" Then
       txt = atrim(forminputbox.cresposta)
         Else: txt = ""
    End If
    Unload forminputbox
    'txt = InputBox("Entra la observació que vols entrar." + Chr(10) + "Fes cancelar o no escriguis per acabar d'afegir.", "Observacions")
    gran = 30
    If txt <> "" Then
      txt = Mid(txt, 1, 50)
      Set rstg = liniescompra.Database.OpenRecordset("select max(ordre) as gran from liniesdescripcio where ordre<48 and idliniacompra=" + atrim(liniescompra.Recordset!idliniacompra))
      If Not rstg.EOF Then gran = cadbl(rstg!gran)
      If gran < 30 Then gran = 29
      If gran > 39 Then MsgBox "No es poden afegir mes de 10 linies d'observacions": Exit Sub
      afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), gran + 1, txt
      liniesdescripcio.Refresh
    End If
   Wend
   Set rstg = Nothing
End Sub
Private Sub combomaterial_DropDown()
 triar_material
 ensenyar_camps_compra
  SendKeys "{tab}"
End Sub

Private Sub combomaterial_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then triar_material
End Sub

Private Sub Command3_Click()
  Dim idlinia As Long
  If combomaterial.tag <> "Unitats" Then MsgBox "LA MODIFICACIÓ DE QUALSEVOL DELS PARAMETRES DE LA COMPRA QUE AFECTIN A LES COMANDES RELACIONADES NO ES COMPROVARAN.", vbCritical + vbOKOnly, "MOLTA ATENCIÓ"
  If Not liniescompra.Recordset.EOF Then idlinia = liniescompra.Recordset!idliniacompra
  If capcalera.Recordset.EditMode > 0 Then capcalera.Recordset.Update
  capcalera.Recordset.Bookmark = capcalera.Recordset.LastModified
  liniescompra.Recordset.FindFirst "idliniacompra=" + atrim(cadbl(idlinia))
  If liniescompra.Recordset.EditMode > 0 Then MsgBox "S'està editant la linia.", vbCritical, "Atenció"
  liniescompra.Recordset.Edit
  fdescmat.Enabled = True
  liniescompra.Recordset!idcompra = capcalera.Recordset!id
  'triar_material
  If itl.visible Then
     itl.SetFocus
       Else: preu.SetFocus
  End If

End Sub

Private Sub Command4_Click()
   Dim kilosxrstoc As Double
   Dim comandavisual As String
  If capcalera.Recordset!enviat Or capcalera.Recordset!imprimit Then
     MsgBox "OJO que aquesta comanda ja està enviada al proveïdor.", vbCritical + vbOKOnly, "ATENCIÓ"
  End If


   If MsgBox("Segur que vols borrar aquesta linia?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
         If comandesxlinia.Recordset.EditMode > 0 Then comandesxlinia.Recordset.CancelUpdate
         kilosxrstoc = cadbl(comandesxlinia.Recordset!kgcompra)
         comandavisual = atrim(comandesxlinia.Recordset!comandavisual)
         comandesxlinia.Recordset.Delete
         kgcompralinia = kilosxrstoc
         If comandavisual <> "ESTOC" Then afegircomandaalinia "ESTOC"
         comandesxlinia.Refresh
         liniadedetalldecomandes
         sumar_kilos
   End If
End Sub
Function kilosdestoc(Optional canviarkgdestoc As Double)
  Dim rstcxlinia As Recordset
   Set rstcxlinia = comandesxlinia.Recordset.Clone
   'rstcxlinia.Refresh
   If Not rstcxlinia.EOF Then rstcxlinia.MoveFirst
   rstcxlinia.FindFirst "comandavisual='ESTOC'"
   kilosdestoc = 0
   If Not rstcxlinia.NoMatch Then
      kilosdestoc = rstcxlinia!kgcompra
      If canviarkgdestoc > 1 Then
          rstcxlinia.Edit
          rstcxlinia!kgcompra = canviarkgdestoc
          rstcxlinia.Update
          kilosdestoc = canviarkgdestoc
      End If
      If canviarkgdestoc = 1 Then
         rstcxlinia.Delete
         kilosdestoc = 0
      End If
   End If
End Function
Function cridarselectordestocoprecomanda() As String
   Unload triarcomandastocopre
    triarcomandastocopre.Show 1
    If cadbl(triarcomandastocopre.comanda) > 0 Then
       cridarselectordestocoprecomanda = triarcomandastocopre.comanda
         Else
           If triarcomandastocopre.estoc.BackColor = &H9AA6FA Then
               cridarselectordestocoprecomanda = "ESTOC"
              Else
                 If triarcomandastocopre.precomanda.BackColor = &H9AA6FA Then cridarselectordestocoprecomanda = "PRECOMANDA"
           End If
    End If
    Unload triarcomandastocopre
End Function
Function comprovarsilacomandajashacomprat(numc As Double) As Boolean
   Dim rstc As Recordset
   Set rstc = capcalera.Database.OpenRecordset("select * from comandesxlinia where numcomanda=" + atrim(numc))
   If Not rstc.EOF Then
       comprovarsilacomandajashacomprat = True
         Else: comprovarsilacomandajashacomprat = False
   End If
   Set rstc = Nothing
End Function
Private Sub Command5_Click()
  Dim numc As String
  Dim kilos As Double
  Dim bk As Double
  Dim descripcio As String
  Dim enviada As Boolean
  Dim kgestoc As Double
  If liniescompra.Recordset.EditMode > 0 Then MsgBox "Primer grava la linia de compra": Exit Sub
  kgestoc = kilosdestoc
  kgcompralinia = 0
  If capcalera.Recordset!enviat Or capcalera.Recordset!imprimit Then
     MsgBox "OJO que aquesta comanda ja està enviada al proveïdor.", vbCritical + vbOKOnly, "ATENCIÓ"
  End If
  numc = cridarselectordestocoprecomanda
  If cadbl(numc) > 0 Then
    If comprovarsilacomandajashacomprat(cadbl(numc)) Then MsgBox "Aquesta comanda ja s'ha comprat": Exit Sub
    If comparasielmaterialcorrespon(atrim(numc), liniescompra.Recordset!idliniacompra) <> 1 Then
       If MsgBox("No coincideixen les carecteristiques d'aquesta comanda amb les de la compra." + Chr(10) + "Vols continuar comprant per aquesta comanda?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then
           GoTo fi
       End If
    End If
    If Not comprovarsiesmaterialexacte(cadbl(numc)) Then MsgBox "Aquest material no coincideix amb el de la comanda.": Exit Sub
    afegircomandaalinia atrim(numc)
    Else
     numc = UCase(numc)
     If numc = "ESTOC" Or numc = "PRECOMANDA" Then
        If numc = "PRECOMANDA" Then descripcio = InputBox("Entra una descripcio per aquest PRECOMANDA", "PRECOMANDA")
        afegircomandaalinia atrim(numc), descripcio
          Else: MsgBox "No hi ha res a afegir": GoTo fi
     End If
     
  End If
  bk = comandesxlinia.Recordset!id
  
  actualitzar_valors_comanda
  sumar_kilos
  comandesxlinia.Recordset.FindFirst "Id=" + atrim(cadbl(bk))
  If comandesxlinia.Recordset.NoMatch Then MsgBox "Error assignant kilos": GoTo fi
  kgestoc = 0
  If kgestoc = 0 Then
   kilos = cadbl(comandesxlinia.Recordset!kgpendents)
   kilos = cadbl(InputBox("Entra els kilos que vols comprar." + Chr(10) + Chr(13) + "Si es ESTOC entra el total de Kg de la comanda i assignaré la diferencia.", "Kilos per comprar", kilos))
   If comandesxlinia.Recordset!comandavisual = "ESTOC" Then kilos = kilos - cadbl(liniescompra.Recordset!quantitatkg)
     Else
       kilos = cadbl(InputBox("Entra els kilos que vols comprar.", "Kilos per comprar", kgestoc))
       '+ Chr(10) + Chr(13) + "NOMES POTS ASSIGNAR-HI " + atrim(kgestoc) + " KG (que estan amb estoc)"
       If kilos = 0 Then
            MsgBox "Massa quilos nomes podies possar " + atrim(kgestoc) + " Kg", vbCritical, "Error"
            comandesxlinia.Recordset.Delete
            GoTo actualitzarlinies
       End If
  End If
  comandesxlinia.Recordset.Edit
  comandesxlinia.Recordset!kgcompra = cadbl(comandesxlinia.Recordset!kgcompra) + kilos
  comandesxlinia.Recordset.Update
  If kgestoc > 0 And kilos > 0 Then
    comandesxlinia.Recordset.FindFirst "comandavisual='ESTOC'"
    If Not comandesxlinia.Recordset.NoMatch Then
        comandesxlinia.Recordset.Edit
        comandesxlinia.Recordset!kgcompra = kgestoc - kilos
        comandesxlinia.Recordset.Update
        If comandesxlinia.Recordset!kgcompra = 0 Then comandesxlinia.Recordset.Delete
    End If
  End If
actualitzarlinies:
  liniadedetalldecomandes
  actualitzar_valors_comanda
  sumar_kilos
  
fi:
End Sub
Function comprovarsiesmaterialexacte(numc As Double) As Boolean
   Dim rstc As Recordset
   materialexacte = False
   comprovarsiesmaterialexacte = True
   Set rstc = dbtmp.OpenRecordset("SELECT comandes_extres.materialexacte, comandes.materialex FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda Where comandes.comanda = " + atrim(numc))
   If Not rstc.EOF Then materialexacte = cabool(rstc!materialexacte)
   If materialexacte Then
    If cadbl(liniescompra.Recordset!codimaterial) <> cadbl(rstc!materialex) Then
         comprovarsiesmaterialexacte = False
          Else: comprovarsiesmaterialexacte = True
    End If
   End If
End Function

Sub liniadedetalldecomandes()
  Dim desc As String
  Dim rstl As Recordset
  Set rstl = comandesxlinia.Recordset.Clone
  If Not rstl.EOF Then
     If cadbl(rstl!numcomanda) > 0 Then desc = "Pedido: " + atrim(rstl!numcomanda): rstl.MoveNext
     If Not rstl.EOF Then
        If cadbl(rstl!numcomanda) > 0 Then desc = desc + ", " + atrim(rstl!numcomanda): rstl.MoveNext
     End If
     If Not rstl.EOF Then
        If cadbl(rstl!numcomanda) > 0 Then desc = desc + ", ...": rstl.MoveNext
     End If
  End If
  afegir_linia_descripcio atrim(cadbl(liniescompra.Recordset!idliniacompra)), 49, desc
  liniesdescripcio.Refresh
End Sub
Sub actualitzar_valors_comanda()
   Dim rstc As Recordset
   If comandesxlinia.Recordset.EditMode > 0 Then
       comandesxlinia.Recordset.CancelUpdate
   End If
   Set rstc = comandesxlinia.Recordset.Clone
   If Not rstc.EOF Then rstc.MoveFirst
   While Not rstc.EOF
     possarvalorscomandaalinia rstc
     rstc.MoveNext
   Wend
   Set rstc = Nothing
   calcular_totals_comanda capcalera.Recordset!numcomanda
   comandesxlinia.Refresh
End Sub
Sub possarvalorscomandaalinia(rstc As Recordset)
    Dim rstcom As Recordset
    Set rstcom = dbtmp.OpenRecordset("select cantitatex,pes1000mtrs from comandes where comanda=" + atrim(rstc!numcomanda))
    If Not rstcom.EOF Then
       rstc.Edit
       rstc!kgcomanda = cadbl((cadbl((rstcom!cantitatex)) / 1000) * cadbl(rstcom!pes1000mtrs), 0)
       rstc!kgreservats = cadbl((metresreservats(rstc!numcomanda) / 1000) * cadbl(rstcom!pes1000mtrs), 0)
       rstc!kgpendents = cadbl(cadbl(rstc!kgcomanda) - cadbl(rstc!kgreservats) - cadbl(rstc!kgcompra), 0)
       If rstc!comandavisual = "ESTOC" Or rstc!comandavisual = "PRECOMANDA" Then rstc!kgpendents = 0
       rstc.Update
    End If
    Set rstcom = Nothing
End Sub
Function metresreservats(numc As Double) As Double
    Dim metres As Double
    metres = 0
    Set rststocks = dbstocks.OpenRecordset("select sum(metres) as total from parcials where comanda='" + atrim(numc) + "' and not utilitzada ")
    If Not rststocks.EOF Then
         metres = cadbl(rststocks!total)
    End If
    Set rststocks = dbstocks.OpenRecordset("select sum(metres) as total from percomandaoclient where numcomanda=" + atrim(numc) + "")
    If Not rststocks.EOF Then
        metres = metres + cadbl(rststocks!total)
    End If
    metresreservats = metres
    Set rststocks = Nothing
End Function
Sub afegircomandaalinia(numc As String, Optional descripcio As String)
  If comandesxlinia.Recordset.EditMode > 0 Then comandesxlinia.Recordset.CancelUpdate
  If numc = "ESTOC" Then
     comandesxlinia.Recordset.FindFirst "comandavisual='ESTOC'"
     If comandesxlinia.Recordset.NoMatch Then
          comandesxlinia.Recordset.AddNew
        Else: comandesxlinia.Recordset.Edit
     End If
    Else
     comandesxlinia.Recordset.AddNew
  End If
  comandesxlinia.Recordset!idliniacompra = liniescompra.Recordset!idliniacompra
  comandesxlinia.Recordset!numcomanda = cadbl(numc)
  comandesxlinia.Recordset!comandavisual = numc
  If kgcompralinia > 0 Then comandesxlinia.Recordset!kgcompra = cadbl(comandesxlinia.Recordset!kgcompra) + kgcompralinia
  comandesxlinia.Recordset!descripcio = UCase(Mid(atrim(descripcio), 1, 100))
  comandesxlinia.Recordset.Update
  comandesxlinia.Recordset.Bookmark = comandesxlinia.Recordset.LastModified
  kgcompralinia = 0
End Sub
Function comparasielmaterialcorrespon(comanda As String, liniacompra As Double) As Byte
   Dim rstcom As Recordset
   Dim rstlinia As Recordset
   Dim rstmaterial As Recordset
   Dim resp As Byte
   Dim micres As Double
   Dim mesuraesp As String
   If vcomprantmaterialcompatible Then comparasielmaterialcorrespon = 1: Exit Function
   Set rstcom = dbtmp.OpenRecordset("select * from comandes where comanda=" + comanda)
   If Not rstcom.EOF Then
      If rstcom!materialex < 500 Then
         MsgBox "El material de la comanda no està actualitzat als nous materials superiors al 500", vbCritical, "Atenció"
         GoTo fi
      End If
      Set rstcom = dbtmp.OpenRecordset("select * from comandes where comanda=" + comanda)
      If Not rstcom.EOF Then
         resp = 1
        If rstcom!materialex < 500 Then resp = 4: GoTo fi
      End If
   End If
   resp = 1
   If Not rstcom.EOF Then
      Set rstlinia = dbtmpb.OpenRecordset("select * from liniescompra where idliniacompra=" + atrim(liniacompra))
      Set rstmaterial = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstcom!materialex)))
      If Not rstlinia.EOF Then
          If cadbl(rstlinia!grmm2) > 0 Then
             mesuraesp = "grmm2"
            Else: mesuraesp = "micres"
          End If
          resp = 2
          If cadbl(rstlinia!familia) = cadbl(rstmaterial!familia) Then
             If cadbl(rstlinia!subfamilia) = cadbl(rstmaterial!subfamilia) Then
               If cadbl(rstlinia!familiacol) = cadbl(rstmaterial!familiacol) Then
                 If cadbl(rstlinia!subfamiliacol) = cadbl(rstmaterial!subfamiliacol) Then
                   If cadbl(rstlinia!familiaad) = cadbl(rstmaterial!familiaad) Then
                     If cadbl(rstlinia!subfamiliaad) = cadbl(rstmaterial!subfamiliaad) Then
                          resp = 1
                     End If
                   End If
                 End If
               End If
             End If
          End If
          micres = comandespendents.micresmaterial(rstcom!mesuraesp, rstcom!espessor, rstcom!tubolam)
          If micres < 0 Then micres = micres * -1
          If resp = 1 Then
             resp = 3
             'If cadbl(rstlinia!ample) >= (cadbl(rstcom!ampleesq) - 1) Then
                 If cadbl(rstlinia!plegat) = cadbl(rstcom!plegatesq) Then
                   If cadbl(rstlinia!solapa) = cadbl(rstcom!solapa) Then
                     If cadbl(rstlinia.Fields(mesuraesp)) = micres Then
                       'If assignarmat.aatrim(rstlinia!carestractat) = assignarmat.aatrim(rstcom!tractatex) Then
                         If rstlinia!obert = rstcom!oberturaex Then
                           If cabool(rstlinia!microperforat) = cabool(rstcom!micropex) Then
                             If atrim(rstlinia!semielaborat) = atrim(rstcom!tubolam) Then
                               resp = 1
                             End If
                           End If
                         End If
                       'End If
                     End If
                   End If
                 End If
              'End If
          End If
      End If
     Else: resp = 0
   End If
fi:
   comparasielmaterialcorrespon = resp
 'aquesta linia s'ha de treure per compravar bé el material
   'comparasielmaterialcorrespon = 1
End Function

Private Sub editarlinia_Click()

End Sub

Private Sub Command6_Click()
   If capcalera.Recordset.EditMode > 0 Or liniescompra.Recordset.EditMode > 0 Then MsgBox "No pots imprimir si estas editant la comanda.", vbCritical, "Error": Exit Sub
  imprimir_comanda cadbl(cnumcomanda)
End Sub
Sub generar_pdf_comanda(numc As Double)
   Dim fitxerpdftemporal As String
   Dim oapp As CRAXDDRT.Application
   Dim oreport As CRAXDDRT.Report
   
   If capcalera.Recordset.EditMode > 0 Then capcalera.Recordset.Update
   passarregistrealataulatemporal numc
   'If existeix(Environ("userprofile") + "\desktop") Then fitxerpdftemporal = Environ("userprofile") + "\desktop\Pedido_" + atrim(numc) + ".pdf"
'   If existeix(Environ("userprofile") + "\escritorio") Then fitxerpdftemporal = Environ("userprofile") + "\escritorio\Pedido_" + atrim(numc) + ".pdf"
   If existeix(Environ("userprofile") + "\downloads") Then fitxerpdftemporal = Environ("userprofile") + "\downloads\Pedido_" + atrim(numc) + ".pdf"
   If fitxerpdftemporal = "" Then fitxerpdftemporal = "c:\temp\Pedido_" + atrim(numc) + ".pdf"
   Set oapp = New CRAXDDRT.Application
   Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "comandescompres.rpt", 1)
   oreport.Database.SetDataSource (dbconsulta)
   oreport.DiscardSavedData
   oreport.ExportOptions.DiskFileName = fitxerpdftemporal
   oreport.ExportOptions.PDFExportAllPages = True
   oreport.ExportOptions.FormatType = crEFTPortableDocFormat
   oreport.ExportOptions.DestinationType = crEDTDiskFile
   oreport.EnableParameterPrompting = False
   oreport.Database.Tables.Item(1).Location = fitxertemp
   oreport.Export False
 '  wait 2
   'If existeix(fitxerpdftemporal) Then obrir_document (fitxerpdftemporal)
   MsgBox fitxerpdftemporal + vbNewLine + "          GENERAT."
End Sub
Sub imprimir_comanda(numc As Double)
    If vimprimint Then Exit Sub
    vimprimint = True
    If capcalera.Recordset.EditMode > 0 Then capcalera.Recordset.Update
    passarregistrealataulatemporal numc
    imprimircomanda numc
    If numc = 0 Then
      GoTo fi
     Else
       capcalera.Recordset.Edit
       capcalera.Recordset!imprimit = True
       capcalera.Recordset.Update
    End If
    AppActivate "Comandes de Compra."
fi:
   vimprimint = False
   'si numc=0 vol dir que no ha acabat la impresio
End Sub
Sub imprimircomanda(numc As Double)
  
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  passarregistrealataulatemporal cnumcomanda
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "comandescompres.rpt", 1)
  'oreport.Database.SetDataSource (dbconsulta)
  oreport.Database.Tables.Item(1).Location = fitxertemp
  oreport.DiscardSavedData
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
    Else
      oreport.PrintOut False, 1
  End If
  Set oapp = Nothing
  Set oreport = Nothing
  comandescompra.SetFocus

End Sub
Sub passarregistrealataulatemporal(numc As Double)
  Dim rstp As Recordset
  Dim i As Byte
  Dim rstm As Recordset
  Dim gran As Integer
  Dim rstf As Recordset
  wait 2
  borrartaulestmp
  
  dbtmpb.Execute "select * into ll_capcalera IN '" + fitxertemp + "' from capcalera where numcomanda=" + atrim(numc)
  dbtmpb.Execute "SELECT liniescompra.* into ll_liniescompra IN '" + fitxertemp + "' FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((capcalera.numcomanda)=" + atrim(numc) + "));"
  dbtmpb.Execute "SELECT liniesdescripcio.* into ll_liniesdescripcio IN '" + fitxertemp + "' FROM (capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra) RIGHT JOIN liniesdescripcio ON liniescompra.idliniacompra = liniesdescripcio.idliniacompra WHERE (((capcalera.numcomanda)=" + atrim(numc) + "));"
  dbtmpb.Execute "select materials.* into materials in '" + fitxertemp + "' from materials"
  
  wait 2
  On Error Resume Next
  dbconsulta.Execute "alter table ll_liniescompra add column desc_unitat text"
  dbconsulta.Execute "create index principal ON ll_liniescompra([idcompra]);"
  dbconsulta.Execute "create index segona ON ll_liniescompra([idliniacompra]);"
  dbconsulta.Execute "create index principal ON ll_liniesdescripcio([idliniacompra]);"
  dbconsulta.Execute "create index principal ON ll_capcalera([id]);"
  On Error GoTo 0
  
  Set rstp = dbconsulta.OpenRecordset("select max(idcompra) as gran from ll_liniescompra")
  If rstp.EOF Then Exit Sub
  dbconsulta.Execute "alter table ll_liniesdescripcio add column tipus byte"
  
  gran = 100
  If Not rstp.EOF Then gran = cadbl(rstp!gran)
  Set rstp = dbtmp.OpenRecordset("select * from proveidors_comercial where codi=" + atrim(cadbl(capcalera.Recordset!codiproveidorcomercial)))
  For i = 0 To 9
    Set rstm = capcalera.Database.OpenRecordset("select * from descripcionsmsgpeu where idmsg=" + atrim(cadbl(rstp.Fields("msg" + atrim(i + 1)))) + " order by ordre")
    If Not rstm.EOF Then dbconsulta.Execute "insert into ll_liniescompra (idcompra) values (" + atrim(gran) + ")"
    Set rstf = dbconsulta.OpenRecordset("select max(idliniacompra) as gran from ll_liniescompra ")
    If Not rstf.EOF Then gran = rstf!gran
    While Not rstm.EOF
      dbconsulta.Execute "insert into ll_liniesdescripcio (idliniacompra,ordre,descripcio,tipus) values (" + atrim(gran) + "," + atrim(rstm!ordre) + ",'" + treure_apostruf(rstm!descripcio) + "',1)"
      rstm.MoveNext
    Wend
    gran = gran + 1
  Next i
  'poso les initats correcte per cada linia
  dbconsulta.Execute "UPDATE ll_liniescompra INNER JOIN materials ON ll_liniescompra.codimaterial = materials.codi SET ll_liniescompra.desc_unitat = [materials].[mesuarespcompra];"

  
  
  dbconsulta.Execute "update ll_liniesdescripcio set tipus=2 where ordre >29 and ordre<40"
  Set rstp = dbconsulta.OpenRecordset("select * from ll_liniesdescripcio ORDER BY IDLINIACOMPRA,ORDRE")
  While Not rstp.EOF
     If rstp!ordre = 49 And atrim(rstp!descripcio) <> "" Then
       dbconsulta.Execute "insert into ll_liniesdescripcio (idliniacompra,ordre,descripcio) values (" + atrim(rstp!idliniacompra) + ",48.9,'')"
     End If
     If rstp!ordre = 50 And atrim(rstp!descripcio) <> "" Then
       dbconsulta.Execute "insert into ll_liniesdescripcio (idliniacompra,ordre,descripcio) values (" + atrim(rstp!idliniacompra) + ",49.9,'')"
     End If
     If rstp!ordre = 30 Then
       dbconsulta.Execute "insert into ll_liniesdescripcio (idliniacompra,ordre,descripcio) values (" + atrim(rstp!idliniacompra) + ",28.9,'')"
     End If
     rstp.MoveNext
  Wend
  Set rstp = Nothing
  Set rstm = Nothing
  Set rstf = Nothing
End Sub
Sub borrartaulestmp()
 On Error Resume Next
  dbconsulta.Execute "drop table ll_capcalera"
  dbconsulta.Execute "drop table ll_liniescompra"
  dbconsulta.Execute "drop table ll_liniesdescripcio"
  dbconsulta.Execute "drop table materials"
  

End Sub
Sub passarcomandaaenviada()
  capcalera.Recordset.Edit
  capcalera.Recordset!enviat = True
  capcalera.Recordset.Update
End Sub
Private Sub Command7_Click()
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim fitxerpdftemporal As String
  Dim email As String
  Dim cosmissatge As String
  Dim venviantpreuzero As Boolean
  If vimprimint Then Exit Sub
  If capcalera.Recordset.EditMode > 0 Or liniescompra.Recordset.EditMode > 0 Then MsgBox "No pots imprimir si estas editant la comanda.", vbCritical, "Error": Exit Sub
  vimprimint = True
reenviarapreuzero:
  borrarpedidostemporalsanteriors
  If capcalera.Recordset.EditMode > 0 Then capcalera.Recordset.Update
  fitxerpdftemporal = "c:\temp\Pedido_" + atrim(capcalera.Recordset!numcomanda) + ".pdf"
  passarregistrealataulatemporal cnumcomanda
  If venviantpreuzero Then possar_preus_Azero_de_la_comanda
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "comandescompres.rpt", 1)
  oreport.Database.SetDataSource (dbconsulta)
  oreport.DiscardSavedData
  oreport.ExportOptions.DiskFileName = fitxerpdftemporal
  oreport.ExportOptions.PDFExportAllPages = True
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  
  oreport.EnableParameterPrompting = False
  
  oreport.Database.Tables.Item(1).Location = fitxertemp
  
  'oreport.ExportOptions.DestinationType = crEDTEMailMAPI
  If venviantpreuzero Then
    ratoli "espera"
    wait 5
    ratoli "normal"
  End If
  
  oreport.Export False
  Set oapp = Nothing
  Set oreport = Nothing
  vimprimint = False
  If venviantpreuzero Then GoTo enviarpreuzero
  
  buscaremailproveidor email
  cosmissatge = vbCrLf + vbCrLf + vbCrLf + "Alícia Miquel" + vbCrLf + "Telf: +34 972 460 190" + vbCrLf + "e-mail: amiquel@inplacsa.com" + vbCrLf + "web: www.inplacsa.com"
  'cosmissatge = " hola"
  If enviaremail(email, "Pedido Inplacsa Nº: " + atrim(cnumcomanda), cosmissatge, fitxerpdftemporal) Then
      passarcomandaaenviada
      capcalera.Recordset.Move 0
      If capcalera.Recordset!enviat = True Then
          MsgBox "Comanda posada a la cua d'enviament per al proveïdor", vbInformation, "Bandeja de sortida"
          If hihatintes Then
            venviantpreuzero = True
            MsgBox "Passarem un correu a tintes per informar-los de la compra amb comanda sense valorar.", vbInformation, "Atenció"
          'passar copia a tintes
           GoTo reenviarapreuzero
          ' oreport.Export False
enviarpreuzero:
           If enviaremail("tintes@inplacsa.com", "Copia de la compra Nº: " + atrim(cnumcomanda), "Adjuntem comanda de compra. ", fitxerpdftemporal, True) Then
             MsgBox "Envio correcte a la cua d'enviament per a tintes@inplacsa.com", vbInformation, "Bandeja de sortida"
               Else: MsgBox "No s'ha podut enviar el mail a la bandeja de sortida", vbCritical, "Error"
           End If
         End If
        Else: r = InputBox("Error... la comanda s'ha enviat però no s'ha guardat com enviada." + Chr(10), "Error de dades")
      End If
       Else: MsgBox "La comanda no s'ha enviat al proveidor o hi ha hagut un error de dades.", vbCritical, "Error de dades"
  End If
  capcalera.Recordset.Move 0
  ratoli "normal"
End Sub
Sub possar_preus_Azero_de_la_comanda()
  Dim rst As Recordset
  Set rst = dbconsulta.OpenRecordset("select* from ll_liniescompra")
  While Not rst.EOF
    rst.Edit
     rst!preu = 0
    rst.Update
    rst.MoveNext
  Wend
  Set rst = dbconsulta.OpenRecordset("select* from ll_capcalera")
  If Not rst.EOF Then
    rst.Edit
    rst!baseimp = 0
    rst.Update
  End If
  Set rst = dbconsulta.OpenRecordset("select* from ll_capcalera")
End Sub
Function hihatintes() As Boolean
   Dim rst As Recordset
   Set rst = dbconsulta.OpenRecordset("select * from ll_liniescompra where tipusmaterialcomprat='T'")
   If Not rst.EOF Then hihatintes = True
   Set rst = Nothing
End Function
Sub borrarpedidostemporalsanteriors()
   On Error Resume Next
   Kill "c:\temp\Pedido_*.pdf"
End Sub
Sub buscaremailproveidor(email As String)
   Dim rstc As Recordset
   Set rstc = dbtmp.OpenRecordset("select emailcomandes from proveidors_comercial where codi=" + atrim(cadbl(capcalera.Recordset!codiproveidorcomercial)))
   If Not rstc.EOF Then
      email = LCase(atrim(rstc!emailcomandes))
   End If
End Sub

Private Sub Command8_Click()
  If capcalera.Recordset.EditMode > 0 Then capcalera.Recordset.CancelUpdate
  If liniescompra.Recordset.EditMode > 0 Then liniescompra.Recordset.CancelUpdate
  comandespendents.Show 1
End Sub

Private Sub Command9_Click()

  If capcalera.Recordset.EditMode > 0 Or liniescompra.Recordset.EditMode > 0 Then MsgBox "No pots imprimir si estas editant la comanda.", vbCritical, "Error": Exit Sub
  generar_pdf_comanda capcalera.Recordset!numcomanda
End Sub

Private Sub consultar_Click()
   Dim numc As Double
   If capcalera.Recordset.EditMode > 0 Or liniescompra.Recordset.EditMode > 0 Then MsgBox "No pots imprimir si estas editant la comanda.", vbCritical, "Error": Exit Sub
   numc = cadbl(InputBox("Entra el Nº de comanda que vols buscar.", "Buscar comanda"))
   If numc > 0 Then
     capcalera.RecordSource = "select * from capcalera order by numcomanda desc"
     capcalera.Refresh
     capcalera.Recordset.FindFirst "numcomanda=" + atrim(numc)
     If capcalera.Recordset.NoMatch Then MsgBox "No he trobat aquesta comanda", vbInformation, "Atenció"
   End If
End Sub

Private Sub diamext_LostFocus()
If IsNumeric(idiamext) Then diamext = cadbl(diamext)
End Sub

Private Sub eliminar_Click()
If UCase(InputBox("Segur que vols eliminar aquesta Comanda de compra?" + Chr(10) + Chr(13) + "escriu ELIMINAR per acceptar l'eliminació.", "Atencio")) = "ELIMINAR" Then
    dbtmpb.Execute "delete * from comandesxlinia where idliniacompra IN (SELECT idliniacompra from liniescompra where idcompra=" + atrim(capcalera.Recordset!id) + ")"
    dbtmpb.Execute "delete * from liniesdescripcio where idliniacompra IN (SELECT idliniacompra from liniescompra where idcompra=" + atrim(capcalera.Recordset!id) + ")"
    dbtmpb.Execute "delete * from liniescompra where idcompra=" + atrim(capcalera.Recordset!id)
    dbtmpb.Execute "delete * from capcalera where id=" + atrim(capcalera.Recordset!id)
    capcalera.Refresh
    
  End If
End Sub

Private Sub eliminarlinia_Click()
  If UCase(InputBox("Segur que vols eliminar aquesta linia de compra?" + Chr(10) + Chr(13) + "escriu ELIMINAR per acceptar l'eliminació.", "Atencio")) = "ELIMINAR" Then
    dbtmpb.Execute "delete * from comandesxlinia where idliniacompra=" + atrim(liniescompra.Recordset!idliniacompra)
    dbtmpb.Execute "delete * from liniesdescripcio where idliniacompra=" + atrim(liniescompra.Recordset!idliniacompra)
    liniescompra.Recordset.Delete
    liniescompra.Refresh
  End If
  
End Sub

Private Sub Form_Activate()
  If comandescompra.tag = "primera" Then
     comprovarcomandesambdiferencialderebutidemanatmassagran
     comprovarsihihaprecomandes
     comandescompra.tag = ""
  End If
End Sub
Sub comprovarcomandesambdiferencialderebutidemanatmassagran()
  Dim rstt As Recordset
  Set rstt = dbtmpb.OpenRecordset("SELECT capcalera.numcomanda,liniescompra.idliniacompra,liniescompra.kgentregats,liniescompra.quantitatkg, liniescompra.[kgentregats]-[quantitatkg] AS mitatdiferencial fROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE not diferencialcorrecte and (((liniescompra.[kgentregats]-[quantitatkg])>([quantitatkg]/4)));")
  While Not rstt.EOF
     If MsgBox("Comanda:" + atrim(rstt!numcomanda) + "  Kg demanats: " + atrim(rstt!quantitatkg) + " Kg entregats: " + atrim(rstt!kgentregats) + Chr(10) + " Fes NO per fer canvis i SI per confirmar que es correcte.", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
       'marcar la linia com correcte
        dbtmpb.Execute "update liniescompra set diferencialcorrecte=true where idliniacompra=" + atrim(rstt!idliniacompra)
          Else: capcalera.Recordset.FindFirst "numcomanda=" + atrim(cadbl(rstt!numcomanda)): GoTo fi
     End If
     rstt.MoveNext
  Wend
fi:
  Set rstt = Nothing
End Sub

Private Sub Form_Click()
Exit Sub
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim fitxerpdftemporal As String
  Dim email As String
  Dim cosmissatge As String
  Dim venviantpreuzero As Boolean
  If vimprimint Then Exit Sub
  If capcalera.Recordset.EditMode > 0 Or liniescompra.Recordset.EditMode > 0 Then MsgBox "No pots imprimir si estas editant la comanda.", vbCritical, "Error": Exit Sub
  vimprimint = True
reenviarapreuzero:
  borrarpedidostemporalsanteriors
  If capcalera.Recordset.EditMode > 0 Then capcalera.Recordset.Update
  fitxerpdftemporal = "c:\temp\Pedido_" + atrim(capcalera.Recordset!numcomanda) + ".pdf"
  passarregistrealataulatemporal cnumcomanda
  If venviantpreuzero Then possar_preus_Azero_de_la_comanda
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "comandescompres_prova.rpt", 1)
  oreport.Database.SetDataSource (dbconsulta)
  oreport.DiscardSavedData
  oreport.ExportOptions.DiskFileName = fitxerpdftemporal
  oreport.ExportOptions.PDFExportAllPages = True
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  
  oreport.EnableParameterPrompting = False
  
  oreport.Database.Tables.Item(1).Location = fitxertemp
  
  'oreport.ExportOptions.DestinationType = crEDTEMailMAPI
  If venviantpreuzero Then
    ratoli "espera"
    wait 5
    ratoli "normal"
  End If
  
  oreport.Export False
  Set oapp = Nothing
  Set oreport = Nothing
If existeix(fitxerpdftemporal) Then obrir_document fitxerpdftemporal
End Sub
Sub comprovasitotentregat(vnumcomanda As Double)
  Dim rstc As Recordset
  Set rstc = dbtmpb.OpenRecordset("SELECT capcalera.numcomanda,capcalera.materialrebut, capcalera.data, capcalera.dataentrega, capcalera.nomprov, liniescompra.* FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra where not totentregat and numcomanda=" + atrim(cadbl(vnumcomanda)) + ";")
  If rstc.EOF Then
     dbtmpb.Execute "update capcalera set materialrebut=true where numcomanda=" + atrim(cadbl(vnumcomanda))
  End If
  
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 112 Then Command1_Click
  If KeyCode = 27 Then
    cancelarcanvis
    fcapcalera.Enabled = False
  End If
  codiKeyCode = KeyCode
End Sub
Sub cancelarcanvis()
   If capcalera.Recordset.EditMode > 0 Then capcalera.Recordset.CancelUpdate
   If liniescompra.Recordset.EOF Then
      capcalera.Recordset.Delete
      capcalera.Refresh
   End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
  If Chr(KeyAscii) = "." And codiKeyCode = 110 Then KeyAscii = Asc(",")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set dbconsulta = Nothing
  Unload comandespendents
  Set dbtmp = Nothing
  End
End Sub

Private Sub iample_LostFocus()
  If IsNumeric(iample) Then iample = cadbl(iample)
End Sub

Private Sub icares_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub iespesor_LostFocus()
If IsNumeric(iespesor) Then iespesor = cadbl(iespesor)
End Sub

Private Sub iobert_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub iobert_LostFocus()
  If itl.Text = "L" Then iobert = "N"
End Sub

Private Sub iplegat_LostFocus()
If IsNumeric(iplegat) Then iplegat = cadbl(iplegat)
End Sub

Private Sub isolapa_LostFocus()
If IsNumeric(isolap) Then isolapa = cadbl(isolapa)
End Sub

Private Sub itl_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub itl_LostFocus()
  If itl.Text = "L" Then iobert = "N"
End Sub

Private Sub kilosxrcomprar_GotFocus()
    kilosxrcomprar.SelStart = 0
    kilosxrcomprar.SelLength = Len(kilosxrcomprar)
End Sub

Private Sub liniescompra_Reposition()
  Dim rstmat As Recordset
  comandesxlinia.RecordSource = "select * from comandesxlinia where idliniacompra=" + atrim(cadbl(liniescompra.Recordset!idliniacompra)) + " order by numcomanda desc"
  liniesdescripcio.RecordSource = "select * from liniesdescripcio where idliniacompra=" + atrim(cadbl(liniescompra.Recordset!idliniacompra)) + " order by ordre"
  liniesdescripcio.Refresh
  comandesxlinia.Refresh
  If Not comandesxlinia.Recordset.EOF Then actualitzar_valors_comanda
  possarmicresogrmm2
  If Not liniescompra.Recordset.EOF Then
     Set rstmat = dbtmp.OpenRecordset("select mesuarespcompra from materials where codi=" + atrim(cadbl(liniescompra.Recordset!codimaterial)))
     If Not rstmat.EOF Then
        combomaterial.tag = atrim(rstmat!mesuarespcompra)
     End If
     ensenyar_camps_compra
  End If
 
End Sub
Sub ensenyar_camps_compra()
   If liniescompra.Recordset!tipusmaterialcomprat <> "M" Then
     possarcampsvisiblesa False
       Else: possarcampsvisiblesa True
   End If
   
End Sub
Sub possarcampsvisiblesa(v As Boolean)
    itl.visible = v
    icares.visible = v
    iobert.visible = v
    iample.visible = v
    iplegat.visible = v
    isolapa.visible = v
    iespesor.visible = v
    imicrop.visible = v
    diamext.visible = v
    mandril.visible = v
    lblLabels(6).visible = v
    lblLabels(5).visible = v
    lblLabels(1).visible = v
    lblLabels(15).visible = v
    lblLabels(3).visible = v
    lblLabels(0).visible = v
    lblLabels(2).visible = v
    Label1(1).visible = v
    Label1(2).visible = v
    Label1(3).visible = v
    If v = False Then
         combomaterial.width = 5000
         preu.Left = 6500
         kilosxrcomprar.Left = 6500 + preu.width + 100
         kilosxrcomprar.Locked = False
         lblLabels(7).Left = preu.Left
         lblLabels(4).Left = kilosxrcomprar.Left
       Else
          combomaterial.width = 3000
           preu.Left = 8850
         kilosxrcomprar.Left = 9420
         kilosxrcomprar.Locked = True
         lblLabels(7).Left = preu.Left
         lblLabels(4).Left = kilosxrcomprar.Left
    End If
End Sub

Private Sub llcpmateriaprimera_Click()
   imprimir_llistatcompresproducte "M"
End Sub

Private Sub llctintes_Click()
   imprimir_llistatcompresproducte "T"
End Sub

Private Sub llcvaris_Click()
 
End Sub

Private Sub llistattoteslescompres_Click()
   imprimir_llistattotalkgsxrefcupu False
End Sub

Private Sub llmateriaprimera_Click()
    imprimir_llistattotalkgs "M"
End Sub

Private Sub llsitatambcupu_Click()
  imprimir_llistattotalkgsxrefcupu True
End Sub

Private Sub Lltintes_Click()
imprimir_llistattotalkgs "T"
End Sub

Private Sub llvaris_Click()
imprimir_llistattotalkgs "V"
End Sub

Private Sub m_entregades_Click()
 capcalera.RecordSource = "select * from capcalera where materialrebut order by numcomanda desc"
  capcalera.Refresh
  
End Sub

Sub imprimir_llistattotalkgsxrefcupu(vfercupu As Boolean)
  Dim rst As Recordset
  Dim vnomfitxer As String
  Dim vlinia As String
  Dim inici As Date
  Dim fi As Date
  Dim v As String
  Dim wdates As String
  Dim mespasat As Date
  Dim vambcupu As String
  Dim vsql As String
  
  mespasat = DateAdd("m", -1, Now)
  vnomfitxer = "c:\temp\llistatkgrefcupu.csv"
  If noespoteliminarnomfitxer(vnomfitxer) Then MsgBox "No es pot eliminar el fitxer temporal CSV del llistat mira que no el tinguis obert.", vbCritical, "Error": Exit Sub
   
  v = InputBox("Entra la data d'inici de la consulta.", "Inici consulta", atrim(DateSerial(Year(Now), Month(Now), 1)))
  If Not IsDate(v) Then MsgBox "La data no es correcte.": Exit Sub
  inici = CVDate(v)
  v = InputBox("Entra la data de fi de la consulta.", "Inici consulta", atrim(DateSerial(Year(Now), Month(DateAdd("m", 1, Now)), 0)))
  If Not IsDate(v) Then MsgBox "La data no es correcte.": Exit Sub
   fi = CVDate(v)
   wdates = " capcalera.data>=#" + Format(inici, "mm/dd/yy") + "# and capcalera.data<=#" + Format(fi, "mm/dd/yy") + "# "
   If vfercupu Then
        vsql = "SELECT capcalera.data, capcalera.numcomanda, capcalera.codiproveidorcomercial, capcalera.nomprovcomercial, liniescompra.nommaterial, liniescompra.semielaborat, liniescompra.Ample, liniescompra.micres, liniescompra.quantitatkg, liniesdescripcio.descripcio AS codicupu, familiesmaterials.descripcio AS descfam, subfamiliesmaterials.descripcio AS descsubfam, familiescolorants.descripcio AS desccol, subfamiliescolorants.descripcio AS descsubcol, familiesaditius.descripcio AS descaditiu, subfamiliesaditius.descripcio AS descsubaditiu"
        vsql = vsql + " FROM (capcalera RIGHT JOIN (liniescompra RIGHT JOIN liniesdescripcio ON liniescompra.idliniacompra = liniesdescripcio.idliniacompra) ON capcalera.id = liniescompra.idcompra) LEFT JOIN ((((((materials LEFT JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) LEFT JOIN familiescolorants ON materials.familiacol = familiescolorants.codi) LEFT JOIN familiesaditius ON materials.familiaad = familiesaditius.codi) LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi) LEFT JOIN subfamiliescolorants ON materials.subfamiliacol = subfamiliescolorants.codi) LEFT JOIN subfamiliesaditius ON materials.subfamiliaad = subfamiliesaditius.codi) ON liniescompra.codimaterial = materials.codi "
        vambcupu = "(((liniesdescripcio.descripcio) Like 'COD. *')) and"
         Else
           vsql = "SELECT capcalera.data, capcalera.numcomanda, capcalera.codiproveidorcomercial, capcalera.nomprovcomercial, liniescompra.nommaterial, liniescompra.semielaborat, liniescompra.Ample, liniescompra.micres, liniescompra.quantitatkg, familiesmaterials.descripcio AS descfam, subfamiliesmaterials.descripcio AS descsubfam, familiescolorants.descripcio AS desccol, subfamiliescolorants.descripcio AS descsubcol, familiesaditius.descripcio AS descaditiu, subfamiliesaditius.descripcio AS descsubaditiu "
           vsql = vsql + " FROM (capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra) LEFT JOIN ((((((materials LEFT JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) LEFT JOIN familiescolorants ON materials.familiacol = familiescolorants.codi) LEFT JOIN familiesaditius ON materials.familiaad = familiesaditius.codi) LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi) LEFT JOIN subfamiliescolorants ON materials.subfamiliacol = subfamiliescolorants.codi) LEFT JOIN subfamiliesaditius ON materials.subfamiliaad = subfamiliesaditius.codi) ON liniescompra.codimaterial = materials.codi "
   End If
   'Set rst = capcalera.Database.OpenRecordset("SELECT capcalera.data, capcalera.numcomanda, capcalera.codiproveidorcomercial, capcalera.nomprovcomercial, liniescompra.nommaterial, liniescompra.semielaborat, liniescompra.Ample, liniescompra.micres, liniescompra.quantitatkg,liniesdescripcio.descripcio AS codicupu FROM capcalera RIGHT JOIN (liniescompra RIGHT JOIN liniesdescripcio ON liniescompra.idliniacompra = liniesdescripcio.idliniacompra) ON capcalera.id = liniescompra.idcompra WHERE " + vambcupu + wdates)
   Set rst = capcalera.Database.OpenRecordset(vsql + " where " + vambcupu + wdates)
   If rst.EOF Then MsgBox "No hi ha dades per exportar. Comprova la sel.lecció", vbCritical + vbOKOnly, "Atenció": GoTo fi
   Open vnomfitxer For Output As #1
   Print #1, "CODIREFCUPU;DATA;COMANDA;CODIPROVEIDOR;NOMPROVEIDOR;NOMMATERIAL;T/L;AMPLE;ESPESOR;KG;FAMILIA;SUBFAMILIA;FAMILIA COL;SUBFAMILIA COL;FAMILIA ADITIU;SUBFAMILIA ADITIU"
   While Not rst.EOF
      If vfercupu Then
          vlinia = IIf(vfercupu, atrim(rst!codicupu), "") + ";" + atrim(Format(rst!data, "dd/mm/yy")) + ";" + atrim(rst!numcomanda) + ";" + atrim(rst!codiproveidorcomercial) + ";" + atrim(rst!nomprovcomercial) + ";" + atrim(rst!nommaterial) + ";" + atrim(rst!semielaborat) + ";" + atrim(rst!ample) + ";" + atrim(rst!micres) + ";" + atrim(rst!quantitatkg) + ";" + atrim(rst!descfam) + ";" + atrim(rst!descsubfam) + ";" + atrim(rst!desccol) + ";" + atrim(rst!descsubcol) + ";" + atrim(rst!descaditiu) + ";" + atrim(rst!descsubaditiu)
           Else: vlinia = ";" + atrim(Format(rst!data, "dd/mm/yy")) + ";" + atrim(rst!numcomanda) + ";" + atrim(rst!codiproveidorcomercial) + ";" + atrim(rst!nomprovcomercial) + ";" + atrim(rst!nommaterial) + ";" + atrim(rst!semielaborat) + ";" + atrim(rst!ample) + ";" + atrim(rst!micres) + ";" + atrim(rst!quantitatkg) + ";" + atrim(rst!descfam) + ";" + atrim(rst!descsubfam) + ";" + atrim(rst!desccol) + ";" + atrim(rst!descsubcol) + ";" + atrim(rst!descaditiu) + ";" + atrim(rst!descsubaditiu)
      End If
      Print #1, vlinia
      rst.MoveNext
   Wend
   Close #1
fi:
   Set rst = Nothing
   If existeix(vnomfitxer) Then obrir_document vnomfitxer
End Sub
Function noespoteliminarnomfitxer(vnomfitxer As String) As Boolean
  On Error GoTo fi
  If existeix(vnomfitxer) Then Kill vnomfitxer
  Exit Function
fi:
  noespoteliminarnomfitxer = True
End Function

Sub imprimir_llistattotalkgs(vtipus As String)
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim inici As Date
  Dim fi As Date
  Dim v As String
  Dim vtitol As String
  Dim mespasat As Date
  mespasat = DateAdd("m", -1, Now)
  v = InputBox("Entra la data d'inici de la consulta.", "Inici consulta", atrim(DateSerial(Year(mespasat), Month(mespasat), 1)))
  If Not IsDate(v) Then MsgBox "La data no es correcte.": Exit Sub
  inici = CVDate(v)
  v = InputBox("Entra la data de fi de la consulta.", "Inici consulta", atrim(DateSerial(Year(mespasat), Month(Now), 0)))
  If Not IsDate(v) Then MsgBox "La data no es correcte.": Exit Sub
  fi = CVDate(v)
  If vtipus = "M" Then vtitol = "Llistat de Kg de film comprats per proveidor entre " + atrim(inici) + " i " + atrim(fi)
  If vtipus = "T" Then vtitol = "Llistat de Kg de Tinta comprats per proveidor entre " + atrim(inici) + " i " + atrim(fi)
  If vtipus = "V" Then vtitol = "Llistat de unitats comprades per proveidor entre " + atrim(inici) + " i " + atrim(fi)
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "comprestotalskgproveidor.rpt", 1)
  oreport.Database.Tables.Item(1).Location = capcalera.DatabaseName
  
  oreport.RecordSelectionFormula = "{capcalera.data} in DateTime (" + atrim(Year(inici)) + "," + atrim(Month(inici)) + "," + atrim(Day(inici)) + ", 0,0, 0) to DateTime (" + atrim(Year(fi)) + "," + atrim(Month(fi)) + "," + atrim(Day(fi)) + ", 23,59, 0)" + " AND {liniescompra.tipusmaterialcomprat}='" + vtipus + "' "
  oreport.FormulaFields.GetItemByName("titol").Text = "'" + vtitol + "'"
  
  oreport.EnableParameterPrompting = False
 ' If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
 '   Else
       
 '     oreport.PrintOut False, 1
 ' End If
  
  comandescompra.SetFocus
End Sub
Sub triar_material_llistat(vcodi As String, vdesc As String)
  Dim rstmat As Recordset
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  'formseleccio.Data1.DatabaseName = cami
  Set formseleccio.Data1.Recordset = dbtmp.OpenRecordset("select * from materials where codi>499 order by descripcio")
  formseleccio.width = 7000
  'formseleccio.Data1.RecordSource = "select * from proveidors"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   vcodi = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   vdesc = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
End Sub
Sub triar_tintes_llistat(vcodiproducte As String, vdescripcioproducte As String)
  Dim rstmat As Recordset
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
  'formseleccio.Data1.RecordSource = "SELECT tintes.codi, tintes.descripcio, tintes.referenciacolor, tintesreferencies.referencia, tipusbidons.nombido,tipusbidons.litrescompres,tintesreferencies.id FROM (tintesreferencies INNER JOIN tintes ON tintesreferencies.idtinta = tintes.idtinta) INNER JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id where tintesreferencies.codiproveidor=" + atrim(capcalera.Recordset!codiproveidor)
  formseleccio.Data1.RecordSource = "SELECT tintes.codi, tintes.descripcio, tintes.referenciacolor, tintesreferencies.referencia, tipusbidons.nombido,tipusbidons.litrescompres,tintesreferencies.id FROM (tintesreferencies INNER JOIN tintes ON tintesreferencies.idtinta = tintes.idtinta) INNER JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id "
'  Clipboard.SetText formseleccio.Data1.RecordSource
  formseleccio.refrescar
  formseleccio.width = 9200
  formseleccio.DBGrid2.Columns(0).width = 800
  formseleccio.DBGrid2.Columns(1).width = 3000
  formseleccio.DBGrid2.Columns(2).width = 2000
  formseleccio.DBGrid2.Columns(3).width = 1200
  formseleccio.DBGrid2.Columns(6).width = 0
  formseleccio.Show 1
  If seleccioret = 1 Then
    vcodiproducte = atrim(cadbl(formseleccio.Data1.Recordset!codi))
    vdescripcioproducte = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
End Sub
Sub imprimir_llistatcompresproducte(vtipus As String)
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim inici As Date
  Dim fi As Date
  Dim v As String
  Dim vtitol As String
  Dim mespasat As Date
  Dim vdescripcioproducte As String
  Dim vcodiproducte As String
  mespasat = DateAdd("m", -1, Now)
  
  If vtipus <> "T" Then
     triar_material_llistat vcodiproducte, vdescripcioproducte
       Else: triar_tintes_llistat vcodiproducte, vdescripcioproducte
  End If
  
  If cadbl(vcodiproducte) = 0 Then Exit Sub
  'If vtipus = "M" Then
  vtitol = "Llistat de compres de " + vcodiproducte + " - " + vdescripcioproducte
  'If vtipus = "T" Then vtitol = "Llistat de Kg de Tinta comprats per proveidor entre " + atrim(inici) + " i " + atrim(fi)
  'If vtipus = "V" Then vtitol = "Llistat de unitats comprades per proveidor entre " + atrim(inici) + " i " + atrim(fi)
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "compresperarticle.rpt", 1)
  oreport.Database.Tables.Item(1).Location = capcalera.DatabaseName
  
  oreport.RecordSelectionFormula = "{liniescompra.codimaterial}=" + vcodiproducte + "  AND {liniescompra.tipusmaterialcomprat}" + IIf(vtipus <> "T", "<>", "=") + "'T' "
  oreport.FormulaFields.GetItemByName("titol").Text = "'" + treure_apostruf(vtitol) + "'"
  
  oreport.EnableParameterPrompting = False
 ' If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
 '   Else
       
 '     oreport.PrintOut False, 1
 ' End If
  
  comandescompra.SetFocus
End Sub
Private Sub mandril_LostFocus()
If IsNumeric(mandril) Then mandril = cadbl(mandril)
End Sub

Private Sub mcomandesclientconcret_Click()
  Dim vproveidor As Double
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  'formseleccio.Data1.DatabaseName = cami
  Set formseleccio.Data1.Recordset = dbtmp.OpenRecordset("select * from proveidors")
  'formseleccio.Data1.RecordSource = "select * from proveidors"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   vproveidor = atrim(cadbl(formseleccio.Data1.Recordset!codi))
  End If
  Unload formseleccio
  If cadbl(vproveidor) > 0 Then
        capcalera.RecordSource = "select * from capcalera where codiproveidor=" + atrim(vproveidor) + " order by numcomanda desc"
        capcalera.Refresh
  End If
End Sub

Private Sub menullistatcompresdetot_Click()
   Dim rst As Recordset
   Dim rstmat As Recordset
   Dim rstkg As Recordset
   Dim vnomfitxer As String
   Dim vlinia As String
   Dim inici As Date
  Dim fi As Date
  Dim v As String
  Dim wdates As String
  Dim vfam As String
  Dim mespasat As Date
  Dim vkgxcomandes As Double
  Dim vkgxestoc As Double
  
  mespasat = DateAdd("m", -1, Now)
   vnomfitxer = "c:\temp\llistatkgcompratsentredates.csv"
   If noespoteliminarnomfitxer(vnomfitxer) Then MsgBox "No es pot eliminar el fitxer temporal CSV del llistat mira que no el tinguis obert.", vbCritical, "Error": Exit Sub
   
    v = InputBox("Entra la data d'inici de la consulta.", "Inici consulta", atrim(DateSerial(Year(Now), Month(Now), 1)))
  If Not IsDate(v) Then MsgBox "La data no es correcte.": Exit Sub
  inici = CVDate(v)
  v = InputBox("Entra la data de fi de la consulta.", "Inici consulta", atrim(DateSerial(Year(Now), Month(DateAdd("m", 1, Now)), 0)))
  If Not IsDate(v) Then MsgBox "La data no es correcte.": Exit Sub
  fi = CVDate(v)
   wdates = " capcalera.data>=#" + Format(inici, "mm/dd/yy") + "# and capcalera.data<=#" + Format(fi, "mm/dd/yy") + " 23:59# "
   Set rst = capcalera.Database.OpenRecordset("SELECT capcalera.data, capcalera.numcomanda, capcalera.codiproveidorcomercial, capcalera.nomprovcomercial, liniescompra.nommaterial, liniescompra.idliniacompra as idliniac,liniescompra.semielaborat, liniescompra.kgentregats,liniescompra.totentregat,liniescompra.Ample, liniescompra.micres, liniescompra.quantitatkg, liniescompra.codimaterial,liniescompra.tipusmaterialcomprat,liniescompra.codimaterial FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra Where " + wdates)
   If rst.EOF Then MsgBox "No hi ha dades per exportar. Comprova la sel.lecció", vbCritical + vbOKOnly, "Atenció": GoTo fi
   Open vnomfitxer For Output As #1
   Print #1, "TIPUSMAT;DATA;COMANDA;CODIPROVEIDOR;NOMPROVEIDOR;CODIMATERIAL;NOMMATERIAL;T/L;AMPLE;ESPESOR;KG;KGENTREGATS;TOTENTREGAT;KGxCOMANDES;KGxESTOC;FAMILIA;SUBFAMILIA;FAMILIA_COLOR;SUBFAMILIA_COLOR;FAMILIA_ADITIU;SUBFAMILIA_ADITIU"
   While Not rst.EOF
     vkgxestoc = 0
     vkgxcomandes = 0
     'sumo les compres de estoc
     Set rstkg = capcalera.Database.OpenRecordset("select sum(kgcompra) as sumakgcompra from comandesxlinia where idliniacompra=" + atrim(rst!idliniac) + " and comandavisual='ESTOC' GROUP BY idliniacompra")
     vkgxestoc = cadbl(rstkg!sumakgcompra)
     'sumo les compres per comandes
     Set rstkg = capcalera.Database.OpenRecordset("select sum(kgcompra) as sumakgcompra from comandesxlinia where idliniacompra=" + atrim(rst!idliniac) + " and numcomanda>0 GROUP BY idliniacompra")
     vkgxcomandes = cadbl(rstkg!sumakgcompra)
     
     If rst!tipusmaterialcomprat <> "T" Then
         Set rstmat = dbtmp.OpenRecordset("select * from [llistat materials] where codi=" + atrim(rst!codimaterial))
         vfam = ";" + atrim(rstmat.Fields(8)) + ";" + atrim(rstmat.Fields(9)) + ";" + atrim(rstmat.Fields(10)) + ";" + atrim(rstmat.Fields(11)) + atrim(rstmat.Fields(12)) + ";" + atrim(rstmat.Fields(13))
        Else
          Set rstmat = dbtintes.OpenRecordset("select * from [tintes_tot] where codi='" + atrim(rst!codimaterial) + "'")
          vfam = ";" + atrim(rstmat.Fields(18)) + ";" + atrim(rstmat.Fields(19)) + ";" + atrim(rstmat.Fields(20)) + ";" + atrim(rstmat.Fields(21)) + ";;"
     End If
     vlinia = atrim(rst!tipusmaterialcomprat) + ";" + atrim(Format(rst!data, "dd/mm/yy")) + ";" + atrim(rst!numcomanda) + ";" + atrim(rst!codiproveidorcomercial) + ";" + atrim(rst!nomprovcomercial) + ";" + atrim(rst!codimaterial) + ";" + atrim(rst!nommaterial) + ";" + atrim(rst!semielaborat) + ";" + atrim(rst!ample) + ";" + atrim(rst!micres) + ";" + atrim(rst!quantitatkg) + ";" + atrim(rst!kgentregats) + ";" + atrim(cabool(rst!totentregat)) + ";" + atrim(vkgxcomandes) + ";" + atrim(vkgxestoc) + vfam
     Print #1, vlinia
     rst.MoveNext
   Wend
   Close #1
fi:
   Set rst = Nothing
   If existeix(vnomfitxer) Then obrir_document vnomfitxer

End Sub

Private Sub mllistatpendentderebre_Click()
Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim inici As Date
  Dim fi As Date
  Dim v As String
  Dim vtitol As String
  Dim mespasat As Date
  Dim vdescripcioproducte As String
  Dim vcodiproducte As String
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatcomprespendents.rpt", 1)
  oreport.Database.Tables.Item(1).Location = capcalera.DatabaseName
  
'  oreport.FormulaFields.GetItemByName("titol").Text = "'" + vtitol + "'"
  
  oreport.EnableParameterPrompting = False
 ' If existeix("c:\ordprog.ini") Then
  If Not vllistatcomprespendents Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
    Else
       oreport.ExportOptions.DestinationType = crEDTDiskFile
       oreport.ExportOptions.FormatType = crEFTPortableDocFormat
       oreport.ExportOptions.DiskFileName = "c:\temp\Llistat_comprespendents.pdf"
       oreport.Export False
  End If
 
 '   Else
       
 '     oreport.PrintOut False, 1
 ' End If
  
  If comandescompra.visible Then comandescompra.SetFocus
End Sub

Private Sub modificar_Click()
  If capcalera.Recordset.EditMode = 0 Then
     capcalera.Recordset.Edit
     fcapcalera.Enabled = True
     Text2.SetFocus
       Else
         MsgBox "Ja estàs editant la capçalera.", vbCritical, "Error"
  End If
  
End Sub
Private Sub Command1_Click()
  Dim bk As Variant
  If capcalera.Recordset.EditMode > 0 Then
     bk = capcalera.Recordset.Bookmark
     capcalera.Recordset.Update: capcalera.Recordset.Bookmark = bk
     calcular_totals_comanda capcalera.Recordset!numcomanda
     capcalera.Recordset.Bookmark = bk
  End If
  oklinia
  fcapcalera.Enabled = False
End Sub
Sub calcular_totals_comanda(numc As Double)
   Dim rstc As Recordset
   Dim rstl As Recordset
   Dim iva As Double
   iva = cadbl(llegir_ini("General", "iva", "comandes.ini"))
   If iva = 0 Then
      escriure_ini "General", "iva", 21, "comandes.ini"
      iva = 21
      MsgBox "No he trobat el valor de l'iva configurat i he possat el 21% si vols canviar-lo ves al menu de comandes a Entrades de dades-Taules Auxiliars-Canvi de tipus d'iva"
   End If
   Set rstc = dbtmpb.OpenRecordset("select * from capcalera where numcomanda=" + atrim(numc))
   If Not rstc.EOF Then
      rstc.Edit
      rstc![%iva] = iva
      Set rstl = dbtmpb.OpenRecordset("select sum(quantitatkg*preu) as total from liniescompra where idcompra=" + atrim(rstc!id))
      If Not rstl.EOF Then
        rstc!baseimp = rstl!total
        rstc!iva = (rstc!baseimp * rstc![%iva]) / 100
        rstc!total = cadbl(rstc!baseimp) + cadbl(rstc!iva)
      End If
      If Not rstc!enviat And Not IsDate(rstc!precomandafins) Then
          rstc!precomandafins = Format(DateAdd("d", 1, rstc!data), "dd/mm/yy")
      End If
      rstc.Update
   End If
   capcalera.UpdateControls
End Sub
Private Sub DBGrid1_Click()

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
 ' MsgBox DatePart("ww", Now)
 Dim arguments As Variant
 arguments = ObtenerLíneaComando
  
 If App.PrevInstance Then MsgBox "El programa ja està obert.", vbCritical, "Atenció": End
 If llegir_ini("General", "parar", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini") = "si" Then MsgBox "Ara no es pot entrar al programa s'està actualitzant, espera 5 MINUTS, Gràcies", vbCritical, "Actualització": End
  assignardecimalipunt
  cami = llegir_ini("General", "cami", "comandes.ini")
  centerscreen Me
  capcalera.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
  capcalera.RecordSource = "select * from capcalera where not materialrebut order by numcomanda desc"
  liniescompra.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
  liniesdescripcio.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
  comandesxlinia.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
  liniescompra.RecordSource = ""
  liniesdescripcio.RecordSource = ""
  comandesxlinia.RecordSource = ""
   Set dbtmp = OpenDatabase(cami)
   Set dbtmpb = OpenDatabase(liniescompra.DatabaseName)
   Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb", , True)
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb", , True)
  crearfitxertemp
  
  capcalera.Refresh
  dbstocks.Execute "delete * from pendentsdereservar"
  If arguments(1) = "llistatcomprespendents" Then enviar_llistat_comprespendents: End
End Sub
Sub enviar_llistat_comprespendents()
   vllistatcomprespendents = True
  If existeix("c:\temp\Llistat_comprespendents.pdf") Then Kill "c:\temp\Llistat_comprespendents.pdf"
  mllistatpendentderebre_Click
End Sub
Sub comprovarsihihaprecomandes()
    Dim rstp As Recordset
    
    Set rstp = dbtmpb.OpenRecordset("select * from capcalera where not materialrebut and not enviat and precomandafins<=now")
    If Not rstp.EOF Then
       MsgBox "Hi ha comandes que ja haurien d'estar enviades a proveïdor." + Chr(10) + "LES CARREGO A PANTALLA PER REVISAR-LES", vbCritical + vbOKOnly, "PRECOMANDES"
       capcalera.RecordSource = "select * from capcalera where not materialrebut and not enviat and precomandafins<=now order by numcomanda desc"
       capcalera.Refresh
    End If
    Set rstp = Nothing
End Sub
Sub crearfitxertemp()
    fitxertemp = "c:\temp\comprestmp.mdb"
    If existeix(fitxertemp) And Not existeix("c:\ordprog.ini") Then Kill fitxertemp
    If Not existeix(fitxertemp) Then
       crearfitxertemporal
       vinculolestaules
    End If
    Set dbconsulta = OpenDatabase(fitxertemp)
End Sub
Sub vinculolestaules()
  Dim tdfproductos As TableDef
  Dim camicompres As String
  camicompres = comandescompra.capcalera.DatabaseName
  'comandes
   Set tdfproductos = dbconsulta.CreateTableDef("comandes")
   tdfproductos.Connect = ";DATABASE=" & cami & ";"
   tdfproductos.SourceTableName = "comandes"
   dbconsulta.TableDefs.Append tdfproductos
  'compres de comandes fetes
    Set tdfproductos = dbconsulta.CreateTableDef("comandesxlinia")
   tdfproductos.Connect = ";DATABASE=" & camicompres & ";"
   tdfproductos.SourceTableName = "comandesxlinia"
   dbconsulta.TableDefs.Append tdfproductos
   Set tdfproductos = Nothing
End Sub

Sub crearfitxertemporal()
    borrartemps
    'fitxertemp = "c:\temp\comprestmp.mdb"
    '"c:\temp\~compres" + Format(Now, "ddmmhhnnss") + ".mdb"
    If Not existeix(fitxertemp) Then
       DBEngine.CreateDatabase fitxertemp, dbLangGeneral
    End If
    Set dbconsulta = DBEngine.OpenDatabase(fitxertemp)
    creartaula
End Sub
Sub creartaula()
  Dim i As Integer
  Dim camps(100, 2) As String
  
  i = 1
  camps(i, 1) = "comanda": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "mtrscomanda": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "aor": camps(i, 2) = "string(1)": i = i + 1
  camps(i, 1) = "mtrsassignats": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "mtrspendents": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "kgpendents": camps(i, 2) = "double": i = i + 1
  'camps(i, 1) = "mtrsdisponibles": camps(i, 2) = "double": i = i + 1
  'camps(i, 1) = "kgdisponibles": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "kgcomprats": camps(i, 2) = "double": i = i + 1
  'camps(i, 1) = "kgcompralliures": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "pesx1000": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "mtrscomprats": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "perlinkar": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "compatible": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "material": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "nommaterial": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "ample": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "plegat": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "solapa": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "semielaborat": camps(i, 2) = "string(2)": i = i + 1
  camps(i, 1) = "obert": camps(i, 2) = "string(2)": i = i + 1
  camps(i, 1) = "tractat": camps(i, 2) = "string(2)": i = i + 1
  camps(i, 1) = "microperforat": camps(i, 2) = "string(2)": i = i + 1
  camps(i, 1) = "familiamat": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "subfamiliamat": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "familiacol": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "subfamiliacol": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "familiaad": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "subfamiliaad": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "espesor": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "mesuraesp": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "client": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "texteimp": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "cont": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "seleccionat": camps(i, 2) = "string(1)": i = i + 1
  
  dbconsulta.Execute ("create table comprescomandes (id counter)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbconsulta.Execute ("alter table comprescomandes add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
       
        Else: i = 1000
    End If
    
  Next i
  
  'creo la segona taula
  
  i = 1
  camps(i, 1) = "migelaborat": camps(i, 2) = "string(1)": i = i + 1
  camps(i, 1) = "material": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "descripcio": camps(i, 2) = "string(50)": i = i + 1
  camps(i, 1) = "micres": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "metresdisponibles": camps(i, 2) = "double": i = i + 1
  
  dbconsulta.Execute ("create table familiescomprescomandes (id counter)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbconsulta.Execute ("alter table familiescomprescomandes add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
       
        Else: i = 1000
    End If
  Next i
  
End Sub
    
Sub borrartemps()
   On Error Resume Next
   Kill "c:\temp\~compres*.*"
   Kill "c:\temp\compres*.mdb"
End Sub


Sub triar_proveidor()
  If capcalera.Recordset.EditMode = 0 Then Exit Sub
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  'formseleccio.Data1.DatabaseName = cami
  Set formseleccio.Data1.Recordset = dbtmp.OpenRecordset("select * from proveidors where databaixa=null")
  'formseleccio.Data1.RecordSource = "select * from proveidors"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   capcalera.Recordset!codiproveidor = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   capcalera.Recordset!nomprov = atrim(formseleccio.Data1.Recordset!nom)
   proveidor = atrim(formseleccio.Data1.Recordset!nom)
   capcalera.Recordset!codiproveidorcomercial = 0
   Unload formseleccio
   capcalera.Recordset!codiproveidorcomercial = triar_proveidor_comercial(capcalera.Recordset!codiproveidor)
  End If
  Unload formseleccio
End Sub
Function triar_proveidor_comercial(codipro As Long) As Long
  If capcalera.Recordset.EditMode = 0 Then Exit Function
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  'formseleccio.Data1.DatabaseName = cami
  Set formseleccio.Data1.Recordset = dbtmp.OpenRecordset("select * from proveidors_comercial where codiproduccio=" + atrim(codipro))
  'formseleccio.Data1.RecordSource = "select * from proveidors"
  formseleccio.refrescar
  If Not formseleccio.Data1.Recordset.EOF Then
     formseleccio.Data1.Recordset.MoveLast
     If formseleccio.Data1.Recordset.RecordCount = 1 Then seleccioret = 1: GoTo fi
  End If
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).visible = False
  formseleccio.DBGrid2.Columns(2).width = 3500
  formseleccio.Show 1
fi:
  If seleccioret = 1 Then
   triar_proveidor_comercial = atrim(cadbl(formseleccio.Data1.Recordset!codi))
  End If
  Unload formseleccio
End Function


Private Sub mpendentsdenviar_Click()
capcalera.RecordSource = "select * from capcalera where not materialrebut and not enviat and precomandafins<=now"
  capcalera.Refresh

End Sub

Private Sub mpendentsentrega_Click()
capcalera.RecordSource = "select * from capcalera where not materialrebut order by numcomanda desc"
  capcalera.Refresh
End Sub

Private Sub novalinia_Click()
  formselecciotipuscompra.Show 1
  If formselecciotipuscompra.tag <> "" Then novaliniacompra
  
  
End Sub
Sub novaliniacompra()
  If capcalera.Recordset.EditMode > 0 Then capcalera.Recordset.Update
  ' capcalera.Recordset.Bookmark = capcalera.Recordset.LastModified
  If liniescompra.Recordset.EditMode > 0 Then MsgBox "S'està editant la linia.", vbCritical, "Atenció"
  liniescompra.Recordset.AddNew
  fdescmat.Enabled = True
  liniescompra.Recordset!idcompra = capcalera.Recordset!id
  liniescompra.Recordset!tipusmaterialcomprat = formselecciotipuscompra.tag
  If liniescompra.Recordset!tipusmaterialcomprat <> "T" Then
     If liniescompra.Recordset!tipusmaterialcomprat = "M" Then
        diamext = cadbl("80")
        mandril = cadbl("15.2")
     End If
     triar_material
  End If
  If liniescompra.Recordset!tipusmaterialcomprat = "T" Then triar_tintes
  ensenyar_camps_compra
  If cadbl(liniescompra.Recordset!codimaterial) = 0 Then
     liniescompra.Recordset.CancelUpdate
     actualitzar_valors_comanda
     fdescmat.Enabled = False
       Else:
          possar_preu_del_producte
          If itl.visible Then
             itl.SetFocus
               Else: preu.SetFocus
          End If
  End If
End Sub
Sub possar_preu_del_producte()
   Dim rst As Recordset
   If liniescompra.Recordset!tipusmaterialcomprat <> "M" Then
       Set rst = dbtmpb.OpenRecordset("select top 1 preu from albaransbip where article='" + atrim(cadbl(liniescompra.Recordset!codimaterial)) + "' order by data DESC", , ReadOnly)
       If Not rst.EOF Then
           preu = atrim(cadbl(rst!preu))
       End If
   End If
   Set rst = Nothing
End Sub
Private Sub preu_GotFocus()
   preu.SelStart = 0
   preu.SelLength = Len(preu)
End Sub

Private Sub preu_LostFocus()
  If IsNumeric(preu) Then preu = cadbl(preu)
End Sub

Private Sub proveidor_DropDown()
 If capcalera.Recordset.EditMode > 0 Then
   If liniescompra.Recordset.RecordCount > 0 Then
       MsgBox "No pots canviar de proveidor si ja hi ha linies de compra." + vbNewLine + " Primer hauries de borrar les linies.", vbCritical, "Atenció"
        Else
          triar_proveidor
          refrescar_proveidor
   End If
  SendKeys "{tab}"
 End If
End Sub
Sub refrescar_proveidor()
  Dim rstprov As Recordset
 
 
  Set rstprov = dbtmp.OpenRecordset("select * from proveidors_comercial where codi=" + atrim(cadbl(capcalera.Recordset!codiproveidorcomercial)))
  If Not rstprov.EOF Then
     Label6 = atrim(rstprov!nom)
     Label7 = atrim(rstprov!direccio)
     Label8 = atrim(rstprov!codipostal) + "-" + atrim(rstprov!poblacio)
     Label9 = atrim(rstprov!provinciapais)
     capcalera.Recordset!descripcioformadepago = atrim(rstprov!descripciopagament)
     capcalera.Recordset!formadepago = atrim(rstprov!formadepagament)
     capcalera.Recordset!nif = atrim(rstprov!nif)
     capcalera.Recordset!tel = atrim(rstprov!tel)
     capcalera.Recordset!fax = atrim(rstprov!fax)
       Else:
         Label6 = "": Label7 = "": Label8 = "": Label9 = ""
         capcalera.Recordset!descripcioformadepago = ""
         capcalera.Recordset!formadepago = ""
         capcalera.Recordset!nif = ""
         capcalera.Recordset!tel = ""
         capcalera.Recordset!fax = ""
  End If
  Set rstprov = Nothing
 
  
  
End Sub
Private Sub proveidor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triar_proveidor
End Sub

Private Sub reixacomandes_DblClick()
  Dim desc As String
  Dim kgestoc As Double
  
  kgestoc = kilosdestoc()
  If LCase(reixacomandes.Columns(reixacomandes.col).DataField) = "descripcio" Then
     desc = UCase(InputBox("Entra la descripció... ", "Descripció"))
     If comandesxlinia.Recordset.EditMode > 0 Then comandesxlinia.Recordset.CancelUpdate
     comandesxlinia.Recordset.Edit
     comandesxlinia.Recordset!descripcio = desc
     comandesxlinia.Recordset.Update
     reixacomandes.col = 0
  End If
  If LCase(reixacomandes.Columns(reixacomandes.col).DataField) = "kgcompra" And reixacomandes.Columns("comandavisual") = "ESTOC" Then
     desc = InputBox("Entra els kilos que vols d'ESTOC", "CANVI DE KG D'ESTOC")
     If cadbl(desc) > 0 Then reixacomandes.Columns("kgcompra") = cadbl(desc)
     comandesxlinia.Refresh
     sumar_kilos
  End If
  If (reixacomandes.Columns(reixacomandes.col).DataField) = "kgcompra" And cadbl(reixacomandes.Columns("comandavisual")) > 100000 Then
     kgcomprats = cadbl(reixacomandes.Text)
     desc = InputBox("Entra els kilos que vols per aquesta comanda", "CANVI DE KGs ")
     If cadbl(desc) > 0 Then
      If (cadbl(desc) - kgcomprats) <= kgestoc Then
         reixacomandes.Columns("kgcompra") = cadbl(desc)
         comandesxlinia.Refresh
         kgestoc = kgestoc - (cadbl(desc) - kgcomprats)
         If kgestoc < 1 Then kgestoc = 1
         kgestoc = kilosdestoc(kgestoc)
           Else: MsgBox "No hi ha prou estoc per canviar.", vbCritical, "Atencio": kgestoc = 0
      End If
     End If
     comandesxlinia.Refresh
     sumar_kilos
  End If
End Sub

Private Sub reixacomandes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   Dim colum As Integer
   colum = reixacomandes.col
   If reixacomandes.col < 5 And LastCol = 5 Then reixacomandes.col = 0: reixacomandes.col = colum
End Sub

Private Sub reixalinies_DblClick()
   Dim kg As Double
   Dim bk As Long
   If liniescompra.Recordset.EOF Then Exit Sub
   bk = liniescompra.Recordset!idliniacompra
   If reixalinies.Columns(reixalinies.col).DataField = "kgentregats" Then
      kg = cadbl(InputBox("Entra els Kilos entregats.", "Modificació de Kg Entregats"))
      If kg > 0 Then
          liniescompra.Recordset.Edit
          liniescompra.Recordset!kgentregats = kg
          If MsgBox("Està tot enviat?", vbInformation + vbYesNo, "Enviat?") = vbYes Then
             liniescompra.Recordset!totenviat = True
            Else: liniescompra.Recordset!totenviat = False
          End If
          If MsgBox("Ja ha arrivat tot el material a fàbrica?", vbInformation + vbYesNo, "Entregat?") = vbYes Then
             liniescompra.Recordset!totentregat = True
            Else: liniescompra.Recordset!totentregat = False
          End If
          liniescompra.Recordset.Update
          liniescompra.Refresh
          liniescompra.Recordset.FindFirst "idliniacompra=" + atrim(bk)
      End If
   End If
End Sub

Private Sub sortir_Click()
  Set dbconsulta = Nothing
  Unload comandespendents
  Set dbtmp = Nothing
  End
End Sub

Private Sub Timer1_Timer()
   If capcalera.Recordset.EditMode = 1 Then estat = "Editant..."
   If capcalera.Recordset.EditMode = 2 Then estat = "Afegint..."
   If capcalera.Recordset.EditMode = 0 Then estat = ""
   If estat.tag = "1" Then
      estat = "": estat.tag = "0"
     Else: estat.tag = "1"
   End If
End Sub

Private Sub Timer2_Timer()
  Static vcomprovarenvioemailsok As Byte
  Static vemailspendents As Boolean
  vcomprovarenvioemailsok = vcomprovarenvioemailsok + 1
  comprovarsihihacomandaxrobrirdeplanificacio
  If vemailspendents Then
     Command7.BackColor = IIf(Command7.BackColor = QBColor(12), &H8000000F, QBColor(12))
  End If
  If vcomprovarenvioemailsok > 4 Then
    vcomprovarenvioemailsok = 0
    vemailspendents = Not comprovarenvioemailsok
    If Not vemailspendents Then Command7.BackColor = &H8000000F
  End If
End Sub
Function comprovarenvioemailsok() As Boolean
  Dim vnomcarpeta As String
  Dim v As String
  vnomcarpeta = "\\serverprodu\Dades\progcomandes\dades\spoolerenviament\"
  v = Dir(vnomcarpeta + nomordinador + "*.", vbDirectory)
  comprovarenvioemailsok = True
  While v <> ""
    If v <> "." And v <> ".." Then
      If InStr(1, v, "#Error#") > 0 Then comprovarenvioemailsok = False: GoTo fi
    End If
    v = Dir(, vbDirectory)
  Wend
fi:
End Function
Sub comprovarsihihacomandaxrobrirdeplanificacio()
  Dim comanda As Double
  comanda = cadbl(llegir_ini("Planificacio", "comandacompraxrobrir", "comandes.ini"))
  escriure_ini "Planificacio", "comandacompraxrobrir", "0", "comandes.ini"
  If comanda > 0 Then
     comandespendents.Hide
     capcalera.RecordSource = "select * from capcalera where numcomanda=" + atrim(comanda)
     capcalera.Refresh
  End If
End Sub
Function enviaremail2(sSendTo As String, sSubject As String, sText As String, adjunt As String) As Boolean
  Dim usuarim As String
  Dim contrasenyam As String
  Command7.tag = ""
   enviaremail2 = False
  '#remitent#
  '#destinatari#
  '#cosdelmisatge#
  '#asumpte#
  '#fitxeradjunt#
  '#usuarigmail#
  '#contrasenyagmail#
  'Kill "c:\temp\enviomailcompra.vbs"
   Load formenviomails
   formenviomails.destinatari = sSendTo
   formenviomails.asumpte = sSubject
   formenviomails.nomfitxeradjunt = adjunt
   formenviomails.cosdelmissatge = sText
   formenviomails.Show 1
   If Command7.tag <> "enviar" Then Exit Function
   usuarim = llegir_ini("Enviomails", "usuari", "comandes.ini")
   contrasenyam = llegir_ini("Enviomails", "contrasenya", "comandes.ini")
   If usuarim = "{[}]" Or contrasenyam = "{[}]" Then MsgBox "L'usuari o la contrasenya no estan entrades", vbCritical, "Error": Exit Function
   
'creo el fitxer de cos de missatge
   Open "c:\temp\cosmissatge.txt" For Output As #2
   Print #2, formenviomails.cosdelmissatge
   Close #2
   
   
    Set objMessage = CreateObject("CDO.Message")
    objMessage.Subject = formenviomails.asumpte
    objMessage.From = "miquel.inplacsa@gmail.com"
    objMessage.To = formenviomails.destinatari
    objMessage.TextBody = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\temp\cosmissatge.txt", 1).ReadAll
    objMessage.AddAttachment formenviomails.nomfitxeradjunt
    
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = usuarim
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = contrasenyam
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    objMessage.Configuration.Fields.Update
    
    
    '==End remote SMTP server configuration section==
    If cadbl(objMessage.Send) = 0 Then enviaremail2 = True
   
   
    Unload formenviomails
End Function
Function enviaremail(sSendTo As String, sSubject As String, sText As String, adjunt As String, Optional noensenyarinterficie As Boolean) As Boolean
  Dim usuarim As String
  Dim contrasenyam As String
  Dim destinatari As String
  Dim vnomcarpeta As String
  Dim vadjunt As String
  
  Command7.tag = ""
   enviaremail = False
   Load formenviomails
   formenviomails.destinatari = sSendTo
   formenviomails.asumpte = sSubject
   formenviomails.nomfitxeradjunt = adjunt
   formenviomails.cosdelmissatge = sText
   If Not noensenyarinterficie Then
     formenviomails.Show 1
     If Command7.tag <> "enviar" Then Exit Function
   End If
   usuarim = llegir_ini("Enviomails", "usuari", "comandes.ini")
   contrasenyam = llegir_ini("Enviomails", "contrasenya", "comandes.ini")
   If usuarim = "{[}]" Or contrasenyam = "{[}]" Then MsgBox "L'usuari o la contrasenya no estan entrades", vbCritical, "Error": Exit Function
   vadjunt = adjunt
   vnomcarpeta = "\\serverprodu\Dades\progcomandes\dades\spoolerenviament\" + nomordinador + "_" + Format(Now, "yymmdd_hhnnss")
   
   If Not existeix(vnomcarpeta) Then MkDir vnomcarpeta
   escriure_ini "Capcalera", "apuntperenviar", "No", vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "data", Now, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "nomordinador", nomordinador, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "usuari", usuarim, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "contrasenya", contrasenyam, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "destinatari", formenviomails.destinatari, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "remitent", usuarim, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "assumpte", treure_apostruf(formenviomails.asumpte), vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "adjunt", vnomcarpeta + "\" + substituirtot(vadjunt, rutadelfitxer(vadjunt), ""), vnomcarpeta + "\dadesmail.txt"
   Copiar_Fitxer adjunt, vnomcarpeta
   Open "c:\temp\cosmissatge.txt" For Output As #2
   Print #2, formenviomails.cosdelmissatge
   Close #2
   Copiar_Fitxer "c:\temp\cosmissatge.txt", vnomcarpeta
   Kill "c:\temp\cosmissatge.txt"
  
   escriure_ini "Capcalera", "apuntperenviar", "Si", vnomcarpeta + "\dadesmail.txt"
   enviaremail = True
   Unload formenviomails
End Function



Function enviaremail_old(sSendTo As String, sSubject As String, sText As String, adjunt As String, Optional noensenyarinterficie As Boolean) As Boolean
  Dim usuarim As String
  Dim contrasenyam As String
  Command7.tag = ""
   enviaremail_old = False
  '#remitent#
  '#destinatari#
  '#cosdelmisatge#
  '#asumpte#
  '#fitxeradjunt#
  '#usuarigmail#
  '#contrasenyagmail#
  'Kill "c:\temp\enviomailcompra.vbs"
   Load formenviomails
   formenviomails.destinatari = sSendTo
   formenviomails.asumpte = sSubject
   formenviomails.nomfitxeradjunt = adjunt
   formenviomails.cosdelmissatge = sText
   If Not noensenyarinterficie Then
     formenviomails.Show 1
     If Command7.tag <> "enviar" Then Exit Function
   End If
   usuarim = llegir_ini("Enviomails", "usuari", "comandes.ini")
   contrasenyam = llegir_ini("Enviomails", "contrasenya", "comandes.ini")
   If usuarim = "{[}]" Or contrasenyam = "{[}]" Then MsgBox "L'usuari o la contrasenya no estan entrades", vbCritical, "Error": Exit Function
   Open llegir_ini("General", "rutallistats", "comandes.ini") + "enviomail.vbs" For Input As #1
   linia.Text = Input(LOF(1), #1)
   Close #1
   
   substituir "#remitent#", usuarim
   substituir "#destinatari#", formenviomails.destinatari
   substituir "#asumpte#", treure_apostruf(formenviomails.asumpte)
   'substituir "#cosdelmisatge#", treure_apostruf(formenviomails.cosdelmissatge)
   vfitxercos = rutadelfitxer(cami) + "spoolerenviament\" + Environ("computername") + ".txt"
   substituir "#cosdelmisatge#", "CreateObject(""Scripting.FileSystemObject"").OpenTextFile(""" + vfitxercos + """, 1).ReadAll"
   substituir "#fitxeradjunt#", rutadelfitxer(cami) + "spoolerenviament\" + substituir_per(adjunt, "c:\temp\", "")
   substituir "#usuarigmail#", usuarim
   substituir "#contrasenyagmail#", contrasenyam
   
   Open "c:\temp\cosmissatge.txt" For Output As #2
   Print #2, formenviomails.cosdelmissatge
   Close #2
   
   Open "c:\temp\enviomailcompra.vbs" For Output As #2
   Print #2, linia.Text
   Close #2
   
  ' executarelvbs ("c:\temp\enviomailcompra.vbs")
  ' wait 2
   If Not enviar_desde_el_servidor("c:\temp\enviomailcompra.vbs", adjunt) Then Exit Function
  
   enviaremail_old = True
   Unload formenviomails
End Function
Sub eliminar_fitxers_spooler(vruta As String)
    On Error Resume Next
    Kill vruta + "\*.*"
End Sub
Function enviar_desde_el_servidor(vfitxer As String, vadjunt As String) As Boolean
   Dim vruta As String
   Dim vdir As String
   Dim vinici As Date
   Dim vnomusuari As String
   Dim vfitxerdesti As String
   vnomusuari = substituir_per(llegir_ini("Enviomails", "usuari", "comandes.ini"), "@", "_")
   vfitxerdesti = rutadelfitxer(cami) + "spoolerenviament\envioemail_" + Environ("computername") + ".vbs"
   
   vruta = rutadelfitxer(cami) + "spoolerenviament"
   If Not existeix(vruta) Then MkDir rutadelfitxer(cami) + "spoolerenviament"
'   If existeix(vfitxerdesti) Then Kill vfitxerdesti
   eliminar_fitxers_spooler vruta
   If existeix(rutadelfitxer(cami) + "spoolerenviament\" + substituir_per(vadjunt, "c:\temp\", "")) Then Kill rutadelfitxer(cami) + "spoolerenviament\" + substituir_per(vadjunt, "c:\temp\", "")
   If existeix(rutadelfitxer(cami) + "spoolerenviament\" + Environ("computername") + ".txt") Then Kill rutadelfitxer(cami) + "spoolerenviament\" + Environ("computername") + ".txt"
   Copiar_Fitxer "c:\temp\cosmissatge.txt", rutadelfitxer(cami) + "spoolerenviament\" + Environ("computername") + ".txt"
   Copiar_Fitxer vfitxer, vfitxerdesti
   Copiar_Fitxer vadjunt, rutadelfitxer(cami) + "spoolerenviament\" + substituir_per(vadjunt, "c:\temp\", "")
   vinici = Now
   While DateDiff("s", vinici, Now) < 20
      If Not existeix(vfitxerdesti) Then
         If existeix(rutadelfitxer(cami) + "spoolerenviament\" + substituir_per(vadjunt, "c:\temp\", "")) Then Kill rutadelfitxer(cami) + "spoolerenviament\" + substituir_per(vadjunt, "c:\temp\", "")
         GoTo enviocorrecte
      End If
   Wend
   Exit Function
enviocorrecte:
   enviar_desde_el_servidor = True

End Function
Sub executarelvbs(lin As String)

 Dim objShell

        Set objShell = CreateObject("shell.application")

        objShell.ShellExecute lin, "", "", "open", 0

        Set objShell = Nothing
End Sub
Sub substituir(buscar As String, canviar As String)
   comença = InStr(1, linia, buscar) - 1
   If comença < 1 Then Exit Sub
   acaba = comença + Len(buscar) + 1
   linia = Mid(linia, 1, comença) + canviar + Mid(linia, acaba)
   'MsgBox linia
End Sub
Function enviaremail_gestor(sSendTo As String, sSubject As String, sText As String, adjunt As String) As Boolean
   
    On Error GoTo ErrHandler
     If MiMAPISession.SessionID <> 0 Then MiMAPISession.SignOff
     
     MiMAPISession.SignOn
    With MiMAPISession
        If .SessionID = 0 Then
            .DownLoadMail = False
            .LogonUI = True
            .SignOn
            .NewSession = True
            MAPIMessages1.SessionID = .SessionID
        End If
    End With
    MiMAPIMessages.SessionID = MiMAPISession.SessionID
    With MiMAPIMessages
        .Compose
        .RecipAddress = sSendTo
        .AddressResolveUI = True
        
        .ResolveName
        .MsgSubject = sSubject
        .MsgNoteText = sText
        
        
    MiMAPIMessages.AttachmentPathName = adjunt
        
        .Send True
    End With
    MailSend = True
    Exit Function
ErrHandler:
    'MsgBox err.Description
    MsgBox "Error enviant el mail.", vbCritical, "Atenció"
    MailSend = False
End Function
Sub generar_reserva_corresponent()
   Dim rstr As Recordset
   Dim rstxc As Recordset
   Dim rstmc As Recordset
   Dim rstc As Recordset
   Dim metrescomanda As Double
   metrescomanda = 0
   Set rstc = liniescompra.Recordset
   If Not existeixreserva(rstc, rstr) Then
       novareserva rstc, rstr
   End If
   Set rstr = Nothing
   Set rstc = Nothing
End Sub




Sub novareserva(rstc As Recordset, rstr As Recordset)
  rstr.AddNew
  rstr!ample = cadbl(rstc!ample)
  rstr!plegat = cadbl(rstc!plegat)
  rstr!carestractat = atrim(rstc!carestractat)
  rstr!obert = atrim(rstc!obert)
  rstr!microperforat = rstc!microperforat
  rstr!semielaborat = rstc!semielaborat
  rstr!espesor = IIf(cadbl(rstc!grmm2) > 0, cadbl(rstc!grmm2) * -1, cadbl(rstc!micres))
  rstr!familia = rstc!familia
  rstr!subfamilia = rstc!subfamilia
  rstr!familiacol = rstc!familiacol
  rstr!subfamiliacol = rstc!subfamiliacol
  rstr!familiaad = rstc!familiaad
  rstr!subfamiliaad = rstc!subfamiliaad
  rstr.Update
End Sub
Function existeixreserva(rstc As Recordset, rstr As Recordset) As Boolean
  Dim r As String
  Dim r2 As String
      r = "ample=" + passaradecimalpunt(rstc!ample) + " and plegat=" + passaradecimalpunt(rstc!plegat)
      r = r + " and solapa=" + passaradecimalpunt(rstc!solapa) + " and carestractat='" + atrim(rstc!carestractat + "'")
      r = r + " and obert='" + atrim(rstc!obert) + "' and microperforat=" + IIf(cabool(rstc!microperforat), "True", "False")
      r = r + " and semielaborat='" + atrim(rstc!semielaborat) + "' and espesor=" + passaradecimalpunt(IIf(rstc!grmm2 > 0, rstc!grmm2 * -1, rstc!micres))
      r2 = "and familia=" + atrim(cadbl(rstc!familia)) + " and subfamilia=" + atrim(cadbl(rstc!subfamilia))
      r2 = r2 + " and familiacol=" + atrim(cadbl(rstc!familiacol)) + " and subfamiliacol=" + atrim(cadbl(rstc!subfamiliacol))
      r2 = r2 + " and familiaad=" + atrim(cadbl(rstc!familiaad)) + " and subfamiliaad=" + atrim(cadbl(rstc!subfamiliaad))
      
      Set rstr = dbstocks.OpenRecordset("select * from reserves where " + r + r2)
      If Not rstr.EOF Then
           existeixreserva = True
         Else: existeixreserva = False
      End If
End Function

Function substituir_per(ByVal cadena As String, buscar As String, canviar As String) As String
   If buscar = canviar Then GoTo fi
   cadena = " " + cadena
   While InStr(1, cadena, buscar) > 0
    comença = InStr(1, cadena, buscar) - 1
    If comença < 1 Then substituir_per = cadena: Exit Function
    acaba = comença + Len(buscar) + 1
    cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   Wend
fi:
   substituir_per = atrim(cadena)
   'MsgBox linia
End Function

