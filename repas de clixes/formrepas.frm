VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form formrepas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Repàs de Clixes"
   ClientHeight    =   11130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
   Icon            =   "formrepas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11130
   ScaleWidth      =   15360
   Begin VB.CommandButton Command14 
      Caption         =   "Agrupar Treballs"
      Height          =   615
      Left            =   12150
      Picture         =   "formrepas.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   119
      ToolTipText     =   "Agrupar Treballs"
      Top             =   480
      Width           =   900
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   13965
      Picture         =   "formrepas.frx":237C
      Style           =   1  'Graphical
      TabIndex        =   104
      ToolTipText     =   "Eliminar totes les dades del repàs actiu."
      Top             =   480
      Width           =   900
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0080FF80&
      Caption         =   "Modifi."
      Height          =   615
      Left            =   13965
      Picture         =   "formrepas.frx":2906
      Style           =   1  'Graphical
      TabIndex        =   97
      ToolTipText     =   "Imprimeix la fulla de modificacions"
      Top             =   2370
      Width           =   900
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0080FF80&
      Caption         =   "Escull Revisador/a:"
      Height          =   300
      Left            =   7125
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   465
      Width           =   1980
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Sep.Colors"
      Height          =   615
      Left            =   13965
      Picture         =   "formrepas.frx":2E90
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   1110
      Width           =   900
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Complert"
      Height          =   615
      Left            =   13065
      Picture         =   "formrepas.frx":31D2
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   1110
      Width           =   900
   End
   Begin VB.CommandButton Command7 
      Caption         =   "crear tots els pdf de repas"
      Height          =   1140
      Left            =   8010
      TabIndex        =   93
      Top             =   4755
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Imp Etiq."
      Height          =   615
      Left            =   13065
      Picture         =   "formrepas.frx":3514
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "Imprimeix l'Etiqueta de la bossa del Clixé"
      Top             =   2370
      Width           =   900
   End
   Begin VB.CommandButton botoescullpersona 
      BackColor       =   &H00FF8080&
      Caption         =   "Escull Repasador/a"
      Height          =   300
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   0
      Width           =   1980
   End
   Begin VB.TextBox cpersona 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5430
      Locked          =   -1  'True
      TabIndex        =   89
      Top             =   30
      Width           =   1545
   End
   Begin VB.CommandButton imprimirrepas 
      BackColor       =   &H0080FF80&
      Caption         =   "Guardar PDFa l' arxiu"
      Height          =   615
      Left            =   13065
      Picture         =   "formrepas.frx":3A9E
      Style           =   1  'Graphical
      TabIndex        =   88
      ToolTipText     =   "Guarda el PDF i imprimeix la fulla de repàs de clixes."
      Top             =   3000
      Width           =   1785
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   13140
      Picture         =   "formrepas.frx":4428
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4845
      Width           =   1770
   End
   Begin VB.CheckBox repasarprova 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Repasar la prova de clixé sobre el pdf o cromalin  (Textes, codi de barres, nº de referència, etc)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   555
      TabIndex        =   85
      Top             =   3330
      Width           =   9915
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Veure L'IMP"
      Height          =   615
      Left            =   13065
      Picture         =   "formrepas.frx":49B2
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   1740
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Comanda"
      Height          =   615
      Left            =   13065
      Picture         =   "formrepas.frx":4EA4
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   480
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F3B378&
      Caption         =   "Començar a repasar"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   0
      Width           =   3090
   End
   Begin VB.CommandButton botonstinters 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEE4D7&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   13170
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3660
      Width           =   1845
   End
   Begin VB.CommandButton botonstinters 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEE4D7&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   11325
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3660
      Width           =   1845
   End
   Begin VB.CommandButton botonstinters 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEE4D7&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3660
      Width           =   1845
   End
   Begin VB.CommandButton botonstinters 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEE4D7&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   7635
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3660
      Width           =   1845
   End
   Begin VB.CommandButton botonstinters 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEE4D7&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   5790
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3660
      Width           =   1845
   End
   Begin VB.CommandButton botonstinters 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEE4D7&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3660
      Width           =   1845
   End
   Begin VB.CommandButton botonstinters 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEE4D7&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3660
      Width           =   1845
   End
   Begin VB.CommandButton botonstinters 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3B378&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   255
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3660
      Width           =   1845
   End
   Begin VB.Frame framedadestintes 
      BackColor       =   &H00EAD9CE&
      Height          =   7050
      Left            =   180
      TabIndex        =   59
      Top             =   4080
      Width           =   14820
      Begin VB.ComboBox Combotipusfoam 
         Height          =   315
         ItemData        =   "formrepas.frx":542E
         Left            =   7800
         List            =   "formrepas.frx":5430
         TabIndex        =   8
         Top             =   2745
         Width           =   960
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EEE4D7&
         Caption         =   "Traspas micropunts a la muntadora nova"
         Height          =   2640
         Left            =   10935
         TabIndex        =   107
         Top             =   1440
         Width           =   3675
         Begin VB.TextBox cnumbandes 
            Height          =   345
            Left            =   1395
            TabIndex        =   117
            Top             =   270
            Width           =   375
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00F1B75F&
            Height          =   795
            Left            =   2730
            Picture         =   "formrepas.frx":5432
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Enviar fitxer de parametres a la muntadora nova."
            Top             =   210
            Width           =   780
         End
         Begin VB.TextBox centremicropunts 
            Height          =   315
            Left            =   1860
            TabIndex        =   112
            Top             =   1740
            Width           =   810
         End
         Begin VB.TextBox cmicropuntmotiu 
            Height          =   315
            Left            =   1380
            TabIndex        =   110
            Top             =   1155
            Width           =   810
         End
         Begin VB.TextBox cmicropuntstotal 
            Height          =   315
            Left            =   1380
            TabIndex        =   108
            Top             =   660
            Width           =   810
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Bandes:"
            Height          =   270
            Left            =   195
            TabIndex        =   118
            Top             =   345
            Width           =   945
         End
         Begin VB.Label etenviatamuntadora 
            Alignment       =   2  'Center
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
            Height          =   270
            Left            =   525
            TabIndex        =   116
            Top             =   2280
            Width           =   2415
         End
         Begin VB.Label etdistanciaentremicropunts 
            BackStyle       =   0  'Transparent
            Caption         =   "Distancia entre motius:                       mm     (de micropunt a micropunt del motiu)"
            Height          =   450
            Left            =   165
            TabIndex        =   113
            Top             =   1815
            Width           =   3060
         End
         Begin VB.Label etmidadunmotiu 
            BackStyle       =   0  'Transparent
            Caption         =   "Mida d'un motiu:                       mm     (de micropunt a micropunt del motiu)"
            Height          =   450
            Left            =   165
            TabIndex        =   111
            Top             =   1245
            Width           =   2760
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Distancia total:                       mm       (micropunt esquerra al dret)"
            Height          =   450
            Left            =   150
            TabIndex        =   109
            Top             =   765
            Width           =   2535
         End
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Borrar i recarregar"
         Height          =   615
         Left            =   11970
         Picture         =   "formrepas.frx":568B
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   765
         Width           =   945
      End
      Begin VB.TextBox ampladatotalmotiu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2175
         TabIndex        =   13
         Top             =   4260
         Width           =   555
      End
      Begin VB.TextBox midadesarroll 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4350
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   5175
         Width           =   555
      End
      Begin VB.TextBox numerodemotius 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8940
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   5190
         Width           =   555
      End
      Begin VB.TextBox midacilindre 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7005
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   5145
         Width           =   555
      End
      Begin VB.CheckBox micropunts 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Repassar que hi hagi micropunts."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   390
         TabIndex        =   3
         Top             =   1755
         Width           =   3870
      End
      Begin VB.Frame frameposiciosang 
         BackColor       =   &H00FDDECE&
         Caption         =   "Posició de la Sang"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5925
         TabIndex        =   60
         Top             =   3435
         Width           =   1710
         Begin VB.ComboBox comboposiciosang 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "formrepas.frx":5777
            Left            =   60
            List            =   "formrepas.frx":5784
            TabIndex        =   12
            Top             =   255
            Width           =   1620
         End
      End
      Begin VB.CheckBox defectuos 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Comprovar que no està defectuos o marcat."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   420
         TabIndex        =   4
         Top             =   2235
         Width           =   5310
      End
      Begin VB.CheckBox lecturarelleu 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Es llegeix pel relleu."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3225
         TabIndex        =   9
         Top             =   3105
         Width           =   2715
      End
      Begin VB.CheckBox lecturallisa 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Es llegeix per la zona llisa."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3210
         TabIndex        =   10
         Top             =   3420
         Width           =   2715
      End
      Begin VB.TextBox ample 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3735
         TabIndex        =   11
         Top             =   3720
         Width           =   630
      End
      Begin VB.TextBox motiuamotiu 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5565
         TabIndex        =   26
         Top             =   4740
         Width           =   555
      End
      Begin VB.ComboBox combobandaseguiment 
         DataField       =   "bandaseguiment"
         DataSource      =   "modificacions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "formrepas.frx":57A9
         Left            =   4665
         List            =   "formrepas.frx":57B6
         TabIndex        =   30
         Top             =   5625
         Width           =   1515
      End
      Begin VB.TextBox amplebanda 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6315
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   5625
         Width           =   555
      End
      Begin VB.ComboBox macula 
         DataField       =   "macula"
         DataSource      =   "modificacions"
         Height          =   315
         ItemData        =   "formrepas.frx":57D5
         Left            =   2745
         List            =   "formrepas.frx":57E8
         TabIndex        =   32
         Top             =   6120
         Width           =   1515
      End
      Begin VB.CheckBox repasclixe 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Repassar el clixé amb el pdf  (Textes, codi de barres, nº de referència, etc)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   375
         TabIndex        =   2
         Top             =   1245
         Width           =   9915
      End
      Begin VB.CheckBox clixesllencats 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Verificar que dels clixes modificats es llencen els vells."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   405
         TabIndex        =   33
         Top             =   6645
         Width           =   6225
      End
      Begin VB.Frame framedescripciotinter 
         BackColor       =   &H00EAD9CE&
         Height          =   570
         Left            =   90
         TabIndex        =   73
         Top             =   120
         Width           =   14655
         Begin VB.Label descripciotinter 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   90
            TabIndex        =   74
            Top             =   150
            Width           =   14475
         End
      End
      Begin VB.OptionButton gruixplimer 
         BackColor       =   &H00EAD9CE&
         Caption         =   "1,14 mm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2910
         TabIndex        =   5
         Tag             =   "1,14"
         Top             =   2805
         Width           =   1020
      End
      Begin VB.OptionButton gruixplimer 
         BackColor       =   &H00EAD9CE&
         Caption         =   "2,54 mm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   4020
         TabIndex        =   6
         Tag             =   "2,54"
         Top             =   2805
         Width           =   1020
      End
      Begin VB.OptionButton gruixplimer 
         BackColor       =   &H00EAD9CE&
         Caption         =   "2,84 mm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   5145
         TabIndex        =   7
         Tag             =   "2,84"
         Top             =   2805
         Width           =   1020
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 1"
         Height          =   660
         Index           =   0
         Left            =   5760
         TabIndex        =   72
         Top             =   4080
         Visible         =   0   'False
         Width           =   690
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   90
            TabIndex        =   14
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 2"
         Height          =   660
         Index           =   1
         Left            =   6510
         TabIndex        =   71
         Top             =   4080
         Visible         =   0   'False
         Width           =   690
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   150
            TabIndex        =   15
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 3"
         Height          =   660
         Index           =   2
         Left            =   7245
         TabIndex        =   70
         Top             =   4080
         Visible         =   0   'False
         Width           =   690
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   150
            TabIndex        =   16
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 4"
         Height          =   660
         Index           =   3
         Left            =   7980
         TabIndex        =   69
         Top             =   4080
         Visible         =   0   'False
         Width           =   690
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   150
            TabIndex        =   17
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 5"
         Height          =   660
         Index           =   4
         Left            =   8700
         TabIndex        =   68
         Top             =   4080
         Visible         =   0   'False
         Width           =   690
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   150
            TabIndex        =   18
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 6"
         Height          =   660
         Index           =   5
         Left            =   9420
         TabIndex        =   67
         Top             =   4080
         Visible         =   0   'False
         Width           =   690
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   5
            Left            =   150
            TabIndex        =   19
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 7"
         Height          =   660
         Index           =   6
         Left            =   10140
         TabIndex        =   66
         Top             =   4080
         Visible         =   0   'False
         Width           =   690
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   6
            Left            =   150
            TabIndex        =   20
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 8"
         Height          =   660
         Index           =   7
         Left            =   10860
         TabIndex        =   65
         Top             =   4080
         Visible         =   0   'False
         Width           =   690
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   7
            Left            =   150
            TabIndex        =   21
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 9"
         Height          =   660
         Index           =   8
         Left            =   11580
         TabIndex        =   64
         Top             =   4080
         Visible         =   0   'False
         Width           =   690
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   8
            Left            =   150
            TabIndex        =   22
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 10"
         Height          =   660
         Index           =   9
         Left            =   12315
         TabIndex        =   63
         Top             =   4080
         Visible         =   0   'False
         Width           =   810
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   9
            Left            =   150
            TabIndex        =   23
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 11"
         Height          =   660
         Index           =   10
         Left            =   13155
         TabIndex        =   62
         Top             =   4080
         Visible         =   0   'False
         Width           =   765
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   10
            Left            =   150
            TabIndex        =   24
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.Frame framemotiu 
         BackColor       =   &H00FDDECE&
         Caption         =   "Motiu 12"
         Height          =   660
         Index           =   11
         Left            =   13965
         TabIndex        =   61
         Top             =   4080
         Visible         =   0   'False
         Width           =   765
         Begin VB.TextBox amplemotiu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   11
            Left            =   150
            TabIndex        =   25
            Top             =   210
            Width           =   510
         End
      End
      Begin VB.CheckBox creuvermella 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Possar creu vermella al clixe vell."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   390
         TabIndex        =   1
         Top             =   750
         Width           =   3495
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipus de Foam:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6345
         TabIndex        =   115
         Top             =   2760
         Width           =   1560
      End
      Begin VB.Label etTOTSELCCLIXES 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   720
         Left            =   3990
         TabIndex        =   106
         Top             =   720
         Width           =   7920
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2805
         TabIndex        =   103
         Top             =   4305
         Width           =   510
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total del clixe sobre la lamina"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   420
         TabIndex        =   102
         Top             =   4215
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4995
         TabIndex        =   101
         Top             =   5190
         Width           =   510
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Entra la mida de motiu a motiu per desarroll:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   100
         Top             =   5235
         Width           =   4110
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   12
         Left            =   45
         Top             =   5178
         Width           =   315
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Motius:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7665
         TabIndex        =   99
         Top             =   5220
         Width           =   1245
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Mida Cilindre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5655
         TabIndex        =   98
         Top             =   5205
         Width           =   1245
      End
      Begin VB.Label et_resposta_motiuamotiu 
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
         ForeColor       =   &H008080FF&
         Height          =   345
         Left            =   6630
         TabIndex        =   87
         Top             =   4725
         Width           =   900
      End
      Begin VB.Label et_resposta_ample 
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
         ForeColor       =   &H008080FF&
         Height          =   345
         Left            =   4905
         TabIndex        =   86
         Top             =   3750
         Width           =   900
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   11
         Left            =   45
         Top             =   6660
         Width           =   315
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   10
         Left            =   45
         Top             =   6162
         Width           =   315
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   9
         Left            =   45
         Top             =   5670
         Width           =   315
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   8
         Left            =   45
         Top             =   4686
         Width           =   315
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   7
         Left            =   45
         Top             =   4194
         Width           =   315
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   6
         Left            =   45
         Top             =   3702
         Width           =   315
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   5
         Left            =   45
         Top             =   3210
         Width           =   315
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   4
         Left            =   45
         Top             =   2718
         Width           =   315
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   3
         Left            =   45
         Top             =   2226
         Width           =   315
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   2
         Left            =   45
         Top             =   1734
         Width           =   315
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   1
         Left            =   45
         Top             =   1242
         Width           =   315
      End
      Begin VB.Image estat 
         Height          =   300
         Index           =   0
         Left            =   45
         Top             =   750
         Width           =   315
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Escull el gruix del polimer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   84
         Top             =   2760
         Width           =   2445
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Diga'm per on llegeixes el texte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   83
         Top             =   3255
         Width           =   2910
      End
      Begin VB.Label etampleisang 
         BackStyle       =   0  'Transparent
         Caption         =   "Entra l'ample + sang que correspon:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   82
         Top             =   3765
         Width           =   3360
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4455
         TabIndex        =   81
         Top             =   3750
         Width           =   510
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ample de motiu a motiu en mm: (de esquerra a dreta)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3315
         TabIndex        =   80
         Top             =   4215
         Width           =   2400
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Entra la mida de motiu a motiu per desarroll  (amb reducció):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   420
         TabIndex        =   79
         Top             =   4770
         Width           =   5235
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6210
         TabIndex        =   78
         Top             =   4755
         Width           =   510
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Escull on està la banda d'arrastre i entra el guix:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   77
         Top             =   5685
         Width           =   4320
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6975
         TabIndex        =   76
         Top             =   5640
         Width           =   510
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Escull on està la màcula:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   75
         Top             =   6180
         Width           =   2415
      End
      Begin VB.Image nocorrecte 
         Height          =   240
         Left            =   13620
         Picture         =   "formrepas.frx":581B
         Top             =   855
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image correcte 
         Height          =   240
         Left            =   13950
         Picture         =   "formrepas.frx":5DA5
         Top             =   840
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Dades del Treball"
      Enabled         =   0   'False
      Height          =   3360
      Left            =   210
      TabIndex        =   0
      Top             =   285
      Width           =   14805
      Begin Crystal.CrystalReport llistat 
         Left            =   9270
         Top             =   1860
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EEE4D7&
         Caption         =   "Informació del Treball i de la Comanda"
         Height          =   2040
         Left            =   285
         TabIndex        =   50
         Top             =   1005
         Width           =   8205
         Begin VB.Label lmaterialcomanda 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00FF8080&
            Height          =   450
            Left            =   105
            TabIndex        =   55
            Top             =   1545
            Width           =   7935
         End
         Begin VB.Label lmodificatonouidescproducte 
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
            Height          =   240
            Left            =   90
            TabIndex        =   54
            Top             =   1230
            Width           =   7935
         End
         Begin VB.Label lcodidebarres 
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
            Height          =   240
            Left            =   90
            TabIndex        =   53
            Top             =   915
            Width           =   7935
         End
         Begin VB.Label lmarcailinia 
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
            Height          =   240
            Left            =   90
            TabIndex        =   52
            Top             =   600
            Width           =   7935
         End
         Begin VB.Label lnomclient 
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
            Height          =   240
            Left            =   90
            TabIndex        =   51
            Top             =   285
            Width           =   7935
         End
      End
      Begin VB.TextBox ccomanda 
         BackColor       =   &H00EEE4D7&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   660
         Width           =   900
      End
      Begin VB.TextBox cversio 
         BackColor       =   &H00EEE4D7&
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   645
         Width           =   405
      End
      Begin VB.TextBox cidtreball 
         BackColor       =   &H00EEE4D7&
         Height          =   285
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   645
         Width           =   900
      End
      Begin VB.Label etnomrevisador 
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
         Height          =   255
         Left            =   9030
         TabIndex        =   96
         Top             =   225
         Width           =   3705
      End
      Begin VB.Image botototok 
         Height          =   1545
         Left            =   10785
         Picture         =   "formrepas.frx":632F
         Stretch         =   -1  'True
         Top             =   1710
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comanda:"
         Height          =   195
         Left            =   2025
         TabIndex        =   49
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lcomandespendents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comandes Pendents:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2145
         TabIndex        =   47
         Top             =   135
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Treball / Versió"
         Height          =   195
         Left            =   525
         TabIndex        =   45
         Top             =   390
         Width           =   1080
      End
   End
   Begin VB.Label operariaqueharevisat 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   7230
      TabIndex        =   91
      Top             =   30
      Width           =   3525
   End
End
Attribute VB_Name = "formrepas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vavisarREPASADORA  As Boolean
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Sub treurecolorbotons()
   Dim i As Byte
   For i = 0 To 7
     If botonstinters(i).HelpContextID = 0 Then botonstinters(i).BackColor = &HEEE4D7
     If botonstinters(i).HelpContextID = 1 Then botonstinters(i).BackColor = &H80FF80
     If botonstinters(i).HelpContextID = 2 Then botonstinters(i).BackColor = &H8080FF
   Next i
End Sub

Private Sub ampladatotalmotiu_Change()
   If Screen.ActiveControl.Name <> "ampladatotalmotiu" Then Exit Sub
   If escorrecteampleamotiu And escorrectetotaldelmotiu(cadbl(ampladatotalmotiu)) Then
     posar_ok True, 8
    Else: posar_ok False, 8
  End If
End Sub
Function escorrecteampleamotiu() As Boolean
  escorrecteampleamotiu = True
  If cadbl(ampladatotalmotiu) > (cadbl(ampladatotalmotiu.HelpContextID) - 4) Then
     escorrecteampleamotiu = False
  End If
End Function

Private Sub ampladatotalmotiu_GotFocus()
  ampladatotalmotiu.SelStart = 0
  ampladatotalmotiu.SelLength = Len(ample)
End Sub

Private Sub ample_Change()
   If Screen.ActiveControl.Name <> "ample" And Screen.ActiveControl.Name <> "comboposiciosang" Then Exit Sub
   If escorrecteampleisang(comboposiciosang, cadbl(ample)) Then
        posar_ok True, 7
         Else: posar_ok False, 7
   End If
End Sub
Function escorrecteampleisang(posiciosang As String, ample As Double) As Boolean
   Dim rstmodifi As Recordset
   Dim mirarsang As Boolean
   Dim amplesang As Double
   Dim amplecorrecte As Double
   mirarsang = frameposiciosang.visible
   comboposiciosang.BackColor = &HFFFFFF
   Set rstmodifi = dbclixes.OpenRecordset("select amplelamina,portasang,amplesang from modificacions where id_treball=" + atrim(cidtreball) + " and ordre=" + atrim(cversio))
   If rstmodifi.EOF Then Exit Function
   If mirarsang Then
      If atrim(rstmodifi!portasang) = "" Then MsgBox "Aquesta tinta esta marcada amb sang però el treball no diu quina sang." + Chr(10) + "S'HA DE CORREGIR EL TREBALL PER PODER REPASAR BÉ.", vbCritical, "ATENCIÓ"
      amplesang = cadbl(rstmodifi!amplesang)
   End If
   
   If mirarsang And atrim(posiciosang) = atrim(rstmodifi!portasang) Then
      escorrecteampleisang = True
       Else: comboposiciosang.BackColor = &HC0C0FF ': Exit Function
   End If
   'amplelamina pasat a milimetres multiplicat per motius horitzontals 'i tot sumar-hi la sang
   amplecorrecte = Redondejar(cadbl(rstmodifi!amplelamina) * 10 * cadbl(framemotiu(0).tag), 0) '+ amplesang
   If amplecorrecte = ample Then
     escorrecteampleisang = True
     et_resposta_ample = ""
    Else:
      escorrecteampleisang = False
      et_resposta_ample = atrim(amplecorrecte) + " mm"
   End If
   Set rstmodifi = Nothing
End Function

Private Sub ample_GotFocus()
  ample.SelStart = 0
  ample.SelLength = Len(ample)
End Sub

Private Sub amplebanda_Change()
  combobandaseguiment_Click
End Sub

Private Sub amplebanda_GotFocus()
  amplebanda.SelStart = 0
  amplebanda.SelLength = Len(amplebanda)
End Sub

Private Sub amplemotiu_Change(Index As Integer)
 If Screen.ActiveControl.Name = "amplemotiu" Then
    If escorrecteamplemotiu And escorrectetotaldelmotiu(cadbl(ampladatotalmotiu)) Then
       posar_ok True, 8
      Else: posar_ok False, 8
    End If
   End If
End Sub
Function escorrecteamplemotiu() As Boolean
    Dim rstmodifi As Recordset
    Dim ultim As Byte
    Dim inici As Byte
    Dim fi As Byte
    Dim i As Byte
    inici = 0
    ultim = cadbl(framemotiu(0).tag) - 1
    fi = ultim
    Set rstmodifi = dbclixes.OpenRecordset("select amplelamina,portasang,amplesang from modificacions where id_treball=" + atrim(cidtreball) + " and ordre=" + atrim(cversio))
    If rstmodifi.EOF Then Exit Function
   ' If atrim(rstmodifi!portasang) <> "" And comboposiciosang.Visible Then
   '    If InStr(1, atrim(rstmodifi!portasang), "Esquerra") Then
   '     inici = 1
   '     If cadbl(amplemotiu(0)) = (cadbl(rstmodifi!amplelamina) * 10) + cadbl(rstmodifi!amplesang) Then
   '          escorrecteamplemotiu = True
   '       Else: escorrecteamplemotiu = False: GoTo fi
   '     End If
   '    End If
   '    If InStr(1, atrim(rstmodifi!portasang), "Dreta") Then
   '     If cadbl(amplemotiu(ultim)) = (cadbl(rstmodifi!amplelamina) * 10) + cadbl(rstmodifi!amplesang) Then
   '        escorrecteamplemotiu = True
   '          Else: escorrecteamplemotiu = False: GoTo fi
   '     End If
   '     fi = ultim - 1
   '
   '    End If
    'End If
    For i = inici To fi
        If cadbl(amplemotiu(i)) = (cadbl(rstmodifi!amplelamina) * 10) Then
           escorrecteamplemotiu = True
          Else: escorrecteamplemotiu = False: GoTo fi
        End If
    Next i
fi:
    Set rstmodifi = Nothing
End Function

Private Sub amplemotiu_GotFocus(Index As Integer)
  amplemotiu(Index).SelStart = 0
  amplemotiu(Index).SelLength = Len(amplemotiu(Index))
End Sub

Private Sub botoescullpersona_Click()
  cpersona.tag = ""
   While cadbl(cpersona.tag) = 0
     escullirpersona
     If cadbl(cpersona.tag) = 0 Then MsgBox "Has d'escullir un operari per treballar"
   Wend
End Sub

Private Sub botonstinters_Click(Index As Integer)
    If repasarprova = 0 Then MsgBox "Primer has de repassar la prova de clixé.", vbInformation, "Atenció": Exit Sub
  '  If Not guardat And Screen.ActiveControl.Name = "botonstinters" Then If MsgBox("No has guardat el polimer anterior, vols guardar-lo ara?", vbInformation + vbYesNo, "Atenció") = vbYes Then Command5_Click
    treurecolorbotons
    If botonstinters(Index).tag = "" Then botonstinters(Index).HelpContextID = 0: GoTo cont
    botonstinters(Index).BackColor = &HF3B378
    descripciotinter.tag = botonstinters(Index).caption
    framedescripciotinter.tag = cadbl(botonstinters(Index).tag)
    carregardescripciotinta cadbl(cidtreball), cadbl(cversio), cadbl(botonstinters(Index).tag)
    descripciotinter.tag = botonstinters(Index).caption
    ampladatotalmotiu.BackColor = QBColor(15)
    'If botonstinters(Index).HelpContextID = 0 The
    posar_informacio_fitxermuntadora
    comprovacions_correctes
    guardat = False
cont:
End Sub
Sub posar_informacio_fitxermuntadora()
   If cadbl(cnumbandes) = 0 Then cnumbandes = atrim(cadbl(framemotiu(0).tag))
    etdistanciaentremicropunts.visible = False: centremicropunts.visible = False
    etmidadunmotiu.visible = False: cmicropuntmotiu.visible = False
    If cnumbandes > 1 Then etmidadunmotiu.visible = True: cmicropuntmotiu.visible = True
    If cnumbandes > 2 Then etdistanciaentremicropunts.visible = True: centremicropunts.visible = True
End Sub
Sub comprovacions_correctes()
   Dim i As Byte
   Dim totsok As Boolean
   Dim botonstotsok As Boolean
   Dim vnomfitxerpdf As String
   botonstotsok = True
   totsok = True
   For i = 0 To 11
      If estat(i).Picture = nocorrecte.Picture Then totsok = False
   Next i
   If cadbl(framedescripciotinter.tag) > 0 Then
    If totsok Then
       botonstinters(cadbl(framedescripciotinter.tag) - 1).HelpContextID = 1
        Else: botonstinters(cadbl(framedescripciotinter.tag) - 1).HelpContextID = 2
    End If
   End If
   For i = 0 To 7
      If botonstinters(i).HelpContextID <> 1 And botonstinters(i).caption <> "" Then botonstotsok = False
   Next i
   
   If botonstotsok Then
       botototok.visible = True
      Else: botototok.visible = False
   End If
   
End Sub
Sub mirarsihihaelPDFguardat()
  Dim fitxerpdftemporal As String
  fitxerpdftemporal = ruta_documentacio_clixes + "\" + Format(cadbl(cidtreball), "00000") + "\Arxiu_documentacio_relacionada" + "\v" + atrim(cadbl(cversio))
  fitxerpdftemporal = fitxerpdftemporal + "\Repasdeclixes.pdf"
  If Not existeix(fitxerpdftemporal) Then
      If MsgBox("AQUEST REPÀS ESTÀ TOT OK" + vbNewLine + "VOLS GUARDAR EL PDF A LA CARPETA DE DOCUMENTACIÓ DEL CLIXÉ?", vbExclamation + vbDefaultButton1 + vbYesNo, "ATENCIÓ") = vbYes Then
          imprimirrepas_Click
          wait 2
          If cadbl(cversio) > 1 Then
             MsgBox "Ara imprimiré també la fulla de modificacions d'impresió.", vbInformation, "Atenció"
             If etnomrevisador = "" Then
                MsgBox "No hi ha Revisador escullit, escull-ne un.", vbCritical, "Atenció"
                Command9_Click
                wait 2
             End If
             Command10_Click
          End If
      End If
  End If
End Sub
Sub ELIMINARelPDFguardat()
  Dim fitxerpdftemporal As String
  fitxerpdftemporal = ruta_documentacio_clixes + "\" + Format(cadbl(cidtreball), "00000") + "\Arxiu_documentacio_relacionada" + "\v" + atrim(cadbl(cversio))
  fitxerpdftemporal = fitxerpdftemporal + "\Repasdeclixes.pdf"
  If existeix(fitxerpdftemporal) Then
     Kill fitxerpdftemporal
  End If
End Sub

Sub desactivar_novinculantsenaquestatinta(rst As Recordset)
   
   creuvermella.Enabled = True
   clixesllencats.Enabled = True
   clixesllencats.Value = 1
   If cadbl(cversio) = 1 Then posar_ok True, 1: posar_ok True, 12: creuvermella.Enabled = False: clixesllencats.Enabled = False
   If rst.EOF Then Exit Sub
   macula.Enabled = True
   If Not rst!macula Then macula.Enabled = False: posar_ok True, 11
   combobandaseguiment.Enabled = True
   If Not rst!arrastre Then combobandaseguiment.Enabled = False: posar_ok True, 10
End Sub
Sub carregardescripciotinta(ntreball As Double, cversio As Double, tinter As Byte)
   Dim rst As Recordset
   Dim rstc As Recordset
   descripciotinter = ""
   
   Set rst = dbclixes.OpenRecordset("select * from repasdadestintes where  id_repas=0")
   If Not rst.EOF Then carregar_dades_tintes rst
   Set rst = dbclixes.OpenRecordset("select * from tintes where ordretinter=" + atrim(tinter) + " and id_treball=" + atrim(ntreball) + " and ordremodificacio=" + atrim(cversio))
   If rst.EOF Then Exit Sub
   descripciotinter = atrim(rst!ordretinter) + ": " + atrim(rst!Color) + " Anilox: " + atrim(rst!anilox) + " Desar: " + atrim(rst!desarroll) + IIf(rst!portasang, " Porta Sang", "")
   midacilindre.tag = ""
   numerodemotius.tag = ""
   If cadbl(rst!desarroll) > 0 Then
      motiuamotiu.tag = atrim(Int(cadbl(rst!cilindre) / cadbl(rst!desarroll)))
      midacilindre.tag = atrim(cadbl(rst!cilindre))
      numerodemotius.tag = atrim(cadbl(rst!desarroll))
        Else: MsgBox "Atenció el desarrol d'aquest tinter està a zero", vbCritical, "Atenció"
   End If
   motiuamotiu.HelpContextID = cadbl(rst!desarroll)
   If rst!portasang Then
       frameposiciosang.visible = True
       etampleisang.caption = "Entra l'ample + sang que correspon:"
        Else
           frameposiciosang.visible = False
           comboposiciosang = ""
           etampleisang.caption = "Entra l'ample que correspon: "
   End If
   
   Set rstc = dbclixes.OpenRecordset("select * from repasdadestintes where ordretinter=" + atrim(tinter) + " and id_repas=" + atrim(cadbl(framedadestintes.tag)))
   If Not rstc.EOF Then carregar_dades_tintes rstc
   desactivar_novinculantsenaquestatinta rst
   If rst!arrastre Or rst!portasang Then
      ampladatotalmotiu.Enabled = True
        Else: ampladatotalmotiu.Enabled = False
   End If
   Set rst = Nothing
   Set rstc = Nothing
End Sub

Private Sub clixesllencats_Click()
 If Screen.ActiveControl.Name = "clixesllencats" Then
    If clixesllencats.Value = 1 Then
       posar_ok True, 12
      Else: posar_ok False, 12
    End If
   End If
End Sub

Private Sub cnumbandes_LostFocus()
  posar_informacio_fitxermuntadora
End Sub

Private Sub combobandaseguiment_Click()
  If Screen.ActiveControl.Name <> "combobandaseguiment" And Screen.ActiveControl.Name <> "amplebanda" Then Exit Sub
   If escorrectebandaseguiment(combobandaseguiment, cadbl(amplebanda)) Then
        posar_ok True, 10
         Else: posar_ok False, 10
   End If
End Sub

Private Sub combobandaseguiment_GotFocus()
  combobandaseguiment.SelStart = 0
  combobandaseguiment.SelLength = Len(combobandaseguiment)
End Sub

Private Sub comboposiciosang_Click()
   ample_Change
   
End Sub

Private Sub comboposiciosang_GotFocus()
  comboposiciosang.SelStart = 0
  comboposiciosang.SelLength = Len(comboposiciosang)
End Sub

Private Sub Combotipusfoam_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 13 Then KeyAscii = 0
End Sub

Private Sub Combotipusfoam_LostFocus()
  'If Screen.ActiveControl.Name <> "Combotipusfoam" Then Exit Sub
  If escorrectegruixpolimerifoam(cadbl(gruixplimer(Index).tag)) Then
     posar_ok True, 5
    Else: posar_ok False, 5
  End If
End Sub

Private Sub Command1_Click()
    If cadbl(cpersona.tag) = 0 Then botoescullpersona_Click
    demanartreballperrepasar
End Sub
Sub demanartreballperrepasar(Optional ntreball As Double, Optional nversio As Double)
   Dim rstc As Recordset
   Dim rstclixe As Recordset
   Dim rstmodi As Recordset
   Dim ncomanda As Double
   If cadbl(ntreball) = 0 Or cadbl(nversio) = 0 Then
        ntreball = cadbl(InputBox("Entra el numero de treball:", "Treball"))
        If ntreball = 0 Then Exit Sub
        nversio = cadbl(InputBox("Entra la versió:", "Versió del treball"))
        If nversio = 0 Then Exit Sub
   End If
   netejarbotons
   centremicropunts = 0: cmicropuntmotiu = 0: cmicropuntstotal = 0: cnumbandes = 0
   botototok.visible = False
   If Not carregarrepasdeltreball(ntreball, nversio) Then
     ncomanda = triar_comandapendent(ntreball, nversio)
     If ncomanda = 0 Then Exit Sub
   End If
   If ncomanda = 0 Then ncomanda = cadbl(ccomanda)  'faig aqueta linia perque si carrega els valors guardats d'actualitzi la varible ncomanda
   cidtreball = atrim(ntreball)
   cversio = atrim(nversio)
   ccomanda = atrim(ncomanda)
   Set rstc = dbcomandes.OpenRecordset("SELECT comandes.amplereb, comandes.simulteneitatreb,comandes.client,comandes.impressio,comandes.comanda, clients.nom, productes.descripcio,comandes.ampleesq  FROM (comandes LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN productes ON comandes.producte = productes.codi Where comanda = " + atrim(ncomanda))
   Set rstclixe = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(ntreball))
   If rstc.EOF Or rstclixe.EOF Then Exit Sub
   Set rstmodi = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(ntreball) + " and ordre=" + atrim(nversio))
   If rstmodi.EOF Then MsgBox "Aquesta modificació no exiteix", vbCritical, "Error": Exit Sub
   ensenyarmotius cadbl(rstmodi!bandes)
   ampladatotalmotiu.HelpContextID = (cadbl(rstc!ampleesq) * 10) * cadbl(rstmodi!bandes)
   If cadbl(rstc!amplereb) > 0 Then
       ampladatotalmotiu.tag = (cadbl(rstc!amplereb) * cadbl(rstmodi!bandes)) + 23
         Else:
            ampladatotalmotiu.tag = atrim(cadbl(rstc!ampleesq) * 10) + 23
   End If
   framemotiu(0).tag = atrim(cadbl(rstmodi!bandes))
   lnomclient = "Nom del Client: " + UCase(rstc!nom)
   lnomclient.tag = atrim(rstc!client) + " - " + UCase(rstc!nom)
   lmarcailinia = "Marca i Linia: " + UCase(atrim(rstclixe!marca)) + " - " + UCase(atrim(rstclixe!linia))
   lmarcailinia.tag = UCase(atrim(rstclixe!marca)) + " - " + UCase(atrim(rstclixe!linia))
   lcodidebarres = "Codi de Barres: " + atrim(rstclixe!codidebarres) + " Arxiu XL: " + atrim(rstclixe!arxiu)
   If rstc!impressio = "N" Then lmodificatonouidescproducte = "Nova"
   If rstc!impressio = "M" Then lmodificatonouidescproducte = "Modificada"
   If rstc!impressio = "F" Then lmodificatonouidescproducte = "Falta Autoritzar"
   If rstc!impressio = "R" Then lmodificatonouidescproducte = "Repetida"
   If rstc!impressio <> "N" And rstc!impressio <> "M" Then MsgBox "El tipus de impressió d'aquesta comanda no es correcte, revisa que no hi hagi algun error.", vbCritical, "Atenció"
   lmodificatonouidescproducte = lmodificatonouidescproducte + " - " + atrim(rstc!descripcio)
   possarnommaterialdelacomanda ncomanda
   carregarbotonscolors ntreball, nversio
   comprovar_simaterialblanc ntreball, nversio
   If atrim(rstmodi!observacionsrepasclixes) <> "" Then
     MsgBox atrim(rstmodi!observacionsrepasclixes), vbInformation, "Observacions per repasar el clixé"
   End If
   
   Set rstmodi = Nothing
   Set rstc = Nothing
   Set rstclixe = Nothing
   carregarrepasdeltreball ntreball, nversio
   
   guardat = True
   formannex.carregarcomanda cadbl(ccomanda)
   If Not hihaafectatspelcanvi(ntreball, nversio) Then
            botototok.visible = True
            passar_treball_a_revisat
   End If
   wait 1
   activarprimercolor
End Sub

Sub activarprimercolor()
  Dim i As Byte
  i = 0
  While botonstinters(i).tag = "" And i + 1 < botonstinters.Count
     i = i + 1
  Wend
  If botonstinters(i).tag <> "" Then botonstinters(i).SetFocus: botonstinters_Click cadbl(i)
End Sub
Sub ensenyarmotius(bandes As Double)
   Dim i As Byte
    For i = 0 To framemotiu.Count - 1
      framemotiu(i).visible = False
    Next i
    If bandes = 0 Then bandes = 1
   For i = 0 To bandes - 1
       framemotiu(i).visible = True
   Next i
End Sub
Sub comprovar_simaterialblanc(ntreball As Double, nversio As Double)
  Dim i As Byte
  Dim hihablanc As Boolean
  Dim rst As Recordset
  Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(ntreball) + " and ordremodificacio=" + atrim(nversio))
  If rst.EOF Then Exit Sub
  While Not rst.EOF
     If InStr(1, UCase(rst!Color), "BLANC") > 0 Then hihablanc = True
     rst.MoveNext
  Wend
  'material transparent
  If InStr(1, UCase(lmaterialcomanda), "TRANSPA") > 0 Or InStr(1, UCase(lmaterialcomanda), "METALI") > 0 Then
     If InStr(1, UCase(lmaterialcomanda), "BLANC") = 0 Then
      If Not hihablanc Then MsgBox "Aquesta comanda està preparada amb material TRANSPARENT/METALITZAT i no hi ha el color BLANC revisa que sigui correcte.", vbCritical, "Atenció"
     End If
  End If
  
  'material blanc
  If InStr(1, UCase(lmaterialcomanda), "BLANC") > 0 Then
      If hihablanc Then MsgBox "Aquesta comanda té el material BLANC i has afegir color BLANC revisa que sigui correcte.", vbCritical, "Atenció"
  End If
  Set rst = Nothing
End Sub
Sub netejarbotons()
   
   For i = 0 To 7
      botonstinters(i).caption = ""
      botonstinters(i).tag = ""
      botonstinters(i).HelpContextID = 0
   Next i
   treurecolorbotons
End Sub
Function hihaafectatspelcanvi(vtreball As Double, vversio As Double, Optional rst As Recordset) As Boolean
   Set rst = dbclixes.OpenRecordset("select * from tintes where afectatspelcanvi and id_treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vversio))
   If Not rst.EOF Then hihaafectatspelcanvi = True
End Function
Sub carregarbotonscolors(ntreball As Double, nversio As Double)
   Dim rst As Recordset
   Dim rstanterior As Recordset
   Dim i As Byte
   netejarbotons
   If Not hihaafectatspelcanvi(ntreball, nversio, rst) Then
      If operariaqueharevisat = "" Then
       MsgBox "No hi ha tintes per aquets treball/modificacio." + Chr(10) + " AFECTATS PEL CANVI A CLIXES.", vbCritical, "Atenció"
       guardar_dades
      End If
      Exit Sub
   End If
   etTOTSELCCLIXES = ""
   If nversio > 1 Then
      Set rstanterior = dbclixes.OpenRecordset("select * from tintes where afectatspelcanvi and id_treball=" + atrim(ntreball) + " and ordremodificacio=" + atrim(nversio - 1))
      If Not rstanterior.EOF Then
           If mirarsitoteslestintes(rst, rstanterior) Then etTOTSELCCLIXES = "TOTS ELS CLIXES AFECTATS " + vbNewLine + " S'HAN D'ELIMINAR FISICAMENT ELS ANTICS."
      End If
   End If
   If Not rst.EOF Then rst.MoveFirst
   While Not rst.EOF
      botonstinters(rst!ordretinter - 1).caption = "(" + atrim(rst!ordretinter) + ") " + atrim(rst!Color) + IIf(atrim(rst!detalltinter) <> "", "(" + atrim(rst!detalltinter) + ")", "")
      botonstinters(rst!ordretinter - 1).tag = atrim(rst!ordretinter)
      rst.MoveNext
   Wend
   
End Sub
Function triar_comandapendent(ntreball, nversio) As Double
  Dim rstc As Recordset
  Dim vsql As String
  vsql = "SELECT comandes.comanda, clients.nom FROM comandes LEFT JOIN clients ON comandes.client = clients.codi where (producte<>'PC' and producte<>'PCP' and producte<>'PC2') and numordremodificacio=" + atrim(cadbl(nversio)) + " and numtreball=" + atrim(cadbl(ntreball)) + " and (proximaseccio <>'T') order by comanda Desc"
bucle:
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = vsql
   formseleccio.DBGrid2.AllowDelete = False
   

   formseleccio.DBGrid2.Columns(0).width = 2000
   formseleccio.DBGrid2.Columns(1).width = 6000
   formseleccio.width = 9000
   formseleccio.caption = "TRIA LA COMANDA ON VOLS TREBALLAR."
   formseleccio.refrescar
   formseleccio.Show 1
   lcomandespendents = "Comandes Pendents: "
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           triar_comandapendent = cadbl(formseleccio.DBGrid2.Columns("comanda"))
            formseleccio.Data1.Recordset.MoveFirst
            While Not formseleccio.Data1.Recordset.EOF
               lcomandespendents = lcomandespendents + atrim(formseleccio.Data1.Recordset!comanda) + "  "
               formseleccio.Data1.Recordset.MoveNext
            Wend
        End If
   End If
    If triar_comandapendent = 0 And InStr(1, vsql, "<>'T'") > 0 Then
      If MsgBox("No has escullit cap comanda s'aquest treball o no h'hi cap de pendent, vols escullir una d'anterior?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
         vsql = "SELECT comandes.comanda, clients.nom FROM comandes LEFT JOIN clients ON comandes.client = clients.codi where (producte<>'PC' and producte<>'PCP' and producte<>'PC2') and numordremodificacio=" + atrim(cadbl(nversio)) + " and numtreball=" + atrim(cadbl(ntreball)) + " order by comanda Desc"
         GoTo bucle
      End If
    End If
      
        
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
  
End Function

Private Sub Command10_Click()
  If etnomrevisador = "" Then MsgBox "Primer has d'escullir el revisador", vbCritical, "Atenció": Exit Sub
  If hihaafectatspelcanvi(cadbl(cidtreball), cadbl(cversio)) Then
        If mirar_si_faltaalgunadata Then MsgBox "Falta entrar la data de validació de colors.", vbCritical, "Error": Exit Sub
  End If
  If cadbl(cidtreball.tag) > 0 Then dbclixes.Execute "update  repasclixes set numoprevisat=" + atrim(cadbl(etnomrevisador.tag)) + ",nomoprevisat='" + atrim(etnomrevisador) + "' where id_repas=" + atrim(cadbl(cidtreball.tag))
  imprimir_fulla_modificacions
End Sub
Function mirarsitoteslestintes(rsttintes2 As Recordset, rsttintesanteriors As Recordset) As Boolean
   rsttintesanteriors.MoveFirst
   mirarsitoteslestintes = True
   While Not rsttintesanteriors.EOF
      rsttintes2.FindFirst "color='" + atrim(rsttintesanteriors!Color) + "'"
      If Not rsttintes2.NoMatch Then If rsttintes2!afectatspelcanvi = False Then mirarsitoteslestintes = False
      rsttintesanteriors.MoveNext
   Wend
End Function
Function possarvalorsalreport(oreport As CRAXDDRT.Report) As Boolean
   Dim rstcap As Recordset
   Dim rstclixe As Recordset
   Dim rstmodifi As Recordset
   Dim rsttintes As Recordset
   Dim rsttintes2 As Recordset
   Dim rsttintescontador As Recordset
   Dim rsttintesanteriors As Recordset
   Dim totselsclixes As Boolean
   Dim subconsulta As String
   Dim marcarnou As String
   Dim i As Byte
   
   Set rstcap = dbclixes.OpenRecordset("select * from repasclixes where id_treball=" + atrim(cadbl(cidtreball)) + " and nummodificacio=" + atrim(cadbl(cversio)))
   If rstcap.EOF Then Exit Function
   Set rstclixe = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(cidtreball)))
   Set rstmodifi = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(cadbl(cidtreball)) + " and ordre=" + atrim(cadbl(cversio)))
   Set rsttintes = dbclixes.OpenRecordset("select * from tintes where afectatspelcanvi=true and id_treball=" + atrim(cadbl(cidtreball)) + " and ordremodificacio=" + atrim(cadbl(cversio)))
   Set rsttintes2 = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(cadbl(cidtreball)) + " and ordremodificacio=" + atrim(cadbl(cversio)) + " and trim(color)<>''")
   subconsulta = "select color from tintes where id_treball=" + atrim(cadbl(cidtreball)) + " and ordremodificacio=" + atrim(IIf(cadbl(cversio) > 1, cadbl(cversio) - 1, cadbl(cversio))) + " and trim(color)<>''"
   Set rsttintesanteriors = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(cadbl(cidtreball)) + " and ordremodificacio=" + atrim(cadbl(cversio)) + " and trim(color)<>''")
   If rstmodifi.EOF Or rstclixe.EOF Or rsttintesanteriors.EOF Then GoTo fi
   If atrim(rstmodifi!descripcio) = "" Then MsgBox "La descripció de modificació d'aquest clixé està en blanc, abans hauries de possar-hi quelcom.", vbCritical, "Error": Exit Function
   rsttintes2.MoveLast
   rsttintesanteriors.MoveLast
   If Not rsttintes.EOF Then rsttintes.MoveLast
   If mirarsitoteslestintes(rsttintes2, rsttintesanteriors) Then totselsclixes = True
   oreport.FormulaFields.GetItemByName("comandes").Text = """" + comandesrelacionades + """"
   oreport.FormulaFields.GetItemByName("data").Text = """" + Format(rstcap!Data, "dd/mm/yy") + """"
   oreport.FormulaFields.GetItemByName("codiclient").Text = """" + atrim(rstclixe!codiclienttemporal) + """"
   oreport.FormulaFields.GetItemByName("nomclient").Text = """" + atrim(rstclixe!nomclienttemporal) + """"
   oreport.FormulaFields.GetItemByName("marcailinia").Text = """" + atrim(rstclixe!marca) + " - " + atrim(rstclixe!linia) + """"
   oreport.FormulaFields.GetItemByName("nommodificacio").Text = """" + treure_apostruf(atrim(rstmodifi!descripcio)) + IIf(totselsclixes, "| SE HAN HECHO NUEVOS TODOS LOS CLICHÉS.", "") + """"
   If totselsclixes Then
      oreport.FormulaFields.GetItemByName("rcolor1").Text = """ TODOS LOS CLICHÉS"""
      oreport.FormulaFields.GetItemByName("rmarcar1").Text = """X"""
      oreport.FormulaFields.GetItemByName("marcacomprovar").Text = """"""
     Else
        rsttintesanteriors.MoveFirst
        oreport.FormulaFields.GetItemByName("marcacomprovar").Text = """X"""
        While Not rsttintesanteriors.EOF
           rsttintes2.FindFirst "color='" + atrim(rsttintesanteriors!Color) + "' and detalltinter='" + atrim(rsttintesanteriors!detalltinter) + "'"
           If rsttintes2.NoMatch Then
afegirrcolor:
               oreport.FormulaFields.GetItemByName("rcolor" + atrim(rsttintesanteriors!ordretinter)).Text = """" + atrim(rsttintesanteriors!Color) + """"
               oreport.FormulaFields.GetItemByName("rmarcar" + atrim(rsttintesanteriors!ordretinter)).Text = """X"""
                 Else: If rsttintes2!afectatspelcanvi Then GoTo afegirrcolor
           End If
           rsttintesanteriors.MoveNext
        Wend
   End If
    rsttintes2.MoveFirst
    While Not rsttintes2.EOF
       If rsttintes2!afectatspelcanvi Then
          oreport.FormulaFields.GetItemByName("color" + atrim(rsttintes2!ordretinter)).Text = """" + atrim(rsttintes2!Color) + """"
          oreport.FormulaFields.GetItemByName("marcar" + atrim(rsttintes2!ordretinter)).Text = """X"""
       End If
       rsttintes2.MoveNext
    Wend
    
   oreport.FormulaFields.GetItemByName("soportetextos").Text = """" + atrim(rstmodifi!textevalidaciotexte) + """"
   oreport.FormulaFields.GetItemByName("datatextos").Text = """" + Format(rstmodifi!datavalidaciotexte, "dd/mm/yy") + """"
   oreport.FormulaFields.GetItemByName("soportemedidas").Text = """" + atrim(rstmodifi!textevalidaciomides) + """"
   oreport.FormulaFields.GetItemByName("datamedidas").Text = """" + Format(rstmodifi!datavalidaciomides, "dd/mm/yy") + """"
   oreport.FormulaFields.GetItemByName("soportecolores").Text = """" + atrim(rstmodifi!textevalidaciocolors) + """"
   oreport.FormulaFields.GetItemByName("datacolores").Text = """" + Format(rstmodifi!datavalidaciocolors, "dd/mm/yy") + """"
   oreport.FormulaFields.GetItemByName("nomfirma1").Text = """" + "Firma: " + atrim(rstcap!operaria) + """"
   oreport.FormulaFields.GetItemByName("nomfirma2").Text = """" + "Firma: " + atrim(rstcap!nomoprevisat) + """"
   
fi:
   possarvalorsalreport = True
   Set rstcap = Nothing
   Set rstclixe = Nothing
   Set rstmodifi = Nothing
   Set rsttintes = Nothing
End Function
Function comandesrelacionades() As String
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda,comandes.numtreball,comandes.numordremodificacio FROM comandes LEFT JOIN clients ON comandes.client = clients.codi where (producte<>'PC' and producte<>'PCP' and producte<>'PC2') and numordremodificacio=" + atrim(cadbl(cversio)) + " and numtreball=" + atrim(cadbl(cidtreball)) + " and (proximaseccio <>'T') order by comanda Desc")
   While Not rst.EOF
       comandesrelacionades = comandesrelacionades + IIf(comandesrelacionades <> "", ", ", "") + atrim(rst!comanda) + " [" + atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio) + "]"
       rst.MoveNext
   Wend
   Set rst = Nothing
End Function
Function rutamodifispdftreball(vidtreball As Double, vordre As Double) As String
   On Error Resume Next
   MkDir ruta_documentacio_clixes + "\" + Format(vidtreball, "00000")
   rutamodifispdftreball = ruta_documentacio_clixes + "\" + Format(vidtreball, "00000") + "\MODIFI" + Format(vidtreball, "00000") + "-" + Format(vordre, "000") + ".pdf"
   If existeix(rutamodifispdftreball) Then Kill rutamodifispdftreball
End Function

Sub imprimir_fulla_modificacions()
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vnomfitxermodifispdf As String
  vnomfitxermodifispdf = rutamodifispdftreball(cadbl(cidtreball), cadbl(cversio))
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "fullamodificacionsclixe.rpt", 1)
  If Not possarvalorsalreport(oreport) Then Exit Sub
  oreport.DiscardSavedData
  'If existeix("c:\ordprog.ini") Then
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DiskFileName = vnomfitxermodifispdf
  oreport.Export False
  oreport.ExportOptions.DestinationType = crEDTNoDestination
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
   ' Else
   '   oreport.PrintOut False, 1
 ' End If
End Sub
Sub eliminar_toteslesdadesdaquestrepas(vidtreball As Double, vordre As Double)
   Dim rst As Recordset
   Dim rstcap As Recordset
   
   Set rstcap = dbclixes.OpenRecordset("select * from repasclixes where id_treball=" + atrim(vidtreball) + " and nummodificacio=" + atrim(vordre))
   If rstcap.EOF Then Exit Sub
   Set rst = dbclixes.OpenRecordset("select * from repasdadestintes where id_repas=" + atrim(rstcap!id_repas))
   While Not rst.EOF
      dbclixes.Execute "delete * from repasamplesmotius where id_repasdades=" + atrim(rst!id_repasdades)
      rst.MoveNext
   Wend
   Set rst = Nothing
   dbclixes.Execute "delete * from repasdadestintes where id_repas=" + atrim(rstcap!id_repas)
   dbclixes.Execute "delete * from repasclixes where id_repas=" + atrim(rstcap!id_repas)
   Set rstcap = Nothing
End Sub
Private Sub Command11_Click()
  Dim v As String
  If cadbl(cidtreball) = 0 Then Exit Sub
  v = UCase(InputBox("Estàs eliminant TOTES les dades d'aquest repàs." + vbNewLine + "ESCRIU [ELIMINAR] PER ELIMINAR AQUESTES DADES.", "ELIMINAR DADES DEL REPÀS"))
  If StrPtr(v) = 0 Then Exit Sub
  If v = "ELIMINAR" Then
      ELIMINARelPDFguardat
      eliminar_toteslesdadesdaquestrepas cadbl(cidtreball), cadbl(cversio)
      netejarbotons
      lnomclient = ""
      lnomclient.tag = ""
      lmarcailinia = ""
      lmarcailinia.tag = ""
      lcodidebarres = ""
      lmodificatonouidescproducte = ""
      lmaterialcomanda = ""
      carregarrepasdeltreball 0, 0
      botototok.visible = False
  End If
End Sub

Private Sub Command12_Click()
   Dim rstcap As Recordset
   Dim rstdades As Recordset
   Dim rstamples As Recordset
   Set rstcap = dbclixes.OpenRecordset("repasclixes")
   Set rstdades = dbclixes.OpenRecordset("repasdadestintes")
   Set rstamples = dbclixes.OpenRecordset("repasamplesmotius")
   If MsgBox("Aixó borrarà les dades del repas d'aquest tinter." + vbNewLine + "ESTAS SEGUR QUE VOLS FER-HO?", vbCritical + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbNo Then Exit Sub
   borrar_dadesguardades cadbl(cidtreball), cadbl(cversio), cadbl(framedescripciotinter.tag)
   wait 1
   carregardescripciotinta cadbl(cidtreball), cadbl(cversio), cadbl(botonstinters(Index).tag)
   
   Set rstcap = Nothing
   Set rstdades = Nothing
   Set rstamples = Nothing
End Sub
Sub borrar_ultims5diesMuntadora()
  Dim v As String
  Dim vruta As String
  vruta = "\\pc-vision-02\xmlinput"
  If Not existeix(vruta) Then Exit Sub
  v = Dir(vruta + "\*.xml")
  While v <> ""
    vdata = FileDateTime(vruta + "\" + v)
    If DateDiff("d", vdata, Now) > 7 Then Kill vruta + "\" + v
    v = Dir
  Wend
End Sub
Private Sub Command13_Click()
  Dim vnummotius As Byte
  Dim vmidacamisa As Double
  Dim x1 As Double
  Dim x2 As Double
  Dim vnomfitxerxml As String
  Command5_Click
  vnummotius = cadbl(cnumbandes)
  If cadbl(midacilindre.tag) = 0 Then MsgBox "La mida del cilindre d'aquest treball està a zero.", vbCritical, "Error": GoTo imprimiretiqueta
  If cadbl(cmicropuntstotal) = 0 Then MsgBox "Si la distancia de micropunts total es zero no s'envia el fitxer a muntadora.": GoTo imprimiretiqueta
  'If Not verificar_micropuntsibandes Then Exit Sub
  borrar_ultims5diesMuntadora
  If existeix(vnomfitxerxml) Then Kill vnomfitxerxml
  vmidacamisa = 1240
  x1 = (vmidacamisa - cadbl(cmicropuntstotal)) / 2
  x2 = x1 + IIf(vnummotius = 1, cadbl(cmicropuntstotal), cadbl(cmicropuntmotiu))
  vnomfitxerxml = "\\pc-vision-02\XmlInput\" + cidtreball + "-" + cversio + ".xml"
  ' els altre micropunts*xrquantitatdemotius-1 seran vsegonmicropunt + espaientremicropunts   i el laltra +amplemotiu
  'fer el fitxer xml
  generar_xml vnomfitxerxml, x1, x2
  If existeix(vnomfitxerxml) Then etenviatamuntadora.tag = "True": MsgBox "FITXER CREAT CORRECTAMENT A LA MUNTADORA.", vbExclamation, "ATENCIÓ"
  Command5_Click
  
  If atrim(Mid(lcodidebarres, InStr(1, lcodidebarres, "Arxiu XL:") + 10)) = "" Then
       vXL = demanararxiu(True)
       If vXL <> "" Then
        dbclixes.Execute "update clixes set arxiu='" + atrim(vXL) + "' where id_treball=" + atrim(cidtreball)
        carregarrepasdeltreball cadbl(cidtreball), cadbl(cversio)
        lcodidebarres = lcodidebarres + " " + vXL
         Else: Exit Sub
       End If
  End If
imprimiretiqueta:
  If MsgBox("Vols imprimir la bossa del treball?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
      imprimiretiquetabossaclixesdemanantimpresora cadbl(cidtreball), cadbl(cversio), llistat, False
  End If
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
  'Clipboard.Clear
  'Clipboard.SetText "SELECT count(Clixes.arxiu) as Q,first(clixes.arxiu) as Parxiu, cdbl(mid(trim(first(Clixes.arxiu)),4)) as NumXL FROM Clixes WHERE " + vXL + " GROUP BY Clixes.arxiu " + vcriteri + " ORDER BY Count(Clixes.arxiu) asc;"
  
  If Not rst.EOF Then suggerirXL = rst!NumXL
fi:
  Set rst = Nothing
End Function

Sub generar_xml(vnomfitxerxml As String, x1 As Double, x2 As Double)
     Open vnomfitxerxml For Output As #1
     genera_xml_capçalera
     genera_xml_linies x1, x2
     Print #1, vbTab + "</XML_Montacliche>" + vbNewLine + "</DocumentElement>" + vbNewLine
     Close #1
End Sub

Sub genera_xml_linies(x1 As Double, x2 As Double)
   vlinia = vbTab + vbTab + "<Colore_0>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<Colore>QUALSEVOL COLOR</Colore>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<Note> </Note>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<TipoBiadesivo></TipoBiadesivo>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<TipoAnilox> </TipoAnilox>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<NrManica> </NrManica>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + vbTab + "<NumStep>1</NumStep>" + vbNewLine
   vsteps = 1
   For i = 0 To vsteps - 1
        vlinia = vlinia + vbTab + vbTab + vbTab + "<Step_" + atrim(i) + ">" + vbNewLine
        vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + "<W>0</W>" + vbNewLine
        vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + "<NumPosizioni>" + atrim(cadbl(cnumbandes)) + "</NumPosizioni>" + vbNewLine
        vposicions = cadbl(cnumbandes)
        For j = 0 To vposicions - 1
            x1 = Redondejar(x1, 0)
            x2 = Redondejar(x2, 0)
            vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + "<Posizione_" + atrim(j) + ">" + vbNewLine
            vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + vbTab + "<X1>" + atrim(x1) + "</X1>" + vbNewLine
            vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + vbTab + "<X2>" + atrim(x2) + "</X2>" + vbNewLine
            vlinia = vlinia + vbTab + vbTab + vbTab + vbTab + "</Posizione_" + atrim(j) + ">" + vbNewLine
            If cnumbandes > 2 Then
                 x1 = x2 + cadbl(centremicropunts): x2 = x1 + cadbl(cmicropuntmotiu)
                   Else:
                     x1 = (cadbl(cmicropuntstotal) + x1) - cadbl(cmicropuntmotiu)
                     x2 = x1 + cadbl(cmicropuntmotiu)
            End If
        Next j
        
fistep:
        vlinia = vlinia + vbTab + vbTab + vbTab + "</Step_" + atrim(i) + ">" + vbNewLine
   Next i
   vlinia = vlinia + vbTab + vbTab + "</Colore_0>"
   Print #1, vlinia
End Sub
Sub genera_xml_capçalera()
   Dim vlinia As String
   vlinia = "<?xml version=""1.0""?>" + vbNewLine
   vlinia = vlinia + "<DocumentElement Version=""1.0"">" + vbNewLine
   vlinia = vlinia + vbTab + "<XML_Montacliche>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + "<PR>" + atrim(cadbl(midacilindre.tag)) + "</PR>" + vbNewLine
   vlinia = vlinia + vbTab + vbTab + "<Descrizione>" + treure_simbolsextranys(atrim(lmarcailinia.tag)) + "</Descrizione>" + vbNewLine
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
     v = substituirtot(v, Mid(UCase(vaccents), i, 1), Mid(UCase(vnoaccents), i, 1))
  Next i
  treure_simbolsextranys = v
  'v = TextoSinAcentos(v)
     
End Function
Private Sub Command14_Click()
   Shell "\\SERVERPRODU\Dades\progcomandes\aplicacio\Manteniment tintes.exe agrupartreballs", vbNormalFocus
End Sub

Private Sub Command2_Click()
  veurelacomanda cadbl(ccomanda)
End Sub

Private Sub Command3_Click()
  veureelpdf cadbl(ccomanda), "TOT"
End Sub

Private Sub Command4_Click()
    veureelimp cadbl(ccomanda)
End Sub

Private Sub Command5_Click()
   Dim esticaltinter As Byte
   If descripciotinter = "" Then Exit Sub
'   If Not verificar_micropuntsibandes Then Exit Sub
   esticaltinter = cadbl(framedescripciotinter.tag)
   comprovacions_correctes
   guardar_dades
   guardat = True
   netejaridsicolors
   comprovacions_correctes
   carregarrepasdeltreball cadbl(cidtreball), cadbl(cversio)
   If esticaltinter > 0 Then botonstinters_Click esticaltinter - 1
   guardat = True
   If botototok.visible = True Then
       passar_treball_a_revisat
   End If
End Sub
Function verificar_micropuntsibandes() As Boolean
  Dim vnummotius As Double
  vnummotius = cadbl(cnumbandes)
  verificar_micropuntsibandes = True
  If vnummotius = 0 Or cadbl(cmicropuntstotal) = 0 Then MsgBox "No hi ha el numero de motius o la distancia de micropunts.", vbCritical, "Error": verificar_micropuntsibandes = False: GoTo fi
  If vnummotius = 2 And cadbl(cmicropuntmotiu) = 0 Then MsgBox "No hi ha el distancia entre motius.", vbCritical, "Error": verificar_micropuntsibandes = False: GoTo fi
  If vnummotius = 3 And cadbl(centremicropunts) = 0 Then MsgBox "No hi ha el distancia entre micropunts.", vbCritical, "Error": verificar_micropuntsibandes = False: GoTo fi
fi:
End Function
Function treure_entersinici(r) As String
   Dim i As Byte
   Dim v As String
   v = r
   While Len(v) > 1
      If Asc(Mid(v, 1, 1)) < 32 Then
              v = Mid(v, 2)
          Else: GoTo fi
      End If
   Wend
fi:
   treure_entersinici = v
End Function
Sub demanar_observacions_tintes(vidtreball As Double, vversio As Double)
     Dim rant As String
     Dim v As String
     Load obsidtreball
     obsidtreball.caption = "Observacions pel treball " + atrim(vidtreball)
     Set rst = dbtmpb.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(vidtreball) + " and ordre=" + atrim(vversio) + " order by id")
     While Not rst.EOF
            obsidtreball.obsid = obsidtreball.obsid + atrim(rst!observacio) + vbNewLine
            rst.MoveNext
     Wend
     rant = obsidtreball.obsid
   '  r = rant
     obsidtreball.Show 1
     If r <> rant Then
         
         dbtmpb.Execute "delete * from tintes_observacions where id_treball=" + atrim(vidtreball) + " and ordre=" + atrim(vversio)
         While InStr(1, r, Chr(13)) > 0
            r = treure_entersinici(r)
            v = treure_apostruf(Mid(r, 1, InStr(1, r, Chr(13))))
            v = treure_entersinici(v)
            If Mid(v, 1, 1) <> Chr(10) And atrim(v) <> "" Then
                dbtmpb.Execute "insert into tintes_observacions (id_treball,ordre,observacio) values (" + atrim(vidtreball) + "," + atrim(vversio) + ",'" + atrim(v) + "')"
            End If
            r = atrim(Mid(r, InStr(1, r, Chr(13)) + 1))
         Wend
     End If
     Set rst = Nothing
End Sub
Sub passar_treball_a_revisat()
     Dim vpassaraFET As Boolean
     Dim vnomfitxerpdf As String
     If mirar_si_faltaalgunadata Then MsgBox "Falta entrar la data de validació de colors.", vbCritical, "Error": Exit Sub
     If MsgBox("Vols passar aquest treball a FET i posar-lo a Taula-3?", vbInformation + vbDefaultButton2 + vbYesNo, "Fet?") = vbYes Then vpassaraFET = True
     If vpassaraFET Then demanar_observacions_tintes cidtreball, IIf(cversio > 1, cversio - 1, 1)
     vavisarREPASADORA = False
     passarrepasalasecciodIMPRESORES vpassaraFET
     vnomfitxerpdf = ruta_documentacio_clixes + "\" + Format(cadbl(cidtreball), "00000") + "\Arxiu_documentacio_relacionada" + "\v" + atrim(cadbl(cversio)) + "\Repasdeclixes.pdf"
     If Not existeix(vnomfitxerpdf) Then imprimirrepas.SetFocus: imprimirrepas_Click
     If vavisarREPASADORA And cversio > 1 Then treballrevisat_avisarREPASADOR
     'mirarsihihaelPDFguardat
End Sub
Function mirar_si_faltaalgunadata() As Boolean
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("SELECT modificacions.datavalidaciocolors,textevalidaciocolors FROM modificacions Where modificacions.id_treball = " + atrim(cidtreball) + " And ordre = " + atrim(cversio))
   If rst.EOF Then GoTo fi
   If atrim(rst!textevalidaciocolors) <> "MOSTRA VERIFICACIÓ" Then Exit Function
   If IsNull(rst!datavalidaciocolors) Then
       vdata = InputBox("Entra la data de validació de colors.", "Data validació de colors")
       If StrPtr(vdata) = 0 Then GoTo fi
       If IsDate(vdata) Then
            rst.Edit: rst!datavalidaciocolors = vdata: rst.Update
            While UCase(InputBox("En aquest treball has d'enganxar una etiqueta de mostra." + vbNewLine + "ESCRIU [ETIQUETA MOSTRA] PER PODER CONTINUAR.", "ETIQUETA MOSTRA")) <> "ETIQUETA MOSTRA"
               DoEvents
            Wend
            mirar_si_faltaalgunadata = False
              Else: mirar_si_faltaalgunadata = True
       End If
         Else: mirar_si_faltaalgunadata = False
   End If
fi:
   Set rst = Nothing
End Function

Sub treballrevisat_avisarREPASADOR()
   If MsgBox("El treball està revisat, vols avisar al REPASADOR per fer la firma?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
       enviaremailgeneric "ellinas@inplacsa.com", "Repàs de TREBALL " + atrim(cidtreball) + "/" + cversio + "  -Comanda: " + ccomanda + " ->" + cpersona, "El Treball: " + atrim(cidtreball) + "/" + cversio + " de la comanda: " + ccomanda + " ja s'ha repasat falta la firma de la Revisora." + vbNewLine + "Data: " + atrim(Now)
   End If
End Sub
Sub netejaridsicolors()
   
   For i = 0 To 7
      botonstinters(i).HelpContextID = 0
   Next i
   treurecolorbotons
End Sub

Private Sub Command6_Click()
   imprimiretiquetabossaclixesdemanantimpresora cadbl(cidtreball), cadbl(cversio), llistat, False
End Sub
Sub escullirpersona()
 Dim numoptmp As Integer
 Dim nomoptmp As String
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='C' and actiu<>0"
  formseleccio.caption = "Selecció d'Operari"
  formseleccio.refrescar
  formseleccio.width = 5000
  formseleccio.DBGrid2.Columns(1).width = 3000
  formseleccio.Show 1
  If seleccioret = 1 Then
   If Not contrasenyacorrecte(formseleccio.Data1.Recordset!codi) Then MsgBox "La contrasenya no es correcte", vbCritical, "Error": GoTo fi
   cpersona.tag = cadbl(formseleccio.Data1.Recordset!codi)
   cpersona = atrim(formseleccio.Data1.Recordset!descripcio)
   'If InStr(1, nomoperari.Caption, "MARTINEZ") Then
   '    Command12.Visible = True
   '   Else: Command12.Visible = False
   'End If
  End If
fi:
  If seleccioret = 0 Then End
  Unload formseleccio
  
End Sub
Sub escullirrevisador()
 Dim numoptmp As Integer
 Dim nomoptmp As String
 Dim vcodioperari As Long
 vcodioperari = cadbl(cpersona.tag)
 If vcodioperari = 1 Then vcodioperari = -1  'si es l'EVA pot ser repasador i Revisor
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='C' and actiu<>0 and codi<>" + atrim(vcodioperari)
  formseleccio.caption = "Selecció del Revisador"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   If Not contrasenyacorrecte(formseleccio.Data1.Recordset!codi) Then MsgBox "La contrasenya no es correcte", vbCritical, "Error": GoTo fi
   etnomrevisador.tag = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   etnomrevisador = atrim(formseleccio.Data1.Recordset!descripcio)
   'If InStr(1, nomoperari.Caption, "MARTINEZ") Then
   '    Command12.Visible = True
   '   Else: Command12.Visible = False
   'End If
  End If
fi:
  Unload formseleccio
  
End Sub
Function contrasenyacorrecte(vcodi As Long) As Boolean
   Dim rst As Recordset
   Dim vcontrasenya As String
   contrasenyacorrecte = False
   Set rst = dbcomandes.OpenRecordset("select * from operaris_contrasenyes where seccio='C' and operari=" + atrim(vcodi))
   If rst.EOF Then Exit Function
   vcontrasenya = InputBoxEx("Entra la contrasenya de l'operari", "Contrasenya", , , , , , SPassword)
   If vcontrasenya = rst!contrasenya Then
        contrasenyacorrecte = True
   End If
End Function
Sub imprimir_repas()
 ' Dim rst As Recordset
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
 
  
 ' llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "incidenciescomandaitreball.rpt"
 ' If Not existeix("c:\ordprog.ini") Then llistat.Destination = crptToPrinter
 ' llistat.DataFiles(0) = rutadelfitxer(cami) + "clixesnous.mdb"
 ' llistat.SelectionFormula = "{diferenciescomandaitreball.comanda}=" + atrim(numc)
 ' llistat.Action = 1
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "impresiorepasdeclixes.rpt", 1)
  oreport.Database.Tables.item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
  oreport.RecordSelectionFormula = "{repasclixes.id_repas}=" + atrim(cadbl(framedadestintes.tag)) + " and {repasdadestintes.ordretinter}>0"
  possarlafirma cadbl(framedadestintes.tag)
  wait 2
  oreport.DiscardSavedData
  'If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
   ' Else
   '   oreport.PrintOut False, 1
 ' End If
End Sub

Private Sub Command7_Click()
  Dim rst As Recordset
  Dim vnompdf As String
  Set rst = dbclixes.OpenRecordset("select * from repasclixes")
  While Not rst.EOF
    If rst!id_treball > 0 Then
     cidtreball = rst!id_treball
     cversio = rst!nummodificacio
     If cversio = "0" Then cversio = "1"
     framedadestintes.tag = atrim(rst!id_repas)
     exportar_carpetatreball vnompdf
    End If
     rst.MoveNext
  Wend
  Set rst = Nothing
End Sub

Private Sub Command8_Click()
    veureelpdf cadbl(ccomanda), "SC"
End Sub

Private Sub Command9_Click()
   etnomrevisador.tag = ""
   While cadbl(etnomrevisador.tag) = 0
     escullirrevisador
     If cadbl(etnomrevisador.tag) = 0 Then MsgBox "Has d'escullir un operari per treballar"
     If etnomrevisador.tag = "" Then Exit Sub
     Command5.Enabled = False
   Wend
   'guardar_dades
End Sub

Private Sub creuvermella_Click()
   On Error Resume Next
   If Screen.ActiveControl.Name = "creuvermella" Then
    If creuvermella.Value = 1 Then
       posar_ok True, 1
      Else: posar_ok False, 1
    End If
   End If
End Sub

Private Sub defectuos_Click()
If Screen.ActiveControl.Name = "defectuos" Then
    If defectuos.Value = 1 Then
       posar_ok True, 4
      Else: posar_ok False, 4
    End If
   End If
End Sub

Function carregarrepasdeltreball(id_treball As Double, nummodificacio As Double) As Boolean
   Dim rstcap As Recordset
   etenviatamuntadora = ""
   etenviatamuntadora.tag = ""
   framedadestintes.tag = ""
   cidtreball.tag = ""
   descripciotinter = ""
   Set rstcap = dbclixes.OpenRecordset("select * from repasclixes where id_treball=" + atrim(id_treball) + " and nummodificacio=" + atrim(nummodificacio))
   If rstcap.EOF Then carregarrepasdeltreball = False: Exit Function
   framedadestintes.tag = atrim(rstcap!id_repas)
   carregar_dades id_treball, nummodificacio
   carregarrepasdeltreball = True
   operariaqueharevisat = "Repasat per: " + atrim(rstcap!operaria)
   centremicropunts = atrim(cadbl(rstcap!micropuntsdistanciaentremotius))
   cmicropuntmotiu = atrim(cadbl(rstcap!micropuntsmidamotiu))
   cmicropuntstotal = atrim(cadbl(rstcap!micropuntsdistanciatotal))
   cnumbandes = atrim(cadbl(rstcap!numbandes))
   etenviatamuntadora.tag = atrim(rstcap!enviatamuntadora)
   If rstcap!enviatamuntadora Then etenviatamuntadora = "ENVIAT a muntadora"
   If Not rstcap!enviatamuntadora Then etenviatamuntadora = "NO enviat a muntadora"
   'operariaqueharevisat.tag = atrim(rstcap!numop)
'   etnomrevisador = atrim(rstcap!nomoprevisat)
'   etnomrevisador.tag = atrim(rstcap!numoprevisat)
   Set rstcap = Nothing
End Function
Sub carregar_dades(id_treball As Double, nummodificacio As Double)
  Dim id_repas As Long
  id_repas = carregar_dades_capcalera(id_treball, nummodificacio)
  carregar_dades_tintes_botons id_repas
  cidtreball.tag = atrim(id_repas)
End Sub
Sub carregar_dades_tintes_botons(id_repas As Long)
   Dim rstdades As Recordset
  
   
   Set rstdades = dbclixes.OpenRecordset("select * from repasdadestintes where id_repas=" + atrim(id_repas) + " order by ordretinter DESC")
   If rstdades.EOF Then GoTo nohihares
   While Not rstdades.EOF
    If cadbl(rstdades!ordretinter) > 0 Then
     botonstinters(rstdades!ordretinter - 1) = atrim(cadbl(rstdades!nomtinter))
     botonstinters(rstdades!ordretinter - 1).tag = atrim(cadbl(rstdades!ordretinter))
     framedescripciotinter.tag = atrim(cadbl(rstdades!ordretinter))
     carregar_dades_tintes rstdades
    comprovacions_correctes
    End If
     rstdades.MoveNext
   Wend
   carregar_tintes_enblanc
   
nohihares:
   Set rstdades = Nothing
   treurecolorbotons
End Sub
Sub carregar_tintes_enblanc()
   Dim rstdades As Recordset
   Dim idcolorboto As Long
   If cadbl(framedescripciotinter.tag) < 1 Then Exit Sub
   idcolorboto = botonstinters(cadbl(framedescripciotinter.tag) - 1).HelpContextID
   Set rstdades = dbclixes.OpenRecordset("select * from repasdadestintes where id_repas=0")
   carregar_dades_tintes rstdades
   botonstinters(cadbl(framedescripciotinter.tag) - 1).HelpContextID = idcolorboto
End Sub
Sub carregar_dades_tintes(rstdades As Recordset)
   
   If rstdades.EOF Then botonstinters(cadbl(framedescripciotinter.tag) - 1).HelpContextID = 0: Exit Sub
   With rstdades
   If cadbl(!ordretinter) = 0 Then
       Set rstdades = dbclixes.OpenRecordset("select * from repasdadestintes where id_repas=0")
       
        Else
          framedescripciotinter.tag = atrim(cadbl(!ordretinter))
   End If
   descripciotinter.tag = atrim(!nomtinter)
   'descripciotinter = atrim(!descripciotinter)
   creuvermella = IIf(!creuvermella, 1, 0)
   repasclixe = IIf(!repasclixe, 1, 0)
   micropunts = IIf(!micropunts, 1, 0)
   defectuos = IIf(!defectuos, 1, 0)
   gruixplimer(0) = IIf(!gruixpol = 1.14, True, False)
   gruixplimer(1) = IIf(!gruixpol = 2.54, True, False)
   gruixplimer(2) = IIf(!gruixpol = 2.84, True, False)
   lecturarelleu = IIf(!lecturarelleu, 1, 0)
   lecturallisa = IIf(!lecturallisa, 1, 0)
   ample = CDbl(!ample)
   If tintaportasang(!ordretinter) Then
      comboposiciosang = atrim(!posiciosang)
      frameposiciosang.visible = True
      etampleisang.caption = "Entra l'ample + sang que correspon:"
     Else:
       comboposiciosang = ""
       frameposiciosang.visible = False
       etampleisang.caption = "Entra l'ample que correspon:"
   End If
   motiuamotiu = cadbl(!motiuamotiu)
   ampladatotalmotiu = cadbl(!ampladatotalmotiu)
   midadesarroll = cadbl(!midadesarroll)
   midacilindre = cadbl(!midacilindre)
   numerodemotius = cadbl(!numerodemotius)
   combobandaseguiment = atrim(!bandaseguiment)
   amplebanda = atrim(cadbl(!amplebanda))
   macula = atrim(!macula)
   Combotipusfoam = atrim(!tipusdefoam)
   clixesllencats = IIf(!clixesllencats, 1, 0)
   carregar_oks rstdades
   If !ordretinter = 0 Then
        botonstinters(cadbl(framedescripciotinter.tag) - 1).HelpContextID = 0
       Else: botonstinters(cadbl(!ordretinter) - 1).HelpContextID = 2
   End If
   End With
   If escorrectemidadesarroll(cadbl(midadesarroll)) Then
     posar_ok True, 13
    Else: posar_ok False, 13
  End If
End Sub
Function tintaportasang(tinter As Byte) As Boolean
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from tintes where ordretinter=" + atrim(tinter) + " and id_treball=" + atrim(cadbl(cidtreball)) + " and ordremodificacio=" + atrim(cversio))
   If rst.EOF Then Exit Function
   If rst!portasang Then tintaportasang = True
End Function
Sub carregar_amples(rstdades As Recordset)
   Dim rst As Recordset
   netejarmotius
   Set rst = dbclixes.OpenRecordset("select * from repasamplesmotius where id_repasdades=" + atrim(rstdades!id_repasdades))
   If rst.EOF Then Exit Sub
   While Not rst.EOF
     amplemotiu(rst!nummotiu - 1).visible = True
     amplemotiu(rst!nummotiu - 1).Text = rst!ample
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub netejarmotius()
   Dim i As Byte
   For i = 0 To 11
     amplemotiu(i) = ""
   Next i
End Sub
Function carregar_dades_capcalera(id_treball, nummodificacio) As Long
   Dim rstcap As Recordset
   Set rstcap = dbclixes.OpenRecordset("select * from repasclixes where id_treball=" + atrim(id_treball) + " and nummodificacio=" + atrim(nummodificacio))
   If rstcap.EOF Then Exit Function
   cidtreball = cadbl(rstcap!id_treball)
   cversio = cadbl(rstcap!nummodificacio)
   ccomanda = cadbl(rstcap!comanda)
   repasarprova = IIf(rstcap!repasprova, 1, 0)
   carregar_dades_capcalera = rstcap!id_repas
   Set rstcap = Nothing
End Function
Sub guardar_dades()
   Dim rstcap As Recordset
   Dim rstdades As Recordset
   Dim rstamples As Recordset
'   If etnomrevisador.tag <> "" Then MsgBox "No es pot guardar un repàs si hi ha un REVISOR escullit."
   Set rstcap = dbclixes.OpenRecordset("repasclixes")
   Set rstdades = dbclixes.OpenRecordset("repasdadestintes")
   Set rstamples = dbclixes.OpenRecordset("repasamplesmotius")
   borrar_dadesguardades cadbl(cidtreball), cadbl(cversio), cadbl(framedescripciotinter.tag)
   gravar_capcalera rstcap
   gravar_dadestintes rstcap!id_repas, rstdades, rstamples
   
   Set rstcap = Nothing
   Set rstdades = Nothing
   Set rstamples = Nothing
End Sub
Sub borrar_dadesguardades(id_treball As Double, ordremodificacio As Double, vordretinter As Double)
   Dim rstcap As Recordset
   Dim rstdades As Recordset
   Dim idrepasdades As Long
   Set rstcap = dbclixes.OpenRecordset("select * from repasclixes where id_treball=" + atrim(id_treball) + " and nummodificacio=" + atrim(ordremodificacio))
   If rstcap.EOF Then Exit Sub
   Set rstdades = dbclixes.OpenRecordset("select * from repasdadestintes where ordretinter=" + atrim(vordretinter) + " and id_repas=" + atrim(rstcap!id_repas))
   If rstdades.EOF Then Exit Sub
   idrepasdades = cadbl(rstdades!id_repasdades)
   Set rstdades = Nothing
   dbclixes.Execute "delete * from repasamplesmotius where id_repasdades=" + atrim(idrepasdades)
   dbclixes.Execute "delete * from repasdadestintes where id_repasdades=" + atrim(idrepasdades)
   Set rstcap = Nothing
End Sub
Sub gravar_dadestintes(id_repas As Long, rstdades As Recordset, rstamples As Recordset)
   With rstdades
   .AddNew
   !id_repas = id_repas
   !ordretinter = cadbl(framedescripciotinter.tag)
   !nomtinter = Mid(descripciotinter.tag, 1, 50)
   !descripciotinter = Mid(descripciotinter, 1, 80)
   !creuvermella = IIf(creuvermella = 1, True, False)
   !repasclixe = IIf(repasclixe = 1, True, False)
   !micropunts = IIf(micropunts = 1, True, False)
   !defectuos = IIf(defectuos = 1, True, False)
   !gruixpol = IIf(gruixplimer(0), 1.14, IIf(gruixplimer(1), 2.54, IIf(gruixplimer(2), 2.84, 0)))
   !lecturarelleu = IIf(lecturarelleu = 1, True, False)
   !lecturallisa = IIf(lecturallisa = 1, True, False)
   !ample = cadbl(ample)
   !posiciosang = atrim(comboposiciosang)
   !motiuamotiu = cadbl(motiuamotiu)
   !ampladatotalmotiu = cadbl(ampladatotalmotiu)
   !midadesarroll = cadbl(numerodemotius.tag)
   !midacilindre = cadbl(midacilindre.tag)
   !numerodemotius = cadbl(numerodemotius)
   !bandaseguiment = atrim(combobandaseguiment)
   !amplebanda = cadbl(amplebanda)
   !macula = atrim(macula)
   !tipusdefoam = atrim(Combotipusfoam)
   !clixesllencats = IIf(clixesllencats = 1, True, False)
   gravar_oks rstdades
   .Update
   rstdades.Bookmark = rstdades.LastModified
   gravaramples rstdades!id_repasdades, rstamples
   
   End With
End Sub
Sub gravar_oks(rstdades As Recordset)
   Dim i As Byte
   For i = 0 To 12
      rstdades.Fields("ok" + Trim(i + 1)) = IIf(estat(i).Picture = correcte.Picture, True, False)
   Next i
End Sub
Sub posar_ok(vestat As Boolean, item As Byte)
    estat(item - 1).Picture = IIf(vestat, correcte.Picture, nocorrecte.Picture)
End Sub
Sub carregar_oks(rstdades As Recordset)
   Dim i As Byte
   Dim totsok As Boolean
   totsok = True
   For i = 0 To 12
      estat(i).Picture = IIf(rstdades.Fields("ok" + Trim(i + 1)), correcte.Picture, nocorrecte.Picture)
      If estat(i).Picture = nocorrecte.Picture Then totsok = False
   Next i
   If cadbl(framedescripciotinter.tag) > 0 Then
    If totsok Then
       botonstinters(cadbl(framedescripciotinter.tag) - 1).HelpContextID = 1
        Else: botonstinters(cadbl(framedescripciotinter.tag) - 1).HelpContextID = 2
    End If
   End If
End Sub

Sub gravaramples(id_repasdades As Long, rstamples As Recordset)
    Dim i As Byte
    With rstamples
    For i = 0 To amplemotiu.Count - 1
       If cadbl(amplemotiu(i)) > 0 Then
           .AddNew
           !id_repasdades = id_repasdades
           !nummotiu = i + 1
           !ample = cadbl(amplemotiu(i))
           .Update
       End If
    Next i
    End With
    
End Sub
Sub gravar_capcalera(rstcap As Recordset)
   Set rstcap = dbclixes.OpenRecordset("select  * from repasclixes where id_treball=" + atrim(cadbl(cidtreball)) + " and nummodificacio=" + atrim(cadbl(cversio)))
   If rstcap.EOF Then
        rstcap.AddNew
       Else: rstcap.Edit
   End If
   rstcap!id_treball = cadbl(cidtreball)
   rstcap!nummodificacio = cadbl(cversio)
   rstcap!comanda = cadbl(ccomanda)
   rstcap!repasprova = IIf(repasarprova = 1, True, False)
   rstcap!marcailinia = Mid(lmarcailinia.tag, 1, 100)
   rstcap!nomclient = lnomclient.tag
   rstcap!operaria = cpersona
   rstcap!numop = cpersona.tag
   rstcap!numoprevisat = cadbl(etnomrevisador.tag)
   rstcap!nomoprevisat = atrim(etnomrevisador)
   rstcap!totsok = botototok.visible
   rstcap!micropuntsdistanciaentremotius = cadbl(centremicropunts)
   rstcap!micropuntsmidamotiu = cadbl(cmicropuntmotiu)
   rstcap!micropuntsdistanciatotal = cadbl(cmicropuntstotal)
   rstcap!numbandes = cadbl(cnumbandes)
   rstcap!enviatamuntadora = CBool(IIf(etenviatamuntadora.tag = "", False, etenviatamuntadora.tag))
   rstcap.Update
   Set rstcap = dbclixes.OpenRecordset("select  * from repasclixes where id_treball=" + atrim(cadbl(cidtreball)) + " and nummodificacio=" + atrim(cadbl(cversio)))
End Sub

Private Sub Form_Activate()
   Dim vidtreball As Double
   Dim vordre As Double
   borrar_ultims5diesMuntadora
   'If cadbl(cpersona.tag) = 0 Then botoescullpersona_Click
   vidtreball = cadbl(llegir_ini("clixes", "repasidtreball", "comandes.ini"))
   vordre = cadbl(llegir_ini("clixes", "repasordre", "comandes.ini"))
   escriure_ini "clixes", "repasidtreball", "0", "comandes.ini"
   escriure_ini "clixes", "repasordre", "0", "comandes.ini"
   If vidtreball <> 0 Then demanartreballperrepasar vidtreball, vordre
End Sub

Private Sub Form_Click()
demanar_observacions_tintes 1879, 2
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  fitxerini = "comandes.ini"
  cami = llegir_ini("General", "cami", fitxerini)
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  '"c:\misdoc~1\commandes\comandes.mdb"
  If existeix("c:\ordprog.ini") Then cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
  inicidragover = 0
  hora = Now
 
  camiclixes = rutadelfitxer(cami) + "clixesnous.mdb"
  Set dbclixes = DBEngine.OpenDatabase(camiclixes)
  Set dbcomandes = DBEngine.OpenDatabase(cami)
  Set dbtmpb = DBEngine.OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
  Set dbtintes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  Set dbtmp = DBEngine.OpenDatabase(cami)
  botoescullpersona_Click
  carregarcombofoam
  Load formannex
'  centerscreen Me
  Me.Top = 1
  Me.Left = 1
  formannex.Top = 80
  formannex.Left = 0
  formannex.Left = Me.width
  formannex.Show
  formrepas.Top = 1
  formrepas.Left = 1
  
End Sub

Sub carregarcombofoam()
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select distinct inicialsfoam from adhesiusmuntadora order by inicialsfoam")
   While Not rst.EOF
     Combotipusfoam.AddItem rst!inicialsfoam
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub possarestat(Index As Byte, vestat As Boolean)
     estat(Index - 1).Picture = IIf(vestat, correcte.Picture, nocorrecte.Picture)
End Sub

Sub possarnommaterialdelacomanda(ncomanda As Double)
   Dim rst As Recordset
   Dim rstmat As Recordset
   Dim subseleccio As String
   lmaterialcomanda = ""
   Set rst = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(ncomanda))
   If rst.EOF Then Exit Sub
   Set rst = dbcomandes.OpenRecordset("select materialex,producte from comandes where comanda=" + atrim(rst!comanda) + "or comanda=" + atrim(cadbl(rst!linkcomanda1)) + " or comanda=" + atrim(cadbl(rst!linkcomanda2)))
   While Not rst.EOF
        Set rstmat = dbcomandes.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rst!materialex)))
        If rst!producte <> "PC" And rst!producte <> "PC2" Then lmaterialcomanda = "(1): "
        If rst!producte = "PC" Then lmaterialcomanda = lmaterialcomanda + " (2):"
        If rst!producte = "PC2" Then lmaterialcomanda = lmaterialcomanda + " (3):"
        
        lmaterialcomanda = lmaterialcomanda + "  " + descripciomaterial_curta(rstmat)
        rst.MoveNext
   Wend
   Set rst = Nothing
   Set rstmat = Nothing
End Sub
Function descripciomaterial_curta(rstmat As Recordset) As String
  Dim desc As String
  Dim rstfam As Recordset
  If rstmat.EOF Then Exit Function
  Set rstfam = dbcomandes.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)))
  If Not rstfam.EOF Then desc = desc + atrim(rstfam!descripcio)
 ' Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliesmaterials where codi=" + atrim(cadbl(rstmat!subfamilia)))
 ' If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbcomandes.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(rstmat!familiacol)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  'Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliescolorants where codi=" + atrim(cadbl(rstmat!subfamiliacol)))
  'If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  'Set rstfam = dbtmpb.OpenRecordset("select descripcio from familiesaditius where codi=" + atrim(cadbl(rstmat!familiaad)))
  'If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  'Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliesaditius where codi=" + atrim(cadbl(rstmat!subfamiliaad)))
  'If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  descripciomaterial_curta = desc
End Function

Function af(v As Variant) As String
  v = atrim(v)
  If Len(v) > 1 Then
     v = " + " + v
    Else: v = ""
  End If
  af = v
End Function

Private Sub veurelacomanda(comanda As Double)
  escriure_ini "Baixes", "imprimircomanda", cadbl(comanda), "comandes.ini"
  Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", "comandes.ini")) + "comandes.exe - imprimir", vbHide
   missatgevist.Show 1
End Sub
Sub veureelimp(comanda As Double)
   Dim rstc As Recordset
  Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(comanda)))
  obrir_imp_treball cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), cadbl(rstc!client), cadbl(rstc!direnvio)
End Sub
Sub obrir_imp_treball(treball As Double, modificacio As Double, codiclient As Double, direnvio As Double)
   Dim generarfitxer_imp As String
   If modificacio = 0 Then modificacio = 1
   generarfitxer_imp = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\IMP" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "-" + Format(codiclient, "000000") + "_" + atrim(direnvio) + ".doc"
   If Not existeix(generarfitxer_imp) Then generarfitxer_imp = generarfitxer_imp + "x"
   If existeix(generarfitxer_imp) Then
     obrir_document generarfitxer_imp
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + generarfitxer_imp, vbCritical, "Error"
  End If
End Sub

Sub obrir_pdf_treball(treball As Double, modificacio As Double, quin As String)
   Dim generarfitxer_pdf As String
   If modificacio = 0 Then modificacio = 1
   If quin = "SC" Then generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "_SC.pdf"
   If quin = "TOT" Then generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + ".pdf"
   If existeix(generarfitxer_pdf) Then
     obrir_document generarfitxer_pdf
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + generarfitxer_pdf, vbCritical, "Error"
  End If
End Sub

Sub veureelpdf(comanda As Double, quin As String)
  Dim rstc As Recordset
  Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(comanda)))
  obrir_pdf_treball cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), quin
  
End Sub

Private Sub Form_Resize()
   If formrepas.WindowState <> 0 Then
          formannex.visible = False
        Else: wait 1: formannex.visible = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload formannex
End Sub

Private Sub gruixplimer_Click(Index As Integer)
  If Screen.ActiveControl.Name <> "gruixplimer" Then Exit Sub
  If escorrectegruixpolimerifoam(cadbl(gruixplimer(Index).tag)) Then
     posar_ok True, 5
    Else: posar_ok False, 5
  End If
End Sub
Function escorrectegruixpolimerifoam(gruix As Double) As Boolean
   Dim rstmodifi As Recordset
   Set rstmodifi = dbclixes.OpenRecordset("select gruixpolimer from modificacions where id_treball=" + atrim(cidtreball) + " and ordre=" + atrim(cversio))
   If rstmodifi.EOF Then Exit Function
   If cadbl(rstmodifi!gruixpolimer) = atrim(gruix) Then escorrectegruixpolimerifoam = True
   If Combotipusfoam.Text = "" Then escorrectegruixpolimerifoam = False
   Set rstmodifi = Nothing
End Function


Private Sub imprimirrepas_Click()
  Dim vnompdf As String
  If cadbl(framedadestintes.tag) = 0 Then Exit Sub
  'If etnomrevisador = "" Then MsgBox "No hi ha Revisador escullit, escull-ne un.", vbCritical, "Atenció": Exit Sub
  If Screen.ActiveControl.Name = "imprimirrepas" Then
      If Not guardat Then Command5_Click
    '  imprimir_repas
      If botototok.visible = False Then MsgBox "Atenció no està tot correcte o revisat." + vbNewLine + "POTS IMPRIMIR IGUALMENT PERÒ REVISA-HO.", vbCritical, "ATENCIÓ"
      exportar_carpetatreball vnompdf
      If existeix(vnompdf) Then
         obrir_document vnompdf
      End If
  End If
  'If botototok.visible = True Then passarrepasalasecciodIMPRESORES
  
End Sub
Sub passarrepasalasecciodIMPRESORES(vpassaraFET As Boolean)
   Dim rst As Recordset
   Dim rstclixes As Recordset

   'Set rst = dbclixes.OpenRecordset("select * from clixesentrats_control where format(dataentrada,'dd/mm/yy')=format('" + atrim(Date) + "','dd/mm/yy') and numtreball=" + atrim(cidtreball) + " and versio=" + atrim(cversio))
   Set rst = dbclixes.OpenRecordset("select * from clixesentrats_control where numtreball = " + atrim(cidtreball) + " And versio = " + atrim(cversio) + " order by id desc")
   If Not rst.EOF Then
      rst.Edit
      If IsNull(rst!dataentrada) Then rst!dataentrada = Now
      If IsNull(rst!datarepas) Then rst!datarepas = Now: vavisarREPASADORA = True
      If vpassaraFET Then
           rst!datafet = Now
           rst!numtaula = "T-3"
           dbclixes.Execute "update Clixes set ubicacio='T-3' Where id_treball = " + atrim(cidtreball)
      End If
      rst.Update
      GoTo fi
   End If
   Set rstclixes = dbclixes.OpenRecordset("SELECT Clixes.ubicacio, modificacions.desarroll FROM Clixes RIGHT JOIN modificacions ON Clixes.id_treball = modificacions.id_treball Where modificacions.id_treball = " + atrim(cidtreball) + " And ordre = " + atrim(cversio))
   rst.AddNew
   rst!numtreball = cadbl(cidtreball)
   rst!versio = cadbl(cversio)
   If Not rstclixes.EOF Then rst!desarroll = cadbl(rstclixes!desarroll)
   rst!dataentrada = Now
   rst!datarepas = Now
   rst!numtaula = atrim(rstclixes!ubicacio)
   rst.Update
fi:
   Set rst = Nothing
   Set rstclixes = Nothing
End Sub
Sub exportar_carpetatreball(vnompdf As String)
  Dim fitxerpdftemporal As String
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vcont As Byte
  
  fitxerpdftemporal = ruta_documentacio_clixes + "\" + Format(cadbl(cidtreball), "00000") + "\Arxiu_documentacio_relacionada" + "\v" + atrim(cadbl(cversio))
  creartotalarutadeldirectori fitxerpdftemporal, 12  'poso el dotze per possar un numero gran per començar a crear la ruta mes endavant
  fitxerpdftemporal = fitxerpdftemporal + "\Repasdeclixes.pdf"
  
  'If existeix(fitxerpdftemporal) Then Exit Sub
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "impresiorepasdeclixes.rpt", 1)
  oreport.Database.Tables.item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
  oreport.RecordSelectionFormula = "{repasclixes.id_repas}=" + atrim(cadbl(framedadestintes.tag))
  possarlafirma cadbl(framedadestintes.tag)
  wait 2
  
  oreport.ExportOptions.DiskFileName = fitxerpdftemporal
  oreport.ExportOptions.PDFExportAllPages = True
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  
  oreport.Export False
  vnompdf = fitxerpdftemporal
  While Not existeix(vnompdf) And vcont < 5
    wait 1
    vcont = vcont + 1
  Wend
End Sub
Sub possarlafirma(idrepas As Double)
  Dim rst As Recordset
  Dim rst2 As Recordset
  
  Set rst = dbclixes.OpenRecordset("select numop from repasclixes where id_repas=" + atrim(idrepas))
  If rst.EOF Then Exit Sub
  dbclixes.Execute "delete * from valorsgenerals"
  Set rst2 = dbclixes.OpenRecordset("Select * from valorsgenerals")
  rst2.AddNew
  rst2!codiop = idrepas 'rst!numop
  copiafoto nomfitxerfirma(rst!numop), rst2!imatge
  rst2.Update
  Set rst = Nothing
  Set rst2 = Nothing
End Sub
Function nomfitxerfirma(numop As Long) As String
  Dim vnomfirma As String
  vnomfirma = Dir(rutadelfitxer(llegir_ini("General", "cami", fitxerini)) + "firmes\C" + Format(numop, "00") + "-*.jpg")
  If vnomfirma = "" Then vnomfirma = "sensefirma.jpg"
  nomfitxerfirma = rutadelfitxer(llegir_ini("General", "cami", fitxerini)) + "firmes\" + vnomfirma
End Function
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

Sub creartotalarutadeldirectori(ByVal ruta As String, Optional pos As Integer)
    Dim vruta As String
    If pos = 0 Then pos = 3
    If Right$(ruta, 1) <> "\" Then ruta = ruta + "\"
    vruta = Mid(ruta, 1, pos) + "\"
    While Len(vruta) <> Len(ruta)
        vruta = Mid(ruta, 1, InStr(Len(vruta) + 1, ruta, "\"))
        If InStr(3, vruta, "\") = Len(vruta) Then
          vruta = Mid(ruta, 1, InStr(Len(vruta) + 1, ruta, "\"))
        End If
        
        If Not existeix(vruta) Then MkDir vruta
    Wend
    

End Sub
Public Function Split(ByVal sIn As String, Optional sDelim As _
            String, Optional nLimit As Long = -1, Optional bCompare As _
             VbCompareMethod = vbBinaryCompare) As Variant
          Dim sRead As String, sOut() As String, nC As Integer
          If sDelim = "" Then
              Split = sIn
          End If
          sRead = ReadUntil(sIn, sDelim, bCompare)
          Do
              ReDim Preserve sOut(nC)
              sOut(nC) = sRead
              nC = nC + 1
              If nLimit <> -1 And nC >= nLimit Then Exit Do
              sRead = ReadUntil(sIn, sDelim)
          Loop While sRead <> ""
          ReDim Preserve sOut(nC)
          sOut(nC) = sIn
          Split = sOut
      End Function
 
Public Function ReadUntil(ByRef sIn As String, _
            sDelim As String, Optional bCompare As VbCompareMethod _
          = vbBinaryCompare) As String
          Dim nPos As String
          nPos = InStr(1, sIn, sDelim, bCompare)
          If nPos > 0 Then
              ReadUntil = Left(sIn, nPos - 1)
              sIn = Mid(sIn, nPos + Len(sDelim))
          End If
      End Function

Private Sub lecturarelleu_Click()
   If Screen.ActiveControl.Name = "lecturarelleu" Then
    lecturallisa.Value = 0
    lecturarelleu.Value = 1
    If (lecturarelleu.Value = 1 Or lecturallisa.Value = 1) And escorrectelectura("T") Then
       posar_ok True, 6
      Else: posar_ok False, 6
    End If
   End If
End Sub

Private Sub lecturallisa_Click()
   If Screen.ActiveControl.Name = "lecturallisa" Then
    lecturarelleu.Value = 0
    lecturallisa.Value = 1
    If (lecturarelleu.Value = 1 Or lecturallisa.Value = 1) And escorrectelectura("N") Then
       posar_ok True, 6
      Else: posar_ok False, 6
    End If
   End If
End Sub
Function escorrectelectura(tipuslectura As String) As Boolean
Dim rstmodifi As Recordset
   Set rstmodifi = dbclixes.OpenRecordset("select formaimpresio from modificacions where id_treball=" + atrim(cidtreball) + " and ordre=" + atrim(cversio))
   If rstmodifi.EOF Then Exit Function
   If atrim(rstmodifi!formaimpresio) = atrim(tipuslectura) Then escorrectelectura = True
   Set rstmodifi = Nothing
End Function

Function escorrectebandaseguiment(bandaseguiment As String, gruix As Double) As Boolean
Dim rstmodifi As Recordset
   
   Set rstmodifi = dbclixes.OpenRecordset("select bandaseguiment,amplebandaseguiment from modificacions where id_treball=" + atrim(cidtreball) + " and ordre=" + atrim(cversio))
   If rstmodifi.EOF Then Exit Function
   amplebanda = cadbl(rstmodifi!amplebandaseguiment)
   gruix = cadbl(rstmodifi!amplebandaseguiment)
   If atrim(rstmodifi!bandaseguiment) = atrim(bandaseguiment) And cadbl(rstmodifi!amplebandaseguiment) = gruix Then escorrectebandaseguiment = True
   Set rstmodifi = Nothing
End Function

Private Sub macula_Click()
  If Screen.ActiveControl.Name <> "macula" Then Exit Sub
  If escorrectelamacula(macula) Then
     posar_ok True, 11
    Else: posar_ok False, 11
  End If
End Sub
Function escorrectelamacula(vmacula As String) As Boolean
   Dim rstmodifi As Recordset
   Set rstmodifi = dbclixes.OpenRecordset("select macula from modificacions where id_treball=" + atrim(cidtreball) + " and ordre=" + atrim(cversio))
   If rstmodifi.EOF Then Exit Function
   If atrim(rstmodifi!macula) = atrim(vmacula) Then escorrectelamacula = True
   Set rstmodifi = Nothing
End Function
Function escorrecteelsmicropunts(vmicropunts As Double) As Boolean
   Dim rst As Recordset
   Dim micropuntscorrecte As Double
   micropuntscorrecte = (cadbl(framemotiu(0).tag) * cadbl(motiuamotiu.tag)) * 4
   If vmicropunts = micropuntscorrecte Then escorrecteelsmicropunts = True
End Function




Private Sub macula_GotFocus()
  macula.SelStart = 0
  macula.SelLength = Len(macula)
End Sub

Private Sub micropunts_Click()
  If Screen.ActiveControl.Name = "micropunts" Then
    If micropunts.Value = 1 Then
       posar_ok True, 3
      Else: posar_ok False, 3
    End If
   End If
End Sub

Private Sub midacilindre_Change()
  If Screen.ActiveControl.Name <> "midacilindre" Then Exit Sub
  If escorrectemidadesarroll(cadbl(midadesarroll)) Then
     posar_ok True, 13
    Else: posar_ok False, 13
  End If
End Sub

Private Sub midacilindre_GotFocus()
  midacilindre.SelStart = 0
  midacilindre.SelLength = Len(midacilindre)
End Sub

Private Sub midadesarroll_Change()
  If Screen.ActiveControl.Name <> "midadesarroll" Then Exit Sub
  If escorrectemidadesarroll(cadbl(midadesarroll)) Then
     posar_ok True, 13
    Else: posar_ok False, 13
  End If
End Sub

Private Sub midadesarroll_GotFocus()
  midadesarroll.SelStart = 0
  midadesarroll.SelLength = Len(midadesarroll)
End Sub

Private Sub motiuamotiu_Change()
  If Screen.ActiveControl.Name <> "motiuamotiu" Then Exit Sub
  If escorrectemotiuamotiu(cadbl(motiuamotiu)) Then
     posar_ok True, 9
    Else: posar_ok False, 9
  End If
End Sub
Function escorrectemidadesarroll(vdesarroll As Double) As Boolean
   Dim motiusverticals As Double
   motiusverticals = cadbl(motiuamotiu.tag)
   midacilindre = cadbl(midacilindre.tag)
   numerodemotius = cadbl(numerodemotius.tag)
   
   If cadbl(numerodemotius.tag) = vdesarroll Then
      escorrectemidadesarroll = True
   End If
   numerodemotius.BackColor = QBColor(15)
   midacilindre.BackColor = QBColor(15)
   If cadbl(midacilindre.tag) <> cadbl(midacilindre) Then
      midacilindre.BackColor = QBColor(12)
      escorrectemidadesarroll = False
   End If
   If cadbl(numerodemotius.tag) > 0 Then
      numerodemotius = Int(cadbl(midacilindre.tag) / cadbl(numerodemotius.tag))
      escorrectemidadesarroll = True
       Else: escorrectemidadesarroll = False
   End If
End Function
Function escorrectetotaldelmotiu(vamplemotiu As Double) As Boolean
  If ampladatotalmotiu.Enabled = False Then escorrectetotaldelmotiu = True: Exit Function
  If vamplemotiu > (cadbl(ampladatotalmotiu.tag) * 10) Or vamplemotiu < 1 Then
     ampladatotalmotiu.BackColor = QBColor(12)
     escorrectetotaldelmotiu = False
       Else:
         ampladatotalmotiu.BackColor = QBColor(15)
         escorrectetotaldelmotiu = True
  End If

End Function
Function escorrectemotiuamotiu(vmotiuamotiu As Double) As Boolean
   Dim motiusverticals As Double
   motiusverticals = cadbl(motiuamotiu.tag)
   et_resposta_motiuamotiu = ""
   If motiusverticals = 0 Then Exit Function
   If distorsio = 0 Then MsgBox "No hi ha cap distrosio escullida." + vbNewLine + "ASSEGURA QUE HI HA UN GRUIX DE POLIMER ESCULLIT.", vbCritical, "Atenció": Exit Function
   If Redondejar(motiuamotiu.HelpContextID - (distorsio / motiusverticals), 0) = vmotiuamotiu Then
      escorrectemotiuamotiu = True
     Else: et_resposta_motiuamotiu = "" 'atrim(Redondejar(motiuamotiu.HelpContextID - (distorsio / motiusverticals), 0))
   End If
   
End Function
Function distorsio() As Double
   distorsio = 0
   If gruixplimer(0).Value Then distorsio = 6
   If gruixplimer(1).Value Then distorsio = 15
   If gruixplimer(2).Value Then distorsio = 18
End Function

Private Sub motiuamotiu_GotFocus()
  motiuamotiu.SelStart = 0
  motiuamotiu.SelLength = Len(motiuamotiu)
End Sub

Private Sub motiuamotiu_LostFocus()
  et_resposta_motiuamotiu = ""
End Sub

Private Sub numerodemotius_Change()
  If Screen.ActiveControl.Name <> "numerodemotius" Then Exit Sub
  If escorrectemidadesarroll(cadbl(midadesarroll)) Then
     posar_ok True, 13
    Else: posar_ok False, 13
  End If
End Sub

Private Sub numerodemotius_GotFocus()
  numerodemotius.SelStart = 0
  numerodemotius.SelLength = Len(numerodemotius)
End Sub

Private Sub repasclixe_Click()
 If Screen.ActiveControl.Name = "repasclixe" Then
    If repasclixe.Value = 1 Then
       posar_ok True, 2
      Else: posar_ok False, 2
    End If
   End If
End Sub
