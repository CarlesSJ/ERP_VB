VERSION 5.00
Begin VB.Form formaniloxos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manteniment Aniloxos i Densitats"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14475
   Icon            =   "aniloxosutilitzats.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   14475
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cpostit 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   5910
      Locked          =   -1  'True
      MouseIcon       =   "aniloxosutilitzats.frx":058A
      TabIndex        =   233
      Top             =   1725
      Visible         =   0   'False
      Width           =   5430
   End
   Begin VB.Frame framecandau 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   8715
      TabIndex        =   232
      Top             =   990
      Width           =   360
      Begin VB.Image fotocandau 
         Height          =   375
         Left            =   15
         Picture         =   "aniloxosutilitzats.frx":0B14
         Stretch         =   -1  'True
         Top             =   15
         Width           =   360
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F3B378&
      Caption         =   "Consums Tinta"
      Height          =   300
      Left            =   12195
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   630
      Width           =   1275
   End
   Begin VB.Frame frameconsums 
      BackColor       =   &H00F3B378&
      Caption         =   "Consums tinta"
      Height          =   4050
      Left            =   10560
      TabIndex        =   32
      Top             =   1000
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton bborrarlot 
         Height          =   255
         Left            =   1440
         Picture         =   "aniloxosutilitzats.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Borrar el lot d'aquesta tinta"
         Top             =   420
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton boto_nou 
         Appearance      =   0  'Flat
         BackColor       =   &H00F3B378&
         Height          =   315
         Index           =   0
         Left            =   2310
         Picture         =   "aniloxosutilitzats.frx":118A
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Afegir llaunes i lots"
         Top             =   135
         Width           =   555
      End
      Begin VB.TextBox kbpantone 
         BackColor       =   &H00C0E0FF&
         DataField       =   "kg10"
         Height          =   285
         Index           =   9
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   54
         Tag             =   "1"
         Top             =   3690
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00C0C0FF&
         DataField       =   "lot10"
         Height          =   285
         Index           =   9
         Left            =   615
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   53
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   3690
         Width           =   1110
      End
      Begin VB.TextBox kbpantone 
         BackColor       =   &H00C0E0FF&
         DataField       =   "kg9"
         Height          =   285
         Index           =   8
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   52
         Tag             =   "1"
         Top             =   3345
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00C0C0FF&
         DataField       =   "lot9"
         Height          =   285
         Index           =   8
         Left            =   630
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   51
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   3345
         Width           =   1095
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg8"
         Height          =   285
         Index           =   7
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   50
         Tag             =   "1"
         Top             =   2925
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot8"
         Height          =   285
         Index           =   7
         Left            =   75
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   49
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   2925
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg7"
         Height          =   285
         Index           =   6
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   48
         Tag             =   "1"
         Top             =   2565
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot7"
         Height          =   285
         Index           =   6
         Left            =   75
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   47
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   2565
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg6"
         Height          =   285
         Index           =   5
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   46
         Tag             =   "1"
         Top             =   2205
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot6"
         Height          =   285
         Index           =   5
         Left            =   75
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   45
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   2205
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg5"
         Height          =   285
         Index           =   4
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   44
         Tag             =   "1"
         Top             =   1815
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot5"
         Height          =   285
         Index           =   4
         Left            =   75
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   43
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   1815
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg4"
         Height          =   285
         Index           =   3
         Left            =   1725
         MaxLength       =   8
         TabIndex        =   42
         Tag             =   "1"
         Top             =   1440
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot4"
         Height          =   285
         Index           =   3
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   41
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   1440
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg3"
         Height          =   285
         Index           =   2
         Left            =   1725
         MaxLength       =   8
         TabIndex        =   40
         Tag             =   "1"
         Top             =   1080
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot3"
         Height          =   285
         Index           =   2
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   1080
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg2"
         Height          =   285
         Index           =   1
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   38
         Tag             =   "1"
         Top             =   750
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot2"
         Height          =   285
         Index           =   1
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   37
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   750
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg1"
         Height          =   285
         Index           =   0
         Left            =   1725
         MaxLength       =   8
         TabIndex        =   36
         Tag             =   "1"
         Top             =   390
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot1"
         Height          =   285
         Index           =   0
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   35
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   390
         Width           =   1680
      End
      Begin VB.Label etkgteorics 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   7
         Left            =   2325
         TabIndex        =   68
         ToolTipText     =   "Kg teorics que es gastaran."
         Top             =   2940
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label etkgteorics 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   6
         Left            =   2325
         TabIndex        =   67
         ToolTipText     =   "Kg teorics que es gastaran."
         Top             =   2577
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label etkgteorics 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   5
         Left            =   2325
         TabIndex        =   66
         ToolTipText     =   "Kg teorics que es gastaran."
         Top             =   2220
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label etkgteorics 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   4
         Left            =   2325
         TabIndex        =   65
         ToolTipText     =   "Kg teorics que es gastaran."
         Top             =   1863
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label etkgteorics 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   3
         Left            =   2325
         TabIndex        =   64
         ToolTipText     =   "Kg teorics que es gastaran."
         Top             =   1506
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label etkgteorics 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   2
         Left            =   2325
         TabIndex        =   63
         ToolTipText     =   "Kg teorics que es gastaran."
         Top             =   1149
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label etkgteorics 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   1
         Left            =   2325
         TabIndex        =   62
         ToolTipText     =   "Kg teorics que es gastaran."
         Top             =   792
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label etkgteorics 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   0
         Left            =   2325
         TabIndex        =   61
         ToolTipText     =   "Kg teorics que es gastaran."
         Top             =   435
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "80/20"
         Height          =   240
         Left            =   90
         TabIndex        =   57
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   75
         TabIndex        =   56
         Top             =   3390
         Width           =   675
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg consum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1470
         TabIndex        =   34
         Top             =   195
         Width           =   915
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Lots"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   33
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.CommandButton bconsumsr25 
      Height          =   330
      Left            =   10665
      Picture         =   "aniloxosutilitzats.frx":1714
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4635
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton bconsumetoxi 
      Height          =   330
      Left            =   10680
      Picture         =   "aniloxosutilitzats.frx":1C9E
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4290
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CheckBox checkeditar 
      Caption         =   "Editar"
      Height          =   195
      Left            =   13590
      TabIndex        =   25
      Top             =   435
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   30
      TabIndex        =   20
      Top             =   -30
      Width           =   13485
      Begin VB.CommandButton bveurepdf 
         Height          =   525
         Left            =   12900
         Picture         =   "aniloxosutilitzats.frx":2228
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Veure el PDF"
         Top             =   75
         Width           =   555
      End
      Begin VB.Label etqualitat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Molt Bona"
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
         Height          =   225
         Left            =   11520
         TabIndex        =   31
         Top             =   345
         Width           =   1125
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Resultat Impresió"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11490
         TabIndex        =   30
         Top             =   150
         Width           =   1470
      End
      Begin VB.Label ettreball 
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
         Height          =   240
         Left            =   45
         TabIndex        =   24
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label etlinia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1410
         TabIndex        =   23
         Top             =   390
         Width           =   7650
      End
      Begin VB.Label etcomanda 
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
         Height          =   240
         Left            =   45
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label etclient 
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
         Left            =   1425
         TabIndex        =   21
         Top             =   120
         Width           =   7665
      End
   End
   Begin VB.Frame fbotonsok 
      Height          =   3450
      Left            =   13650
      TabIndex        =   0
      Top             =   780
      Width           =   690
      Begin VB.CommandButton bnotots 
         BackColor       =   &H00C0C0FF&
         Height          =   300
         Left            =   330
         Picture         =   "aniloxosutilitzats.frx":2D32
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Marcar tots no"
         Top             =   255
         Width           =   315
      End
      Begin VB.CommandButton boktots 
         BackColor       =   &H00C0FFC0&
         Height          =   300
         Left            =   30
         Picture         =   "aniloxosutilitzats.frx":32BC
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Marcar tots Si"
         Top             =   255
         Width           =   315
      End
      Begin VB.CommandButton bno 
         Height          =   300
         Index           =   7
         Left            =   345
         Picture         =   "aniloxosutilitzats.frx":3846
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Anular canvis"
         Top             =   3090
         Width           =   300
      End
      Begin VB.CommandButton bno 
         Height          =   300
         Index           =   6
         Left            =   345
         Picture         =   "aniloxosutilitzats.frx":3DD0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Anular canvis"
         Top             =   2727
         Width           =   300
      End
      Begin VB.CommandButton bno 
         Height          =   300
         Index           =   5
         Left            =   345
         Picture         =   "aniloxosutilitzats.frx":435A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Anular canvis"
         Top             =   2365
         Width           =   300
      End
      Begin VB.CommandButton bno 
         Height          =   300
         Index           =   4
         Left            =   345
         Picture         =   "aniloxosutilitzats.frx":48E4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Anular canvis"
         Top             =   2003
         Width           =   300
      End
      Begin VB.CommandButton bno 
         Height          =   300
         Index           =   3
         Left            =   345
         Picture         =   "aniloxosutilitzats.frx":4E6E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Anular canvis"
         Top             =   1641
         Width           =   300
      End
      Begin VB.CommandButton bno 
         Height          =   300
         Index           =   2
         Left            =   345
         Picture         =   "aniloxosutilitzats.frx":53F8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Anular canvis"
         Top             =   1279
         Width           =   300
      End
      Begin VB.CommandButton bno 
         Height          =   300
         Index           =   1
         Left            =   345
         Picture         =   "aniloxosutilitzats.frx":5982
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Anular canvis"
         Top             =   917
         Width           =   300
      End
      Begin VB.CommandButton bok 
         Height          =   300
         Index           =   7
         Left            =   30
         Picture         =   "aniloxosutilitzats.frx":5F0C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Acceptar canvis"
         Top             =   3090
         Width           =   300
      End
      Begin VB.CommandButton bok 
         Height          =   300
         Index           =   6
         Left            =   30
         Picture         =   "aniloxosutilitzats.frx":6496
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Acceptar canvis"
         Top             =   2727
         Width           =   300
      End
      Begin VB.CommandButton bok 
         Height          =   300
         Index           =   5
         Left            =   30
         Picture         =   "aniloxosutilitzats.frx":6A20
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Acceptar canvis"
         Top             =   2365
         Width           =   300
      End
      Begin VB.CommandButton bok 
         Height          =   300
         Index           =   4
         Left            =   30
         Picture         =   "aniloxosutilitzats.frx":6FAA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Acceptar canvis"
         Top             =   2003
         Width           =   300
      End
      Begin VB.CommandButton bok 
         Height          =   300
         Index           =   3
         Left            =   30
         Picture         =   "aniloxosutilitzats.frx":7534
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Acceptar canvis"
         Top             =   1641
         Width           =   300
      End
      Begin VB.CommandButton bok 
         Height          =   300
         Index           =   2
         Left            =   30
         Picture         =   "aniloxosutilitzats.frx":7ABE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Acceptar canvis"
         Top             =   1279
         Width           =   300
      End
      Begin VB.CommandButton bok 
         Height          =   300
         Index           =   1
         Left            =   30
         Picture         =   "aniloxosutilitzats.frx":8048
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Acceptar canvis"
         Top             =   917
         Width           =   300
      End
      Begin VB.CommandButton bno 
         Height          =   300
         Index           =   0
         Left            =   345
         Picture         =   "aniloxosutilitzats.frx":85D2
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Anular canvis"
         Top             =   555
         Width           =   300
      End
      Begin VB.CommandButton bok 
         Height          =   300
         Index           =   0
         Left            =   30
         Picture         =   "aniloxosutilitzats.frx":8B5C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Acceptar canvis"
         Top             =   555
         Width           =   300
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tots"
         Height          =   225
         Left            =   165
         TabIndex        =   18
         Top             =   90
         Width           =   435
      End
   End
   Begin VB.Frame frametinters 
      Enabled         =   0   'False
      Height          =   3855
      Left            =   30
      TabIndex        =   69
      Top             =   525
      Width           =   13530
      Begin VB.Frame Frame1 
         Height          =   3495
         Left            =   60
         TabIndex        =   70
         Top             =   315
         Width           =   13335
         Begin VB.TextBox densitat 
            BackColor       =   &H00C0C0C0&
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
            Left            =   4575
            Locked          =   -1  'True
            TabIndex        =   206
            Top             =   3045
            Width           =   380
         End
         Begin VB.TextBox densitat 
            BackColor       =   &H00C0C0C0&
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
            Index           =   6
            Left            =   4575
            Locked          =   -1  'True
            TabIndex        =   205
            Top             =   2685
            Width           =   380
         End
         Begin VB.TextBox densitat 
            BackColor       =   &H00C0C0C0&
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
            Index           =   5
            Left            =   4575
            Locked          =   -1  'True
            TabIndex        =   204
            Top             =   2325
            Width           =   380
         End
         Begin VB.TextBox densitat 
            BackColor       =   &H00C0C0C0&
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
            Index           =   4
            Left            =   4575
            Locked          =   -1  'True
            TabIndex        =   203
            Top             =   1965
            Width           =   380
         End
         Begin VB.TextBox densitat 
            BackColor       =   &H00C0C0C0&
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
            Index           =   3
            Left            =   4575
            Locked          =   -1  'True
            TabIndex        =   202
            Top             =   1590
            Width           =   380
         End
         Begin VB.TextBox densitat 
            BackColor       =   &H00C0C0C0&
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
            Index           =   2
            Left            =   4575
            Locked          =   -1  'True
            TabIndex        =   201
            Top             =   1230
            Width           =   380
         End
         Begin VB.TextBox densitat 
            BackColor       =   &H00C0C0C0&
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
            Index           =   1
            Left            =   4575
            Locked          =   -1  'True
            TabIndex        =   200
            Top             =   870
            Width           =   380
         End
         Begin VB.TextBox densitat 
            BackColor       =   &H00C0C0C0&
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
            Index           =   0
            Left            =   4575
            Locked          =   -1  'True
            TabIndex        =   199
            Top             =   510
            Width           =   380
         End
         Begin VB.TextBox anilox 
            BackColor       =   &H00C0C0C0&
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
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   198
            Top             =   3045
            Width           =   380
         End
         Begin VB.TextBox anilox 
            BackColor       =   &H00C0C0C0&
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
            Index           =   6
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   197
            Top             =   2685
            Width           =   380
         End
         Begin VB.TextBox anilox 
            BackColor       =   &H00C0C0C0&
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
            Index           =   5
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   196
            Top             =   2325
            Width           =   380
         End
         Begin VB.TextBox anilox 
            BackColor       =   &H00C0C0C0&
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
            Index           =   4
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   195
            Top             =   1965
            Width           =   380
         End
         Begin VB.TextBox anilox 
            BackColor       =   &H00C0C0C0&
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
            Index           =   3
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   194
            Top             =   1590
            Width           =   380
         End
         Begin VB.TextBox anilox 
            BackColor       =   &H00C0C0C0&
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
            Index           =   2
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   193
            Top             =   1230
            Width           =   380
         End
         Begin VB.TextBox anilox 
            BackColor       =   &H00C0C0C0&
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
            Index           =   1
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   192
            Top             =   870
            Width           =   380
         End
         Begin VB.TextBox anilox 
            BackColor       =   &H00C0C0C0&
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
            Index           =   0
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   191
            Top             =   510
            Width           =   380
         End
         Begin VB.TextBox color 
            BackColor       =   &H00C0C0C0&
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
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   190
            Top             =   3045
            Width           =   2685
         End
         Begin VB.TextBox color 
            BackColor       =   &H00C0C0C0&
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
            Index           =   6
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   189
            Top             =   2685
            Width           =   2685
         End
         Begin VB.TextBox color 
            BackColor       =   &H00C0C0C0&
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
            Index           =   5
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   188
            Top             =   2325
            Width           =   2685
         End
         Begin VB.TextBox color 
            BackColor       =   &H00C0C0C0&
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
            Index           =   4
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   187
            Top             =   1965
            Width           =   2685
         End
         Begin VB.TextBox color 
            BackColor       =   &H00C0C0C0&
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
            Index           =   3
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   186
            Top             =   1590
            Width           =   2685
         End
         Begin VB.TextBox color 
            BackColor       =   &H00C0C0C0&
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
            Index           =   2
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   185
            Top             =   1230
            Width           =   2685
         End
         Begin VB.TextBox color 
            BackColor       =   &H00C0C0C0&
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
            Index           =   1
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   184
            Top             =   870
            Width           =   2685
         End
         Begin VB.TextBox color 
            BackColor       =   &H00C0C0C0&
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
            Index           =   0
            Left            =   615
            Locked          =   -1  'True
            TabIndex        =   183
            Top             =   510
            Width           =   2685
         End
         Begin VB.TextBox ordre 
            Height          =   315
            Index           =   7
            Left            =   240
            TabIndex        =   182
            Top             =   3045
            Width           =   360
         End
         Begin VB.TextBox ordre 
            Height          =   315
            Index           =   6
            Left            =   240
            TabIndex        =   181
            Top             =   2685
            Width           =   360
         End
         Begin VB.TextBox ordre 
            Height          =   315
            Index           =   5
            Left            =   240
            TabIndex        =   180
            Top             =   2310
            Width           =   360
         End
         Begin VB.TextBox ordre 
            Height          =   315
            Index           =   4
            Left            =   240
            TabIndex        =   179
            Top             =   1950
            Width           =   360
         End
         Begin VB.TextBox ordre 
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   178
            Top             =   1590
            Width           =   360
         End
         Begin VB.TextBox ordre 
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   177
            Top             =   1230
            Width           =   360
         End
         Begin VB.TextBox ordre 
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   176
            Top             =   855
            Width           =   360
         End
         Begin VB.TextBox ordre 
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   175
            Top             =   495
            Width           =   360
         End
         Begin VB.TextBox aniloxcomanda 
            Height          =   330
            Index           =   0
            Left            =   10920
            TabIndex        =   174
            Top             =   525
            Width           =   495
         End
         Begin VB.TextBox densitatcomanda 
            Height          =   330
            Index           =   0
            Left            =   12300
            TabIndex        =   173
            Top             =   525
            Width           =   495
         End
         Begin VB.TextBox aniloxcomanda 
            Height          =   330
            Index           =   1
            Left            =   10920
            TabIndex        =   172
            Top             =   885
            Width           =   495
         End
         Begin VB.TextBox aniloxcomanda 
            Height          =   330
            Index           =   2
            Left            =   10920
            TabIndex        =   171
            Top             =   1245
            Width           =   495
         End
         Begin VB.TextBox aniloxcomanda 
            Height          =   330
            Index           =   3
            Left            =   10920
            TabIndex        =   170
            Top             =   1605
            Width           =   495
         End
         Begin VB.TextBox aniloxcomanda 
            Height          =   330
            Index           =   4
            Left            =   10920
            TabIndex        =   169
            Top             =   1980
            Width           =   495
         End
         Begin VB.TextBox aniloxcomanda 
            Height          =   330
            Index           =   5
            Left            =   10920
            TabIndex        =   168
            Top             =   2340
            Width           =   495
         End
         Begin VB.TextBox aniloxcomanda 
            Height          =   330
            Index           =   6
            Left            =   10920
            TabIndex        =   167
            Top             =   2700
            Width           =   495
         End
         Begin VB.TextBox aniloxcomanda 
            Height          =   330
            Index           =   7
            Left            =   10920
            TabIndex        =   166
            Top             =   3060
            Width           =   495
         End
         Begin VB.TextBox aniloxfotogravador 
            BackColor       =   &H00C0C0C0&
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
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   165
            Top             =   3045
            Width           =   380
         End
         Begin VB.TextBox aniloxfotogravador 
            BackColor       =   &H00C0C0C0&
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
            Index           =   6
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   164
            Top             =   2685
            Width           =   380
         End
         Begin VB.TextBox aniloxfotogravador 
            BackColor       =   &H00C0C0C0&
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
            Index           =   5
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   163
            Top             =   2325
            Width           =   380
         End
         Begin VB.TextBox aniloxfotogravador 
            BackColor       =   &H00C0C0C0&
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
            Index           =   4
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   162
            Top             =   1965
            Width           =   380
         End
         Begin VB.TextBox aniloxfotogravador 
            BackColor       =   &H00C0C0C0&
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
            Index           =   3
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   161
            Top             =   1590
            Width           =   380
         End
         Begin VB.TextBox aniloxfotogravador 
            BackColor       =   &H00C0C0C0&
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
            Index           =   2
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   160
            Top             =   1230
            Width           =   380
         End
         Begin VB.TextBox aniloxfotogravador 
            BackColor       =   &H00C0C0C0&
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
            Index           =   1
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   159
            Top             =   870
            Width           =   380
         End
         Begin VB.TextBox aniloxfotogravador 
            BackColor       =   &H00C0C0C0&
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
            Index           =   0
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   158
            Top             =   510
            Width           =   380
         End
         Begin VB.ComboBox tintacomanda 
            Height          =   315
            Index           =   0
            Left            =   5820
            TabIndex        =   157
            Top             =   555
            Width           =   3660
         End
         Begin VB.ComboBox tintacomanda 
            Height          =   315
            Index           =   1
            Left            =   5820
            TabIndex        =   156
            Top             =   915
            Width           =   3660
         End
         Begin VB.ComboBox tintacomanda 
            Height          =   315
            Index           =   2
            Left            =   5820
            TabIndex        =   155
            Top             =   1275
            Width           =   3660
         End
         Begin VB.ComboBox tintacomanda 
            Height          =   315
            Index           =   3
            Left            =   5820
            TabIndex        =   154
            Top             =   1635
            Width           =   3660
         End
         Begin VB.ComboBox tintacomanda 
            Height          =   315
            Index           =   4
            Left            =   5820
            TabIndex        =   153
            Top             =   2010
            Width           =   3660
         End
         Begin VB.ComboBox tintacomanda 
            Height          =   315
            Index           =   5
            Left            =   5820
            TabIndex        =   152
            Top             =   2370
            Width           =   3660
         End
         Begin VB.ComboBox tintacomanda 
            Height          =   315
            Index           =   6
            Left            =   5820
            TabIndex        =   151
            Top             =   2730
            Width           =   3660
         End
         Begin VB.ComboBox tintacomanda 
            Height          =   315
            Index           =   7
            Left            =   5820
            TabIndex        =   150
            Top             =   3090
            Width           =   3660
         End
         Begin VB.CommandButton bconsums 
            Height          =   330
            Index           =   1
            Left            =   12915
            Picture         =   "aniloxosutilitzats.frx":90E6
            Style           =   1  'Graphical
            TabIndex        =   149
            ToolTipText     =   "Consums de tinta (Llaunes)"
            Top             =   872
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.CommandButton bconsums 
            Height          =   330
            Index           =   2
            Left            =   12915
            Picture         =   "aniloxosutilitzats.frx":9670
            Style           =   1  'Graphical
            TabIndex        =   148
            ToolTipText     =   "Consums de tinta (Llaunes)"
            Top             =   1234
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.CommandButton bconsums 
            Height          =   330
            Index           =   3
            Left            =   12915
            Picture         =   "aniloxosutilitzats.frx":9BFA
            Style           =   1  'Graphical
            TabIndex        =   147
            ToolTipText     =   "Consums de tinta (Llaunes)"
            Top             =   1596
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.CommandButton bconsums 
            Height          =   330
            Index           =   4
            Left            =   12915
            Picture         =   "aniloxosutilitzats.frx":A184
            Style           =   1  'Graphical
            TabIndex        =   146
            ToolTipText     =   "Consums de tinta (Llaunes)"
            Top             =   1958
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.CommandButton bconsums 
            Height          =   330
            Index           =   5
            Left            =   12915
            Picture         =   "aniloxosutilitzats.frx":A70E
            Style           =   1  'Graphical
            TabIndex        =   145
            ToolTipText     =   "Consums de tinta (Llaunes)"
            Top             =   2320
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.CommandButton bconsums 
            Height          =   330
            Index           =   6
            Left            =   12915
            Picture         =   "aniloxosutilitzats.frx":AC98
            Style           =   1  'Graphical
            TabIndex        =   144
            ToolTipText     =   "Consums de tinta (Llaunes)"
            Top             =   2682
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.CommandButton bconsums 
            Height          =   330
            Index           =   7
            Left            =   12915
            Picture         =   "aniloxosutilitzats.frx":B222
            Style           =   1  'Graphical
            TabIndex        =   143
            ToolTipText     =   "Consums de tinta (Llaunes)"
            Top             =   3045
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.ComboBox detalltinter 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Index           =   0
            ItemData        =   "aniloxosutilitzats.frx":B7AC
            Left            =   9540
            List            =   "aniloxosutilitzats.frx":B7AE
            Locked          =   -1  'True
            TabIndex        =   142
            Top             =   540
            Width           =   1365
         End
         Begin VB.ComboBox detalltinter 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Index           =   1
            Left            =   9540
            Locked          =   -1  'True
            TabIndex        =   141
            Top             =   900
            Width           =   1365
         End
         Begin VB.ComboBox detalltinter 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Index           =   2
            Left            =   9540
            Locked          =   -1  'True
            TabIndex        =   140
            Top             =   1260
            Width           =   1365
         End
         Begin VB.ComboBox detalltinter 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Index           =   3
            Left            =   9540
            Locked          =   -1  'True
            TabIndex        =   139
            Top             =   1620
            Width           =   1365
         End
         Begin VB.ComboBox detalltinter 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Index           =   4
            Left            =   9540
            Locked          =   -1  'True
            TabIndex        =   138
            Top             =   1995
            Width           =   1365
         End
         Begin VB.ComboBox detalltinter 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Index           =   5
            Left            =   9540
            Locked          =   -1  'True
            TabIndex        =   137
            Top             =   2355
            Width           =   1365
         End
         Begin VB.ComboBox detalltinter 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Index           =   6
            Left            =   9540
            Locked          =   -1  'True
            TabIndex        =   136
            Top             =   2715
            Width           =   1365
         End
         Begin VB.ComboBox detalltinter 
            BackColor       =   &H0080C0FF&
            Height          =   315
            Index           =   7
            Left            =   9540
            Locked          =   -1  'True
            TabIndex        =   135
            Top             =   3075
            Width           =   1365
         End
         Begin VB.TextBox cdetall 
            BackColor       =   &H0080C0FF&
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
            Index           =   0
            Left            =   1905
            TabIndex        =   134
            Top             =   510
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox cdetall 
            BackColor       =   &H0080C0FF&
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
            Index           =   1
            Left            =   1905
            TabIndex        =   133
            Top             =   872
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox cdetall 
            BackColor       =   &H0080C0FF&
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
            Index           =   2
            Left            =   1905
            TabIndex        =   132
            Top             =   1234
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox cdetall 
            BackColor       =   &H0080C0FF&
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
            Index           =   3
            Left            =   1905
            TabIndex        =   131
            Top             =   1596
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox cdetall 
            BackColor       =   &H0080C0FF&
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
            Index           =   4
            Left            =   1905
            TabIndex        =   130
            Top             =   1958
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox cdetall 
            BackColor       =   &H0080C0FF&
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
            Index           =   5
            Left            =   1905
            TabIndex        =   129
            Top             =   2320
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox cdetall 
            BackColor       =   &H0080C0FF&
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
            Index           =   6
            Left            =   1905
            TabIndex        =   128
            Top             =   2682
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox cdetall 
            BackColor       =   &H0080C0FF&
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
            Left            =   1905
            TabIndex        =   127
            Top             =   3045
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox viscositatcomanda 
            Height          =   330
            Index           =   0
            Left            =   11850
            TabIndex        =   126
            Top             =   525
            Width           =   405
         End
         Begin VB.TextBox viscositatcomanda 
            Height          =   330
            Index           =   1
            Left            =   11850
            TabIndex        =   125
            Top             =   885
            Width           =   405
         End
         Begin VB.TextBox viscositatcomanda 
            Height          =   330
            Index           =   2
            Left            =   11850
            TabIndex        =   124
            Top             =   1245
            Width           =   405
         End
         Begin VB.TextBox viscositatcomanda 
            Height          =   330
            Index           =   3
            Left            =   11850
            TabIndex        =   123
            Top             =   1605
            Width           =   405
         End
         Begin VB.TextBox viscositatcomanda 
            Height          =   330
            Index           =   4
            Left            =   11850
            TabIndex        =   122
            Top             =   1965
            Width           =   405
         End
         Begin VB.TextBox viscositatcomanda 
            Height          =   330
            Index           =   5
            Left            =   11850
            TabIndex        =   121
            Top             =   2325
            Width           =   405
         End
         Begin VB.TextBox viscositatcomanda 
            Height          =   330
            Index           =   6
            Left            =   11850
            TabIndex        =   120
            Top             =   2685
            Width           =   405
         End
         Begin VB.TextBox viscositatcomanda 
            Height          =   330
            Index           =   7
            Left            =   11850
            TabIndex        =   119
            Top             =   3045
            Width           =   405
         End
         Begin VB.TextBox volumcomanda 
            Height          =   330
            Index           =   0
            Left            =   11460
            TabIndex        =   118
            Top             =   525
            Width           =   360
         End
         Begin VB.TextBox volumcomanda 
            Height          =   330
            Index           =   1
            Left            =   11460
            TabIndex        =   117
            Top             =   915
            Width           =   360
         End
         Begin VB.TextBox volumcomanda 
            Height          =   330
            Index           =   2
            Left            =   11460
            TabIndex        =   116
            Top             =   1260
            Width           =   360
         End
         Begin VB.TextBox volumcomanda 
            Height          =   330
            Index           =   3
            Left            =   11460
            TabIndex        =   115
            Top             =   1620
            Width           =   360
         End
         Begin VB.TextBox volumcomanda 
            Height          =   330
            Index           =   4
            Left            =   11460
            TabIndex        =   114
            Top             =   1980
            Width           =   360
         End
         Begin VB.TextBox volumcomanda 
            Height          =   330
            Index           =   5
            Left            =   11460
            TabIndex        =   113
            Top             =   2325
            Width           =   360
         End
         Begin VB.TextBox volumcomanda 
            Height          =   330
            Index           =   6
            Left            =   11460
            TabIndex        =   112
            Top             =   2685
            Width           =   360
         End
         Begin VB.TextBox volumcomanda 
            Height          =   330
            Index           =   7
            Left            =   11460
            TabIndex        =   111
            Top             =   3045
            Width           =   360
         End
         Begin VB.TextBox observacions 
            Height          =   315
            Index           =   7
            Left            =   12720
            MaxLength       =   30
            TabIndex        =   110
            Top             =   3045
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox observacions 
            Height          =   315
            Index           =   6
            Left            =   12720
            MaxLength       =   30
            TabIndex        =   109
            Top             =   2685
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox observacions 
            Height          =   315
            Index           =   5
            Left            =   12720
            MaxLength       =   30
            TabIndex        =   108
            Top             =   2325
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox observacions 
            Height          =   315
            Index           =   4
            Left            =   12720
            MaxLength       =   30
            TabIndex        =   107
            Top             =   1965
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox observacions 
            Height          =   315
            Index           =   3
            Left            =   12720
            MaxLength       =   30
            TabIndex        =   106
            Top             =   1605
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox observacions 
            Height          =   315
            Index           =   2
            Left            =   12720
            MaxLength       =   30
            TabIndex        =   105
            Top             =   1245
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox observacions 
            Height          =   315
            Index           =   1
            Left            =   12720
            MaxLength       =   30
            TabIndex        =   104
            Top             =   885
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.TextBox volum 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   0
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   103
            Top             =   510
            Width           =   380
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   0
            Left            =   4155
            Locked          =   -1  'True
            TabIndex        =   102
            Top             =   510
            Width           =   380
         End
         Begin VB.TextBox volum 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   1
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   101
            Top             =   872
            Width           =   380
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   1
            Left            =   4155
            Locked          =   -1  'True
            TabIndex        =   100
            Top             =   872
            Width           =   380
         End
         Begin VB.TextBox volum 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   2
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   99
            Top             =   1234
            Width           =   380
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   2
            Left            =   4155
            Locked          =   -1  'True
            TabIndex        =   98
            Top             =   1234
            Width           =   380
         End
         Begin VB.TextBox volum 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   3
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   97
            Top             =   1596
            Width           =   380
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   3
            Left            =   4155
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   1596
            Width           =   380
         End
         Begin VB.TextBox volum 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   4
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   95
            Top             =   1958
            Width           =   380
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   4
            Left            =   4155
            Locked          =   -1  'True
            TabIndex        =   94
            Top             =   1958
            Width           =   380
         End
         Begin VB.TextBox volum 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   5
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   93
            Top             =   2320
            Width           =   380
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   5
            Left            =   4155
            Locked          =   -1  'True
            TabIndex        =   92
            Top             =   2320
            Width           =   380
         End
         Begin VB.TextBox volum 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   6
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   91
            Top             =   2682
            Width           =   380
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   6
            Left            =   4155
            Locked          =   -1  'True
            TabIndex        =   90
            Top             =   2682
            Width           =   380
         End
         Begin VB.TextBox volum 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   7
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   89
            Top             =   3045
            Width           =   380
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Index           =   7
            Left            =   4155
            Locked          =   -1  'True
            TabIndex        =   88
            Top             =   3045
            Width           =   380
         End
         Begin VB.CommandButton bconsums 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Index           =   0
            Left            =   12915
            Picture         =   "aniloxosutilitzats.frx":B7B0
            Style           =   1  'Graphical
            TabIndex        =   87
            ToolTipText     =   "Consums de tinta (Llaunes)"
            Top             =   510
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.TextBox observacions 
            Height          =   315
            Index           =   0
            Left            =   12750
            MaxLength       =   50
            TabIndex        =   86
            Top             =   525
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox densitatcomanda 
            Height          =   330
            Index           =   7
            Left            =   12300
            TabIndex        =   85
            Top             =   3045
            Width           =   495
         End
         Begin VB.TextBox densitatcomanda 
            Height          =   330
            Index           =   6
            Left            =   12300
            TabIndex        =   84
            Top             =   2685
            Width           =   495
         End
         Begin VB.TextBox densitatcomanda 
            Height          =   330
            Index           =   5
            Left            =   12300
            TabIndex        =   83
            Top             =   2325
            Width           =   495
         End
         Begin VB.TextBox densitatcomanda 
            Height          =   330
            Index           =   4
            Left            =   12300
            TabIndex        =   82
            Top             =   1965
            Width           =   495
         End
         Begin VB.TextBox densitatcomanda 
            Height          =   330
            Index           =   3
            Left            =   12300
            TabIndex        =   81
            Top             =   1605
            Width           =   495
         End
         Begin VB.TextBox densitatcomanda 
            Height          =   330
            Index           =   2
            Left            =   12300
            TabIndex        =   80
            Top             =   1245
            Width           =   495
         End
         Begin VB.TextBox densitatcomanda 
            Height          =   330
            Index           =   1
            Left            =   12300
            TabIndex        =   79
            Top             =   885
            Width           =   495
         End
         Begin VB.TextBox numextensio 
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   8265
            TabIndex        =   78
            Text            =   "Ext:1234/5"
            Top             =   570
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox numextensio 
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   8265
            TabIndex        =   77
            Text            =   "Ext:1234/5"
            Top             =   915
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox numextensio 
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   8265
            TabIndex        =   76
            Text            =   "Ext:1234/5"
            Top             =   1275
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox numextensio 
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   8265
            TabIndex        =   75
            Text            =   "Ext:1234/5"
            Top             =   1650
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox numextensio 
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   8265
            TabIndex        =   74
            Text            =   "Ext:1234/5"
            Top             =   2055
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox numextensio 
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   8265
            TabIndex        =   73
            Text            =   "Ext:1234/5"
            Top             =   2370
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox numextensio 
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   8295
            TabIndex        =   72
            Text            =   "Ext:1234/5"
            Top             =   2745
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox numextensio 
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   7
            Left            =   8265
            TabIndex        =   71
            Text            =   "Ext:1234/5"
            Top             =   3090
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Densitat"
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
            Left            =   3675
            TabIndex        =   229
            Top             =   255
            Width           =   630
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Anilox"
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
            Left            =   2550
            TabIndex        =   228
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nom del Color"
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
            Left            =   990
            TabIndex        =   227
            Top             =   255
            Width           =   1560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Ordre Tinter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   226
            Top             =   285
            Width           =   930
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Densitat"
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
            Left            =   12375
            TabIndex        =   225
            Top             =   315
            Width           =   630
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anilox"
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
            Left            =   10905
            TabIndex        =   224
            Top             =   315
            Width           =   630
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   5400
            Picture         =   "aniloxosutilitzats.frx":BD3A
            Top             =   600
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   5400
            Picture         =   "aniloxosutilitzats.frx":C2C4
            Top             =   960
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   2
            Left            =   5400
            Picture         =   "aniloxosutilitzats.frx":C84E
            Top             =   1320
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   3
            Left            =   5400
            Picture         =   "aniloxosutilitzats.frx":CDD8
            Top             =   1680
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   4
            Left            =   5400
            Picture         =   "aniloxosutilitzats.frx":D362
            Top             =   2055
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   5
            Left            =   5400
            Picture         =   "aniloxosutilitzats.frx":D8EC
            Top             =   2415
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   6
            Left            =   5400
            Picture         =   "aniloxosutilitzats.frx":DE76
            Top             =   2775
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   7
            Left            =   5400
            Picture         =   "aniloxosutilitzats.frx":E400
            Top             =   3150
            Width           =   240
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Observacions"
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
            Left            =   12150
            TabIndex        =   223
            Top             =   150
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Treball"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4200
            TabIndex        =   222
            Top             =   285
            Width           =   405
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Liniatura"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4125
            TabIndex        =   221
            Top             =   90
            Width           =   1020
         End
         Begin VB.Label etordre 
            Caption         =   "1"
            Height          =   270
            Index           =   0
            Left            =   75
            TabIndex        =   220
            Top             =   525
            Width           =   180
         End
         Begin VB.Label etordre 
            Caption         =   "2"
            Height          =   270
            Index           =   1
            Left            =   75
            TabIndex        =   219
            Top             =   885
            Width           =   180
         End
         Begin VB.Label etordre 
            Caption         =   "3"
            Height          =   270
            Index           =   2
            Left            =   75
            TabIndex        =   218
            Top             =   1260
            Width           =   180
         End
         Begin VB.Label etordre 
            Caption         =   "4"
            Height          =   270
            Index           =   3
            Left            =   75
            TabIndex        =   217
            Top             =   1620
            Width           =   180
         End
         Begin VB.Label etordre 
            Caption         =   "5"
            Height          =   270
            Index           =   4
            Left            =   75
            TabIndex        =   216
            Top             =   1980
            Width           =   180
         End
         Begin VB.Label etordre 
            Caption         =   "6"
            Height          =   270
            Index           =   5
            Left            =   75
            TabIndex        =   215
            Top             =   2340
            Width           =   180
         End
         Begin VB.Label etordre 
            Caption         =   "7"
            Height          =   270
            Index           =   6
            Left            =   75
            TabIndex        =   214
            Top             =   2715
            Width           =   180
         End
         Begin VB.Label etordre 
            Caption         =   "8"
            Height          =   270
            Index           =   7
            Left            =   75
            TabIndex        =   213
            Top             =   3075
            Width           =   180
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Nom del Color de la Comanda"
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
            Left            =   6450
            TabIndex        =   212
            Top             =   300
            Width           =   2700
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Detall Tinter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9615
            TabIndex        =   211
            Top             =   315
            Width           =   915
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Volum"
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
            Left            =   11445
            TabIndex        =   210
            Top             =   315
            Width           =   480
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Viscos."
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
            Left            =   11895
            TabIndex        =   209
            Top             =   315
            Width           =   630
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Volum"
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
            Left            =   2970
            TabIndex        =   208
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Viscos."
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
            Left            =   3360
            TabIndex        =   207
            Top             =   90
            Width           =   630
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   3390
            Left            =   15
            Top             =   90
            Width           =   5640
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   3375
            Left            =   5640
            Top             =   90
            Width           =   7665
         End
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Dades master per aquesta comanda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   720
         TabIndex        =   231
         Top             =   90
         Width           =   4110
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Dades utilitzades en aquesta comanda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   6075
         TabIndex        =   230
         Top             =   90
         Width           =   4110
      End
   End
   Begin VB.Label eterrorlectura 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   285
      TabIndex        =   58
      Top             =   4305
      Visible         =   0   'False
      Width           =   10260
   End
   Begin VB.Label cobservacions 
      Height          =   615
      Left            =   45
      TabIndex        =   29
      Top             =   4305
      Width           =   10590
   End
   Begin VB.Menu mconsultaaltracomanda 
      Caption         =   "Consulta un altra comanda"
   End
   Begin VB.Menu maniloxoscanviats 
      Caption         =   "Entrar aniloxos canviats"
   End
End
Attribute VB_Name = "formaniloxos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jaheentrat As Boolean
Function comprovarrepetits() As Boolean
  Dim trobat As Byte
  For i = 0 To 7
    trobat = 0
    For j = 0 To 7
       If cadbl(ordre(i)) = cadbl(ordre(j)) And cadbl(ordre(i)) > 0 Then trobat = trobat + 1
    Next j
    If trobat > 1 Then
       comprovarrepetits = True
       ordre(i).BackColor = QBColor(12)
      Else: ordre(i).BackColor = QBColor(15)
    End If
  Next i
End Function

Private Sub alta_Click()
 If dataliniesobs.Recordset.RecordCount >= 2 Then MsgBox "Només es poden entrar dues linies d'observació", vbInformation, "Atenció": Exit Sub
   dataliniesobs.Recordset.AddNew
   dataliniesobs.Recordset!id_treball = id_treball
   dataliniesobs.Recordset!ordre = ordremodificacio
   dataliniesobs.Recordset.Update
   dataliniesobs.Refresh
   If dataliniesobs.Recordset.EOF Then Exit Sub
   dataliniesobs.Recordset.MoveLast
   reixa.SetFocus
End Sub

Private Sub bborrarlot_Click()
  Dim vnumc As Double
  Dim vtinter As Double
  vtinter = cadbl(bborrarlot.tag) + 1
  borrar_lots_tinta_consumida vtinter
  wait 1
  carregar_lots_tinters
End Sub
Sub borrar_lots_tinta_consumida(vtinter As Double, Optional vnopreguntar As Boolean)
  Dim vnumc As Double
  If vtinter < 1 Or vtinter > 8 Then Exit Sub
  vnumc = cadbl(formaniloxos.tag) 'formaniloxos.tag guarda la comanda que treballem
  If Not vnopreguntar Then If MsgBox("Segur que vols borrar el lot d'aquest tinter?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
  dbtmpb.Execute "delete * from impresores_llaunesgastades where comanda=" + atrim(vnumc) + " and tinter=" + atrim(vtinter) + " and tipus='I' "
  kbpantone(vtinter - 1) = "0"
End Sub

Private Sub bconsums_Click(Index As Integer)
    If color(Index).tag = "" Then MsgBox "No hi ha codi de tinta relacionada.", vbCritical, "Error": Exit Sub
    Load formconsumllaunes
    formconsumllaunes.tag = formaniloxos.tag 'es la comanda
    formconsumllaunes.nomtinta.tag = color(Index).tag  'es el codi de la tinta
    formconsumllaunes.Show 1
    Unload formconsumllaunes
End Sub

Private Sub bno_Click(Index As Integer)
  passar_modificacioano Index + 1, cadbl(formaniloxos.tag)
  carregartintes cadbl(formaniloxos.tag)
End Sub
Sub passar_modificacioano(ntinter As Integer, numc As Double)
  Dim rsttintescomanda As Recordset
   Set rsttintescomanda = dbbaixes.OpenRecordset("select * from impresores_aniloxos where comanda=" + atrim(numc) + " and ordretinter=" + atrim(ntinter))
   If Not rsttintescomanda.EOF Then
    rsttintescomanda.Edit
    rsttintescomanda!okcanvi = 2
    rsttintescomanda.Update
   End If
End Sub

Private Sub bnotots_Click()
  For i = 0 To 7
   If bno(i).visible And bno(i).Enabled Then
      passar_modificacioano i + 1, cadbl(formaniloxos.tag)
   End If
  Next i
  carregartintes cadbl(formaniloxos.tag)
End Sub

Private Sub bok_Click(Index As Integer)
  guardar_modificacio Index, cadbl(formaniloxos.tag)
  carregartintes cadbl(formaniloxos.tag)
End Sub
Sub guardar_modificacio(ntinter As Integer, numc As Double)
   Dim rsttintescomanda As Recordset
   Dim rsttreball As Recordset
   Set rsttintescomanda = dbbaixes.OpenRecordset("select * from impresores_aniloxos where comanda=" + atrim(numc) + " and id_tinter=" + atrim(ordre(ntinter).tag))
   Set rsttreball = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(ordre(ntinter).tag))
   If Not rsttintescomanda.EOF And Not rsttreball.EOF Then
    rsttreball.Edit
    If atrim(rsttintescomanda!coditinta_comanda) <> "" And atrim(rsttintescomanda!coditinta_comanda) <> "0" Then
       rsttreball!color = atrim(rsttintescomanda!tinta_comanda)
       rsttreball!coditinta = atrim(rsttintescomanda!coditinta_comanda)
    End If
    rsttreball!detalltinter = atrim(rsttintescomanda!detalltinter_comanda)
    rsttreball!ordretinter = rsttintescomanda!ordretinter
    rsttreball!anilox = rsttintescomanda!anilox_comanda
    rsttreball!volum = rsttintescomanda!volum_comanda
    rsttreball!viscositat = rsttintescomanda!viscositat_comanda
    rsttreball!densitatutilitzada = caadbl(rsttintescomanda!densitat_comanda)
    If atrim(rsttintescomanda!observacions_comanda) <> "" Then
        rsttreball!observacions = atrim(rsttintescomanda!observacions_comanda)
    End If
    rsttreball.Update
    rsttintescomanda.Edit
    rsttintescomanda!okcanvi = 2
    rsttintescomanda.Update
   End If
End Sub

Private Sub boktots_Click()
  guardar_totselscanvisdanilox
End Sub
Sub guardar_totselscanvisdanilox()
  For i = 0 To 7
   If bok(i).visible And bok(i).Enabled Then
      guardar_modificacio i, cadbl(formaniloxos.tag)
   End If
  Next i
  carregartintes cadbl(formaniloxos.tag)
End Sub

Private Sub boto_nou_Click(Index As Integer)
formescanejarllaunes.Show 1
carregar_lots_tinters True
End Sub

Private Sub bveurepdf_Click()
  obrir_pdf_treball id_treball, ordremodificacio
End Sub



Sub ensenyarperescanejarllaunes()
  Command1_Click
  boto_nou_Click 0
End Sub

Private Sub Command1_Click()
 
  ensenyar_els_lots

End Sub
Sub ensenyar_els_lots()
   DoEvents
   carregar_lots_tinters
   possar_elslotsperdefecte
   frameconsums.Left = 10500
   frameconsums.Top = 1000
   frameconsums.visible = Not frameconsums.visible

End Sub
Sub carregar_lots_tinters(Optional vavisarsialgunestaenblanc As Boolean)
   Dim i As Byte
   Dim rst As Recordset
   Dim vlots As String
   Dim vtinter As Integer
   Dim vmsg As String
   
   For i = 0 To 7
      vlots = ""
      kbpantone(i).Enabled = IIf(color(i).tag = "", False, True)
      vtinter = i
      Set rst = dbtmpb.OpenRecordset("select * from impresores_llaunesgastades where comanda=" + atrim(form1.comanda) + " and tinter=" + atrim(i + 1))
      If Not rst.EOF Then vtinter = IIf(rst!id_tinter <> 0, buscar_tinter(rst!id_tinter), i)
      While Not rst.EOF
        vlots = vlots + IIf(vlots <> "", "+", "") + atrim(rst!numllauna)
        rst.MoveNext
      Wend
      'compantone(i) = vlots + "     "
      'If tintacomanda(i) <> "" And atrim(compantone(i)) = "" Then vmsg = "La tinta " + tintacomanda(i) + " no te lots." + Chr(10)
      compantone(vtinter) = vlots + "     "
      If tintacomanda(vtinter) <> "" And atrim(compantone(vtinter)) = "" Then vmsg = "La tinta " + tintacomanda(vtinter) + " no te lots." + Chr(10)
   Next i
   Label8 = form1.pantone(8)
   compantone(8) = form1.compantone(8)
   kbpantone(8) = form1.kbpantone(8)
   Label9 = form1.pantone(9)
   compantone(9) = form1.compantone(9)
   kbpantone(9) = form1.kbpantone(9)
   If vavisarsialgunestaenblanc And vmsg <> "" Then MsgBox vmsg, vbCritical, "Atenció"
   Set rst = Nothing
End Sub
Function buscar_tinter(vidtinter) As Integer
  Dim j As Byte
  For j = 0 To 7
     If cadbl(ordre(j).tag) = vidtinter Then buscar_tinter = j
  Next j
End Function
Sub possar_elslotsperdefecte()
   Dim rst As Recordset
   Dim vquant As Double
   
   '//// EN MIRALLES HA DIT QUE APARTIR DE 01/06/21 NO ETOXI NI R25 TOT JUNT A 80/20
   'If compantone(8) = "" Then compantone(8) = saber_lotactualdelcomponent(101)
   If compantone(9) = "" Then
      compantone(9) = saber_lotactualdelcomponent(100)
      Label9 = "80/20"
      Label8 = ""
   End If
   If cadbl(kbpantone(9)) = 0 Then
       Set rst = dbtmp.OpenRecordset("select cantitatex from comandes where comanda=" + atrim(cadbl(formaniloxos.tag)))
       If Not rst.EOF Then
           vquant = Redondejar(cadbl(rst!cantitatex) / 1000, 0)
           'kbpantone(8) = Redondejar(vquant / 1.9, 1)
           'posso el total del etoxi + r25 al 9 que es 80/20
           kbpantone(9) = Redondejar(vquant / 2.1, 1) + Redondejar(vquant / 1.9, 1)
       End If
   End If
End Sub

Private Sub compantone_GotFocus(Index As Integer)
  If Index > 7 Then Exit Sub
  bborrarlot.Top = compantone(Index).Top
  bborrarlot.Left = compantone(Index).Left + compantone(Index).width - bborrarlot.width
  bborrarlot.tag = atrim(Index)
  bborrarlot.visible = True
End Sub

Private Sub compantone_LostFocus(Index As Integer)
  If Screen.ActiveControl.Name <> "bborrarlot" Then bborrarlot.visible = False
  
End Sub

Private Sub densitatcomanda_LostFocus(Index As Integer)
   If cadbl(densitatcomanda(Index)) = 0 Then Exit Sub
   If cadbl(densitatcomanda(Index)) < 0.5 Or (densitatcomanda(Index)) > 2 Then
      MsgBox "Els valors vàlids són de 0.5 a 2 ", vbCritical, "Atenció"
      densitatcomanda(Index) = 0
      'densitatcomanda(Index).SetFocus
   End If
End Sub

Sub possardetallstinter()
   Dim i As Byte
   For i = o To 7
     If cdetall(i) <> "" Then
        cdetall(i).visible = True
         Else: cdetall(i).visible = False
     End If
     If detalltinter(i) = "" Then
         detalltinter(i).width = 500: detalltinter(i).Left = 7245: tintacomanda(i).width = 2500
           Else
             detalltinter(i).width = 1000: detalltinter(i).Left = 6745: tintacomanda(i).width = 2000
     End If
   Next i
End Sub

Function demanardetalltinter(vvaloranterior As String) As String
  Sendkeys "{TAB}"
  Load formseleccionou
  formseleccionou.caption = "Selecciona DETALL TINTER"
  formseleccionou.Data1.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
  formseleccionou.Data1.RecordSource = "select detall from detallsdelstinters order by detall"
  formseleccionou.refrescar
  formseleccionou.DBGrid2.Columns(0).width = 3500
  formseleccionou.sortirs.tag = "filtre"
  formseleccionou.Show 1
  If seleccioret = 1 Then
   demanardetalltinter = atrim(formseleccionou.Data1.Recordset!detall)
  End If
  
  If seleccioret = 9 Then demanardetalltinter = ""
  If seleccioret = 0 Then demanardetalltinter = vvaloranterior
  Unload formseleccionou
End Function

Private Sub detalltinter_DropDown(Index As Integer)
  detalltinter(Index) = demanardetalltinter(detalltinter(Index))
  'possardetallstinter
End Sub

Private Sub etqualitat_Click()
  formqualitatimpresio.Show 1
  possaretiquetaqualitat cadbl(formaniloxos.tag)
End Sub

Function saber_lotactualdelcomponent(vnumdosificador As Byte) As String
   Dim vsql As String
   Dim vsubsql As String
   Dim rst As Recordset
   vsubsql = "SELECT Max(detallnumeroslotsbase.data) AS MáxDedata FROM Componentsbase INNER JOIN detallnumeroslotsbase ON Componentsbase.idcomponent = detallnumeroslotsbase.idcomponent "
   vsubsql = vsubsql + "GROUP BY Componentsbase.numdosificador HAVING (((Componentsbase.numdosificador)=" + atrim(vnumdosificador) + "))"
   vsql = "SELECT detallnumeroslotsbase.data, detallnumeroslotsbase.numerodelot From detallnumeroslotsbase "
   vsql = vsql + " WHERE detallnumeroslotsbase.data=(" + vsubsql + ");"
  ' Clipboard.Clear
  ' Clipboard.SetText vsql
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
   Set rst = dbtintes.OpenRecordset(vsql)
   If Not rst.EOF Then saber_lotactualdelcomponent = atrim(rst!numerodelot)
End Function

Private Sub Form_Activate()
  Me.caption = "Obrint clixes..."
  
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  eterrorlectura = "": eterrorlectura.visible = False
  If jaheentrat Then Exit Sub
  jaheentrat = True
  Me.width = 14565
  Me.Height = 5715
  If cadbl(formaniloxos.tag) = 0 Then MsgBox "Error valor intern de comanda=0"
  Me.caption = "Dades capçalera..."
  posardadescapcalera cadbl(formaniloxos.tag)
  DoEvents
  If etcomanda.tag = "noexisteix" Then
     MsgBox "La comanda " + formaniloxos.tag + " no existeix.", vbCritical, "Error"
     etcomanda.tag = ""
     Exit Sub
  End If
  
  Me.caption = "Posar colors a les tintes..."
  If fbotonsok.tag = "" Then
     amagarbotons True
     possarcolorsalestintes cadbl(formaniloxos.tag)
      Else:
         amagarbotons False
  End If
  Me.caption = "Carregar tintes..."
  carregartintes cadbl(formaniloxos.tag)
  If fbotonsok.tag = "" Then
     ensenyar_els_lots
  End If
  frameconsums.visible = False
  If vtipusimpresio = "R" Then
     campsodreactius False
     frametinters.Enabled = False: framecandau.visible = True: ensenyar_els_lots
      Else: frametinters.Enabled = True: framecandau.visible = False: campsodreactius True
  End If
  If boto_nou(0).tag = "llegirllaunes" Then wait 1: ensenyarperescanejarllaunes: boto_nou(0).tag = ""
End Sub
Sub campsodreactius(v As Boolean)
   Dim i As Byte
   For i = 0 To 7
     ordre(i).Enabled = v
   Next i
End Sub
Sub possarcolorsalestintes(numc As Double)
    Dim rst As Recordset
    Dim i As Byte
    Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, tintes.codi, historiallauna.* FROM tintes LEFT JOIN (Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) ON tintes.idtinta = Llaunes.idtinta WHERE (((historiallauna.tipusmoviment)='I') AND ((historiallauna.comanda)=" + atrim(numc) + "));")
    For i = 0 To 7
       rst.FindFirst "codi='" + atrim(color(i).tag) + "'"
       If rst.NoMatch Then
            bconsums(i).BackColor = QBColor(15)
              Else: bconsums(i).BackColor = QBColor(12)
       End If
    Next i
    Set rst = Nothing
End Sub
Sub amagarbotons(amagar As Boolean)
   Dim i As Byte
   If amagar Then
      fbotonsok.visible = False
      formaniloxos.width = formaniloxos.width - fbotonsok.width
      checkeditar.visible = False
      
       Else
         formaniloxos.Height = 5000
         For i = 0 To 7
            observacions(i).width = 2550
            bconsums(i).visible = False
         Next i
         Command1.visible = False
   End If
End Sub
Sub posardadescapcalera(numc As Double)
  Dim rst As Recordset
  If numc = 0 Then etcomanda = "ERROR: COMANDA A ZERO": Exit Sub
  
  Set rst = dbclixes.OpenRecordset("SELECT Clixes.id_treball, Clixes.marca, Clixes.linia, comandes.numordremodificacio,clients.codi, clients.nom, comandes.comanda FROM (comandes INNER JOIN clients ON comandes.client = clients.codi) INNER JOIN Clixes ON comandes.numtreball = Clixes.id_treball WHERE (((comandes.comanda)=" + atrim(numc) + "));")
  If rst.EOF Then etcomanda.tag = "noexisteix": Exit Sub
  etcomanda = "Lot: " + atrim(numc)
  etclient = "Client: " + atrim(rst!codi) + " - " + atrim(rst!nom)
  ettreball = "Treball: " + atrim(cadbl(rst!id_treball)) + "/" + atrim(CDbl(rst!numordremodificacio))
  etlinia = "Texte: " + atrim(rst!marca) + " - " + atrim(rst!linia)
  id_treball = cadbl(rst!id_treball)
  ordremodificacio = cadbl(rst!numordremodificacio)
End Sub
Sub netejar_camps()
  Dim i As Byte
  For i = 0 To 7
    ordre(i).tag = ""
    ordre(i) = ""
    observacions(i) = ""
    color(i) = ""
    anilox(i) = ""
    volum(i) = ""
    tintacomanda(i) = ""
    tintacomanda(i).tag = ""
    aniloxfotogravador(i) = ""
    densitat(i) = ""
    aniloxcomanda(i) = ""
    densitatcomanda(i) = ""
    observacions(i) = ""
    kbpantone(i) = ""
    compantone(i) = ""
    compantone(i).tag = ""
  Next i
End Sub
 Sub carregartintes(numc As Long)
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim rstlink As Recordset
   Dim rsttintescomanda As Recordset
   Dim numtinter As Long
   Dim i As Byte
   Dim dbt As Database
   Dim vkgtinta As Double
   
   DoEvents
   If numc = 0 Or id_treball = 0 Then Exit Sub
   If fbotonsok.tag = "activats" Then
        Set dbt = dbbaixes
       Else: Set dbt = dbtmpb: Set dbbaixes = dbtmpb
   End If
'   Set dbt = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb", , True)
   If fbotonsok.tag <> "activats" Then Set dbbaixes = dbt
   Me.caption = "Netejant camps..."
   netejar_camps
   
   cobservacions = ""
   Me.caption = "Carregant aniloxos..."
   Set rsttintescomanda = dbt.OpenRecordset("select * from impresores_aniloxos where comanda=" + atrim(numc) + " order by ordretinter_original", , ReadOnly)
   Me.caption = "Carregant dades treball..."
   DoEvents
   Set rstc = dbt.OpenRecordset("select numtreball,numordremodificacio,cantitatex from comandes where comanda=" + atrim(numc), , ReadOnly)
   If rstc.EOF Then MsgBox "La comanda no existeix", vbCritical, "Error": Exit Sub
   Me.caption = "Possant observacions..."
   DoEvents
   If Not rstc.EOF Then cobservacions = possarobservacionstintes(cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio))
   Me.caption = "Carregant dades tintes..."
   DoEvents
   Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(rstc!numtreball) + " and ordremodificacio=" + atrim(rstc!numordremodificacio) + " order by ordretinter ASC", , ReadOnly)
   i = 0
   While Not rst.EOF
      Me.caption = "Passant dades treball linia " + atrim(i + 1)
      numtinter = cadbl(IIf(rst!tinterlinkambid_treball > 0, rst!tinterlinkambid_treball, rst!id_tinter))
      rsttintescomanda.FindFirst "id_tinter=" + atrim(numtinter)
      If rsttintescomanda.NoMatch Then
         Me.caption = "Passant dades treball linia " + atrim(i + 1) + " NO MATCH"
         dbt.Execute "delete * from impresores_aniloxos where comanda=" + atrim(numc) + " and ordretinter=" + atrim(rst!ordretinter)
         Me.caption = "Passant dades treball linia " + atrim(i + 1) + " CREANT TINTA"
         crear_tintacomanda numtinter, numc, rsttintescomanda, cadbl(rst!ordretinter)
         Set rsttintescomanda = dbt.OpenRecordset("select * from impresores_aniloxos where comanda=" + atrim(numc) + " order by ordretinter_original", , ReadOnly)
      End If
      rsttintescomanda.FindFirst "id_tinter=" + atrim(numtinter)
      If cadbl(rst!tinterlinkambid_treball) > 0 Then
        Set rstlink = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(rst!tinterlinkambid_treball), , ReadOnly)
        If Not rstlink.EOF Then
            Me.caption = "Passant dades treball linia " + atrim(i + 1) + " PASSANT DATOS A LA LINIA"
            passardatosaliniatinta rsttintescomanda, i, True, rstlink
           'Else: MsgBox "Hi ha hagut un error al carregar el tinter linkat al tinter " + atrim(i + 1), vbCritical, "Atenció"
        End If
         Else: passardatosaliniatinta rsttintescomanda, i, True, rst
      End If
      Set rstlink = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(rsttintescomanda!id_tinter), , ReadOnly)
      If fbotonsok.tag <> "activats" Then
             vkgtinta = Redondejar(calcular_kgmetreteoric(rstlink) * cadbl(rstc!cantitatex), 1)
             etkgteorics(i) = atrim(vkgtinta) + "Kg"
             If cadbl(kbpantone(i)) = 0 Then kbpantone(i) = atrim(vkgtinta)
      End If
      Me.caption = "Passant dades treball linia " + atrim(i + 1) + " MIRAR SI TE EXTENSIO FETA"
      mirarsiteextensiofeta cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), cadbl(rsttintescomanda!coditinta_comanda), i
      i = i + 1
      rst.MoveNext
      DoEvents
   Wend
   
   form1.passarlotsaprincipal
   Me.caption = "Carregant etiqueta qualitat..."
   possaretiquetaqualitat numc
   Set rst = Nothing
   Set rstlink = Nothing
   Set dbt = Nothing
   Me.caption = "Manteniment Aniloxos i Densitats"
   'possardetallstinter
End Sub
Sub mirarsiteextensiofeta(vnumtreball As Double, vnumordre As Byte, vcoditinta As Double, i As Byte)
  Dim rsttintes As Recordset
  Dim j As Byte
 ' For j = 0 To 7
    numextensio(i) = ""
 ' Next j
  
  Set rsttintes = dbbaixes.OpenRecordset("select codiextensio from extensions_treballsrelacionats where numtreball=" + atrim(vnumtreball) + " and numordremodificacio=" + atrim(vnumordre) + " and coditinta=" + atrim(vcoditinta), , ReadOnly)
   If rsttintes.EOF Then
      numextensio(i).visible = False
       Else
         numextensio(i).text = "Ext:" + atrim(rsttintes!codiextensio)
         numextensio(i).visible = True
         numextensio(i).ZOrder 0
   End If
  Set rsttintes = Nothing
End Sub
Sub possaretiquetaqualitat(numc As Long)
   Dim rst As Recordset
   Dim dbt As Database
   Dim vqualitats As Variant
   vqualitats = Array("", "Dolenta", "Bona", "Molt Bona")
   If fbotonsok.tag = "activats" Then
        Set dbt = dbbaixes
       Else: Set dbt = dbtmpb
   End If
   etqualitat = ""
   Set rst = dbt.OpenRecordset("select * from  impressorestot where comanda=" + atrim(cadbl(numc)))
   If Not rst.EOF Then
      etqualitat = vqualitats(cadbl(rst!qualitatimpresio))
   End If
   Set rst = Nothing
End Sub
Function possarobservacionstintes(vidtreball As Double, vordre As Double) As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(vidtreball) + " and ordre=" + atrim(vordre), , ReadOnly)
   If rst.EOF Then Exit Function
   While Not rst.EOF
     possarobservacionstintes = atrim(rst!observacio) + Chr(10) + Chr(13)
     rst.MoveNext
   Wend
   If possarobservacionstintes <> "" Then possarobservacionstintes = "Observacions: " + possarobservacionstintes
End Function
Sub crear_tintacomanda(idtinter As Long, numc As Long, rstc As Recordset, vordretinter As Byte)
    Dim rstlink As Recordset
    Set rstlink = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(idtinter))
    If rstlink.EOF Then Exit Sub
    rstc.AddNew
    rstc!id_tinter = idtinter
    rstc!comanda = numc
    rstc!ordretinter = cadbl(vordretinter)
    rstc!ordretinter_original = cadbl(vordretinter)
    rstc!tinta_original = atrim(rstlink!color)
    rstc!coditinta_original = cadbl(rstlink!coditinta)
    rstc!anilox_original = cadbl(rstlink!anilox)
    rstc!densitat_original = cadbl(rstlink!densitatutilitzada)
    rstc!detalltinter_original = atrim(rstlink!detalltinter)
    rstc!anilox_comanda = cadbl(rstlink!anilox)
    rstc!densitat_comanda = cadbl(rstlink!densitatutilitzada)
    rstc!tinta_comanda = atrim(rstlink!color)
    rstc!detalltinter_comanda = atrim(rstlink!detalltinter)
    rstc!coditinta_comanda = rstlink!coditinta
    rstc!observacions_comanda = atrim(rstlink!observacions)
    rstc.Update
End Sub
Sub passardatosaliniatinta(ByVal rstdatos As Recordset, i As Byte, estalinkat As Boolean, Optional rstlink As Recordset)
     
    ordre(i).BackColor = QBColor(15)
    ordre(i).tag = atrim(cadbl(rstdatos!id_tinter))
    ordre(i) = atrim(cadbl(rstdatos!ordretinter))
    observacions(i) = atrim(rstdatos!observacions_comanda)
    aniloxcomanda(i) = atrim(cadbl(rstdatos!anilox_comanda))
    tintacomanda(i) = atrim(rstdatos!tinta_comanda)
    tintacomanda(i).tag = atrim(rstdatos!coditinta_comanda)
    volumcomanda(i) = atrim(cadbl(rstdatos!volum_comanda))
    viscositatcomanda(i) = atrim(cadbl(rstdatos!viscositat_comanda))
    detalltinter(i) = atrim(rstdatos!detalltinter_comanda)
    densitatcomanda(i) = atrim((rstdatos!densitat_comanda))
    kbpantone(i) = atrim((rstdatos!kgconsumits))
    compantone(i).tag = atrim(rstdatos!id)
    compantone(i) = possartotselslots(rstdatos!id)
    posarelsoks rstdatos, i
    If estalinkat Then Set rstdatos = rstlink
    color(i) = atrim(rstdatos!color)
    color(i).tag = atrim(rstdatos!coditinta)
    cdetall(i) = atrim(rstdatos!detalltinter)
    anilox(i) = atrim(cadbl(rstdatos!anilox))
    volum(i) = atrim(cadbl(rstdatos!volum))
    aniloxfotogravador(i) = atrim(cadbl(rstdatos!aniloxclixe))
    densitat(i) = atrim((rstdatos!densitatutilitzada))
    observacions(i).tag = atrim(rstdatos!observacions)
    
    If cadbl(ordre(i)) <> i + 1 Then ordre(i).BackColor = QBColor(14)
End Sub
Function possartotselslots(vid As Long) As String
    Dim rst As Recordset
    Set rst = dbbaixes.OpenRecordset("select * from impresores_lotsdetinta where id=" + atrim(vid))
    While Not rst.EOF
       possartotselslots = possartotselslots + IIf(possartotselslots <> "", ",", "") + atrim(rst!numerodelot)
       rst.MoveNext
    Wend
    Set rst = Nothing
End Function
Sub posarelsoks(rstdatos As Recordset, i As Byte)
If rstdatos!okcanvi > 0 Then
        bok(i).visible = True
        bno(i).visible = True
        If checkeditar.Value = 1 Then
            bok(i).Enabled = True
            bno(i).Enabled = True
              Else
               If rstdatos!okcanvi = 2 Then
                bok(i).Enabled = False
                bno(i).Enabled = False
               End If
        End If
         Else
            bok(i).visible = False
            bno(i).visible = False
    End If
End Sub


Function toteslesllaunesentrades() As Boolean
   toteslesllaunesentrades = True
   For i = 0 To 7
       If atrim(tintacomanda(i)) <> "" And compantone(i) = "" Then toteslesllaunesentrades = False
   Next i
End Function

Private Sub Form_Click()
'fotocandau_DblClick
'   enviaremailgeneric "miquel.inplacsa@gmail.com;miquel.inplacsa@gmail.com", "ACCES AL CANDAU A MÀQUINA. ", "Comanda " + atrim(numc) + "  " + nommaq + " - " + nomoperari + Chr(10) + Chr(13) + "S'HA ACCEDIT A LA SECCIÓ TINTERS PER CANVIAR ALGUNA TINTA."
End Sub
Function posiciocontrol(vControl As ComboBox, vmida As String) As Double
   If vmida = "left" Then
      posiciocontrol = vControl.Left + frametinters.Left + Frame1.Left
   End If
   If vmida = "top" Then
        posiciocontrol = vControl.Top + frametinters.Top + Frame1.Top
   End If
End Function
Sub ensenyapostittinta(vControl As ComboBox)
   cpostit = ""
   cpostit.text = vControl.text
   cpostit.Left = posiciocontrol(vControl, "left") + 100
   cpostit.Top = posiciocontrol(vControl, "top") + vControl.Height + 50
   If cpostit <> "" Then cpostit.visible = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim i As Byte
   Dim vensenya As Boolean
   For i = 0 To tintacomanda.Count - 1
     If X > posiciocontrol(tintacomanda(i), "left") And X < (posiciocontrol(tintacomanda(i), "left") + tintacomanda(i).width) Then
       If Y > posiciocontrol(tintacomanda(i), "top") And Y < (posiciocontrol(tintacomanda(i), "top") + tintacomanda(i).Height) Then
              ensenyapostittinta tintacomanda(i)
              vensenya = True
       End If
     End If
   Next i
   If Not vensenya Then cpostit.visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

   If comprovarrepetits Then MsgBox "Hi ha ordre de tinter repetit, primer arregla-ho", vbCritical, "Atenció": Cancel = 1: Exit Sub
   
   If fbotonsok.tag <> "activats" Then
      'If Not toteslesllaunesentrades Then If UCase(InputBox("Encara no hi ha totes les llaunes entrades." + Chr(10) + "SI VOLS SORTIR SENSE ACABAR ESCRIU [SORTIR].", "No hi ha totes les llaunes")) <> "SORTIR" Then Cancel = 1: Exit Sub
      guardar_tintes cadbl(formaniloxos.tag)
      form1.passarlotsaprincipal
      If etqualitat = "" And checkeditar.tag <> "fingerprint" Then formqualitatimpresio.Show 1
   End If
   'dbclixes.Close
   jaheentrat = False
   Set dbclixes = Nothing
   id_treball = 0
  ordremodificacio = 0
End Sub
Function existeixformulari(vnomformulari As Form) As Boolean
  Dim oCtrl As Form
  For Each oCtrl In Forms
    If LCase(oCtrl.Name) = "form1" Then existeixformulari = True
  Next
  
End Function


Sub guardar_tintes(numc As Double)
    Dim i As Byte
    Dim rst As Recordset
    Dim dbt As Database
    Dim rstc As Recordset
    If numc = 0 Then Exit Sub
    If fbotonsok.tag = "activats" Then
        Set dbt = dbbaixes
       Else: Set dbt = dbtmpb
   End If
   Set rstc = dbt.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
    Set rst = dbt.OpenRecordset("select * from impresores_aniloxos where comanda=" + atrim(numc))
    rstc.Edit
    For i = 0 To 7
       If cadbl(ordre(i).tag) > 0 Then
        rst.FindFirst "id_tinter=" + atrim(cadbl(ordre(i).tag))
        If Not rst.NoMatch Then
          rst.Edit
           rst!tinta_comanda = tintacomanda(i)
           rst!coditinta_comanda = tintacomanda(i).tag
           If (cadbl(ordre(i)) > 0 And cadbl(ordre(i)) < 9) Then rst!ordretinter = cadbl(ordre(i))
           rst!anilox_comanda = cadbl(aniloxcomanda(i))
           rst!densitat_comanda = caadbl(densitatcomanda(i))
           rst!detalltinter_comanda = atrim(detalltinter(i))
           rst!volum_comanda = atrim(cadbl(volumcomanda(i)))
           rst!viscositat_comanda = atrim(cadbl(viscositatcomanda(i)))
           rst!kgconsumits = cadbl(kbpantone(i))
           rst!observacions_comanda = treure_apostruf(observacions(i))
           If rst!okcanvi < 2 Then
            If cadbl(rst!volum_comanda) <> cadbl(rst!volum_original) Or cadbl(rst!viscositat_comanda) <> cadbl(rst!viscositat_original) Or atrim(rst!detalltinter_comanda) <> atrim(rst!detalltinter_original) Or rst!observacions_comanda <> atrim(observacions(i).tag) Or rst!anilox_comanda <> rst!anilox_original Or caadbl(rst!densitat_comanda) <> caadbl(rst!densitat_original) Or atrim(rst!tinta_original) <> atrim(rst!tinta_comanda) Or atrim(IIf(rst!coditinta_original = "0", "", rst!coditinta_original)) <> atrim(IIf(rst!coditinta_comanda = "0", "", rst!coditinta_comanda)) Or rst!ordretinter <> rst!ordretinter_original Then
                 rst!okcanvi = 1
                Else: rst!okcanvi = 0
            End If
           End If
          rst.Update
        End If
        rstc.Fields("tinta" + atrim(i + 1) + "a") = atrim(tintacomanda(i))
        rstc.Fields("lin" + atrim(i + 1)) = cadbl(aniloxcomanda(i))
       End If
    Next i
    rstc.Update
    
End Sub

Private Sub fotocandau_DblClick()
   Dim v As String
   Dim i As Byte
   v = InputBoxEx("Aquesta comanda es Repetida no es pot canviar l'ordre." + Chr(10) + "Escriu la contrasenya per editar-la." + Chr(10) + Chr(10) + "ACCEDIR A AQUESTES MODIFICACIONS ENVIARÀ AUTOMATICAMENT UN CORREU A IMPRESORES I TINTES, NO ENTREU SI NO CAL FER-HO.", "Comanda repetida.", , , , , , SPassword)
   If UCase(v) = "CONTROLTINTES" Then
      enviaremailgeneric "impresores@inplacsa.com;tintes@inplacsa.com", "ACCES AL CANDAU A MÀQUINA. ", "Comanda " + atrim(form1.comanda) + "  " + nommaq + " - " + form1.nomoperari + Chr(10) + Chr(13) + "S'HA ACCEDIT A LA SECCIÓ TINTERS PER CANVIAR ALGUNA TINTA."
      frametinters.Enabled = True: framecandau.visible = False
      For i = 0 To 7
        aniloxcomanda(i).Enabled = False
        detalltinter(i).Enabled = False
        volumcomanda(i).Enabled = False
        viscositatcomanda(i).Enabled = False
        densitatcomanda(i).Enabled = False
      Next i
   End If
End Sub

Private Sub frameconsums_Click()
 ' Dim i As Byte
 '  For i = 100 To 110
 '     MsgBox saber_lotactualdelcomponent(i)
 '  Next i
End Sub
Function calcular_kgmetreteoric(rsttinter As Recordset) As Double
   Dim vresultat As Double
   Dim vvolum As Double
   Dim vaporte As Double
   Dim rst As Recordset
   Dim rsta As Recordset
   Dim vample As Double
   'amplelamina
   Set rst = dbclixes.OpenRecordset("select amplelamina,bandes from modificacions where id_treball=" + atrim(rsttinter!id_treball) + " and ordre=" + atrim(rsttinter!ordremodificacio))
   If rst.EOF Then Exit Function
   vample = (cadbl(rst!amplelamina) * 10) * cadbl(rst!bandes)
   Set rsta = dbtmp.OpenRecordset("select max(volum) as volummesgran from aniloxos where lineatura=" + atrim(cadbl(rsttinter!anilox)))
   'vvolum = IIf(cadbl(rsttinter!volum) > 0, cadbl(rsttinter!volum), 20)
   vvolum = IIf(rsta.EOF, 0, cadbl(rsta!volummesgran))
   vaporte = 30 'no se quin valor es aquest es per defecte
   vresultat = (vaporte * vvolum) / 100000
   vresultat = ((cadbl(rsttinter!tanx100cobertura) / 100) * (vample / 1000)) * vresultat
   vresultat = vresultat * 0.95
   calcular_kgmetreteoric = vresultat
End Function

Private Sub kbpantone_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim i As Byte
  If KeyCode = 40 And Index < kbpantone.Count - 1 Then
    i = 1
    While Not kbpantone(Index + i).Enabled
       i = i + 1
       If (Index + i) >= (kbpantone.Count - 1) Then GoTo cont
    Wend
    kbpantone(Index + i).SetFocus
  End If
  If KeyCode = 38 And Index > 0 Then
   i = 1
   While Not kbpantone(Index - i).Enabled
       i = i + 1
       If (Index - i) < 0 Then GoTo cont
    Wend
    kbpantone(Index - i).SetFocus
  End If
cont:
End Sub

Private Sub kbpantone_LostFocus(Index As Integer)
   form1.passarlotsaprincipal
End Sub

Private Sub maniloxoscanviats_Click()
   formcanvisanilox.Show 1: Unload formcanvisanilox
End Sub

Private Sub mconsultaaltracomanda_Click()
   Dim vnumc As String
   vnumc = InputBox("Entra la comanda que vols consultar.", "Consulta")
   If cadbl(vnumc) < 1 Then MsgBox "Comanda no vàlida": Exit Sub
   formaniloxos.tag = vnumc
   jaheentrat = False
   fbotonsok.tag = "activats" ' per evita que gravi les dades al surtir del formulari
   Form_Activate
End Sub

Private Sub observacions_GotFocus(Index As Integer)
    observacions(Index).Left = 7740
    observacions(Index).width = 2875
    observacions(Index).BackColor = QBColor(10)
    
End Sub

Private Sub observacions_LostFocus(Index As Integer)
   observacions(Index).width = 675
   observacions(Index).Left = 9960
   observacions(Index).BackColor = QBColor(15)
End Sub

Private Sub ordre_Change(Index As Integer)
 If Not formaniloxos.visible Then GoTo fi
 If Screen.ActiveControl.Name = "ordre" Then
  If Not comprovarrepetits Then
'    borrar_dades_tintes_consumides
    guardar_tintes cadbl(formaniloxos.tag)
    'carregartintes cadbl(Form1.comanda)
  End If
 End If
fi:
End Sub
Sub borrar_dades_tintes_consumides()
  Dim i As Double
  Dim vhihadades As Boolean
  For i = 0 To 7
     If atrim(compantone(i)) <> "" Then vhihadades = True
  Next i
  If Not vhihadades Then GoTo fi
  MsgBox "Canviar l'ordre de tinters comporta l'eliminació dels LOTS de tintes i consums." + vbNewLine + "S'HAURAN DE TORNAR A ESCANEJAR.", vbInformation, "ATENCIÓ"
  For i = 0 To 7
    borrar_lots_tinta_consumida i + 1, True
  Next i
  wait 1
  carregar_lots_tinters
fi:
End Sub

Function caadbl(valor As Variant) As Double
  If IsNull(valor) Then caadbl = 0
  If IsNumeric(valor) Then caadbl = valor
End Function

Private Sub Timer1_Timer()
 
End Sub

Private Sub tintacomanda_DropDown(Index As Integer)
    If atrim(compantone(Index)) = "" Then
       triartinta Index
        Else: MsgBox tintacomanda(Index) + vbNewLine + "No pots canviar de tinta si ja has escanejat les llaunes." + Chr(10) + "Primer borra el lot escanejat i despres canvia'l.", vbCritical, "Error"
    End If
End Sub

Sub triartinta(pos As Integer)
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  Dim fseleccio As Form
  Set fseleccio = formseleccionou
  sql = "SELECT  codi,descripcio,referenciacolor,refproveidor from tintes_tot "
  'per filtra per families les tintes a escullir
'  Set rst2 = dbtintes.OpenRecordset("Select * from tintes_tot where codi='" + atrim(tintacomanda(pos).tag) + "'")
'  If rst2.EOF Then
  '    sql = "SELECT  codi,descripcio,referenciacolor from tintes_tot "
 '      Else:
  '       sql = "SELECT  codi,descripcio,referenciacolor from tintes_tot where idfamilia=" + atrim(cadbl(rst2!idfamilia)) '+ " and idserie=" + atrim(cadbl(rst2!idserie)) + " and idsubfamilia=" + atrim(cadbl(rst2!idsubfamilia))
   '      If InStr(1, rst2!descripcio, "PRIMAR") > 0 Then
    '         sql = "SELECT  codi,descripcio,referenciacolor from tintes_tot where idfamilia=" + atrim(cadbl(rst2!idfamilia)) + " and idfamcolor=" + atrim(cadbl(rst2!idfamcolor)) + " and idsubfamcolor=" + atrim(cadbl(rst2!idsubfamcolor))
  '       End If
 ' End If
     
  were = " order by descripcio"
  Load fseleccio
  fseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
  fseleccio.Data1.RecordSource = sql + were
  fseleccio.width = 12000
  fseleccio.sortirs.tag = "filtre"
 
  fseleccio.refrescar
   fseleccio.DBGrid2.Columns(0).width = 600
  fseleccio.DBGrid2.Columns(1).width = 5000
  fseleccio.DBGrid2.Columns(2).width = 3000
  fseleccio.DBGrid2.Columns(3).width = 2000
  fseleccio.Show 1
  wait 1
  If seleccioret = 1 Then
    tintacomanda(pos) = atrim(fseleccio.Data1.Recordset!descripcio)
    tintacomanda(pos).tag = atrim(fseleccio.Data1.Recordset!codi)
  End If
  If seleccioret = 9 Then
    tintacomanda(pos) = color(pos)
    tintacomanda(pos).tag = color(pos).tag
  End If
 '  Data1.Recordset!client = Text2.Text
 '  nomclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
  
 ' End If
 
  Unload fseleccio
 'SendKeys "{TAB}"
 If aniloxcomanda(pos).Enabled Then aniloxcomanda(pos).SetFocus
End Sub

Private Sub tintacomanda_GotFocus(Index As Integer)
  numextensio(Index).visible = False
End Sub

Private Sub tintacomanda_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub tintacomanda_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub tintacomanda_LostFocus(Index As Integer)
   If numextensio(Index) <> "" Then numextensio(Index).visible = True: numextensio(Index).ZOrder 0
End Sub
