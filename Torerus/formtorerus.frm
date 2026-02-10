VERSION 5.00
Begin VB.Form formtorerus 
   Caption         =   "Torerus"
   ClientHeight    =   11670
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   19200
   Icon            =   "formtorerus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11670
   ScaleWidth      =   19200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Framepassword 
      BackColor       =   &H00EAD9CE&
      Height          =   9165
      Left            =   15540
      TabIndex        =   11
      Top             =   9045
      Visible         =   0   'False
      Width           =   7230
      Begin VB.CommandButton Command14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   5745
         Picture         =   "formtorerus.frx":048A
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   8085
         Width           =   1275
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Index           =   11
         Left            =   3885
         TabIndex        =   48
         Top             =   6390
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   0
         Left            =   135
         TabIndex        =   23
         Top             =   945
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   1
         Left            =   2010
         TabIndex        =   22
         Top             =   945
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   3
         Left            =   135
         TabIndex        =   21
         Top             =   2745
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   4
         Left            =   2010
         TabIndex        =   20
         Top             =   2745
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   6
         Left            =   135
         TabIndex        =   19
         Top             =   4545
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   7
         Left            =   2010
         TabIndex        =   18
         Top             =   4545
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Index           =   9
         Left            =   150
         TabIndex        =   17
         Top             =   6375
         Width           =   3630
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   2
         Left            =   3885
         TabIndex        =   16
         Top             =   945
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   5
         Left            =   3885
         TabIndex        =   15
         Top             =   2745
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   8
         Left            =   3885
         TabIndex        =   14
         Top             =   4545
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         BackColor       =   &H00C0FFC0&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Index           =   10
         Left            =   5715
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   945
         Width           =   1365
      End
      Begin VB.TextBox cpassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   165
         TabIndex        =   12
         Top             =   8070
         Width           =   5505
      End
      Begin VB.Label etmissatgepassword 
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
         ForeColor       =   &H005C31DD&
         Height          =   720
         Left            =   105
         TabIndex        =   74
         Top             =   165
         Width           =   7020
      End
   End
   Begin VB.CommandButton busuari 
      Height          =   855
      Left            =   135
      Picture         =   "formtorerus.frx":135C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   390
      Width           =   990
   End
   Begin VB.Frame cframetotal 
      Enabled         =   0   'False
      Height          =   11610
      Left            =   90
      TabIndex        =   0
      Top             =   -45
      Width           =   19125
      Begin VB.Timer timeractualitzacions 
         Interval        =   60000
         Left            =   1440
         Top             =   990
      End
      Begin VB.Frame chihaalgu 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   4
         Left            =   12300
         TabIndex        =   72
         Top             =   240
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Frame chihaalgu 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   3
         Left            =   10155
         TabIndex        =   71
         Top             =   240
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Frame chihaalgu 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   2
         Left            =   8055
         TabIndex        =   70
         Top             =   240
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Frame chihaalgu 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   1
         Left            =   5910
         TabIndex        =   69
         Top             =   240
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Frame chihaalgu 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   0
         Left            =   3795
         TabIndex        =   68
         Top             =   240
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Frame Framemissatge 
         BackColor       =   &H00EEE4D7&
         Height          =   1875
         Left            =   12105
         TabIndex        =   65
         Top             =   2340
         Visible         =   0   'False
         Width           =   7215
         Begin VB.Label etvermella 
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
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   75
            TabIndex        =   67
            Top             =   1605
            Width           =   6720
         End
         Begin VB.Label etmissatgeframemissatge 
            BackStyle       =   0  'Transparent
            Caption         =   "etmissatgeframemissatge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1545
            Left            =   165
            TabIndex        =   66
            Top             =   180
            Width           =   7005
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Revisar Entrega"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   13380
         Picture         =   "formtorerus.frx":1946
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   165
         Width           =   1245
      End
      Begin VB.ComboBox combodies 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "formtorerus.frx":2210
         Left            =   18105
         List            =   "formtorerus.frx":222C
         TabIndex        =   61
         Text            =   "2"
         Top             =   90
         Width           =   600
      End
      Begin VB.CommandButton Command15 
         Height          =   315
         Left            =   18750
         Picture         =   "formtorerus.frx":2248
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Llista de copies de seguretat"
         Top             =   105
         Width           =   345
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Pack' List"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   14655
         Picture         =   "formtorerus.frx":27D2
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   150
         Width           =   1245
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ESTOC->IMP/LAM"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   10590
         Picture         =   "formtorerus.frx":2DA2
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   150
         Width           =   2010
      End
      Begin VB.CommandButton command12 
         Caption         =   "Forats lliures"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1350
         Left            =   15915
         Picture         =   "formtorerus.frx":346E
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   150
         Width           =   1245
      End
      Begin VB.Frame frameubicacions 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0C000&
         Height          =   1875
         Left            =   5865
         TabIndex        =   39
         Top             =   10785
         Visible         =   0   'False
         Width           =   12000
         Begin VB.CommandButton bubicacio 
            BackColor       =   &H005C31DD&
            Caption         =   "REC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Index           =   6
            Left            =   7455
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   540
            Width           =   1665
         End
         Begin VB.CommandButton bubicacio 
            BackColor       =   &H00FF00FF&
            Caption         =   "G"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Index           =   5
            Left            =   9210
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   525
            Width           =   1290
         End
         Begin VB.CommandButton bubicacio 
            BackColor       =   &H00FF80FF&
            Caption         =   "F"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Index           =   4
            Left            =   10575
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   525
            Width           =   1290
         End
         Begin VB.CommandButton bubicacio 
            BackColor       =   &H00F1B75F&
            Caption         =   "IMP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Index           =   0
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   540
            Width           =   1785
         End
         Begin VB.CommandButton bubicacio 
            BackColor       =   &H00FF8080&
            Caption         =   "LAM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Index           =   1
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   540
            Width           =   1725
         End
         Begin VB.CommandButton bubicacio 
            BackColor       =   &H00F8FDB5&
            Caption         =   "REB"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Index           =   2
            Left            =   3870
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   540
            Width           =   1770
         End
         Begin VB.CommandButton bubicacio 
            BackColor       =   &H00EEE4D7&
            Caption         =   "SOL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Index           =   3
            Left            =   5715
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   540
            Width           =   1665
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H008080FF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   11610
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   30
            Width           =   375
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00808000&
            FillStyle       =   0  'Solid
            Height          =   465
            Left            =   30
            Top             =   0
            Width           =   12015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Escull la secció"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   195
            TabIndex        =   45
            Top             =   60
            Width           =   3330
         End
      End
      Begin VB.Frame framegrups 
         Caption         =   "Grups Estocs"
         Height          =   8865
         Left            =   3705
         TabIndex        =   34
         Top             =   9795
         Visible         =   0   'False
         Width           =   16035
         Begin VB.CommandButton bacceptargrup 
            Height          =   855
            Left            =   780
            Picture         =   "formtorerus.frx":3BB5
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Actualitzar dades."
            Top             =   7710
            Width           =   1650
         End
         Begin VB.ListBox llistagrups 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   33.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7215
            Left            =   180
            TabIndex        =   35
            Top             =   345
            Width           =   15630
         End
      End
      Begin VB.ListBox llistacanvis 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6000
         Left            =   13200
         TabIndex        =   32
         Top             =   5370
         Width           =   5550
      End
      Begin VB.CommandButton bactualitzar 
         Enabled         =   0   'False
         Height          =   885
         Left            =   17235
         Picture         =   "formtorerus.frx":3E8B
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Actualitzar dades."
         Top             =   450
         Width           =   1680
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   1365
         Top             =   360
      End
      Begin VB.ListBox llistapermoure 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5835
         Left            =   555
         TabIndex        =   29
         Top             =   5280
         Width           =   5820
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Pujar de LAM"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   4181
         Picture         =   "formtorerus.frx":4162
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   150
         Width           =   2010
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Baixar a LAM"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   8433
         Picture         =   "formtorerus.frx":483B
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   150
         Width           =   2010
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Baixar a IMP"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   6307
         Picture         =   "formtorerus.frx":4F07
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   150
         Width           =   2010
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Pujar d'IMP"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   2055
         Picture         =   "formtorerus.frx":55D3
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   150
         Width           =   2010
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   13995
         Top             =   480
      End
      Begin VB.Frame frameforats 
         Height          =   2865
         Left            =   60
         TabIndex        =   1
         Top             =   1500
         Width           =   18990
         Begin VB.CommandButton binformacio 
            Height          =   705
            Left            =   8760
            Picture         =   "formtorerus.frx":5CAC
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Informació de la bobina."
            Top             =   210
            Width           =   705
         End
         Begin VB.CommandButton Command7 
            Height          =   705
            Left            =   9540
            Picture         =   "formtorerus.frx":6976
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   210
            Width           =   705
         End
         Begin VB.CommandButton Command11 
            Height          =   705
            Left            =   15345
            Picture         =   "formtorerus.frx":6F00
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Neteja les caselles."
            Top             =   300
            Width           =   1530
         End
         Begin VB.CommandButton Command9 
            Height          =   705
            Left            =   14265
            Picture         =   "formtorerus.frx":7DD2
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Ubicació destí."
            Top             =   225
            Width           =   705
         End
         Begin VB.CommandButton Command8 
            Height          =   705
            Left            =   3735
            Picture         =   "formtorerus.frx":82B0
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Ubicació origen."
            Top             =   195
            Width           =   705
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Acceptar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1380
            Left            =   15330
            Picture         =   "formtorerus.frx":878E
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1035
            Width           =   1560
         End
         Begin VB.TextBox cforatdesti 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   60
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   11130
            TabIndex        =   4
            Top             =   945
            Width           =   3870
         End
         Begin VB.TextBox cbobina 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   50.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   5145
            TabIndex        =   3
            Top             =   945
            Width           =   5160
         End
         Begin VB.TextBox cforat 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   60
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   585
            TabIndex        =   2
            Top             =   945
            Width           =   3870
         End
         Begin VB.Image imatgediametre 
            Height          =   750
            Left            =   5400
            Picture         =   "formtorerus.frx":8A64
            ToolTipText     =   "Preguntarà el diametre al escanejar."
            Top             =   165
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Image Image2 
            Height          =   510
            Left            =   10380
            Picture         =   "formtorerus.frx":8D98
            Stretch         =   -1  'True
            Top             =   1365
            Width           =   525
         End
         Begin VB.Label Label3 
            Caption         =   "Forat destí"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   11355
            TabIndex        =   7
            Top             =   345
            Width           =   3240
         End
         Begin VB.Label Label2 
            Caption         =   "Bobina"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   6540
            TabIndex        =   6
            Top             =   345
            Width           =   2205
         End
         Begin VB.Label Label1 
            Caption         =   "Forat origen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   645
            TabIndex        =   5
            Top             =   315
            Width           =   3450
         End
         Begin VB.Image Image1 
            Height          =   510
            Left            =   4545
            Picture         =   "formtorerus.frx":946D
            Stretch         =   -1  'True
            Top             =   1365
            Width           =   525
         End
      End
      Begin VB.Frame frameforatslliures 
         Caption         =   "Forats lliures"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7035
         Left            =   6735
         TabIndex        =   50
         Top             =   4500
         Visible         =   0   'False
         Width           =   6240
         Begin VB.ListBox llistaforatslliures 
            Columns         =   4
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6000
            Left            =   165
            TabIndex        =   51
            Top             =   345
            Width           =   5895
         End
         Begin VB.Label etforats 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   120
            TabIndex        =   52
            Top             =   6345
            Width           =   6045
         End
      End
      Begin VB.Label etcomptadoractualitzacions 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   225
         Left            =   17790
         TabIndex        =   73
         Top             =   1335
         Width           =   555
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Dies Bobs:"
         Height          =   240
         Left            =   17325
         TabIndex        =   62
         Top             =   180
         Width           =   870
      End
      Begin VB.Label etrecordatori 
         BackStyle       =   0  'Transparent
         Caption         =   "Pensa a traspasar les dades..."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   13290
         TabIndex        =   55
         Top             =   4425
         Visible         =   0   'False
         Width           =   4470
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Bobines canviades de Situació"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13335
         TabIndex        =   33
         Top             =   4920
         Width           =   5355
      End
      Begin VB.Label etllistabobines 
         BackStyle       =   0  'Transparent
         Caption         =   "Llista de bobines per moure"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   555
         TabIndex        =   30
         Top             =   4800
         Width           =   8280
      End
      Begin VB.Label cnumtablet 
         BackStyle       =   0  'Transparent
         Caption         =   "Tablet Nº:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   165
         Width           =   1650
      End
      Begin VB.Label cnomoperari 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sense Operari"
         Height          =   330
         Left            =   -60
         TabIndex        =   8
         Top             =   1305
         Width           =   1335
      End
   End
   Begin VB.Menu mopcions 
      Caption         =   "Opcions"
      Begin VB.Menu m_demanarcmdiametre 
         Caption         =   "Demanar Cm diametre al esccanejar"
      End
      Begin VB.Menu mbobinaalforat 
         Caption         =   "Saber bobines dins d'un forat."
      End
   End
End
Attribute VB_Name = "formtorerus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vdiesbobanterior As Double
Function ultimcanvifamesde4hores() As Boolean
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select * from canvissituacio order by data asc")
   If rst.EOF Then GoTo fi
   If DateDiff("h", rst!Data, Now) > 7 Then
       If MsgBox("ATENCIÓ!!!" + Chr(10) + "HI HA CANVIS DE BOBINA QUE FA MES DE 4 HORES QUE NO S'HAN ACTUALITZAT." + Chr(10) + "ESTAS SEGUR QUE ENCARA SON VÀLIDS? REVISA-HO SISPLAU.", vbCritical + vbDefaultButton2 + vbYesNo, "A T E N C I Ó") = vbYes Then
            ultimcanvifamesde4hores = False
             Else: ultimcanvifamesde4hores = True
       End If
   End If
fi:
   Set rst = Nothing
End Function
Function mirarsihihacanvisdesituacio() As Boolean
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select * from canvissituacio")
   If Not rst.EOF Then mirarsihihacanvisdesituacio = True
   Set rst = Nothing
End Function

Private Sub bacceptargrup_Click()
  acceptargrup
End Sub
Sub copiaseguretatdades()
   Dim i As Integer
   On Error Resume Next
   MkDir App.Path + "\Backups"
   Set dbcomandes = Nothing
   Set dbstocks = Nothing
   On Error GoTo 0
   If existeix(App.Path + "\Backups\Torerus_20.mdb") Then Kill App.Path + "\Backups\Torerus_20.mdb"
   For i = 19 To 1 Step -1
     If existeix(App.Path + "\Backups\Torerus_" + atrim(i) + ".mdb") Then
        Rename App.Path + "\Backups\Torerus_" + atrim(i) + ".mdb", App.Path + "\Backups\Torerus_" + atrim(i + 1) + ".mdb"
     End If
   Next i
   FileCopy App.Path + "\Torerus.mdb", App.Path + "\Backups\Torerus_1.mdb"
End Sub
Sub demanar_ok_encarregat_per_actualitzar()
       formtorerus.SetFocus
       cpassword.Tag = ""
       cpassword = ""
       Framepassword.Tag = "password"
       Framepassword.Visible = True
       Framepassword.Top = 600
       Framepassword.Left = 6000
       While Framepassword.Visible
         DoEvents
       Wend
       If cpassword.Tag = "030201" Then escriure_ini "General", "comptaractualitzacions", "0", "comandes.ini": MsgBox "Comptador resetejat. Torna a provar-ho.", vbInformation, "Atenció"
      Framepassword.Visible = False
      comprovar_actualitzacions
End Sub
Private Sub bactualitzar_Click()
   Dim vcontador As Integer
   Dim vhihacanvisdesituacio As Boolean
   Dim vultimaactualitzacio As String
   
   If ultimcanvifamesde4hores Then Exit Sub
   'vhihacanvisdesituacio = mirarsihihacanvisdesituacio
     'aixó ja no utilitzem perquè potser que no hi hagi canvis pero vulguis actualitzar
   If Not comprovar_actualitzacions Then demanar_ok_encarregat_per_actualitzar: Exit Sub
   vhihacanvisdesituacio = True
       
   Set dbcomandes = Nothing
   Set dbstocks = Nothing
   Shell ("net time \\serverprodu /set /y")
   vresp = llegir_ini("Torerus", "Generartorerus", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini")
   vultimaactualitzacio = llegir_ini("Torerus", "horaultimaactualitzacio", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini")
   If vultimaactualitzacio = "" Or vultimaactualitzacio = "{[}]" Then
     vultimaactualitzacio = Now
     escriure_ini "Torerus", "horaultimaactualitzacio", atrim(Now), rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
     escriure_ini "Torerus", "diesactualitzaciobobs", atrim(combodies), rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
     escriure_ini "Torerus", "usuariTORERUS", atrim(numop), rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
   End If
   If vresp <> "No" And DateDiff("n", vultimaactualitzacio, Now) < 2 Then MsgBox "Hi ha una altra tablet comunicant espera un minut sisplau.", vbCritical, "Error": Set dbcomandes = OpenDatabase(App.Path + "\torerus.mdb"): Exit Sub
   cframetotal.Enabled = False
   formactualitzant.Show
   DoEvents
   escriure_ini "Torerus", "Generartorerus", "Copiant", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
   escriure_ini "Torerus", "horaultimaactualitzacio", atrim(Now), rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
   If vhihacanvisdesituacio Then
     FileCopy App.Path + "\Torerus.mdb", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "Torerus_tablet.mdb"
     copiaseguretatdades
   End If
   escriure_ini "Torerus", "Generartorerus", "Si", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
   vresp = ""
   While vcontador < 90 And (vresp <> "No" And vresp <> "ERROR ELIMINANT FITXER TEMPORAL")
      wait 1
      vresp = llegir_ini("Torerus", "Generartorerus", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini")
      vcontador = vcontador + 1
      If vresp <> "Processant" Then
         formactualitzant.etconnectant.Caption = "Connectant... (" + atrim(vcontador) + ")"
           Else: formactualitzant.etconnectant.Caption = "Processant... (" + atrim(vcontador) + ")"
      End If
   Wend
   If vresp = "ERROR ELIMINANT FITXER TEMPORAL" Then
      MsgBox "Hi ha hagut un error generant les dades al servidor, espera una estona i torna-ho a provar." + Chr(10) + "Gràcies", vbCritical, "Error"
      GoTo fi
   End If
   If vcontador >= 90 Then
      formactualitzant.Timer1.Enabled = False
      Unload formactualitzant
      MsgBox "Error connectant amb el servidor. Temps esgotat." + vbNewLine + "TORNA-HO A PROVAR D'AQUI UN MINUT.", vbCritical, "Error"
      GoTo fi
   End If
   If llegir_ini("Torerus", "ultimresultat", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini") = "OK" Then
     FileCopy rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "Torerus.mdb", App.Path + "\Torerus.mdb"
     Set dbcomandes = OpenDatabase(App.Path + "\torerus.mdb")
     dbcomandes.Execute "create index principal ON bobines([idpalet]);"
     dbcomandes.Execute "create index segon ON bobines([idbobina]);"
     dbcomandes.Execute "create index tercer ON bobines([sit]);"
     dbcomandes.Execute "create UNIQUE index primer ON foratsnous([foratvell]) with PRIMARY;"
     dbcomandes.Execute "create  index segon ON foratsnous([foratnou]);"
     comptar_lactualitzacio  'guardo la actualitzacio com a bona per controlar les actualitzacions al dia
   
   'converteix els numeros de forats vells amb els nous
      'dbcomandes.Execute "UPDATE bobines LEFT JOIN foratsnous ON bobines.Sit = foratsnous.foratvell SET bobines.Sit = [foratnou] where [foratnou]<>'';"
        Else: MsgBox "Hi ha hagut algun error al fer la importació.", vbCritical, "Error"
   End If
fi:
   cframetotal.Enabled = True
   Set dbcomandes = OpenDatabase(App.Path + "\torerus.mdb")
   On Error Resume Next
   dbcomandes.Execute "alter table prestatgesnous add column lliure bit"
   dbcomandes.Execute "alter table bobinesent add column modificat bit"
   dbcomandes.Execute "alter table bobinesgrups add column nomproveidor string"
   On Error GoTo 0
   'formactualitzant.etconnectant = ""
   Unload formactualitzant
   escriure_ini "Torerus", "Generartorerus", "No", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
   escriure_ini "Torerus", "horaultimaactualitzacio", "", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
   carregallistacanvis
   mirar_si_hi_ha_algu
   
End Sub
Function comprovar_actualitzacions() As Boolean
   Dim v As String
   v = llegir_ini("General", "comptaractualitzacions_dia", "comandes.ini")
   If IsDate(v) Then
     If DateDiff("d", v, Now) > 0 Then
        escriure_ini "General", "comptaractualitzacions_dia", Now, "comandes.ini"
        escriure_ini "General", "comptaractualitzacions", "0", "comandes.ini"
        comprovar_actualitzacions = True
        GoTo fi
     End If
       Else: escriure_ini "General", "comptaractualitzacions_dia", Now, "comandes.ini"
   End If
   If cadbl(llegir_ini("General", "comptaractualitzacions", "comandes.ini")) >= 12 Then
          comprovar_actualitzacions = False
           Else: comprovar_actualitzacions = True
   End If
fi:
   etcomptadoractualitzacions = cadbl(llegir_ini("General", "comptaractualitzacions", "comandes.ini"))
End Function
Sub comptar_lactualitzacio()
     escriure_ini "General", "comptaractualitzacions", atrim(cadbl(etcomptadoractualitzacions) + 1), "comandes.ini"
     etcomptadoractualitzacions = cadbl(llegir_ini("General", "comptaractualitzacions", "comandes.ini"))
End Sub
Private Sub binformacio_Click()
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim vpalet As String
  Dim vmetres As Double
  Dim vbob As String
  Dim vsumametres As Double
  Dim vcomandes As String
  Dim vdescmaterial As String
'  cbobina = "55129/10"
  separarpaletibobina cbobina, vpalet, vbob
  If cadbl(vpalet) = 0 Then Exit Sub
  Set rst = dbcomandes.OpenRecordset("select * from bobines where idpalet=" + vpalet + " and idbobina=" + vbob)
  If Not rst.EOF Then
      vsumametres = 0
      vcomandes = ""
      vdescmaterial = atrim(rst!descripcio) + Chr(13) + Chr(10) + "Ample:" + atrim(rst!ample) + " Esp:" + atrim(IIf(cadbl(rst!micres) > 0, rst!micres, cadbl(rst!grmsm2)))
      vmetres = rst!disponible
      'Set rst2 = dbcomandes.OpenRecordset("SELECT Parcials.idpalet, Parcials.idbobina, Sum(Parcials.metres) AS SumaDemetres From Parcials Where (((Parcials.utilitzada) = False)) GROUP BY Parcials.idpalet, Parcials.idbobina HAVING (((Parcials.idpalet)=" + vpalet + ") AND ((Parcials.idbobina)=" + vbob + "))")
      Set rst2 = dbcomandes.OpenRecordset("SELECT Parcials.idpalet, Parcials.idbobina,parcials.comanda, Parcials.metres From Parcials Where (((Parcials.utilitzada) = False)) and  (((Parcials.idpalet)=" + vpalet + ") AND ((Parcials.idbobina)=" + vbob + "))")
      
      While Not rst2.EOF
         vsumametres = vsumametres + cadbl(rst2!metres)
         vcomandes = vcomandes + " [" + atrim(rst2!comanda) + "]"
         rst2.MoveNext
      Wend
       vmetres = vmetres + vsumametres
       vdiametre = calculardiametre(IIf(cadbl(rst!micres) > 0, rst!micres, cadbl(rst!grmsm2)), vmetres, (rst!tamanycanutu))
       
      msgboxex "Informació", "La bobina " + cbobina + " Canutu: " + atrim(rst!tamanycanutu) + Chr(13) + Chr(10) + "té " + Trim(vmetres) + " metres (Ø" + atrim(vdiametre) + "mm)" + Chr(13) + Chr(10) + vdescmaterial + vbNewLine + "Comandes assignades: " + vbNewLine + vcomandes, 20
  End If
  Set rst = Nothing
End Sub
Function calculardiametre(vmicres As Double, vmetres As Double, vcanuto As Double) As Double
   Dim pi As Double
   Dim vdiametre As Double
   If vcanuto < 10 Then vcanuto = vcanuto + 2 'afegeixo l'amplada del cartrò del canutu
   If vcanuto >= 10 Then vcanuto = vcanuto + 2.8 'afegeixo l'amplada del cartrò del canutu
   vcanuto = vcanuto * 10
    pi = 4 * Atn(1)
    vdiametre = Sqr(((vmetres * vmicres * 4) / pi) + (vcanuto * vcanuto))
    calculardiametre = Redondejar(vdiametre, 0)
End Function
Sub msgboxex(vTitolFinestra As String, vEtiquetaMissatge As String, vFontSize)
      Load formmsgbox
      formmsgbox.Caption = vTitolFinestra
      formmsgbox.etiqueta = vEtiquetaMissatge
      formmsgbox.etiqueta.FontSize = vFontSize
      formmsgbox.Show 1
End Sub

Private Sub bubicacio_Click(Index As Integer)
   If frameubicacions.Tag = "origen" Then
      cforat = bubicacio(Index).Caption
   End If
   If frameubicacions.Tag = "desti" Then
      cforatdesti = bubicacio(Index).Caption
   End If
   frameubicacions.Tag = ""
   frameubicacions.Visible = False
   cbobina.SetFocus
End Sub

Private Sub busuari_Click()
  If Framepassword.Visible = True Then Exit Sub
  numop = 0
  cnomoperari = ""
  While numop = 0
   escullir_operari
  Wend
  cframetotal.Enabled = True
  cbobina.SetFocus
   busuari.BackColor = Command2.BackColor
  Timer2.Enabled = False
  mirar_si_hi_ha_algu
End Sub
Sub escullir_operari()
  Dim numoptmp As Integer
  Dim nomoptmp As String
  Dim rstpassword As Recordset
  Dim v As String
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where actiu<>0 AND maquina='T' order by codi "
  formseleccio.Caption = "Selecció d'Operari"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   numoptmp = cadbl(formseleccio.Data1.Recordset!codi)
   nomoptmp = atrim(formseleccio.Data1.Recordset!descripcio)
   Set rstpassword = formseleccio.Data1.Database.OpenRecordset("select * from operaris_contrasenyes where seccio='T' and operari=" + atrim(cadbl(formseleccio.Data1.Recordset!codi)))
   If Not rstpassword.EOF Then
       Unload formseleccio
       formtorerus.SetFocus
       cpassword.Tag = ""
       cpassword = ""
       Framepassword.Tag = "password"
       Framepassword.Visible = True
       Framepassword.Top = 600
       Framepassword.Left = 6000
       While Framepassword.Visible
         DoEvents
       Wend
       If rstpassword!contrasenya <> cpassword.Tag Then GoTo fi
   End If
  End If
  If numoptmp <> 0 Then
     cnomoperari = nomoptmp
     numop = numoptmp
      Else: If cadbl(numop) = 0 Then MsgBox "Has d'escullir un operari per treballar": Exit Sub
  End If
fi:
  Set rstpassword = Nothing
  Framepassword.Visible = False
End Sub

Private Sub cbobina_Change()
 
   posarcolortipusbobina cbobina
  
End Sub
Sub mirar_i_regularitzarDiametrebobina()
   Dim vdiametre As String
   Dim vdiametrenou As Double
   Dim vpalet As String
   Dim vbob As String
   Dim vmetres As Double
   Dim rst As Recordset
   etvermella = ""
   If m_demanarcmdiametre.Checked Or InStr(1, " IMP LAM REB SOL ", " " + cforat + " ") > 0 Then
       separarpaletibobina cbobina, vpalet, vbob
       If cadbl(vpalet) = 0 Then Exit Sub
       vdiametre = calcular_diametre(cadbl(vpalet), cadbl(vbob), vmetres)
       Set rst = dbcomandes.OpenRecordset("select * from comprovacio_diametres_picus where numpalet=" + atrim(vpalet) + " and bobina=" + atrim(vbob) + " and actualitzat=true")
       If Not rst.EOF Then If DateDiff("d", rst!Data, Now) < 6 Then etvermella = "Aquest pico ja el vas fer el dia " + atrim(rst!Data)
       vdiametrenou = demanar_valor_diametre_bobina("Entra el diametre de la bobina." + vbNewLine + "HAURIA DE TENIR. " + vdiametre + "cm " + atrim(vmetres) + " Mtrs")
       If cadbl(vdiametrenou) <> 0 Then
           dbcomandes.Execute "delete * from comprovacio_diametres_picus where numpalet=" + atrim(vpalet) + " and bobina=" + atrim(vbob) + " and actualitzat=false"
           vmetres = Redondejar(calcular_metresambdiametre(cadbl(vpalet), cadbl(vbob), cadbl(vdiametrenou)), 0)
           MsgBox "La bobina " + vpalet + "/" + vbob + " ara té " + atrim(vmetres) + "metres i un diametre de " + atrim(vdiametrenou) + vbNewLine + "AQUEST CANVIS SERAN EFECTIUS A LA NOVA SINCRONITZACIÓ.", vbInformation, "Nous metres"
           dbcomandes.Execute "insert into comprovacio_diametres_picus (numpalet,bobina,data, diametre,diametreanterior,metresnous) values (" + atrim(vpalet) + "," + vbob + ",now," + atrim(passaradecimalpunt(atrim(vdiametrenou))) + "," + passaradecimalpunt(atrim(vdiametre)) + "," + atrim(vmetres) + ")"
       End If
   End If
   etvermella = ""
End Sub
Function demanar_valor_diametre_bobina(vmsg As String) As Double
       Framemissatge.Visible = True
       Framemissatge.ZOrder 0
       etmissatgeframemissatge = vmsg
       cpassword.Tag = ""
       cpassword = ""
       Framepassword.Visible = True
       Framepassword.ZOrder 0
       Framepassword.Top = 1830
       Framepassword.Left = 11145
       Framepassword.Tag = ""
       Framemissatge.Top = Framepassword.Top - Framemissatge.Height + 80
       Framemissatge.Left = Framepassword.Left - 80
       cbotonum(11).Caption = ","
       If cpassword.Visible Then cpassword.SetFocus
       While Framepassword.Visible
         DoEvents
       Wend
        Framemissatge.Visible = False
       demanar_valor_diametre_bobina = cadbl(cpassword.Tag)
       cbotonum(11).Caption = "/"
End Function
Function calcular_mtrsdispreals(palet As Double, bobina As Double) As Double
   Dim rstb As Recordset
   Dim rstp As Recordset
   Dim total As Double
   'dbstocks.Execute "delete * from parcials where metres=0 and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina)
   Set rstb = dbcomandes.OpenRecordset("select mts from bobines where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
   Set rstp = dbcomandes.OpenRecordset("select sum(metres) as tmetres from parcials where utilitzada and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
   If Not rstb.EOF Then total = cadbl(rstb!mts)
   If Not rstp.EOF And Not rstb.EOF Then
      total = cadbl(rstb!mts) - cadbl(rstp!tmetres)
   End If
   calcular_mtrsdispreals = total
   Set rstb = Nothing
   Set rstp = Nothing
End Function
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
  Set rstp = dbcomandes.OpenRecordset("select micres,grmsm2 from palets where idpalet=" + atrim(palet))
  'Set rstb = dbstocks.OpenRecordset("select mts from bobines where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
  If Not rstp.EOF Then
    pi = 4 * Atn(1)
    vdiametre = vdiametre / 100
    canutu = canutu / 100
    micres = cadbl(rstp!micres)
    If micres = 0 Then micres = cadbl(rstp!grmsm2) * -1
    If micres = 0 Then GoTo fi
    If micres < 0 Then
       micres = (micres * -1)
     '  micres = micres / 1.2
    End If
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
Function calcular_diametre(palet As Double, bobina As Double, vmetres As Double, Optional canutu As Double) As String
  Dim rstp As Recordset
  Dim rstb As Recordset
  Dim metres As Double
  Dim micres As Double
  
  Dim diametre As Double
  Dim pi As Double
  If cadbl(canutu) = 0 Then canutu = 15.2
  metres = cadbl(calcular_mtrsdispreals(palet, bobina))
  vmetres = metres
  If metres <= 0 Then calcular_diametre = 0: Exit Function
  Set rstp = dbcomandes.OpenRecordset("select micres,grmsm2 from palets where idpalet=" + atrim(palet))
  Set rstb = dbcomandes.OpenRecordset("select tamanycanutu from bobines where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
  If rstb.EOF Then Exit Function
  If cadbl(rstb!tamanycanutu) > 0 Then canutu = cadbl(rstb!tamanycanutu)
  If canutu < 10 Then canutu = canutu + 2 'afegeixo l'amplada del cartrò del canutu
  If canutu >= 10 Then canutu = canutu + 2.8 'afegeixo l'amplada del cartrò del canutu
  If Not rstp.EOF Then
    pi = 4 * Atn(1)
    'canutu = (canutu / 2) * 10
    canutu = canutu * 10
    micres = cadbl(rstp!micres)
    If micres = 0 Then micres = cadbl(rstp!grmsm2) * -1
    If micres < 0 Then
       micres = (micres * -1)
       'micres = micres / 1.2
    End If
    'micres = (micres * 0.0001) / 100
    'vdiametre = Sqr(((vmetres * vmicres * 4) / pi) + (vcanuto * vcanuto))
    diametre = Sqr(((metres * micres * 4) / pi) + (canutu * canutu))
    'diametre = Sqr(((metres * micres) / pi) + (canutu * canutu)) * 200
    calcular_diametre = Redondejar(diametre, 0)
    calcular_diametre = calcular_diametre / 10
    'If cadbl(calcular_diametre) < 9 Then calcular_diametre = "0"
  End If
  Set rstp = Nothing
  Set rstb = Nothing
End Function

Sub separarpaletibobina(vnumbob As String, vpalet As String, vbob As String)
    If vnumbob = "" Then Exit Sub
    If InStr(1, vnumbob, "/") = 0 Then Exit Sub
    vpalet = cadbl(Mid(vnumbob, 1, InStr(1, vnumbob, "/") - 1))
    vbob = cadbl(substituirtot(vnumbob, vpalet + "/", ""))
End Sub
Sub posarcolortipusbobina(vnumbob As String)
   Dim rst As Recordset
   Dim vpalet As String
   Dim vbob As String
   If InStr(1, vnumbob, "/") = 0 Then cbobina.BackColor = QBColor(15): Exit Sub
   separarpaletibobina vnumbob, vpalet, vbob
   If vbob = 0 Then GoTo fi
   Set rst = dbcomandes.OpenRecordset("SELECT Parcials.idpalet, Parcials.idbobina, Parcials.utilitzada From Parcials WHERE idpalet=" + atrim(vpalet) + " and idbobina=" + vbob + " AND Parcials.utilitzada=True;")
   If Not rst.EOF Then
        cbobina.BackColor = &H8080FF      'vermell
      Else: cbobina.BackColor = &HC0FFC0      'verd
   End If
fi:
   Set rst = Nothing
   Set rst2 = Nothing
End Sub
Private Sub cbobina_GotFocus()
  If cbobina.BackColor = QBColor(15) Then cbobina.BackColor = QBColor(11)
End Sub

Private Sub cbobina_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then: KeyAscii = 0: cforatdesti.SetFocus
End Sub

Private Sub cbobina_LostFocus()
   cbobina = substituirperbarra(cbobina)
   If cbobina.BackColor = QBColor(11) Then cbobina.BackColor = QBColor(15)
   posarforatorigen
   mirar_i_regularitzarDiametrebobina
End Sub
Sub posarforatorigen()
   Dim rst As Recordset
   If cforat <> "" Then Exit Sub
   Set rst = dbcomandes.OpenRecordset("select sit from bobines where trim([idpalet])+'/'+trim([idbobina])='" + atrim(cbobina) + "'")
   If Not rst.EOF Then cforat = atrim(rst!Sit)
   Set rst = Nothing
End Sub

Private Sub cbotonum_Click(Index As Integer)
   If cbotonum(Index).Caption = "OK" Then Framepassword.Visible = False: GoTo fi
   cpassword.Tag = cpassword.Tag + cbotonum(Index).Caption
   If Framepassword.Tag = "password" Then
      cpassword = cpassword + "*"
       Else: cpassword = cpassword.Tag
   End If
   cpassword.SetFocus
fi:
End Sub

Private Sub cforat_GotFocus()
  cforat.BackColor = QBColor(11)
End Sub

Private Sub cforat_LostFocus()
  cforat.BackColor = QBColor(15)
End Sub

Private Sub cforatdesti_GotFocus()
  cforatdesti.BackColor = QBColor(11)
End Sub

Private Sub cforatdesti_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
      If Len(cforatdesti) > 3 Then cforatdesti = "": Exit Sub
       If Not IsNumeric(Mid(cforatdesti + "   ", 1, 1)) And cadbl(Mid(cforatdesti + "   ", 2, 2)) > 0 Then
            escullir_nivell UCase(cforatdesti)
          Else: MsgBox "Format de forat no correcte.", vbCritical, "Error"
       End If
       KeyCode = 0
  End If
End Sub

Private Sub cforatdesti_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub cforatdesti_LostFocus()
   cforatdesti.BackColor = QBColor(15)
End Sub

Private Sub cnumtablet_DblClick()
   Dim vt As String
   vt = InputBox("Entra el numero de tablet.")
   If cadbl(vt) > 0 Then
     escriure_ini "General", "numtablet", atrim(vt), "comandes.ini"
     cnumtablet = "Tablet Nº: " + vt
   End If
End Sub

Private Sub combodies_Click()
   Dim vcontrasenya As String
   vcontrasenya = InputBoxEx("Escriu la contrasenya per poder fer aquest canvi de dies.", "Atenció", , , , , , SPassword)
   If vcontrasenya <> "918273" Then cbobina.SetFocus: Exit Sub
   vdiesbobanterior = cadbl(combodies)
End Sub

Private Sub combodies_GotFocus()
   vdiesbobanterior = cadbl(combodies)
End Sub

Private Sub combodies_LostFocus()
  combodies = atrim(vdiesbobanterior)
End Sub

Private Sub Command1_Click()
  Dim vbobinaescanejada As String
  Dim vbobines(90) As String
  Dim vnometiqueta As String
  vnometiqueta = etllistabobines
  If cbobina = "" Then cbobina.SetFocus: Exit Sub
  If cbobina <> "" And cforatdesti <> "" Then
    If Not existeixlabobina(cbobina) Then MsgBox "Aquesta bobina no existeix a la base de dades.", vbCritical, "Error": Exit Sub
    If jashamogutlabobina(cbobina) Then
       If MsgBox("Aquesta bobina ja l'has canviat d'ubicació." + Chr(10) + "VOLS CANVIAR-LA D'UBICACIO UNA ALTRA VEGADA?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
    End If
    If cforat = cforatdesti Then MsgBox "No pots moure una bobina amb el mateix ORIGEN --> DESTÍ", vbCritical, "ERROR": Exit Sub
    'vbobinaescanejada = InputBox("Escaneja la bobina per assegurar que sigui la correcte.", "Escaneja bobina")
    vbobinaescanejada = atrim(demanarbobina)
    If StrPtr(vbobinaescanejada) = 0 Then Exit Sub
    vbobinaescanejada = substituirperbarra(vbobinaescanejada)
    cbobina = atrim(cbobina)
    If atrim(cbobina) = vbobinaescanejada Then
         vbobines(0) = cbobina
         carregarbobinesdelmateixpalet vbobines
         ensenyarmissatgesihihamesdunabobina vbobines
         guardarcanvidesituacio vbobines
         netejarcamps
         cbobina.SetFocus
           Else: MsgBox "La bobina escanejada no coincideix amb la que es vol moure.", vbCritical, "Error"
    End If
       Else: MsgBox "No hi ha les dades necessaries per canviar el destí d'una bobina", vbCritical, "Error"
  End If
  carregallistacanvis
  carregarllista llistapermoure.Tag
  etllistabobines = vnometiqueta
End Sub
Sub ensenyarmissatgesihihamesdunabobina(vbobines As Variant)
    Dim i As Byte
    Dim vmsg As String
    i = 0
    While vbobines(i) <> ""
      vmsg = vmsg + " " + vbobines(i)
      i = i + 1
    Wend
    If i > 1 Then
       If MsgBox("En aquest palet hi van " + atrim(i) + " bobines." + Chr(10) + vmsg + Chr(10) + "VOLS MOURE-LES TOTES A LA UBICACIÓ " + atrim(cforatdesti) + "?", vbExclamation + vbDefaultButton2 + vbYesNo, "PALET AMB MES D'UNA BOBINA") = vbNo Then
          i = 1
          While vbobines(i) <> ""
             vbobines(i) = ""
             i = i + 1
          Wend
       End If
    End If
    
End Sub
Sub carregarbobinesdelmateixpalet(vbobines As Variant)
   Dim rst As Recordset
   Dim i As Byte
   Set rst = dbcomandes.OpenRecordset("select * from parcials where trim(idpalet)+'/'+trim(idbobina)='" + vbobines(0) + "' and utilitzada=true")
   If rst.EOF Then  'si es jumbo miro les bobines del mateix palet de proveidor
      Set rst = dbcomandes.OpenRecordset("select numpaletpro,idpalet from bobines where trim(idpalet)+'/'+trim(idbobina)='" + vbobines(0) + "'")
      If Not rst.EOF Then
         Set rst = dbcomandes.OpenRecordset("select * from bobines where Numpaletpro='" + atrim(rst!Numpaletpro) + "' and idpalet=" + atrim(rst!idpalet))
         i = 1
         While Not rst.EOF
            If atrim(rst!idpalet) + "/" + atrim(rst!idbobina) <> vbobines(0) Then
               vbobines(i) = atrim(rst!idpalet) + "/" + atrim(rst!idbobina)
               i = i + 1
            End If
            rst.MoveNext
         Wend
      End If
   End If
   Set rst = Nothing
End Sub
Function demanarbobina() As String
       cpassword.Tag = ""
       cpassword = ""
       Framepassword.Visible = True
       Framepassword.Top = 1230
       Framepassword.Left = 11145
       Framepassword.Tag = ""
       cpassword.SetFocus
       While Framepassword.Visible
         DoEvents
       Wend
       demanarbobina = cpassword.Tag
End Function
Function existeixlabobina(vbobina As String) As Boolean
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select trim(idpalet)+'/'+trim(idbobina) as vbobina from bobines")
  rst.FindFirst "vbobina='" + vbobina + "'"
  If Not rst.NoMatch Then existeixlabobina = True
  Set rst = Nothing
End Function
Function jashamogutlabobina(vbobina As String) As Boolean
  Dim rstmogudes As Recordset
  Set rstmogudes = dbcomandes.OpenRecordset("select * from CanvisSituacio")
  rstmogudes.FindFirst "bobina='" + vbobina + "'"
  If Not rstmogudes.NoMatch Then jashamogutlabobina = True
  Set rstmogudes = Nothing
End Function
Sub guardarcanvidesituacio(vbobines As Variant)
   Dim i As Byte
   i = 0
   While vbobines(i) <> ""
    dbcomandes.Execute "insert into CanvisSituacio (sitorigen,sitdesti,bobina,data,operari) values ('" + treure_apostruf(cforat) + "','" + treure_apostruf(cforatdesti) + "','" + treure_apostruf(vbobines(i)) + "',now,'" + Mid(cnomoperari, 1, 10) + "')"
    i = i + 1
   Wend
End Sub
Sub carregallistacanvis()
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select * from canvissituacio order by data desc")
   llistacanvis.Clear
   While Not rst.EOF
       llistacanvis.AddItem rst!bobina + " (" + rst!sitorigen + " --> " + rst!sitdesti + ")"
       llistacanvis.ItemData(llistacanvis.NewIndex) = rst!id
       rst.MoveNext
   Wend
End Sub

Private Sub Command10_Click()
   frameubicacions.Visible = False
   frameubicacions.Tag = ""
   
End Sub

Private Sub Command11_Click()
   netejarcamps
End Sub
Sub netejarcamps()
   cforat = ""
   cbobina = ""
   cforatdesti = ""
   etllistabobines = "Llista de bobines per moure."
   llistapermoure.Clear
   cbobina.SetFocus
   Command2.BackColor = busuari.BackColor: Command3.BackColor = busuari.BackColor: Command4.BackColor = busuari.BackColor
  Command5.BackColor = busuari.BackColor: Command6.BackColor = busuari.BackColor
End Sub
Sub carregarllista(vnomtaula As String)
  Dim rst As Recordset
  Dim rstmogudes As Recordset
  Dim rstbobines As Recordset
  Dim vesungrup As Boolean
  Dim vliniafeta As Boolean
  Dim esurgent As String
  Dim vhihaurgent As Boolean
  llistapermoure.Clear
  If vnomtaula = "" Then Exit Sub
  If cadbl(Mid(vnomtaula, 1, 4)) > 0 Then
    vesungrup = True
    llistapermoure.Tag = llistagrups.Tag 'Mid(vnomtaula, 1, 4)
    vnomtaula = "select * from bobinesgrups where comanda='" + atrim(Mid(llistagrups.Tag, 1, 4)) + "'"
      Else: llistapermoure.Tag = vnomtaula
  End If
  
  Set rst = dbcomandes.OpenRecordset(vnomtaula, , ReadOnly)
  Set rstmogudes = dbcomandes.OpenRecordset("select * from CanvisSituacio")
  Set rstbobines = dbcomandes.OpenRecordset("select * from bobines")
  
  While Not rst.EOF
     rstmogudes.FindFirst "bobina='" + atrim(rst!idpalet) + "/" + atrim(rst!idbobina) + "'"
     If rstmogudes.NoMatch Then
         rstbobines.FindFirst ("idpalet=" + atrim(rst!idpalet) + " and idbobina=" + atrim(rst!idbobina))
         If vesungrup Then
            If rstbobines.NoMatch Then GoTo proxim
            'If Not IsNumeric(Mid(rstbobines!Sit, 2, 1)) Or (UCase(Mid(rstbobines!Sit, 1, 1)) <> "F" And UCase(Mid(rstbobines!Sit, 1, 1)) <> "G") Then GoTo proxim
            If Len(rstbobines!Sit) > 1 And (Not IsNumeric(Mid(rstbobines!Sit, 2, 1)) Or IsNumeric(Mid(rstbobines!Sit, 1, 1))) Then GoTo proxim
         End If
         vsit = ""
         If Not rstbobines.NoMatch Then
            vsit = " (" + UCase(atrim(rstbobines!Sit)) + ")"
            If atrim(rstbobines!Sit) = "" And InStr(1, vnomtaula, "bobinesgrups") = 0 Then vsit = " (" + atrim(rst!nomproveidor) + ")"
         End If
         If InStr(1, vnomtaula, "baixarLAM") > 0 Then If rst!ordre = 999 And Not vliniafeta Then vliniafeta = True: llistapermoure.AddItem "-NO PLAN-"
         If vliniafeta Then vsit = " #" + vsit
         esurgent = IIf(Mid(rst!nomproveidor, 1, 1) = "*", "*", "")
         llistapermoure.AddItem esurgent + atrim(rst!idpalet) + "/" + atrim(rst!idbobina) + vsit
         If esurgent = "*" Then vhihaurgent = True
     End If
proxim:
     rst.MoveNext
  Wend
  If vhihaurgent Then MsgBox "ATENCIÓ... HI HA BOBINES URGENT PER BAIXAR A IMPRESORES." + vbNewLine + vbNewLine + "SON LES BOBINES AMB ASTERISC DAVANT EX: *45678/1 " + vbNewLine + vbNewLine + "BAIXA-LES PRIMER I FER ACTUALITZAR."
  Set rst = Nothing
End Sub

Private Sub Command12_Click()
  If frameforatslliures.Visible Then
        frameforatslliures.Visible = False
        command12.BackColor = busuari.BackColor
         Else:
           frameforatslliures.Visible = True
           command12.BackColor = QBColor(11)
           carregarforatslliures
  End If
End Sub
Sub carregarforatslliures()
   Dim rst As Recordset
   Dim rstp As Recordset
   Dim rstmogudes As Recordset
   
   On Error Resume Next
   dbcomandes.Execute "alter table prestatgesnous add column lliure bit"
   On Error GoTo 0
   etforats = ""
   Set rstmogudes = dbcomandes.OpenRecordset("select * from CanvisSituacio where isnumeric(mid([sitdesti],2,1))")
   Set rst = dbcomandes.OpenRecordset("select * from prestatgesnous")
   Set rstp = dbcomandes.OpenRecordset("select * from bobines where isnumeric(mid([sit],2,1))")
   dbcomandes.Execute "update prestatgesnous set lliure=false"
   dbcomandes.Execute "update prestatgesnous set lliure=true where trim(estanteria)&format(columna,'00')&trim(fila) in (select sitorigen from canvissituacio)"
   vcount = 0
   While Not rst.EOF
     rstp.FindFirst "sit='" + atrim(rst!estanteria) + Format(rst!columna, "00") + atrim(rst!fila) + "'"
     If rstp.NoMatch Then
        rstmogudes.FindFirst "sitdesti='" + atrim(rst!estanteria) + Format(rst!columna, "00") + atrim(rst!fila) + "'"
        If rstmogudes.NoMatch Then
          dbcomandes.Execute "update prestatgesnous set lliure=true where estanteria='" + atrim(rst!estanteria) + "' and columna=" + atrim(rst!columna) + " and fila=" + atrim(rst!fila)
        End If
     End If
     rst.MoveNext
   Wend
   
   'poso els totals
   Set rst = dbcomandes.OpenRecordset("select count(*) as Tforatslliures,estanteria from prestatgesnous where lliure group by estanteria")
   While Not rst.EOF
      etforats = etforats + rst!estanteria + ":" + atrim(rst!Tforatslliures) + " "
      rst.MoveNext
   Wend
   Set rst = dbcomandes.OpenRecordset("select * from prestatgesnous where lliure order by estanteria,columna,fila")
   llistaforatslliures.Clear
   While Not rst.EOF
      llistaforatslliures.AddItem atrim(rst!estanteria) + Format(rst!columna, "00") + atrim(rst!fila)
      rst.MoveNext
   Wend
   Set rst = Nothing
   Set rstp = Nothing
   Set rstmogudes = Nothing
End Sub

Private Sub Command13_Click()
  Dim v As Double
   demanar_password_packing
   v = cadbl(demanarbobina)
   If v > 0 Then
     carregarllista "select * from parcials where comanda='" + atrim(v) + "'"
     etllistabobines = "Llista de bobines Packing-List " + atrim(v)
   End If
End Sub
Sub demanar_password_packing()
   cpassword.Tag = ""
       cpassword = ""
       etmissatgepassword = "PARLA AMB EN PACO / MARC PER PODER TREURE MATERIAL. ENTRA EL PASSWORD"
       Framepassword.Tag = "password"
       Framepassword.Visible = True
       Framepassword.Top = 600
       Framepassword.Left = 6000
       While Framepassword.Visible
         DoEvents
       Wend
       etmissatgepassword = ""
       If cpassword.Tag <> "998876" Then MsgBox "Contrasenya no vàlida", vbCritical, "Error": Exit Sub
       
End Sub
Private Sub Command14_Click()
  If Len(cpassword.Tag) = 0 Then Exit Sub
  cpassword.Tag = Mid(cpassword.Tag, 1, Len(cpassword.Tag) - 1)
  If Framepassword.Tag = "password" Then
     cpassword = Mid(cpassword, 1, Len(cpassword) - 1)
      Else: cpassword = Mid(cpassword, 1, Len(cpassword) - 1)
  End If
  
End Sub

Private Sub Command15_Click()
       cpassword.Tag = ""
       cpassword = ""
       Framepassword.Tag = "password"
       Framepassword.Visible = True
       Framepassword.Top = 600
       Framepassword.Left = 6000
       While Framepassword.Visible
         DoEvents
       Wend
       If cpassword.Tag <> "9999" Then MsgBox "Contrasenya no vàlida", vbCritical, "Error": Exit Sub
       creartaulafitxers
       crearllistadefitxers
       Load formseleccio
       formseleccio.Data1.DatabaseName = cami
       formseleccio.Data1.RecordSource = "select data,nomfitxer from llista_fitxers order by data desc"
       formseleccio.Caption = "Selecció fitxer a recuperar"
       formseleccio.refrescar
       formseleccio.DBGrid2.Columns(1).Visible = False
       formseleccio.DBGrid2.Columns(0).Width = 5000
       formseleccio.Show 1
       If seleccioret = 1 Then
           If MsgBox("Aquesta operació substituirà el fitxer actual de dades per la copia de les" + Chr(10) + Chr(13) + formseleccio.DBGrid2.Columns(0) + Chr(10) + Chr(13) + "Es correcte?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
            Set dbcomandes = Nothing
            Set dbstocks = Nothing
            r = formseleccio.DBGrid2.Columns(1)
            Unload formseleccio
            Kill App.Path + "\Torerus.mdb"
            Copiar_Fitxer r, App.Path + "\Torerus.mdb"
            Set dbcomandes = OpenDatabase(App.Path + "\torerus.mdb")
            carregallistacanvis
            
            MsgBox "Copia restaurada."
           End If
       End If
  
End Sub
Sub crearllistadefitxers()
   Dim r As String
   r = Dir(App.Path + "\Backups\*.*")
   While r <> ""
      If Mid(r, 1, 1) <> "." Then
         dbcomandes.Execute "insert into llista_fitxers (data,nomfitxer) values ('" + Trim(FileDateTime(App.Path + "\Backups\" + r)) + "','" + App.Path + "\Backups\" + r + "')"
      End If
      r = Dir
   Wend
End Sub
Sub creartaulafitxers()
     On Error Resume Next
     dbcomandes.Execute "create table llista_fitxers (data date,nomfitxer text)"
     dbcomandes.Execute "delete * from llista_fitxers"
End Sub

Private Sub Command16_Click()
  formrevisarentregues.Show 1
End Sub

Private Sub Command2_Click()
  carregarllista "llistatperpujarIMP"
   etllistabobines = "Llista de bobines per moure. Pujar IMP"
  Command2.BackColor = busuari.BackColor: Command3.BackColor = busuari.BackColor: Command4.BackColor = busuari.BackColor
  Command5.BackColor = busuari.BackColor: Command6.BackColor = busuari.BackColor
  Command2.BackColor = QBColor(11)
End Sub

Private Sub Command3_Click()
    carregarllista "llistatperbaixarIMP"
    etllistabobines = "Llista de bobines per moure. Baixar IMP"
    Command2.BackColor = busuari.BackColor: Command3.BackColor = busuari.BackColor: Command4.BackColor = busuari.BackColor
  Command5.BackColor = busuari.BackColor: Command6.BackColor = busuari.BackColor
  Command3.BackColor = QBColor(11)
End Sub

Private Sub Command4_Click()

carregarllista "SELECT * FROM llistatperbaixarLAM order by ordre"
etllistabobines = "Llista de bobines per moure. Baixar LAM"
Command2.BackColor = busuari.BackColor: Command3.BackColor = busuari.BackColor: Command4.BackColor = busuari.BackColor
  Command5.BackColor = busuari.BackColor: Command6.BackColor = busuari.BackColor
  Command4.BackColor = QBColor(11)
End Sub

Private Sub Command5_Click()
   carregarllista "llistatperpujarLAM"
  etllistabobines = "Llista de bobines per moure. Pujar LAM"
  Command2.BackColor = busuari.BackColor: Command3.BackColor = busuari.BackColor: Command4.BackColor = busuari.BackColor
  Command5.BackColor = busuari.BackColor: Command6.BackColor = busuari.BackColor
  Command5.BackColor = QBColor(11)
End Sub
Sub escullir_nivell(vcolumna As String, Optional vforatescullit As String)
   Dim rst As Recordset
   Load formnivell
   Set rst = dbcomandes.OpenRecordset("select * from prestatgesnous where estanteria='" + atrim(Mid(vcolumna, 1, 1)) + "' and columna=" + atrim(cadbl(Mid(vcolumna, 2, 2))))
  ' formnivell.Image11(0).Visible = True
   While Not rst.EOF
      formnivell.foratocupat(cadbl(rst!fila) - 1).Visible = True
      'formnivell.Image11(cadbl(rst!fila) + 1).Visible = True
      rst.MoveNext
   Wend
   Set rst = dbcomandes.OpenRecordset("select * from bobines where sit like '" + atrim(vcolumna) + "*'")
     While Not rst.EOF
       If cadbl(Mid(rst!Sit, 4, 1)) > 0 Then formnivell.foratocupat(cadbl(Mid(rst!Sit, 4, 1)) - 1).BackColor = &H8080FF
       rst.MoveNext
     Wend
   If atrim(vforatescullit) <> "" Then formnivell.Frame1.Tag = vforatescullit
   formnivell.Show 1
   If vforatescullit <> "" Then
        vforatescullit = UCase(vforatescullit + formnivell.Tag)
         Else:
           cforatdesti = UCase(cforatdesti + formnivell.Tag)
           If formnivell.Tag = "" Then cforatdesti = ""
   End If
   Unload formnivell
   Set rst = Nothing
End Sub

Private Sub Command6_Click()
  Command2.BackColor = busuari.BackColor: Command3.BackColor = busuari.BackColor: Command4.BackColor = busuari.BackColor
  Command5.BackColor = busuari.BackColor: Command6.BackColor = busuari.BackColor
  Command6.BackColor = QBColor(11)
  ensenyargrups
End Sub
Sub ensenyargrups()
   Dim rst As Recordset
   Dim vsql As String
   Dim rstnomgrup As Recordset
   Dim vnomgrup As String
   Dim vsecciogrup As String
   
   If framegrups.Visible = True Then framegrups.Visible = False: GoTo fi
   llistagrups.Tag = ""
   'vsql = "SELECT DISTINCT bobinesgrups.comanda, nomgrups.nomdelgrup, nomgrups.seccio FROM nomgrups INNER JOIN (bobinesgrups INNER JOIN bobines ON (bobinesgrups.idpalet = bobines.Idpalet) AND (bobinesgrups.idbobina = bobines.Idbobina)) ON nomgrups.numerogrup = cdbl(bobinesgrups.comanda) "
   'vsql = vsql + " WHERE  IsNumeric(Mid([Sit],2,1))<>False or ucase(Mid([Sit]&' ',1,1))=' ' OR ucase(Mid([Sit],1,1))='F' OR ucase(Mid([Sit],1,1))='G';"
   vsql = "SELECT DISTINCT bobinesgrups.comanda FROM bobinesgrups LEFT JOIN bobines ON (bobinesgrups.idbobina = bobines.Idbobina) AND (bobinesgrups.idpalet = bobines.Idpalet) "
   vsql = vsql + " WHERE (((bobines.Sit) Like '[0-9]*'  Or (bobines.Sit) Like '[A-Z][0-9]*' Or (bobines.Sit) Like '[A-Z]')) or (bobines.sit) is null AND ((bobines.Idpalet) Is Not Null);" 'Or (bobines.Sit) Like 'F[0-9]*'

   'Clipboard.Clear
   'Clipboard.SetText vsql
   ratoli "espera"
   Set rst = dbcomandes.OpenRecordset(vsql)
   llistagrups.Clear
   If rst.EOF Then
      MsgBox "No hi ha res per baixar de cap grup", vbInformation, "Estocs"
      GoTo fi
   End If
   Set rstnomgrup = dbcomandes.OpenRecordset("select * from nomgrups")
   While Not rst.EOF
       vnomgrup = ""
       vsecciogrup = ""
       rstnomgrup.FindFirst "numerogrup=" + atrim(cadbl(rst!comanda))
       If Not rstnomgrup.EOF Then vnomgrup = atrim(rstnomgrup!nomdelgrup): vsecciogrup = atrim(rstnomgrup!seccio)
       llistagrups.AddItem IIf(atrim(vsecciogrup) = "I", "IMP", IIf(atrim(vsecciogrup) = "L", "LAM", IIf(atrim(vsecciogrup) = "R", "REB", IIf(atrim(vsecciogrup) = "S", "SOL", "")))) + " " + IIf(atrim(vsecciogrup) = "S" Or atrim(vsecciogrup) = "R", "", atrim(rst!comanda)) + " - " + atrim(vnomgrup)
       llistagrups.ItemData(llistagrups.NewIndex) = cadbl(rst!comanda)
       rst.MoveNext
   Wend
   framegrups.Left = 1750
   framegrups.Top = 1335
   framegrups.Visible = True
   llistagrups.SetFocus
   
fi:
   Set rst = Nothing
   Set rstnomgrup = Nothing
   ratoli "normal"
End Sub


Sub acceptargrup()
  Dim vnomgrup As String
  If llistagrups.ListIndex = -1 Then Exit Sub
  vnomgrup = atrim(llistagrups.ItemData(llistagrups.ListIndex))
  If cadbl(llistagrups.ItemData(llistagrups.ListIndex)) = 5 Or cadbl(llistagrups.ItemData(llistagrups.ListIndex)) = 6 Then
       llistagrups.Tag = atrim(llistagrups.ItemData(llistagrups.ListIndex))
       vnomgrup = IIf(cadbl(llistagrups.Tag) = 5, "REB", IIf(cadbl(llistagrups.Tag) = 6, "SOL", ""))
         Else: llistagrups.Tag = atrim(llistagrups.ItemData(llistagrups.ListIndex)) + "llistatperbaixar" + Mid(atrim(llistagrups.List(llistagrups.ListIndex)), 1, 3)
  End If
  carregarllista atrim(llistagrups.ItemData(llistagrups.ListIndex))
  etllistabobines = "Llista de bobines per moure. (Estoc " + vnomgrup + ")"
  
  framegrups.Visible = False
End Sub

Private Sub Command7_Click()
       cforat = ""
       cbobina = ""
       cpassword.Tag = ""
       cpassword = ""
       Framepassword.Visible = True
       Framepassword.Top = 1230
       Framepassword.Left = 11145
       Framepassword.Tag = ""
       cpassword.SetFocus
       While Framepassword.Visible
         DoEvents
       Wend
       cbobina = cpassword.Tag
       cbobina_LostFocus
       
End Sub

Private Sub Command8_Click()
   If frameubicacions.Visible Then frameubicacions.Visible = False: Exit Sub
   frameubicacions.Left = 3100
   frameubicacions.Top = 2870
   frameubicacions.Visible = True
   frameubicacions.Tag = "origen"
   cbobina.SetFocus
End Sub

Private Sub Command9_Click()
   If frameubicacions.Visible Then frameubicacions.Visible = False: Exit Sub
   frameubicacions.Left = 3100
   frameubicacions.Top = 2870
   frameubicacions.Visible = True
   frameubicacions.Tag = "desti"
   cbobina.SetFocus
End Sub

Private Sub cpassword_KeyPress(KeyAscii As Integer)
  If Framepassword.Tag = "password" Then
     KeyAscii = 0
      Else: If KeyAscii > 21 Then cpassword.Tag = cpassword + Chr(KeyAscii)
  End If
  If KeyAscii = 13 Then
      KeyAscii = 0
      Framepassword.Visible = False
  End If
  
End Sub

Private Sub Form_Activate()
  'MsgBox calcular_metresambdiametre(54890, 2, 54)
     comprovar_actualitzacions
   If cnomoperari = "Sense Operari" Then busuari_Click
   
End Sub

Private Sub Form_DblClick()
 ' MsgBox Trim(Me.Width) + " - " + Trim(Me.Height)
End Sub

Function hihaunerroralabasededadeslocal(vnombd As String)
   Dim db As Database
   On Error GoTo er
   Set db = OpenDatabase(vnombd, , True)
   Set db = Nothing
   Exit Function
er:
   hihaunerroralabasededadeslocal = True
   Set db = Nothing
   
End Function
Private Sub Form_Load()
  'Dim cami_llistats As String
  cami_llistats = llegir_ini("General", "rutallistats", "comandes.ini")
  'cami = llegir_ini("General", "cami", "comandes.ini")
  cami = App.Path + "\torerus.mdb"
  
  numtablet = cadbl(llegir_ini("General", "numtablet", "comandes.ini"))
  If numtablet = 0 Then
    numtablet = cadbl(InputBox("Entra el numero de tablet que treballes", "Tablet"))
    If cadbl(numtablet) = 0 Then End
    escriure_ini "General", "numtablet", atrim(numtablet), "comandes.ini"
  End If
  cnumtablet = "Tablet Nº: " + atrim(numtablet)
  If existeix(App.Path + "\torerus.mdb") Then
      If hihaunerroralabasededadeslocal(App.Path + "\torerus.mdb") Then
          MsgBox "HI HA UN ERROR A LA BASE DE DADES LOCAL, LA ELIMINARÉ PER PODER CONTINUAR TREBALLANT." + Chr(10) + "REVISA QUE TOTS ELS CANVIS ESTIGUESSIN FETS SISPLAU.", vbCritical, "ERROR"
          Rename App.Path + "\torerus.mdb", App.Path + "\Torerus_" + Format(Now, "ddmmyyhhnnss") + ".mdb"
      End If
  End If
  If Not existeix(App.Path + "\torerus.mdb") Then
      If MsgBox("No hi ha la BD local de torerus." + Chr(10) + "Per continuar amb el programa s'ha de copiar del servidor." + Chr(10) + "VOLS QUE LA COPI-HI ARA?", vbInformation + vbDefaultButton2 + vbYesNo, "ATENCIO") = vbNo Then End
      FileCopy rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "Torerus.mdb", App.Path + "\Torerus.mdb"
      wait 5
  End If
  Set dbcomandes = OpenDatabase(App.Path + "\torerus.mdb")
  vdiesbobanterior = 2
  
  'Set dbstocks = OpenDatabase(App.Path + "\Palets.mdb")
'  Set dbbaixes = OpenDatabase(rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "baixes.mdb")
  If ferPing("serverprodu") Then
    If existeix(rutadelfitxer(cami_llistats) + "torerus.exe") Then
      If FileDateTime(rutadelfitxer(cami_llistats) + "torerus.exe") <> FileDateTime(App.Path + "\torerus.exe") Then
        On Error GoTo fi
        If existeix(App.Path + "\torerus2.exe") Then Kill App.Path + "\torerus2.exe"
        Rename App.Path + "\Torerus.exe", App.Path + "\Torerus2.exe"
        FileCopy rutadelfitxer(cami_llistats) + "Torerus.exe", App.Path + "\Torerus.exe"
        MsgBox "S'ha canviat de versió del programa." + Chr(10) + "ARA ES TANCARÀ, TORNA A OBRIR-LO SIUSPLAU.", vbInformation, "NOVA VERSIÓ"
         End
      End If
  End If
  End If
 carregallistacanvis
 Timer1.Enabled = True
 Timer2.Enabled = True
  On Error Resume Next
 dbcomandes.Execute "alter table bobinesent add column modificat bit"
  On Error GoTo 0
 Exit Sub
fi:
 MsgBox "Hi ha un error al actualitzar, prova de reiniciar la tauleta i torna-ho a provar", vbCritical, "Error"
End Sub
Sub Rename(vfitxerorigen As String, vfitxerdesti As String)
   Dim wShell As Object
   Set wShell = CreateObject("Scripting.FileSystemObject")
   wShell.movefile vfitxerorigen, vfitxerdesti
   Set wShell = Nothing
End Sub

Private Sub Form_Paint()
  On Error Resume Next
  Form1.Left = -100
  Form1.Top = 1
  cbobina.SetFocus
End Sub
Sub comprovar_connexio()
   If Not ferPing("serverprodu") Then
       bactualitzar.Enabled = False
         Else: bactualitzar.Enabled = True
   End If
End Sub
Function ferPing(vServer)

  'This function will return TRUE or FALSE after pinging a server and
  'checking it's response.
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com
    Dim oShell, oFSO
    Dim sTemp, sTempFile
    Dim fFile
    Dim sResults
    
    On Error Resume Next
    Const OpenAsDefault = -2
    Const FailIfNotExist = 0
    Const ForReading = 1

    Set oShell = CreateObject("WScript.Shell")
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    sTemp = oShell.ExpandEnvironmentStrings("%TEMP%")
    sTempFile = sTemp & "\runresult.tmp"

    oShell.Run "%comspec% /c ping -n 2 " & vServer & ">" & sTempFile, 0, True

    Set fFile = oFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, _
    OpenAsDefault)

    sResults = fFile.ReadAll
    fFile.Close
    oFSO.DeleteFile (sTempFile)
            
    ferPing = (InStr(sResults, "TTL=") > 0)
    
    Set oShell = Nothing
    Set oFSO = Nothing

End Function

Function ping(strComputer)
   Dim objshell
    ping = False

    Set objshell = CreateObject("WScript.Shell")
    Set objExec = objshell.Exec("%comspec% /c ping.exe " & strComputer & " -n 1 -w 100")
    Do Until objExec.StdOut.AtEndOfStream
        strLine = objExec.StdOut.ReadLine
        If (InStr(strLine, "Reply")) Then
            ping = True
            Exit Function
        End If
    Loop
End Function
Private Sub llistabobines_Click()

End Sub

Private Sub jdslge_Click()

End Sub

Private Sub llistacanvis_DblClick()
If MsgBox("Vols eliminar aquest moviment de bobina?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
       dbcomandes.Execute "delete * from CanvisSituacio where id=" + atrim(llistacanvis.ItemData(llistacanvis.ListIndex))
       carregallistacanvis
       carregarllista llistapermoure.Tag
   End If
End Sub

Private Sub llistagrups_DblClick()
   acceptargrup
End Sub

Private Sub llistagrups_LostFocus()
  If Screen.ActiveControl.Name = "bacceptargrup" Then Exit Sub
  framegrups.Visible = False
End Sub

Private Sub llistapermoure_Click()
   Dim vtaula As String
   cforatdesti = ""
   cforat = ""
   cbobina = ""
   If Mid(llistapermoure.Text, 1, 1) = "-" Then Exit Sub
   If Mid(llistapermoure.Text, 1, 1) = "*" Then
        cbobina = Mid(llistapermoure.Text, 2, InStr(1, llistapermoure.Text, " ") - 1)
         Else: cbobina = Mid(llistapermoure.Text, 1, InStr(1, llistapermoure.Text, " "))
   End If
   vtaula = IIf(cadbl(Mid(llistapermoure.Tag, 1, 4)) > 0, Mid(llistapermoure.Tag, 5), llistapermoure.Tag)
   If InStr(1, vtaula, "llistatperbaixarIMP") > 0 Then
      cforatdesti = "IMP"
      cforat = buscarsituacio(cbobina)
   End If
   
   If InStr(1, vtaula, "llistatperbaixarLAM") > 0 Then
      cforatdesti = "LAM"
      cforat = buscarsituacio(cbobina)
   End If
   
   If InStr(1, vtaula, "llistatperpujarIMP") > 0 Then
      cforat = "IMP"
   End If
   
   If InStr(1, vtaula, "llistatperpujarLAM") > 0 Then
      cforat = "LAM"
   End If
      cforatdesti.SetFocus
End Sub
Function buscarsituacio(vbob As String) As String
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select sit from bobines where trim(idpalet)+'/'+trim(idbobina)='" + atrim(vbob) + "'")
   If Not rst.EOF Then buscarsituacio = atrim(rst!Sit)
   Set rst = Nothing
End Function

Private Sub m_demanarcmdiametre_Click()
' Exit Sub
   If MsgBox("Segur que vols llegir diametres?", vbCritical + vbDefaultButton2 + vbYesNo, "A T E N C I Ó") = vbNo Then Exit Sub
   
   m_demanarcmdiametre.Checked = Not m_demanarcmdiametre.Checked
   imatgediametre.Visible = m_demanarcmdiametre.Checked
End Sub

Private Sub mbobinaalforat_Click()
    Dim cforatdesti As String
    Dim v As String
    Dim rst As Recordset
    
    cforatdesti = InputBox("Escaneja la columna del forat que vols saber quina bobina hi ha.", "Bobina dins del forat")
    If Len(cforatdesti) > 3 Then cforatdesti = "": MsgBox "Columna no vàlida": Exit Sub
    If Not IsNumeric(Mid(cforatdesti + "   ", 1, 1)) And cadbl(Mid(cforatdesti + "   ", 2, 2)) > 0 Then
           escullir_nivell UCase(cforatdesti), cforatdesti
         Else: MsgBox "Format de forat no correcte.", vbCritical, "Error"
    End If
    Set rst = dbcomandes.OpenRecordset("select * from bobines where sit='" + atrim(cforatdesti) + "'")
    If rst.EOF Then MsgBox "NO HI HA CAP BOBINA EN EL FORAT " + cforatdesti, vbExclamation, "ATENCIÓ"
    While Not rst.EOF
       v = atrim(rst!idpalet) + "/" + atrim(rst!idbobina) + "  "
       rst.MoveNext
    Wend
    If v <> "" Then MsgBox "Dins del forat " + cforatdesti + " hi ha les bobines: " + vbNewLine + v, vbInformation, "Bobines"
    Set rst = Nothing
End Sub

Private Sub Timer1_Timer()
  comprovar_connexio
  If formtorerus.Visible = False Then End
  If llistacanvis.ListCount > 15 Then
       llistacanvis.BackColor = QBColor(12)
       etrecordatori.Visible = True
     Else: llistacanvis.BackColor = &HEEE4D7: etrecordatori.Visible = False
  End If
End Sub
Sub mirar_si_hi_ha_algu()
   Dim rst As Recordset
   Dim vsql As String
   Set rst = dbcomandes.OpenRecordset("select * from llistatperpujarIMP")
   If rst.EOF Then chihaalgu(0).Visible = False Else chihaalgu(0).Visible = True
   Set rst = dbcomandes.OpenRecordset("select * from llistatperpujarLAM")
   If rst.EOF Then chihaalgu(1).Visible = False Else chihaalgu(1).Visible = True
   Set rst = dbcomandes.OpenRecordset("select * from llistatperbaixarIMP")
   If rst.EOF Then chihaalgu(2).Visible = False Else chihaalgu(2).Visible = True
   Set rst = dbcomandes.OpenRecordset("select * from llistatperbaixarLAM")
   If rst.EOF Then chihaalgu(3).Visible = False Else chihaalgu(3).Visible = True
   
   'vsql = "SELECT  bobinesgrups.comanda, nomgrups.nomdelgrup, nomgrups.seccio FROM nomgrups INNER JOIN (bobinesgrups INNER JOIN bobines ON (bobinesgrups.idpalet = bobines.Idpalet) AND (bobinesgrups.idbobina = bobines.Idbobina)) ON nomgrups.numerogrup = cdbl(bobinesgrups.comanda) "
   'vsql = vsql + " WHERE  IsNumeric(Mid([Sit],2,1))<>False and ( ucase(Mid([Sit]&' ',1,1))=' ' OR ucase(Mid([Sit],1,1))='F' OR ucase(Mid([Sit],1,1))='G');"
   vsql = "SELECT DISTINCT bobinesgrups.comanda FROM bobinesgrups LEFT JOIN bobines ON (bobinesgrups.idbobina = bobines.Idbobina) AND (bobinesgrups.idpalet = bobines.Idpalet) "
   vsql = vsql + " WHERE (((bobines.Sit) Like '[0-9]*'  Or (bobines.Sit) Like '[A-Z][0-9]*' Or (bobines.Sit) Like '[A-Z]' or bobines.sit is null));"
   Set rst = dbcomandes.OpenRecordset(vsql)
  ' Clipboard.Clear
  ' Clipboard.SetText vsql
   If rst.EOF Then chihaalgu(4).Visible = False Else chihaalgu(4).Visible = True
   Set rst = Nothing
   
End Sub
Private Sub Timer2_Timer()
  If cnomoperari = "Sense Operari" Then
     If busuari.BackColor = QBColor(11) Then
          busuari.BackColor = Command2.BackColor
           Else: busuari.BackColor = QBColor(11)
     End If
  End If
 
End Sub

Private Sub Timer3_Timer()
      
End Sub

Private Sub timeractualitzacions_Timer()
  comprovar_actualitzacions
End Sub
