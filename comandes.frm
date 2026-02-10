VERSION 5.00
Object = "{8C45F041-B87C-11D1-96EF-845C0FC10100}#1.3#0"; "SCROLLBOX.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form formcomandes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comandes"
   ClientHeight    =   15330
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10515
   DrawStyle       =   5  'Transparent
   Icon            =   "comandes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   15330
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   7425
      Left            =   10500
      Max             =   25
      Min             =   1
      TabIndex        =   484
      Top             =   720
      Value           =   1
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox Text32 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1305
      Index           =   12
      Left            =   4860
      MultiLine       =   -1  'True
      TabIndex        =   477
      TabStop         =   0   'False
      Top             =   3270
      Visible         =   0   'False
      Width           =   3945
   End
   Begin VB.CommandButton sortir 
      Height          =   525
      Left            =   9900
      Picture         =   "comandes.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Sortir a Menú"
      Top             =   165
      Width           =   570
   End
   Begin VB.FileListBox fitxers 
      Height          =   285
      Left            =   15
      TabIndex        =   9
      Top             =   1815
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Index           =   0
      Left            =   90
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Tag             =   "100"
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton Command9 
         BackColor       =   &H0025EFAD&
         Height          =   285
         Index           =   8
         Left            =   6645
         Picture         =   "comandes.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   491
         TabStop         =   0   'False
         ToolTipText     =   "Firmes de la comanda"
         Top             =   150
         Width           =   315
      End
      Begin VB.CommandButton Command9 
         Height          =   450
         Index           =   6
         Left            =   4230
         Picture         =   "comandes.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   458
         TabStop         =   0   'False
         ToolTipText     =   "Modificar pedidos client massivament (Crop's)."
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton Command9 
         Height          =   450
         Index           =   5
         Left            =   3795
         Picture         =   "comandes.frx":13E0
         Style           =   1  'Graphical
         TabIndex        =   456
         TabStop         =   0   'False
         ToolTipText     =   "Modificar pedidos client massivament (Crop's)."
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton Command9 
         Height          =   450
         Index           =   4
         Left            =   3375
         Picture         =   "comandes.frx":196A
         Style           =   1  'Graphical
         TabIndex        =   455
         TabStop         =   0   'False
         ToolTipText     =   "Informació de comandes desactivades."
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton Command9 
         Height          =   450
         Index           =   3
         Left            =   2955
         Picture         =   "comandes.frx":1EF4
         Style           =   1  'Graphical
         TabIndex        =   454
         TabStop         =   0   'False
         ToolTipText     =   "Possar preu a la comanda."
         Top             =   225
         Width           =   420
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3750
         Top             =   285
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   2
         Left            =   2535
         Picture         =   "comandes.frx":247E
         Style           =   1  'Graphical
         TabIndex        =   419
         TabStop         =   0   'False
         ToolTipText     =   "Llista de canvis realitzats a la comanda."
         Top             =   135
         Width           =   315
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0C0FF&
         Height          =   450
         Left            =   2010
         Picture         =   "comandes.frx":2A08
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Acceptar els canvis (F1)."
         Top             =   225
         Width           =   465
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Caption         =   "Et. Rebo."
         Height          =   525
         Left            =   6990
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   165
         Width           =   555
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Imprimir"
         Height          =   525
         Index           =   0
         Left            =   7545
         MaskColor       =   &H00C0C0FF&
         Picture         =   "comandes.frx":2F92
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   165
         Width           =   750
      End
      Begin VB.Data data1 
         Caption         =   "Comandes"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   390
         Left            =   4755
         OLEDropMode     =   1  'Manual
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   315
         Width           =   1875
      End
      Begin VB.CheckBox previprint 
         Height          =   195
         Left            =   7200
         TabIndex        =   11
         ToolTipText     =   "Previ d'Impressió de Comanda"
         Top             =   150
         Visible         =   0   'False
         Width           =   240
      End
      Begin Crystal.CrystalReport llistat 
         Left            =   3255
         Top             =   180
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton modificar 
         Height          =   450
         Left            =   480
         Picture         =   "comandes.frx":351C
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Modificar Registres"
         Top             =   225
         Width           =   540
      End
      Begin VB.CommandButton eliminar 
         Height          =   450
         Left            =   1545
         Picture         =   "comandes.frx":3AA6
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   225
         Width           =   465
      End
      Begin VB.CommandButton alta 
         Height          =   450
         Left            =   75
         Picture         =   "comandes.frx":4030
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   225
         Width           =   405
      End
      Begin VB.CommandButton consultar 
         Height          =   450
         Left            =   1020
         Picture         =   "comandes.frx":45BA
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Buscar Registres"
         Top             =   225
         Width           =   525
      End
      Begin VB.CommandButton Command8 
         Height          =   525
         Left            =   8820
         Picture         =   "comandes.frx":4B44
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Duplicar Comanda"
         Top             =   165
         Width           =   450
      End
      Begin VB.CommandButton Command23 
         Height          =   525
         Left            =   8295
         Picture         =   "comandes.frx":50CE
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Avisos Comentaris"
         Top             =   165
         Width           =   525
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Dia"
         Height          =   525
         Index           =   1
         Left            =   9270
         Picture         =   "comandes.frx":5658
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   165
         Width           =   510
      End
      Begin VB.CheckBox impnomescomanda 
         Height          =   195
         Left            =   7200
         TabIndex        =   12
         ToolTipText     =   "Impresió de la comanda sola"
         Top             =   420
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   33
         Left            =   1635
         TabIndex        =   424
         Top             =   120
         Width           =   1410
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
         Left            =   6240
         TabIndex        =   5
         Top             =   30
         Width           =   105
      End
   End
   Begin ScrollBoxCtl.ScrollBox formscrooll 
      Height          =   14565
      Left            =   75
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   675
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   25691
      Resolution      =   0
      ScrollBars      =   2
      Caption         =   ""
      Alignment       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command1 
         BackColor       =   &H005C31DD&
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   5085
         Style           =   1  'Graphical
         TabIndex        =   489
         TabStop         =   0   'False
         ToolTipText     =   "Botó dret canvi de tarifa."
         Top             =   1395
         Width           =   360
      End
      Begin VB.CommandButton Command1 
         Height          =   285
         Index           =   8
         Left            =   855
         Picture         =   "comandes.frx":5BE2
         Style           =   1  'Graphical
         TabIndex        =   473
         TabStop         =   0   'False
         ToolTipText     =   "Documentació de la comanda"
         Top             =   345
         Width           =   315
      End
      Begin VB.CheckBox checkpassaraproduccio 
         Caption         =   "Passar a impresores."
         Height          =   195
         Left            =   135
         TabIndex        =   453
         TabStop         =   0   'False
         ToolTipText     =   "Marcar per passar-ho a impresores."
         Top             =   150
         Width           =   1935
      End
      Begin VB.Frame areadatos 
         Height          =   25000
         Left            =   15
         OLEDropMode     =   1  'Manual
         TabIndex        =   17
         Tag             =   "100"
         Top             =   -60
         Width           =   10065
         Begin VB.CommandButton Command1 
            Caption         =   "Pk'"
            Height          =   255
            Index           =   10
            Left            =   9150
            TabIndex        =   501
            TabStop         =   0   'False
            ToolTipText     =   "Packing list de la comanda."
            Top             =   2925
            Width           =   390
         End
         Begin VB.CommandButton Command26 
            BackColor       =   &H008080FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   7755
            Picture         =   "comandes.frx":616C
            Style           =   1  'Graphical
            TabIndex        =   497
            TabStop         =   0   'False
            ToolTipText     =   "Canvis a la referencia d'Inplacsa."
            Top             =   1485
            Width           =   300
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00ED823A&
            BorderStyle     =   0  'None
            Height          =   2115
            Index           =   1
            Left            =   8055
            OLEDropMode     =   1  'Manual
            TabIndex        =   493
            Tag             =   "100"
            Top             =   1515
            Visible         =   0   'False
            Width           =   345
            Begin VB.CommandButton Command26 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   12
               Left            =   30
               Picture         =   "comandes.frx":66F6
               Style           =   1  'Graphical
               TabIndex        =   503
               TabStop         =   0   'False
               ToolTipText     =   "Disposició de materials a les seccions."
               Top             =   1065
               Width           =   300
            End
            Begin VB.CommandButton Command26 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   11
               Left            =   30
               Picture         =   "comandes.frx":67C8
               Style           =   1  'Graphical
               TabIndex        =   499
               TabStop         =   0   'False
               ToolTipText     =   "Buscar totes les versions d'aquesta referència."
               Top             =   1410
               Width           =   300
            End
            Begin VB.CommandButton Command26 
               BackColor       =   &H8000000E&
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   10
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   498
               TabStop         =   0   'False
               ToolTipText     =   "Comprovar referencies equivocades"
               Top             =   1770
               Width           =   300
            End
            Begin VB.CommandButton Command26 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   9
               Left            =   30
               Picture         =   "comandes.frx":6D52
               Style           =   1  'Graphical
               TabIndex        =   496
               TabStop         =   0   'False
               ToolTipText     =   "Passar la referencia a INACTIVA"
               Top             =   720
               Width           =   300
            End
            Begin VB.CommandButton Command26 
               BackColor       =   &H8000000E&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   8
               Left            =   15
               Picture         =   "comandes.frx":72DC
               Style           =   1  'Graphical
               TabIndex        =   495
               TabStop         =   0   'False
               ToolTipText     =   "Passar la referencia a ACTIVA."
               Top             =   375
               Width           =   300
            End
            Begin VB.CommandButton Command26 
               BackColor       =   &H8000000E&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   7
               Left            =   15
               Picture         =   "comandes.frx":7866
               Style           =   1  'Graphical
               TabIndex        =   494
               TabStop         =   0   'False
               ToolTipText     =   "Canviar el codi de la referencia inplacsa."
               Top             =   15
               Width           =   300
            End
         End
         Begin VB.CommandButton Command26 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   6
            Left            =   795
            Picture         =   "comandes.frx":7A87
            Style           =   1  'Graphical
            TabIndex        =   492
            TabStop         =   0   'False
            ToolTipText     =   "Observacions del PVP"
            Top             =   1785
            Width           =   300
         End
         Begin VB.CheckBox materialexacte 
            Caption         =   "Info"
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   3
            Left            =   8100
            TabIndex        =   487
            TabStop         =   0   'False
            Top             =   390
            Width           =   630
         End
         Begin VB.CommandButton Command9 
            Height          =   330
            Index           =   7
            Left            =   180
            Picture         =   "comandes.frx":8011
            Style           =   1  'Graphical
            TabIndex        =   478
            TabStop         =   0   'False
            ToolTipText     =   "Imprimir bossa soldadores"
            Top             =   18840
            Width           =   810
         End
         Begin VB.CommandButton Command1 
            Height          =   330
            Index           =   7
            Left            =   4695
            Picture         =   "comandes.frx":859B
            Style           =   1  'Graphical
            TabIndex        =   429
            TabStop         =   0   'False
            ToolTipText     =   "Linkar amb PDF pressupost"
            Top             =   1455
            Width           =   360
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Index           =   0
            Left            =   7290
            Picture         =   "comandes.frx":8B25
            Style           =   1  'Graphical
            TabIndex        =   416
            TabStop         =   0   'False
            Top             =   10155
            Width           =   450
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Baixes"
            Height          =   255
            Left            =   9120
            Style           =   1  'Graphical
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   3180
            Width           =   765
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Totals"
            Height          =   255
            Left            =   9135
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   6105
            Width           =   765
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Baixes"
            Height          =   255
            Left            =   9135
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   5850
            Width           =   765
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Totals"
            Height          =   255
            Left            =   9150
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   10470
            Width           =   765
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Baixes"
            Height          =   255
            Left            =   9150
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   10215
            Width           =   765
         End
         Begin VB.CommandButton Command19 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Totals"
            Height          =   255
            Left            =   9120
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   14655
            Width           =   765
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Baixes"
            Height          =   255
            Left            =   9120
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   14400
            Width           =   765
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Totals"
            Height          =   255
            Left            =   9165
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   18150
            Width           =   765
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Baixes"
            Height          =   255
            Left            =   9165
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   17895
            Width           =   765
         End
         Begin VB.CommandButton Command21 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Baixes"
            Height          =   255
            Left            =   9165
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   21180
            Width           =   765
         End
         Begin VB.CommandButton Command20 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Totals"
            Height          =   255
            Left            =   9165
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   21420
            Width           =   765
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   3930
            TabIndex        =   18
            Tag             =   "100"
            Top             =   345
            Width           =   4170
            Begin VB.Label nomclient 
               Caption         =   "Nom del client"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   390
               Left            =   30
               MouseIcon       =   "comandes.frx":90AF
               MousePointer    =   99  'Custom
               TabIndex        =   19
               Top             =   15
               Width           =   4125
            End
         End
         Begin VB.Frame cap 
            Height          =   3435
            Left            =   90
            TabIndex        =   194
            Top             =   105
            Width           =   9915
            Begin VB.TextBox text77 
               Height          =   300
               Index           =   28
               Left            =   3555
               TabIndex        =   508
               ToolTipText     =   "Agrupació de comandes per fer un Pack sempre dins del pressupost"
               Top             =   1980
               Width           =   435
            End
            Begin VB.TextBox Text32 
               DataField       =   "obsext2"
               DataSource      =   "data1"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   14
               Left            =   8040
               MaxLength       =   12
               TabIndex        =   483
               TabStop         =   0   'False
               ToolTipText     =   "Número de proforma"
               Top             =   2535
               Width           =   1425
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00FFFFFF&
               DataSource      =   "data1"
               Height          =   315
               Index           =   3
               ItemData        =   "comandes.frx":94F1
               Left            =   7050
               List            =   "comandes.frx":9501
               TabIndex        =   479
               Top             =   2835
               Width           =   2010
            End
            Begin VB.TextBox Text16 
               DataField       =   "mesurapvp"
               Height          =   285
               Left            =   2565
               TabIndex        =   475
               Top             =   1680
               WhatsThisHelpID =   1
               Width           =   1110
            End
            Begin VB.TextBox Text32 
               BackColor       =   &H00C0FFFF&
               DataField       =   "refilate"
               DataSource      =   "data1"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   11
               Left            =   2010
               TabIndex        =   471
               TabStop         =   0   'False
               Top             =   300
               Width           =   450
            End
            Begin VB.TextBox Text32 
               DataField       =   "obsetiq"
               DataSource      =   "data1"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   2
               Left            =   4500
               TabIndex        =   243
               TabStop         =   0   'False
               Top             =   3120
               Width           =   1380
            End
            Begin VB.TextBox Text32 
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   10
               Left            =   6735
               MaxLength       =   100
               MultiLine       =   -1  'True
               TabIndex        =   469
               TabStop         =   0   'False
               ToolTipText     =   "Observacions a l'albarà"
               Top             =   3135
               Width           =   2190
            End
            Begin VB.CommandButton Command26 
               Caption         =   "Gtin"
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
               Left            =   6000
               TabIndex        =   464
               TabStop         =   0   'False
               ToolTipText     =   "Ensenya el codi GTIN d'aquesta referència"
               Top             =   1155
               Visible         =   0   'False
               Width           =   405
            End
            Begin VB.CommandButton Command26 
               BackColor       =   &H8000000E&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   9360
               Picture         =   "comandes.frx":9531
               Style           =   1  'Graphical
               TabIndex        =   459
               TabStop         =   0   'False
               ToolTipText     =   "Canviar l'estat de la comanda"
               Top             =   1950
               Width           =   300
            End
            Begin VB.CommandButton Command26 
               BackColor       =   &H008080FF&
               Caption         =   ">"
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
               Index           =   2
               Left            =   9660
               Style           =   1  'Graphical
               TabIndex        =   439
               TabStop         =   0   'False
               ToolTipText     =   "Manteniment de Call-offs"
               Top             =   900
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Index           =   2
               ItemData        =   "comandes.frx":9ABB
               Left            =   8265
               List            =   "comandes.frx":9AC8
               TabIndex        =   437
               Top             =   885
               Width           =   1425
            End
            Begin VB.TextBox Text32 
               BackColor       =   &H00FFFF00&
               DataField       =   "refclientdeclient"
               DataSource      =   "data1"
               ForeColor       =   &H00808080&
               Height          =   285
               Index           =   6
               Left            =   8520
               MaxLength       =   15
               TabIndex        =   425
               TabStop         =   0   'False
               Top             =   2250
               Width           =   1320
            End
            Begin VB.TextBox Text32 
               BackColor       =   &H006BEBB1&
               DataField       =   "refinplacsa"
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
               Height          =   285
               Index           =   5
               Left            =   6405
               TabIndex        =   423
               TabStop         =   0   'False
               Top             =   1410
               Width           =   1275
            End
            Begin VB.TextBox Text32 
               BackColor       =   &H00FFFF00&
               DataField       =   "obspedgen2"
               DataSource      =   "data1"
               ForeColor       =   &H00808080&
               Height          =   285
               Index           =   4
               Left            =   6075
               MaxLength       =   15
               TabIndex        =   421
               TabStop         =   0   'False
               Top             =   2280
               Width           =   1725
            End
            Begin MSMask.MaskEdBox Text2 
               DataField       =   "client"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3045
               TabIndex        =   210
               Top             =   210
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16777215
               Format          =   "0"
               PromptChar      =   "_"
            End
            Begin VB.TextBox Text32 
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   3
               Left            =   3045
               TabIndex        =   400
               TabStop         =   0   'False
               Text            =   "43000000100"
               Top             =   -15
               Width           =   4410
            End
            Begin VB.CommandButton Command1 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Canvi Client"
               Height          =   315
               Index           =   5
               Left            =   2535
               Style           =   1  'Graphical
               TabIndex        =   399
               TabStop         =   0   'False
               Top             =   495
               Width           =   1080
            End
            Begin VB.TextBox Text32 
               DataField       =   "obspedgen1"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   240
               Top             =   2850
               Width           =   6030
            End
            Begin VB.TextBox Text13 
               DataField       =   "obspedido1"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   960
               TabIndex        =   238
               Top             =   2565
               Width           =   6165
            End
            Begin VB.Timer Timer1 
               Interval        =   500
               Left            =   -75
               Top             =   2040
            End
            Begin VB.Timer Timer2 
               Interval        =   50
               Left            =   120
               Top             =   1170
            End
            Begin VB.TextBox inventat 
               Height          =   285
               Left            =   75
               TabIndex        =   207
               Text            =   "inventat no tocar es pels llistats"
               Top             =   2250
               Visible         =   0   'False
               Width           =   435
            End
            Begin VB.CommandButton Command24 
               Caption         =   "Command24"
               Height          =   300
               Left            =   915
               TabIndex        =   206
               TabStop         =   0   'False
               ToolTipText     =   "Possa la data d'avui"
               Top             =   825
               Width           =   285
            End
            Begin VB.CommandButton Command25 
               Caption         =   "Command24"
               Height          =   300
               Left            =   3405
               TabIndex        =   205
               TabStop         =   0   'False
               ToolTipText     =   "Possa la data d'avui"
               Top             =   795
               Width           =   285
            End
            Begin VB.CommandButton Command26 
               Caption         =   "Command24"
               Height          =   255
               Index           =   0
               Left            =   6135
               TabIndex        =   204
               TabStop         =   0   'False
               ToolTipText     =   "Possa la data d'avui"
               Top             =   840
               Visible         =   0   'False
               Width           =   285
            End
            Begin VB.TextBox text77 
               DataField       =   "com_representant"
               DataSource      =   "data1"
               Height          =   270
               Index           =   10
               Left            =   4245
               TabIndex        =   203
               Top             =   1680
               Width           =   330
            End
            Begin VB.ComboBox comboenvios 
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":9AE3
               Left            =   4965
               List            =   "comandes.frx":9AF6
               TabIndex        =   202
               Top             =   510
               Visible         =   0   'False
               Width           =   3570
            End
            Begin VB.TextBox text77 
               DataField       =   "linkcomanda1"
               DataSource      =   "data1"
               Height          =   285
               Index           =   11
               Left            =   8640
               MousePointer    =   99  'Custom
               TabIndex        =   201
               TabStop         =   0   'False
               Top             =   120
               Width           =   1155
            End
            Begin VB.TextBox text77 
               DataField       =   "linkcomanda2"
               DataSource      =   "data1"
               Height          =   285
               Index           =   12
               Left            =   8640
               TabIndex        =   200
               TabStop         =   0   'False
               Top             =   360
               Width           =   1155
            End
            Begin VB.TextBox Text32 
               DataField       =   "refclialt"
               DataSource      =   "data1"
               ForeColor       =   &H00808080&
               Height          =   285
               Index           =   1
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   241
               Top             =   3120
               Width           =   2730
            End
            Begin VB.CommandButton Command26 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ñ"
               BeginProperty Font 
                  Name            =   "Symbol"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   7935
               Style           =   1  'Graphical
               TabIndex        =   199
               TabStop         =   0   'False
               ToolTipText     =   "Possa la data d'avui"
               Top             =   1095
               Width           =   285
            End
            Begin VB.TextBox text77 
               BackColor       =   &H0080C0FF&
               Height          =   285
               Index           =   17
               Left            =   9030
               TabIndex        =   198
               Top             =   1980
               Width           =   285
            End
            Begin VB.TextBox dataentrega2 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   7560
               Locked          =   -1  'True
               MouseIcon       =   "comandes.frx":9B09
               TabIndex        =   197
               TabStop         =   0   'False
               Top             =   1680
               Width           =   1050
            End
            Begin VB.TextBox importancia 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   9525
               Locked          =   -1  'True
               MouseIcon       =   "comandes.frx":9C5B
               MousePointer    =   99  'Custom
               TabIndex        =   196
               TabStop         =   0   'False
               Top             =   1680
               Width           =   270
            End
            Begin VB.CheckBox noplanificable 
               Caption         =   "No Planificable."
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   8610
               TabIndex        =   195
               Top             =   660
               Visible         =   0   'False
               Width           =   1260
            End
            Begin MSMask.MaskEdBox Text45 
               DataField       =   "comandaclient"
               DataSource      =   "data1"
               Height          =   285
               Left            =   7920
               TabIndex        =   225
               Top             =   1395
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text44 
               DataField       =   "refclient"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6405
               TabIndex        =   220
               Top             =   1140
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text41 
               DataField       =   "numpressupost"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3240
               TabIndex        =   223
               Top             =   1395
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text11 
               DataField       =   "tipoentrega"
               DataSource      =   "data1"
               Height          =   285
               Left            =   975
               TabIndex        =   235
               Top             =   2235
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text10 
               DataField       =   "pvpcliche"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6405
               TabIndex        =   233
               Top             =   1950
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text8 
               DataField       =   "pvprevisat"
               DataSource      =   "data1"
               Height          =   285
               Left            =   975
               TabIndex        =   230
               Top             =   1950
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   503
               _Version        =   327681
               Enabled         =   0   'False
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text7 
               DataField       =   "mesurapvp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3135
               TabIndex        =   208
               Top             =   1230
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text6 
               DataField       =   "pvp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   990
               TabIndex        =   227
               ToolTipText     =   "Valors en Euros del PVP"
               Top             =   1665
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   503
               _Version        =   327681
               Format          =   "#,##0.00000"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text5 
               DataField       =   "dataentrega"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6405
               TabIndex        =   209
               Top             =   1695
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   503
               _Version        =   327681
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text4 
               DataField       =   "datacomanda"
               DataSource      =   "data1"
               Height          =   285
               Left            =   975
               TabIndex        =   221
               Top             =   1395
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   503
               _Version        =   327681
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text3 
               DataField       =   "producte"
               DataSource      =   "data1"
               Height          =   285
               Left            =   975
               TabIndex        =   218
               Top             =   1110
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   327681
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text1 
               DataField       =   "comanda"
               DataSource      =   "data1"
               Height          =   255
               Left            =   1065
               TabIndex        =   211
               Top             =   300
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   450
               _Version        =   327681
               BackColor       =   8454143
               ForeColor       =   8421631
               Enabled         =   0   'False
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox12 
               DataField       =   "proximaseccio"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8760
               TabIndex        =   212
               TabStop         =   0   'False
               ToolTipText     =   "Si està vermell s'està fabricant aquesta secció"
               Top             =   1980
               Width           =   270
               _ExtentX        =   476
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16777215
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox14 
               DataField       =   "seccioactual"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9240
               TabIndex        =   213
               ToolTipText     =   "Si està vermell s'està fabricant aquesta secció"
               Top             =   2325
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox dataactivacio 
               DataField       =   "dataactivacio"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1215
               TabIndex        =   214
               Top             =   840
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               _Version        =   327681
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox21 
               DataField       =   "datapreu"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3705
               TabIndex        =   215
               Top             =   825
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   327681
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox22 
               DataField       =   "datamaterial"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6405
               TabIndex        =   216
               Top             =   810
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               _Version        =   327681
               AutoTab         =   -1  'True
               MaxLength       =   10
               Format          =   "dd/mm/yy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox11 
               DataField       =   "pvpdolar"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1650
               TabIndex        =   430
               ToolTipText     =   "Valors en Dolars del PVP"
               Top             =   1680
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   327681
               Format          =   "#,##0.00000"
               PromptChar      =   "_"
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Duplicada de: "
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   178
               Left            =   60
               TabIndex        =   502
               Top             =   570
               Width           =   2445
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Nº:Proforma:"
               DataSource      =   "clients"
               Height          =   270
               Index           =   176
               Left            =   7140
               TabIndex        =   482
               Top             =   2565
               Width           =   1020
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               DataField       =   "puntrisc"
               DataSource      =   "data1"
               ForeColor       =   &H80000004&
               Height          =   480
               Index           =   146
               Left            =   9360
               TabIndex        =   224
               Top             =   2565
               Width           =   585
            End
            Begin VB.Label label1 
               Caption         =   $"comandes.frx":9DAD
               DataSource      =   "data1"
               Height          =   255
               Index           =   172
               Left            =   2025
               TabIndex        =   472
               Top             =   105
               Width           =   6765
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Obs. Alb:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   171
               Left            =   5970
               TabIndex        =   470
               Top             =   3150
               Width           =   840
            End
            Begin VB.Label label1 
               Caption         =   "CallOff:"
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
               Index           =   165
               Left            =   7800
               TabIndex        =   438
               Top             =   900
               Width           =   765
            End
            Begin VB.Label label1 
               Caption         =   "$                                                                   %Com:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   161
               Left            =   2415
               TabIndex        =   431
               Top             =   1710
               Width           =   165
            End
            Begin VB.Label label1 
               Caption         =   "Ref Cli de Cli:"
               DataSource      =   "clients"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   35
               Left            =   7830
               TabIndex        =   426
               Top             =   2310
               Width           =   885
            End
            Begin VB.Label label1 
               Caption         =   "Ref.Inplacsa:"
               DataSource      =   "data1"
               Height          =   285
               Index           =   32
               Left            =   5430
               TabIndex        =   422
               Top             =   1470
               Width           =   990
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               DataSource      =   "data1"
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
               Height          =   255
               Index           =   148
               Left            =   4530
               TabIndex        =   282
               ToolTipText     =   "Tarifa activa"
               Top             =   1695
               Width           =   780
            End
            Begin VB.Label label1 
               Caption         =   "Comanda:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   5
               Left            =   45
               TabIndex        =   280
               Top             =   330
               Width           =   765
            End
            Begin VB.Label label1 
               Caption         =   "Client"
               DataSource      =   "data1"
               Height          =   180
               Index           =   0
               Left            =   2535
               TabIndex        =   278
               Top             =   270
               Width           =   765
            End
            Begin VB.Label label1 
               Caption         =   "Producte:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   1
               Left            =   150
               TabIndex        =   276
               Top             =   1185
               Width           =   765
            End
            Begin VB.Label nomproducte 
               Caption         =   "Descripcio del producte"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   240
               Left            =   1950
               TabIndex        =   274
               Top             =   1185
               Width           =   3285
            End
            Begin VB.Label label1 
               Caption         =   "Data Com:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   2
               Left            =   150
               TabIndex        =   272
               Top             =   1470
               Width           =   765
            End
            Begin VB.Label label1 
               Caption         =   "Data Ent 1:                          2:                        Importancia:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   3
               Left            =   5445
               TabIndex        =   270
               Top             =   1740
               Width           =   4245
            End
            Begin VB.Label label1 
               Caption         =   "PVP €:                                                                    %Com:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   4
               Left            =   180
               TabIndex        =   268
               Top             =   1740
               Width           =   4065
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "PVP Rent.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   7
               Left            =   150
               TabIndex        =   266
               Top             =   2025
               Width           =   930
            End
            Begin VB.Label label1 
               Caption         =   "Nº de Pack:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   8
               Left            =   2625
               TabIndex        =   264
               Top             =   2025
               Width           =   1065
            End
            Begin VB.Label label1 
               Caption         =   "Preu del Clixé:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   9
               Left            =   5250
               TabIndex        =   262
               Top             =   2025
               Width           =   1065
            End
            Begin VB.Label label1 
               Caption         =   "T. Entrega:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   10
               Left            =   150
               TabIndex        =   244
               Top             =   2310
               Width           =   1065
            End
            Begin VB.Label Label3 
               Caption         =   "Descripcio del tipus d'entrega"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   315
               Left            =   1650
               TabIndex        =   242
               Top             =   2310
               Width           =   3105
            End
            Begin VB.Label label1 
               Caption         =   "Nº P/Ag:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   34
               Left            =   2520
               TabIndex        =   237
               Top             =   1470
               Width           =   1140
            End
            Begin VB.Label label1 
               Caption         =   "Ref. Client:"
               DataSource      =   "data1"
               Height          =   285
               Index           =   36
               Left            =   5220
               TabIndex        =   236
               Top             =   1185
               Width           =   990
            End
            Begin VB.Label label1 
               Caption         =   "NºCom.Cl.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   37
               Left            =   8520
               TabIndex        =   234
               Top             =   1215
               Width           =   765
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Com. Cli de Cli:"
               DataSource      =   "clients"
               Height          =   270
               Index           =   138
               Left            =   4890
               TabIndex        =   232
               Top             =   2325
               Width           =   1245
            End
            Begin VB.Label label1 
               Caption         =   "Data Activació:"
               DataSource      =   "data1"
               Height          =   420
               Index           =   139
               Left            =   135
               TabIndex        =   231
               Top             =   735
               Width           =   765
            End
            Begin VB.Label label1 
               Caption         =   "Data Preu:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   140
               Left            =   2580
               TabIndex        =   229
               Top             =   870
               Width           =   765
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Data Entrega que vol el client"
               DataSource      =   "data1"
               Height          =   495
               Index           =   141
               Left            =   4995
               TabIndex        =   228
               Top             =   810
               Width           =   1215
            End
            Begin VB.Label label1 
               Caption         =   "Estat Fab:"
               DataSource      =   "data1"
               Height          =   225
               Index           =   142
               Left            =   7965
               TabIndex        =   226
               Top             =   2010
               Width           =   810
            End
            Begin VB.Shape puntrisc 
               BorderColor     =   &H00E0E0E0&
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   390
               Left            =   9450
               Shape           =   3  'Circle
               Top             =   2670
               Width           =   390
            End
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Envio:"
               DataSource      =   "data1"
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   147
               Left            =   3645
               TabIndex        =   222
               Top             =   495
               Width           =   4905
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Ref. Clients Alternatives"
               DataSource      =   "data1"
               Height          =   480
               Index           =   155
               Left            =   105
               TabIndex        =   219
               Top             =   3030
               Width           =   840
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Obs. Etiq:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   156
               Left            =   3780
               TabIndex        =   217
               Top             =   3135
               Width           =   840
            End
            Begin VB.Label label1 
               Caption         =   "Obs.  de la Comanda"
               DataSource      =   "data1"
               Height          =   480
               Index           =   15
               Left            =   135
               TabIndex        =   239
               Top             =   2610
               Width           =   840
            End
         End
         Begin VB.Frame ext 
            Caption         =   "Extrussora"
            Height          =   2940
            Left            =   90
            TabIndex        =   284
            Top             =   3540
            Width           =   9915
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Index           =   4
               ItemData        =   "comandes.frx":9E3F
               Left            =   7695
               List            =   "comandes.frx":9E4C
               TabIndex        =   505
               Top             =   1830
               Width           =   810
            End
            Begin VB.TextBox Text32 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Index           =   15
               Left            =   9510
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   490
               TabStop         =   0   'False
               Top             =   1215
               Visible         =   0   'False
               Width           =   3045
            End
            Begin VB.CheckBox materialexacte 
               BackColor       =   &H000000FF&
               Caption         =   "Utilitzar material  del grup de compatibles."
               Height          =   195
               Index           =   4
               Left            =   1140
               TabIndex        =   488
               TabStop         =   0   'False
               Top             =   585
               Visible         =   0   'False
               Width           =   4890
            End
            Begin VB.TextBox Text32 
               DataField       =   "mesuraquantdemanada"
               DataSource      =   "data1"
               Height          =   285
               Index           =   8
               Left            =   7845
               Locked          =   -1  'True
               TabIndex        =   436
               Top             =   195
               Visible         =   0   'False
               Width           =   435
            End
            Begin VB.TextBox Text32 
               DataField       =   "unitatsquantitatdemanada"
               Height          =   285
               Index           =   7
               Left            =   7680
               Locked          =   -1  'True
               TabIndex        =   435
               Top             =   1410
               Width           =   810
            End
            Begin VB.CheckBox materialexacte 
               Caption         =   "El material ha de ser exactament aquest."
               Height          =   195
               Index           =   0
               Left            =   1800
               TabIndex        =   298
               TabStop         =   0   'False
               Top             =   585
               Width           =   3270
            End
            Begin VB.TextBox Text34 
               DataField       =   "obsextgen1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1335
               Locked          =   -1  'True
               TabIndex        =   261
               Top             =   2580
               Width           =   7500
            End
            Begin VB.TextBox Text36 
               DataField       =   "obsext1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1335
               TabIndex        =   260
               Top             =   2235
               Width           =   7500
            End
            Begin VB.ComboBox Combo8 
               DataField       =   "tubolam"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":9E5E
               Left            =   1125
               List            =   "comandes.frx":9E68
               TabIndex        =   245
               Top             =   180
               WhatsThisHelpID =   1
               Width           =   630
            End
            Begin VB.TextBox grmcm3 
               Height          =   285
               Left            =   5535
               TabIndex        =   288
               Text            =   "grmcm3"
               Top             =   1155
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.CommandButton Command27 
               BackColor       =   &H0000FFFF&
               Caption         =   "PL"
               Height          =   375
               Left            =   9510
               Style           =   1  'Graphical
               TabIndex        =   286
               TabStop         =   0   'False
               Top             =   150
               Width           =   360
            End
            Begin MSMask.MaskEdBox Text33 
               DataField       =   "pes1000mtrs"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6825
               TabIndex        =   254
               Top             =   1110
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text31 
               DataField       =   "mesuracantex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8730
               TabIndex        =   291
               Top             =   210
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text30 
               DataSource      =   "data1"
               Height          =   285
               Left            =   7680
               TabIndex        =   252
               ToolTipText     =   "Mesura de la Quantitat"
               Top             =   810
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text29 
               DataField       =   "cantitatex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6840
               TabIndex        =   251
               Top             =   810
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   503
               _Version        =   327681
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text26 
               DataField       =   "aditiuex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1110
               TabIndex        =   293
               TabStop         =   0   'False
               Top             =   1380
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text25 
               DataField       =   "materialex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1110
               TabIndex        =   294
               TabStop         =   0   'False
               Top             =   1080
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text24 
               DataField       =   "colorex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1125
               TabIndex        =   295
               TabStop         =   0   'False
               Top             =   780
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text22 
               DataSource      =   "data1"
               Height          =   285
               Left            =   7680
               TabIndex        =   296
               TabStop         =   0   'False
               ToolTipText     =   "Mesura de l'Espessor"
               Top             =   510
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text21 
               DataField       =   "espessor"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6840
               TabIndex        =   249
               Top             =   510
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text20 
               DataField       =   "solapa"
               DataSource      =   "data1"
               Height          =   285
               Left            =   5205
               TabIndex        =   248
               Top             =   225
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text19 
               DataField       =   "plegatesq"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3675
               TabIndex        =   247
               Top             =   225
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text18 
               DataField       =   "ampleesq"
               DataSource      =   "data1"
               Height          =   285
               Left            =   2685
               TabIndex        =   246
               Top             =   225
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16777215
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text23 
               DataField       =   "mesuraesp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9210
               TabIndex        =   297
               Top             =   210
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox1 
               DataField       =   "oberturaex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9525
               TabIndex        =   253
               Top             =   795
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   503
               _Version        =   327681
               MaxLength       =   1
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox2 
               DataField       =   "micropex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9525
               TabIndex        =   250
               Top             =   510
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   503
               _Version        =   327681
               MaxLength       =   1
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox3 
               DataField       =   "diamextbobext"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6120
               TabIndex        =   259
               Top             =   1815
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox4 
               DataField       =   "mtrslinbobext"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4680
               TabIndex        =   258
               Top             =   1815
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox5 
               DataField       =   "kilosbobext"
               DataSource      =   "data1"
               Height          =   285
               Left            =   2970
               TabIndex        =   257
               Top             =   1815
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox6 
               DataField       =   "tubbaseext"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6825
               TabIndex        =   256
               Top             =   1410
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox17 
               DataField       =   "tractatex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9525
               TabIndex        =   255
               Top             =   1080
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   503
               _Version        =   327681
               MaxLength       =   1
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(2Clic canvi)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H008080FF&
               Height          =   180
               Index           =   181
               Left            =   6015
               TabIndex        =   507
               Top             =   660
               Width           =   1245
            End
            Begin VB.Label label1 
               Caption         =   "Est o Past:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   180
               Left            =   6885
               TabIndex        =   506
               Top             =   1890
               Width           =   855
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   137
               Left            =   60
               TabIndex        =   500
               Top             =   1410
               Width           =   1020
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(Atenció ha d'esser la mesura del PVP)"
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
               Height          =   285
               Index           =   170
               Left            =   6825
               TabIndex        =   468
               Top             =   1665
               Visible         =   0   'False
               Width           =   3045
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(F2 Escullir mesura)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H008080FF&
               Height          =   285
               Index           =   169
               Left            =   8490
               TabIndex        =   467
               Top             =   1425
               Width           =   1260
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(F3 a Mtrs.)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H008080FF&
               Height          =   285
               Index           =   157
               Left            =   6060
               TabIndex        =   299
               Top             =   990
               Width           =   840
            End
            Begin VB.Label label1 
               Caption         =   "Tubo o Lam:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   12
               Left            =   75
               TabIndex        =   342
               Top             =   300
               Width           =   915
            End
            Begin VB.Label label1 
               Caption         =   "Ample/Pleg:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   13
               Left            =   1800
               TabIndex        =   340
               Top             =   300
               Width           =   915
            End
            Begin VB.Label label1 
               Caption         =   "/"
               DataSource      =   "data1"
               Height          =   255
               Index           =   14
               Left            =   3525
               TabIndex        =   338
               Top             =   300
               Width           =   165
            End
            Begin VB.Label label1 
               Caption         =   "Solapa:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   16
               Left            =   4620
               TabIndex        =   336
               Top             =   300
               Width           =   615
            End
            Begin VB.Label label1 
               Caption         =   "Espessor:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   18
               Left            =   6075
               TabIndex        =   334
               Top             =   495
               Width           =   765
            End
            Begin VB.Label label1 
               Caption         =   "Material:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   21
               Left            =   60
               TabIndex        =   331
               Top             =   1155
               Width           =   915
            End
            Begin VB.Label nomcolor 
               BackStyle       =   0  'Transparent
               Caption         =   "Color:"
               DataSource      =   "data1"
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
               Index           =   23
               Left            =   1800
               TabIndex        =   328
               Top             =   855
               Width           =   4200
            End
            Begin VB.Label nommaterial 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Material:"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   405
               Index           =   23
               Left            =   1785
               TabIndex        =   326
               Top             =   1155
               Width           =   4185
            End
            Begin VB.Label nomadditiu 
               BackStyle       =   0  'Transparent
               Caption         =   "Additiu:"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   23
               Left            =   1785
               TabIndex        =   324
               Top             =   1455
               Width           =   4185
            End
            Begin VB.Label label1 
               Caption         =   "Quantitat:"
               DataSource      =   "data1"
               Height          =   285
               Index           =   25
               Left            =   6075
               TabIndex        =   322
               Top             =   810
               Width           =   840
            End
            Begin VB.Label label1 
               Caption         =   "Pesx1000"
               DataSource      =   "data1"
               Height          =   255
               Index           =   27
               Left            =   6075
               TabIndex        =   320
               Top             =   1185
               Width           =   840
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Obs. del Client"
               DataSource      =   "data1"
               Height          =   480
               Index           =   28
               Left            =   75
               TabIndex        =   319
               Top             =   2580
               Width           =   1245
            End
            Begin VB.Label label1 
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Obs.  Extrussora"
               DataSource      =   "data1"
               Height          =   255
               Index           =   29
               Left            =   75
               TabIndex        =   317
               Top             =   2265
               Width           =   1380
            End
            Begin VB.Label label1 
               Caption         =   "Obert (1,2,N)"
               DataSource      =   "data1"
               Height          =   255
               Index           =   19
               Left            =   8550
               TabIndex        =   311
               Top             =   810
               Width           =   1095
            End
            Begin VB.Label label1 
               Caption         =   "Microperforat:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   26
               Left            =   8550
               TabIndex        =   309
               Top             =   525
               Width           =   1095
            End
            Begin VB.Label label1 
               Caption         =   "Diametre:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   46
               Left            =   5415
               TabIndex        =   307
               Top             =   1845
               Width           =   855
            End
            Begin VB.Label label1 
               Caption         =   "Mtrs Bobina:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   47
               Left            =   3750
               TabIndex        =   305
               Top             =   1860
               Width           =   1020
            End
            Begin VB.Label label1 
               Caption         =   "Kilos Bobina:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   48
               Left            =   1830
               TabIndex        =   303
               Top             =   1860
               Width           =   1020
            End
            Begin VB.Label label1 
               Caption         =   "Quant demanada:"
               DataSource      =   "data1"
               ForeColor       =   &H00000000&
               Height          =   360
               Index           =   49
               Left            =   5520
               TabIndex        =   301
               Top             =   1425
               Width           =   2070
            End
            Begin VB.Label label1 
               Caption         =   "Tractat (1,2,N)"
               DataSource      =   "data1"
               Height          =   255
               Index           =   136
               Left            =   8445
               TabIndex        =   300
               Top             =   1125
               Width           =   1095
            End
            Begin VB.Shape cmarcmaterial 
               BorderColor     =   &H8000000F&
               BorderWidth     =   3
               Height          =   540
               Left            =   1755
               Top             =   1095
               Width           =   3750
            End
         End
         Begin VB.Frame imp1 
            Caption         =   "Impressora-1"
            Height          =   4350
            Left            =   90
            TabIndex        =   31
            Top             =   6495
            Width           =   9915
            Begin VB.TextBox Text32 
               Height          =   285
               Index           =   9
               Left            =   9300
               TabIndex        =   463
               ToolTipText     =   "Dessarroll que el client diu que serà."
               Top             =   120
               Width           =   570
            End
            Begin VB.CheckBox materialexacte 
               Caption         =   "Ok del Client"
               Height          =   345
               Index           =   1
               Left            =   8955
               TabIndex        =   457
               TabStop         =   0   'False
               Top             =   765
               Width           =   840
            End
            Begin VB.TextBox text77 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00EAD9CE&
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               HelpContextID   =   99
               Index           =   26
               Left            =   6180
               TabIndex        =   415
               Top             =   1800
               Width           =   945
            End
            Begin VB.TextBox text77 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00EAD9CE&
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               HelpContextID   =   99
               Index           =   25
               Left            =   6180
               TabIndex        =   414
               Top             =   1575
               Width           =   945
            End
            Begin VB.TextBox text77 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00EAD9CE&
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               HelpContextID   =   99
               Index           =   24
               Left            =   6180
               TabIndex        =   413
               Top             =   1335
               Width           =   945
            End
            Begin VB.TextBox text77 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00EAD9CE&
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               HelpContextID   =   99
               Index           =   23
               Left            =   6180
               TabIndex        =   412
               Top             =   1110
               Width           =   945
            End
            Begin VB.TextBox text77 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "lin8"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Index           =   9
               Left            =   5715
               TabIndex        =   292
               Top             =   1800
               Width           =   450
            End
            Begin VB.TextBox text77 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "lin7"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Index           =   8
               Left            =   5715
               TabIndex        =   290
               Top             =   1560
               Width           =   450
            End
            Begin VB.TextBox text77 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "lin6"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Index           =   7
               Left            =   5715
               TabIndex        =   289
               Top             =   1335
               Width           =   450
            End
            Begin VB.TextBox text77 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "lin5"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Index           =   6
               Left            =   5715
               TabIndex        =   287
               Top             =   1110
               Width           =   450
            End
            Begin VB.TextBox text77 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "lin4"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Index           =   5
               Left            =   5715
               TabIndex        =   285
               Top             =   885
               Width           =   450
            End
            Begin VB.TextBox text77 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "lin3"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Index           =   4
               Left            =   5715
               TabIndex        =   283
               Top             =   675
               Width           =   450
            End
            Begin VB.TextBox text77 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "lin2"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Index           =   3
               Left            =   5715
               TabIndex        =   281
               Top             =   450
               Width           =   450
            End
            Begin VB.TextBox text77 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00EAD9CE&
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               HelpContextID   =   99
               Index           =   22
               Left            =   6180
               TabIndex        =   411
               Top             =   885
               Width           =   945
            End
            Begin VB.TextBox text77 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00EAD9CE&
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               HelpContextID   =   99
               Index           =   21
               Left            =   6180
               TabIndex        =   410
               Top             =   660
               Width           =   945
            End
            Begin VB.TextBox text77 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00EAD9CE&
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               HelpContextID   =   99
               Index           =   20
               Left            =   6180
               TabIndex        =   409
               Top             =   450
               Width           =   945
            End
            Begin VB.TextBox cobs1imp 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               HelpContextID   =   99
               Left            =   435
               TabIndex        =   448
               Top             =   2085
               Width           =   8310
            End
            Begin VB.TextBox Text140 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "tinta8a"
               DataSource      =   "data1"
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
               HelpContextID   =   99
               Left            =   210
               TabIndex        =   277
               Top             =   1785
               Width           =   5460
            End
            Begin VB.TextBox Text141 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "tinta7a"
               DataSource      =   "data1"
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
               HelpContextID   =   99
               Left            =   210
               TabIndex        =   275
               Top             =   1560
               Width           =   5460
            End
            Begin VB.TextBox Text50 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "tinta6a"
               DataSource      =   "data1"
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
               HelpContextID   =   99
               Left            =   210
               TabIndex        =   273
               Top             =   1335
               Width           =   5460
            End
            Begin VB.TextBox Text49 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "tinta5a"
               DataSource      =   "data1"
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
               HelpContextID   =   99
               Left            =   210
               TabIndex        =   271
               Top             =   1110
               Width           =   5460
            End
            Begin VB.TextBox Text48 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "tinta4a"
               DataSource      =   "data1"
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
               HelpContextID   =   99
               Left            =   210
               TabIndex        =   269
               Top             =   885
               Width           =   5460
            End
            Begin VB.TextBox Text47 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "tinta3a"
               DataSource      =   "data1"
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
               HelpContextID   =   99
               Left            =   210
               TabIndex        =   267
               Top             =   660
               Width           =   5460
            End
            Begin VB.TextBox Text46 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "tinta2a"
               DataSource      =   "data1"
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
               HelpContextID   =   99
               Left            =   210
               TabIndex        =   265
               Top             =   450
               Width           =   5460
            End
            Begin MSMask.MaskEdBox MaskEdBox16 
               DataField       =   "cmaquina"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Left            =   7935
               TabIndex        =   45
               TabStop         =   0   'False
               ToolTipText     =   "Codi de la Impressora nova"
               Top             =   3075
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16761024
               PromptChar      =   "_"
            End
            Begin VB.TextBox cdetalltinter 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   5115
               TabIndex        =   447
               Top             =   240
               Width           =   540
            End
            Begin VB.TextBox cdetalltinter 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   5115
               TabIndex        =   446
               Top             =   473
               Width           =   540
            End
            Begin VB.TextBox cdetalltinter 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   4
               Left            =   5115
               TabIndex        =   445
               Top             =   1172
               Width           =   540
            End
            Begin VB.TextBox cdetalltinter 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   5
               Left            =   5115
               TabIndex        =   444
               Top             =   1405
               Width           =   540
            End
            Begin VB.TextBox cdetalltinter 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   6
               Left            =   5115
               TabIndex        =   443
               Top             =   1638
               Width           =   540
            End
            Begin VB.TextBox cdetalltinter 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   7
               Left            =   5115
               TabIndex        =   442
               Top             =   1875
               Width           =   540
            End
            Begin VB.TextBox cdetalltinter 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   5115
               TabIndex        =   441
               Top             =   939
               Width           =   540
            End
            Begin VB.TextBox cdetalltinter 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   5115
               TabIndex        =   440
               Top             =   706
               Width           =   540
            End
            Begin VB.CommandButton bxl 
               Height          =   255
               Left            =   6345
               Picture         =   "comandes.frx":9E72
               Style           =   1  'Graphical
               TabIndex        =   418
               Top             =   2865
               Visible         =   0   'False
               Width           =   285
            End
            Begin VB.TextBox text77 
               DataField       =   "arxiuexp"
               DataSource      =   "data1"
               Height          =   285
               Index           =   27
               Left            =   5805
               Locked          =   -1  'True
               TabIndex        =   417
               TabStop         =   0   'False
               ToolTipText     =   "Arxiu a la comanda"
               Top             =   2850
               Width           =   855
            End
            Begin VB.TextBox text77 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00EAD9CE&
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               HelpContextID   =   99
               Index           =   19
               Left            =   6180
               TabIndex        =   408
               Top             =   240
               Width           =   945
            End
            Begin VB.CommandButton Command1 
               Height          =   315
               Index           =   4
               Left            =   6030
               Picture         =   "comandes.frx":A3FC
               Style           =   1  'Graphical
               TabIndex        =   397
               TabStop         =   0   'False
               Top             =   3660
               Width           =   450
            End
            Begin VB.TextBox text77 
               DataField       =   "obsimpgen1"
               DataSource      =   "data1"
               Height          =   285
               Index           =   0
               Left            =   5310
               Locked          =   -1  'True
               TabIndex        =   323
               TabStop         =   0   'False
               Top             =   3990
               Width           =   3540
            End
            Begin VB.ComboBox ctipusimp 
               BackColor       =   &H00FFC0C0&
               Height          =   315
               HelpContextID   =   99
               ItemData        =   "comandes.frx":A986
               Left            =   8175
               List            =   "comandes.frx":A993
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   1125
               Width           =   1440
            End
            Begin VB.ComboBox cimpressio 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               ItemData        =   "comandes.frx":A9B1
               Left            =   8175
               List            =   "comandes.frx":A9B8
               TabIndex        =   306
               Top             =   465
               Width           =   1425
            End
            Begin VB.TextBox Text142 
               DataField       =   "texteimpressio"
               DataSource      =   "data1"
               Height          =   285
               Left            =   810
               TabIndex        =   36
               TabStop         =   0   'False
               ToolTipText     =   "Texte d'Impressió"
               Top             =   2580
               Width           =   5010
            End
            Begin VB.TextBox text77 
               DataField       =   "obsimp1"
               DataSource      =   "data1"
               Height          =   285
               Index           =   1
               Left            =   510
               TabIndex        =   321
               Top             =   3975
               Width           =   4650
            End
            Begin VB.CheckBox tincclixes 
               Caption         =   "Tinc Clixes"
               DataField       =   "tincclixes"
               Height          =   195
               Left            =   5520
               TabIndex        =   35
               Top             =   3420
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.TextBox text77 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "lin1"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Index           =   2
               Left            =   5715
               TabIndex        =   279
               Top             =   240
               Width           =   450
            End
            Begin VB.CommandButton Command1 
               Height          =   330
               Index           =   1
               Left            =   6225
               Picture         =   "comandes.frx":A9CE
               Style           =   1  'Graphical
               TabIndex        =   34
               TabStop         =   0   'False
               ToolTipText     =   "Visualitzar el Clixe del treball"
               Top             =   3120
               Width           =   315
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00FFC0C0&
               DataField       =   "marques"
               DataSource      =   "data1"
               Height          =   315
               Index           =   1
               ItemData        =   "comandes.frx":AF58
               Left            =   8175
               List            =   "comandes.frx":AF62
               TabIndex        =   308
               Top             =   780
               Width           =   705
            End
            Begin VB.CommandButton Command1 
               Height          =   315
               Index           =   2
               Left            =   8205
               Picture         =   "comandes.frx":AF6E
               Style           =   1  'Graphical
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   3630
               Width           =   450
            End
            Begin VB.CommandButton Command1 
               Height          =   315
               Index           =   3
               Left            =   5895
               Picture         =   "comandes.frx":B4F8
               Style           =   1  'Graphical
               TabIndex        =   32
               TabStop         =   0   'False
               ToolTipText     =   "Canviar el numero i modificacio del treball"
               Top             =   3135
               Width           =   315
            End
            Begin MSMask.MaskEdBox Text63 
               DataField       =   "numerotintes"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Left            =   7665
               TabIndex        =   302
               TabStop         =   0   'False
               Top             =   105
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16761024
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text64 
               DataField       =   "impressio"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8985
               TabIndex        =   37
               Top             =   450
               Visible         =   0   'False
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text65 
               DataField       =   "formaimp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9465
               TabIndex        =   38
               Top             =   465
               Visible         =   0   'False
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text66 
               DataField       =   "dessarroll"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Left            =   8175
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   1470
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16761024
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text67 
               DataField       =   "cilindres"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Left            =   9330
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   1470
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16761024
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text68 
               DataField       =   "obert"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8175
               TabIndex        =   310
               Top             =   1770
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text69 
               DataField       =   "arxiu"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Left            =   6675
               TabIndex        =   42
               TabStop         =   0   'False
               ToolTipText     =   "Arxiu al treball"
               Top             =   2850
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16761024
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text70 
               DataField       =   "arxiumontadora"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Left            =   7785
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   2820
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16761024
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text72 
               DataField       =   "mtrsminut"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6600
               TabIndex        =   312
               Top             =   2565
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text73 
               DataField       =   "impressora"
               DataSource      =   "data1"
               Height          =   285
               Left            =   825
               TabIndex        =   318
               Top             =   3150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox7 
               DataField       =   "diamextbobimp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4785
               TabIndex        =   316
               Top             =   2850
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox8 
               DataField       =   "mtrslinbobimp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3330
               TabIndex        =   315
               Top             =   2850
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox9 
               DataField       =   "kilosbobimp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1890
               TabIndex        =   314
               Top             =   2850
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox10 
               DataField       =   "tubbaseimp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   825
               TabIndex        =   313
               Top             =   2850
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox15 
               DataField       =   "continu"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Left            =   8355
               TabIndex        =   304
               TabStop         =   0   'False
               Top             =   120
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16761024
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox20 
               DataField       =   "gruixpol"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Left            =   9330
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   1770
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16761024
               AutoTab         =   -1  'True
               MaxLength       =   4
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text103 
               Height          =   285
               HelpContextID   =   99
               Index           =   3
               Left            =   4785
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   3120
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16761024
               PromptInclude   =   0   'False
               MaxLength       =   12
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text103 
               DataField       =   "marcailinia"
               DataSource      =   "data1"
               Height          =   285
               HelpContextID   =   99
               Index           =   4
               Left            =   60
               TabIndex        =   345
               Top             =   3690
               Width           =   5385
               _ExtentX        =   9499
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   16761024
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin VB.TextBox Text40 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               DataField       =   "tinta1a"
               DataSource      =   "data1"
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
               HelpContextID   =   99
               Left            =   210
               TabIndex        =   263
               Top             =   240
               WhatsThisHelpID =   2
               Width           =   5460
            End
            Begin MSMask.MaskEdBox Text71 
               DataField       =   "codibarras"
               DataSource      =   "data1"
               Height          =   255
               HelpContextID   =   99
               Left            =   7770
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   2580
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   450
               _Version        =   327681
               BackColor       =   16761024
               PromptChar      =   "_"
            End
            Begin VB.TextBox cobs2imp 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
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
               HelpContextID   =   99
               Left            =   435
               TabIndex        =   449
               Top             =   2325
               Width           =   8310
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               DataSource      =   "data1"
               ForeColor       =   &H000000FF&
               Height          =   225
               Index           =   179
               Left            =   8715
               TabIndex        =   504
               Top             =   2070
               Width           =   1125
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Des.Client:"
               DataSource      =   "data1"
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
               Index           =   22
               Left            =   8640
               TabIndex        =   462
               ToolTipText     =   "Dessarroll que el client diu que serà."
               Top             =   135
               Width           =   750
            End
            Begin VB.Label label1 
               Caption         =   "Obs2:"
               DataSource      =   "data1"
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
               Index           =   167
               Left            =   45
               TabIndex        =   451
               Top             =   2355
               Width           =   915
            End
            Begin VB.Label label1 
               Caption         =   "Obs1:"
               DataSource      =   "data1"
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
               Index           =   166
               Left            =   45
               TabIndex        =   450
               Top             =   2130
               Width           =   915
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF00FF&
               Caption         =   "ATENCIÓ IMPRESIÓ AMB REPRINT"
               DataSource      =   "data1"
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   162
               Left            =   1155
               TabIndex        =   432
               Top             =   15
               WhatsThisHelpID =   1
               Width           =   6000
            End
            Begin VB.Label numtreballnoborrar 
               Caption         =   "numtreballnoborrar"
               DataField       =   "numtreball"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   465
               TabIndex        =   420
               Top             =   4110
               Visible         =   0   'False
               Width           =   1830
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "....."
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   31
               Left            =   3090
               TabIndex        =   407
               Top             =   3480
               Width           =   3630
            End
            Begin VB.Label label1 
               Caption         =   "ImpVell"
               DataSource      =   "clients"
               Height          =   255
               Index           =   23
               Left            =   5475
               TabIndex        =   398
               Top             =   3735
               Width           =   525
            End
            Begin VB.Label label1 
               Caption         =   "8ª"
               DataSource      =   "data1"
               Height          =   255
               Index           =   44
               Left            =   45
               TabIndex        =   83
               Top             =   1800
               Width           =   210
            End
            Begin VB.Label label1 
               Caption         =   "7ª "
               DataSource      =   "data1"
               Height          =   255
               Index           =   45
               Left            =   45
               TabIndex        =   82
               Top             =   1605
               Width           =   210
            End
            Begin VB.Label label1 
               Caption         =   "A.I:"
               DataSource      =   "clients"
               Height          =   255
               Index           =   64
               Left            =   7860
               TabIndex        =   81
               Top             =   3720
               Width           =   345
            End
            Begin VB.Label label1 
               Caption         =   "PDF"
               DataSource      =   "clients"
               Height          =   255
               Index           =   63
               Left            =   6840
               TabIndex        =   80
               Top             =   3720
               Width           =   465
            End
            Begin VB.Label label1 
               Caption         =   "2"
               DataSource      =   "data1"
               Height          =   225
               Index           =   62
               Left            =   5175
               TabIndex        =   79
               Top             =   4020
               Width           =   195
            End
            Begin VB.Label label1 
               Caption         =   "Obs."
               DataSource      =   "data1"
               Height          =   270
               Index           =   61
               Left            =   90
               TabIndex        =   78
               Top             =   4020
               Width           =   420
            End
            Begin VB.Label label1 
               Caption         =   "Impress.:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   60
               Left            =   120
               TabIndex        =   77
               Top             =   3225
               Width           =   795
            End
            Begin VB.Label nomimpressora 
               Caption         =   "nomimpressora"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   1
               Left            =   1290
               TabIndex        =   76
               Top             =   3165
               Width           =   1275
            End
            Begin VB.Label label1 
               Caption         =   "Mtrs/Min.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   59
               Left            =   5850
               TabIndex        =   75
               Top             =   2640
               Width           =   915
            End
            Begin VB.Label label1 
               Caption         =   "C.Bar.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   58
               Left            =   7275
               TabIndex        =   74
               Top             =   2625
               Width           =   570
            End
            Begin VB.Label label1 
               Caption         =   "A.M:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   57
               Left            =   7395
               TabIndex        =   73
               Top             =   2880
               Width           =   435
            End
            Begin VB.Label label1 
               Caption         =   "Arxiu:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   56
               Left            =   5355
               TabIndex        =   72
               Top             =   2895
               Width           =   540
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Obert (N/1/2/C)"
               DataSource      =   "data1"
               Height          =   390
               Index           =   55
               Left            =   7245
               TabIndex        =   71
               Top             =   1680
               Width           =   750
            End
            Begin VB.Label label1 
               Caption         =   "Cilindres:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   54
               Left            =   8685
               TabIndex        =   70
               Top             =   1500
               Width           =   630
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Desarroll:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   53
               Left            =   7245
               TabIndex        =   69
               Top             =   1500
               Width           =   750
            End
            Begin VB.Label label1 
               Caption         =   "Forma Im:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   52
               Left            =   7245
               TabIndex        =   68
               Top             =   1155
               Width           =   750
            End
            Begin VB.Label label1 
               Caption         =   "Impressió:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   51
               Left            =   7245
               TabIndex        =   67
               Top             =   465
               Width           =   750
            End
            Begin VB.Label label1 
               Caption         =   "Tintes:"
               DataSource      =   "data1"
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
               Index           =   50
               Left            =   7245
               TabIndex        =   66
               Top             =   120
               Width           =   600
            End
            Begin VB.Label label1 
               Caption         =   "6ª"
               DataSource      =   "data1"
               Height          =   255
               Index           =   43
               Left            =   45
               TabIndex        =   65
               Top             =   1380
               Width           =   210
            End
            Begin VB.Label label1 
               Caption         =   "5ª"
               DataSource      =   "data1"
               Height          =   255
               Index           =   42
               Left            =   45
               TabIndex        =   64
               Top             =   1155
               Width           =   210
            End
            Begin VB.Label label1 
               Caption         =   "4ª "
               DataSource      =   "data1"
               Height          =   255
               Index           =   41
               Left            =   45
               TabIndex        =   63
               Top             =   930
               Width           =   210
            End
            Begin VB.Label label1 
               Caption         =   "3ª"
               DataSource      =   "data1"
               Height          =   255
               Index           =   40
               Left            =   45
               TabIndex        =   62
               Top             =   705
               Width           =   195
            End
            Begin VB.Label label1 
               Caption         =   "2ª "
               DataSource      =   "data1"
               Height          =   255
               Index           =   39
               Left            =   45
               TabIndex        =   61
               Top             =   450
               Width           =   240
            End
            Begin VB.Label label1 
               Caption         =   "1ª"
               DataSource      =   "data1"
               Height          =   255
               Index           =   38
               Left            =   45
               TabIndex        =   60
               Top             =   240
               Width           =   225
            End
            Begin VB.Label label1 
               Caption         =   "Texte Alt.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   6
               Left            =   60
               TabIndex        =   59
               ToolTipText     =   "Texte d'Impressió"
               Top             =   2610
               Width           =   885
            End
            Begin VB.Label label1 
               Caption         =   "Tub Base:"
               DataSource      =   "data1"
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   129
               Left            =   75
               TabIndex        =   58
               Top             =   2880
               Width           =   825
            End
            Begin VB.Label label1 
               Caption         =   "Kg Bob:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   130
               Left            =   1305
               TabIndex        =   57
               Top             =   2895
               Width           =   705
            End
            Begin VB.Label label1 
               Caption         =   "Mtrs Bob:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   131
               Left            =   2625
               TabIndex        =   56
               Top             =   2895
               Width           =   765
            End
            Begin VB.Label label1 
               Caption         =   "Diametre:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   132
               Left            =   4080
               TabIndex        =   55
               Top             =   2880
               Width           =   855
            End
            Begin VB.Label label1 
               Caption         =   "Cont.:"
               DataSource      =   "data1"
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
               Index           =   134
               Left            =   7950
               TabIndex        =   54
               Top             =   120
               Width           =   465
            End
            Begin VB.Label label1 
               Caption         =   "Reducció x Mtr:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   135
               Left            =   6765
               TabIndex        =   53
               Top             =   3150
               Width           =   1290
            End
            Begin VB.Label label1 
               Caption         =   "Gruix Pol:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   145
               Left            =   8670
               TabIndex        =   52
               Top             =   1815
               Width           =   750
            End
            Begin VB.Label label1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H008080FF&
               Height          =   195
               Index           =   153
               Left            =   135
               TabIndex        =   51
               Top             =   3510
               Width           =   5160
            End
            Begin VB.Label label1 
               Caption         =   "Nº Treball:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   154
               Left            =   3975
               TabIndex        =   50
               Top             =   3165
               Width           =   930
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Flexo Kodak - Clixes 2 Bandes"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   159
               Left            =   6240
               TabIndex        =   49
               Top             =   3450
               Width           =   3630
            End
            Begin VB.Label label1 
               Caption         =   "Canvi Envio:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   11
               Left            =   7245
               TabIndex        =   48
               Top             =   840
               Width           =   915
            End
         End
         Begin VB.Frame lam1 
            Caption         =   "Laminadora-1"
            Height          =   3975
            Left            =   90
            TabIndex        =   119
            Top             =   11010
            Width           =   9885
            Begin VB.CheckBox materialexacte 
               BackColor       =   &H80000005&
               Caption         =   "Cola Exacte"
               Height          =   540
               Index           =   2
               Left            =   165
               TabIndex        =   466
               TabStop         =   0   'False
               Top             =   1095
               Width           =   810
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00FFFFFF&
               DataSource      =   "data1"
               Height          =   315
               Index           =   0
               ItemData        =   "comandes.frx":BA82
               Left            =   1830
               List            =   "comandes.frx":BA8F
               TabIndex        =   405
               Top             =   795
               Width           =   525
            End
            Begin VB.CheckBox primerproces 
               DataSource      =   "data1"
               Height          =   195
               Left            =   180
               TabIndex        =   403
               ToolTipText     =   "Forçar que aquest sigui el primer procès"
               Top             =   870
               Width           =   255
            End
            Begin VB.TextBox text77 
               BackColor       =   &H00FFC0C0&
               DataField       =   "arxiuext"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   18
               Left            =   150
               Locked          =   -1  'True
               TabIndex        =   401
               Top             =   2580
               Width           =   9645
            End
            Begin VB.TextBox adhesiu 
               Height          =   285
               Left            =   1035
               TabIndex        =   126
               TabStop         =   0   'False
               Top             =   1170
               Width           =   3000
            End
            Begin VB.TextBox Text93 
               DataField       =   "obslamgen1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   350
               Top             =   3390
               Width           =   7080
            End
            Begin VB.TextBox Text95 
               DataField       =   "obslam1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   915
               TabIndex        =   349
               Top             =   3090
               Width           =   7065
            End
            Begin VB.CommandButton Command4 
               Height          =   285
               Left            =   9525
               Picture         =   "comandes.frx":BAAA
               Style           =   1  'Graphical
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   3075
               Width           =   285
            End
            Begin VB.TextBox Text97 
               DataField       =   "arxiulaminadora"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   8160
               Locked          =   -1  'True
               MouseIcon       =   "comandes.frx":BE70
               MousePointer    =   99  'Custom
               TabIndex        =   120
               TabStop         =   0   'False
               Top             =   3075
               Width           =   1320
            End
            Begin VB.ComboBox Combo2 
               DataField       =   "simulteneitatlam"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":BFC2
               Left            =   8895
               List            =   "comandes.frx":BFE1
               TabIndex        =   333
               Top             =   480
               Width           =   675
            End
            Begin MSMask.MaskEdBox Text92 
               DataField       =   "mtr/minrodillocola"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8895
               TabIndex        =   348
               Top             =   1710
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text90 
               DataField       =   "rodillocola"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8895
               TabIndex        =   347
               Top             =   1410
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox grmt2 
               DataField       =   "grmt2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   7050
               TabIndex        =   122
               Top             =   1440
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox vadhesiu 
               DataField       =   "tipusadhesiu"
               DataSource      =   "data1"
               Height          =   285
               Left            =   405
               TabIndex        =   343
               Top             =   1170
               Visible         =   0   'False
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               _Version        =   327681
               BackColor       =   13078010
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox pes2 
               DataField       =   "pes2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4800
               TabIndex        =   123
               Top             =   1590
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox pes1 
               DataField       =   "pes1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4800
               TabIndex        =   124
               Top             =   1290
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSFlexGridLib.MSFlexGrid reixaconsums 
               Height          =   870
               Left            =   210
               TabIndex        =   125
               TabStop         =   0   'False
               Tag             =   "1"
               Top             =   3750
               Visible         =   0   'False
               Width           =   9645
               _ExtentX        =   17013
               _ExtentY        =   1535
               _Version        =   393216
               Rows            =   3
               Cols            =   16
               FixedCols       =   0
               BackColor       =   16777215
               ForeColor       =   0
               ForeColorFixed  =   16711680
               ForeColorSel    =   0
               AllowBigSelection=   0   'False
               TextStyle       =   3
               ScrollBars      =   0
            End
            Begin MSMask.MaskEdBox Text91 
               DataField       =   "camisa"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8895
               TabIndex        =   346
               Top             =   1110
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text89 
               DataField       =   "amplelaminar"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8880
               TabIndex        =   341
               TabStop         =   0   'False
               Top             =   810
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               _Version        =   327681
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text87 
               DataField       =   "ampleutil"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8895
               TabIndex        =   329
               Top             =   180
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text86 
               DataField       =   "tensiototal"
               DataSource      =   "data1"
               Height          =   285
               Left            =   7275
               TabIndex        =   339
               Top             =   795
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text85 
               DataField       =   "tensiodesb2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   7275
               TabIndex        =   332
               Top             =   495
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text84 
               DataField       =   "tensiodesb1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   7275
               TabIndex        =   327
               Top             =   195
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text83 
               DataField       =   "mtr/minmaquina"
               DataSource      =   "data1"
               Height          =   285
               Left            =   5115
               TabIndex        =   337
               Top             =   2040
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text82 
               DataField       =   "laminadora"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1080
               TabIndex        =   335
               Top             =   2040
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text81 
               DataField       =   "lotmatdesb2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1050
               TabIndex        =   330
               Top             =   495
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   503
               _Version        =   327681
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text80 
               DataField       =   "lotmatdesb1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1050
               TabIndex        =   325
               Top             =   210
               WhatsThisHelpID =   3
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   503
               _Version        =   327681
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin VB.Label enduridor 
               BackStyle       =   0  'Transparent
               Caption         =   "DESCRIPCIO DE LES FAMILIES"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   210
               TabIndex        =   465
               Top             =   1845
               Width           =   3900
            End
            Begin VB.Label enduridor 
               BackStyle       =   0  'Transparent
               Caption         =   "DESCRIPCIO DE L'ENDURIDOR"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   330
               Index           =   0
               Left            =   1065
               TabIndex        =   146
               Top             =   1500
               Width           =   3000
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Metall:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   30
               Left            =   1290
               TabIndex        =   406
               Top             =   870
               Width           =   600
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "1r. Procès"
               DataSource      =   "data1"
               Height          =   255
               Index           =   24
               Left            =   450
               TabIndex        =   404
               Top             =   855
               Width           =   915
            End
            Begin VB.Label Label2 
               Caption         =   "Laminació:"
               Height          =   270
               Left            =   165
               TabIndex        =   402
               Top             =   2370
               Width           =   780
            End
            Begin VB.Shape fondocoloradhesiu 
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   945
               Left            =   120
               Top             =   1095
               Width           =   4005
            End
            Begin VB.Label label1 
               Caption         =   "Lot Desb 1:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   65
               Left            =   135
               TabIndex        =   158
               Top             =   255
               Width           =   840
            End
            Begin VB.Label label1 
               Caption         =   "Lot Desb 2:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   66
               Left            =   135
               TabIndex        =   157
               Top             =   555
               Width           =   900
            End
            Begin VB.Label label1 
               Caption         =   "Laminadora:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   67
               Left            =   165
               TabIndex        =   156
               Top             =   2100
               Width           =   1005
            End
            Begin VB.Label nomlaminadora 
               Caption         =   "nomlaminadora"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   1845
               TabIndex        =   155
               Top             =   2115
               Width           =   2505
            End
            Begin VB.Label label1 
               Caption         =   "Mtrs/Min.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   68
               Left            =   4380
               TabIndex        =   154
               Top             =   2100
               Width           =   915
            End
            Begin VB.Label label1 
               Caption         =   "Tensió D1:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   69
               Left            =   6450
               TabIndex        =   153
               Top             =   255
               Width           =   915
            End
            Begin VB.Label label1 
               Caption         =   "Tensió D2:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   70
               Left            =   6450
               TabIndex        =   152
               Top             =   555
               Width           =   885
            End
            Begin VB.Label label1 
               Caption         =   "Tensió Tl:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   71
               Left            =   6510
               TabIndex        =   151
               Top             =   840
               Width           =   750
            End
            Begin VB.Label label1 
               Caption         =   "Ample Útil:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   72
               Left            =   7935
               TabIndex        =   150
               Top             =   240
               Width           =   855
            End
            Begin VB.Label label1 
               Caption         =   "Simultaneitat:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   73
               Left            =   7920
               TabIndex        =   149
               Top             =   525
               Width           =   975
            End
            Begin VB.Label label1 
               Caption         =   "Ample Lam.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   74
               Left            =   7905
               TabIndex        =   148
               Top             =   855
               Width           =   975
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Descripció Adhesiu i Enduridor  (F2)"
               DataSource      =   "data1"
               Height          =   255
               Index           =   75
               Left            =   1260
               TabIndex        =   147
               Top             =   1095
               Width           =   2610
            End
            Begin VB.Label label1 
               Caption         =   "Camisa:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   76
               Left            =   8100
               TabIndex        =   145
               Top             =   1170
               Width           =   765
            End
            Begin VB.Label label1 
               Caption         =   "Gr.Cm2"
               DataSource      =   "data1"
               Height          =   255
               Index           =   77
               Left            =   4200
               TabIndex        =   144
               Top             =   1095
               Width           =   615
            End
            Begin VB.Label label1 
               Caption         =   "ºC"
               DataSource      =   "data1"
               Height          =   255
               Index           =   78
               Left            =   5505
               TabIndex        =   143
               Top             =   1095
               Width           =   255
            End
            Begin VB.Label label1 
               Caption         =   "%Litres"
               DataSource      =   "data1"
               Height          =   255
               Index           =   79
               Left            =   5985
               TabIndex        =   142
               Top             =   1095
               Width           =   615
            End
            Begin VB.Label label1 
               Caption         =   "% Pes"
               DataSource      =   "data1"
               Height          =   255
               Index           =   80
               Left            =   4875
               TabIndex        =   141
               Top             =   1095
               Width           =   540
            End
            Begin VB.Label grcm1 
               Alignment       =   2  'Center
               Caption         =   "Grcm1"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   270
               Index           =   0
               Left            =   4170
               TabIndex        =   140
               Top             =   1335
               Width           =   570
            End
            Begin VB.Label ºC1 
               Alignment       =   2  'Center
               Caption         =   "ºC1"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   225
               Index           =   0
               Left            =   5400
               TabIndex        =   139
               Top             =   1335
               Width           =   390
            End
            Begin VB.Label label1 
               Caption         =   "Cola Gr/mt2:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   81
               Left            =   6915
               TabIndex        =   138
               Top             =   1245
               Width           =   1020
            End
            Begin VB.Label grcm2 
               Alignment       =   2  'Center
               Caption         =   "Grcm2"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   270
               Index           =   0
               Left            =   4170
               TabIndex        =   137
               Top             =   1650
               Width           =   570
            End
            Begin VB.Label ºC2 
               Alignment       =   2  'Center
               Caption         =   "ºC2"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   270
               Index           =   0
               Left            =   5415
               TabIndex        =   136
               Top             =   1650
               Width           =   390
            End
            Begin VB.Label litres1 
               Alignment       =   2  'Center
               Caption         =   "--"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   225
               Index           =   1
               Left            =   5895
               TabIndex        =   135
               Top             =   1350
               Width           =   930
            End
            Begin VB.Label litres2 
               Alignment       =   2  'Center
               Caption         =   "--"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   225
               Index           =   2
               Left            =   5955
               TabIndex        =   134
               Top             =   1665
               Width           =   765
            End
            Begin VB.Label label1 
               Caption         =   "Rodillo Cola:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   82
               Left            =   7995
               TabIndex        =   133
               Top             =   1470
               Width           =   960
            End
            Begin VB.Label label1 
               Caption         =   "Mtrs/Min Rodillo:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   83
               Left            =   7680
               TabIndex        =   132
               Top             =   1770
               Width           =   1350
            End
            Begin VB.Label label1 
               Caption         =   "Obs.Cli.:"
               DataSource      =   "data1"
               Height          =   240
               Index           =   84
               Left            =   150
               TabIndex        =   131
               Top             =   3435
               Width           =   675
            End
            Begin VB.Label label1 
               Caption         =   "Obs.Lam:"
               DataSource      =   "data1"
               Height          =   285
               Index           =   85
               Left            =   150
               TabIndex        =   130
               Top             =   3135
               Width           =   750
            End
            Begin VB.Label label1 
               Caption         =   "Arxiu Laminadora:"
               DataSource      =   "clients"
               Height          =   255
               Index           =   86
               Left            =   8235
               TabIndex        =   129
               Top             =   2865
               Width           =   1365
            End
            Begin VB.Label desclot1 
               Caption         =   "descripcio del lot1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   345
               Index           =   0
               Left            =   1875
               TabIndex        =   128
               Top             =   210
               Width           =   4590
            End
            Begin VB.Label desclot2 
               Caption         =   "descripcio del lot2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   300
               Index           =   1
               Left            =   1890
               TabIndex        =   127
               Top             =   555
               Width           =   4515
            End
         End
         Begin VB.Frame reb 
            Caption         =   "Rebobinadora"
            Height          =   3390
            Left            =   75
            TabIndex        =   84
            Top             =   15060
            Width           =   9885
            Begin VB.TextBox Text32 
               DataField       =   "obsreb2"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   13
               Left            =   975
               MaxLength       =   80
               TabIndex        =   481
               Top             =   2775
               Width           =   7110
            End
            Begin VB.ComboBox Combo5 
               DataField       =   "rebmacroperforat"
               DataSource      =   "data1"
               Height          =   315
               Index           =   1
               ItemData        =   "comandes.frx":C000
               Left            =   9210
               List            =   "comandes.frx":C00A
               TabIndex        =   427
               Top             =   690
               Width           =   585
            End
            Begin VB.ComboBox Combo3 
               DataField       =   "simulteneitatreb"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":C014
               Left            =   4050
               List            =   "comandes.frx":C060
               TabIndex        =   353
               Top             =   405
               Width           =   675
            End
            Begin VB.ComboBox Combo4 
               DataField       =   "caratractada"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":C0BB
               Left            =   6240
               List            =   "comandes.frx":C0C8
               TabIndex        =   356
               Top             =   720
               Width           =   615
            End
            Begin VB.ComboBox Combo5 
               DataField       =   "microperforat"
               DataSource      =   "data1"
               Height          =   315
               Index           =   0
               ItemData        =   "comandes.frx":C0D4
               Left            =   7695
               List            =   "comandes.frx":C0E4
               TabIndex        =   357
               Top             =   705
               Width           =   585
            End
            Begin VB.ComboBox Combo6 
               DataField       =   "etiqextbobina"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":C0F4
               Left            =   7950
               List            =   "comandes.frx":C101
               TabIndex        =   363
               Top             =   1050
               Width           =   585
            End
            Begin VB.ComboBox Combo7 
               DataField       =   "etiqintcanutu"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":C123
               Left            =   6255
               List            =   "comandes.frx":C130
               TabIndex        =   362
               Top             =   1065
               Width           =   615
            End
            Begin VB.TextBox Text106 
               DataField       =   "obsreb1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   585
               TabIndex        =   368
               Top             =   2490
               Width           =   7500
            End
            Begin VB.TextBox Text109 
               DataField       =   "arxiureb"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   8130
               Locked          =   -1  'True
               MouseIcon       =   "comandes.frx":C152
               MousePointer    =   99  'Custom
               TabIndex        =   89
               TabStop         =   0   'False
               Top             =   2490
               Width           =   1380
            End
            Begin VB.CommandButton Command5 
               Height          =   285
               Left            =   9555
               Picture         =   "comandes.frx":C2A4
               Style           =   1  'Graphical
               TabIndex        =   88
               TabStop         =   0   'False
               Top             =   2490
               Width           =   285
            End
            Begin VB.TextBox Text108 
               DataField       =   "obsrebgen1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   975
               Locked          =   -1  'True
               TabIndex        =   369
               Top             =   3060
               Width           =   7110
            End
            Begin VB.ComboBox Combo9 
               DataField       =   "migelaborat"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":C66A
               Left            =   1245
               List            =   "comandes.frx":C677
               TabIndex        =   351
               Top             =   345
               WhatsThisHelpID =   4
               Width           =   645
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Guarda Mostra final"
               DataField       =   "rebguardarmostrafinal"
               DataSource      =   "data1"
               Height          =   210
               Index           =   0
               Left            =   7590
               TabIndex        =   87
               TabStop         =   0   'False
               Top             =   1920
               Width           =   1935
            End
            Begin VB.TextBox text77 
               DataSource      =   "data1"
               Height          =   285
               Index           =   13
               Left            =   2205
               Locked          =   -1  'True
               TabIndex        =   366
               TabStop         =   0   'False
               Top             =   1860
               Width           =   2520
            End
            Begin VB.TextBox text77 
               DataField       =   "rebidbobinesembolicades"
               DataSource      =   "data1"
               Height          =   285
               Index           =   14
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   86
               TabStop         =   0   'False
               Top             =   1890
               Visible         =   0   'False
               Width           =   105
            End
            Begin VB.TextBox text77 
               DataField       =   "rebidtipusetiqueta"
               DataSource      =   "data1"
               Height          =   285
               Index           =   15
               Left            =   4755
               Locked          =   -1  'True
               TabIndex        =   85
               TabStop         =   0   'False
               Top             =   1905
               Visible         =   0   'False
               Width           =   105
            End
            Begin VB.TextBox text77 
               DataSource      =   "data1"
               Height          =   285
               Index           =   16
               Left            =   4920
               Locked          =   -1  'True
               TabIndex        =   367
               TabStop         =   0   'False
               Top             =   1875
               Width           =   2520
            End
            Begin MSMask.MaskEdBox Text104 
               DataField       =   "diamextbob"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4410
               TabIndex        =   361
               Top             =   1095
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text103 
               DataField       =   "mtrslinbob"
               DataSource      =   "data1"
               Height          =   285
               Index           =   0
               Left            =   2970
               TabIndex        =   360
               Top             =   1095
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text102 
               DataField       =   "kilosbob"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1260
               TabIndex        =   359
               Top             =   1095
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text101 
               DataField       =   "tubbase"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9360
               TabIndex        =   358
               Top             =   1065
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text100 
               DataField       =   "matintbob"
               DataSource      =   "data1"
               Height          =   285
               Left            =   5145
               TabIndex        =   354
               Top             =   405
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text99 
               DataField       =   "rebobinadora"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1260
               TabIndex        =   355
               Top             =   720
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text98 
               DataField       =   "amplereb"
               DataSource      =   "data1"
               Height          =   285
               Left            =   2475
               TabIndex        =   352
               Top             =   405
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   327681
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text103 
               DataField       =   "rebnumpisosbob"
               DataSource      =   "data1"
               Height          =   285
               Index           =   1
               Left            =   1245
               TabIndex        =   364
               Top             =   1530
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   327681
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text103 
               DataField       =   "rebnumbobxpis"
               DataSource      =   "data1"
               Height          =   285
               Index           =   2
               Left            =   1245
               TabIndex        =   365
               Top             =   1875
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   327681
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text103 
               Height          =   285
               Index           =   5
               Left            =   9375
               TabIndex        =   461
               ToolTipText     =   "Matrial del tub base. Cartró o Pvc"
               Top             =   1410
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               _Version        =   327681
               MaxLength       =   1
               PromptChar      =   "_"
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF00FF&
               Caption         =   "MICRO/MACROPERFORAT"
               DataSource      =   "data1"
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   163
               Left            =   6555
               TabIndex        =   433
               Top             =   0
               Width           =   3210
            End
            Begin VB.Label rebpes 
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
               Height          =   285
               Left            =   1245
               TabIndex        =   100
               Top             =   2235
               Width           =   2820
            End
            Begin VB.Label rebmetres 
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
               Height          =   270
               Left            =   4560
               TabIndex        =   99
               Top             =   2235
               Width           =   1995
            End
            Begin VB.Label rebpcs 
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
               Left            =   7065
               TabIndex        =   90
               Top             =   2250
               Width           =   1335
            End
            Begin VB.Label label1 
               Caption         =   "Obs.Operari:"
               DataSource      =   "data1"
               Height          =   195
               Index           =   175
               Left            =   45
               TabIndex        =   480
               Top             =   2775
               Width           =   870
            End
            Begin VB.Label label1 
               Caption         =   "Material Tub base: (C) Cartró (P) PVC"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   420
               Index           =   20
               Left            =   8085
               TabIndex        =   460
               Top             =   1395
               Width           =   1320
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(Fred/Calent/Laser)"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   168
               Left            =   7050
               TabIndex        =   452
               Top             =   510
               Width           =   1245
            End
            Begin VB.Label label1 
               Caption         =   "MACROperf:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   160
               Left            =   8265
               TabIndex        =   428
               Top             =   780
               Width           =   885
            End
            Begin VB.Label label1 
               Caption         =   "Tubo o Lam:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   87
               Left            =   1140
               TabIndex        =   118
               Top             =   150
               Width           =   1125
            End
            Begin VB.Label label1 
               Caption         =   "Ample Reb:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   88
               Left            =   2460
               TabIndex        =   117
               Top             =   180
               Width           =   855
            End
            Begin VB.Label label1 
               Caption         =   "Simultaneitat:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   89
               Left            =   3885
               TabIndex        =   116
               Top             =   180
               Width           =   975
            End
            Begin VB.Label nomrebobinadora 
               Caption         =   "nomrebobinadora"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   1
               Left            =   1875
               TabIndex        =   115
               Top             =   810
               Width           =   3045
            End
            Begin VB.Label label1 
               Caption         =   "Rebobinadora:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   90
               Left            =   105
               TabIndex        =   114
               Top             =   795
               Width           =   1125
            End
            Begin VB.Label desclot1 
               Caption         =   "descripcio del lot1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   330
               Index           =   1
               Left            =   6150
               TabIndex        =   113
               Top             =   405
               Width           =   3570
            End
            Begin VB.Label label1 
               Caption         =   "Lot Mat. Int. Bob."
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   91
               Left            =   5070
               TabIndex        =   112
               Top             =   180
               Width           =   1725
            End
            Begin VB.Label label1 
               Caption         =   "Cara Tractada:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   92
               Left            =   5130
               TabIndex        =   111
               Top             =   795
               Width           =   1170
            End
            Begin VB.Label label1 
               Caption         =   "Microperf:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   93
               Left            =   6930
               TabIndex        =   110
               Top             =   705
               Width           =   1170
            End
            Begin VB.Label label1 
               Caption         =   "Tub Base:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   270
               Index           =   94
               Left            =   8580
               TabIndex        =   109
               Top             =   1125
               Width           =   825
            End
            Begin VB.Label label1 
               Caption         =   "Kilos Bobina:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   95
               Left            =   120
               TabIndex        =   108
               Top             =   1140
               Width           =   1020
            End
            Begin VB.Label label1 
               Caption         =   "Mtrs Bobina:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   96
               Left            =   2010
               TabIndex        =   107
               Top             =   1125
               Width           =   1020
            End
            Begin VB.Label label1 
               Caption         =   "Diametre:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   97
               Left            =   3705
               TabIndex        =   106
               Top             =   1125
               Width           =   855
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Et. Ext. Bob:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   98
               Left            =   6945
               TabIndex        =   105
               Top             =   1140
               Width           =   1050
            End
            Begin VB.Label label1 
               Caption         =   "Et. Int. Canutu:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   99
               Left            =   5145
               TabIndex        =   104
               Top             =   1140
               Width           =   1170
            End
            Begin VB.Label label1 
               Caption         =   "Obs.Client:"
               DataSource      =   "data1"
               Height          =   195
               Index           =   101
               Left            =   45
               TabIndex        =   102
               Top             =   3105
               Width           =   870
            End
            Begin VB.Label label1 
               Caption         =   "Arxiu Reb:"
               DataSource      =   "clients"
               Height          =   255
               Index           =   102
               Left            =   8505
               TabIndex        =   101
               Top             =   2250
               Width           =   1020
            End
            Begin VB.Label label1 
               Caption         =   "NºPisos Bob:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   149
               Left            =   135
               TabIndex        =   96
               Top             =   1590
               Width           =   1020
            End
            Begin VB.Label label1 
               Caption         =   "Nº Bob. x pis:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   150
               Left            =   135
               TabIndex        =   95
               Top             =   1920
               Width           =   1020
            End
            Begin VB.Label label1 
               Caption         =   "Bobines Embolicades (F2)"
               DataSource      =   "data1"
               Height          =   270
               Index           =   151
               Left            =   2400
               TabIndex        =   94
               Top             =   1650
               Width           =   2280
            End
            Begin VB.Label label1 
               Caption         =   "Tipus d'etiquetes (F2)"
               DataSource      =   "data1"
               Height          =   270
               Index           =   152
               Left            =   5115
               TabIndex        =   93
               Top             =   1665
               Width           =   2280
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Height          =   270
               Left            =   0
               TabIndex        =   91
               Top             =   0
               Width           =   990
            End
            Begin VB.Label label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Obs."
               Height          =   180
               Index           =   100
               Left            =   165
               TabIndex        =   103
               Top             =   2505
               Width           =   420
            End
            Begin VB.Label label1 
               Caption         =   "Pes Kg:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   143
               Left            =   645
               TabIndex        =   98
               Top             =   2250
               Width           =   675
            End
            Begin VB.Label label1 
               Caption         =   " Mtrs:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   144
               Left            =   4065
               TabIndex        =   97
               Top             =   2250
               Width           =   555
            End
            Begin VB.Label label1 
               Caption         =   "Peces:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   158
               Left            =   6480
               TabIndex        =   92
               Top             =   2250
               Width           =   690
            End
         End
         Begin VB.Frame sol 
            Caption         =   "Soldadora"
            Height          =   3255
            Left            =   90
            TabIndex        =   159
            Top             =   18495
            Width           =   9885
            Begin VB.CommandButton Command9 
               BackColor       =   &H0025EFAD&
               Caption         =   "Accessoris Nous i mes"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Index           =   9
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   509
               TabStop         =   0   'False
               ToolTipText     =   "Imprimir bossa soldadores"
               Top             =   690
               Width           =   930
            End
            Begin VB.ComboBox Combo10 
               DataField       =   "migelaboratsol"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":C685
               Left            =   1275
               List            =   "comandes.frx":C695
               TabIndex        =   370
               Top             =   390
               WhatsThisHelpID =   5
               Width           =   645
            End
            Begin VB.CommandButton Command7 
               Height          =   285
               Left            =   8700
               Picture         =   "comandes.frx":C6A6
               Style           =   1  'Graphical
               TabIndex        =   163
               TabStop         =   0   'False
               Top             =   2895
               Width           =   285
            End
            Begin VB.TextBox Text111 
               DataField       =   "arxiusol"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   7380
               Locked          =   -1  'True
               MouseIcon       =   "comandes.frx":CA6C
               MousePointer    =   99  'Custom
               TabIndex        =   396
               TabStop         =   0   'False
               Top             =   2895
               Width           =   1290
            End
            Begin VB.ComboBox Combo14 
               DataField       =   "costatobertsol"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":CBBE
               Left            =   9120
               List            =   "comandes.frx":CBCE
               TabIndex        =   384
               Top             =   975
               Width           =   615
            End
            Begin VB.ComboBox Combo15 
               DataField       =   "simulteneitatsol"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":CBDE
               Left            =   7620
               List            =   "comandes.frx":CBF1
               TabIndex        =   382
               Top             =   960
               Width           =   675
            End
            Begin VB.ComboBox Combo11 
               DataField       =   "microperforatsol"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":CC04
               Left            =   8490
               List            =   "comandes.frx":CC0E
               TabIndex        =   383
               Top             =   975
               Width           =   585
            End
            Begin VB.TextBox Text113 
               DataField       =   "obssol1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   45
               TabIndex        =   391
               Top             =   2340
               Width           =   7305
            End
            Begin VB.TextBox Text17 
               DataField       =   "obssolgen1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   45
               Locked          =   -1  'True
               TabIndex        =   395
               Top             =   2880
               Width           =   7305
            End
            Begin VB.ComboBox Combo13 
               DataField       =   "microperforatsol"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":CC18
               Left            =   8250
               List            =   "comandes.frx":CC22
               TabIndex        =   160
               TabStop         =   0   'False
               Top             =   -6285
               Width           =   585
            End
            Begin MSMask.MaskEdBox Text134 
               DataField       =   "ansa"
               DataSource      =   "data1"
               Height          =   285
               Left            =   960
               TabIndex        =   386
               Top             =   1395
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text135 
               DataField       =   "troquel"
               DataSource      =   "data1"
               Height          =   285
               Left            =   960
               TabIndex        =   385
               Top             =   1110
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text133 
               DataField       =   "cinta"
               DataSource      =   "data1"
               Height          =   285
               Left            =   960
               TabIndex        =   387
               Top             =   1695
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text132 
               DataField       =   "unitatsxcaixa"
               DataSource      =   "data1"
               Height          =   285
               Left            =   7950
               TabIndex        =   393
               Top             =   2340
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text131 
               DataField       =   "unitatsxpaquet"
               DataSource      =   "data1"
               Height          =   285
               Left            =   7350
               TabIndex        =   392
               Top             =   2340
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text130 
               DataSource      =   "data1"
               Height          =   285
               Left            =   6600
               TabIndex        =   161
               TabStop         =   0   'False
               Top             =   1695
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   327681
               BorderStyle     =   0
               Appearance      =   0
               BackColor       =   14737632
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text129 
               DataField       =   "tipusoldadura"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6210
               TabIndex        =   388
               Top             =   1710
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text128 
               DataField       =   "unitatespsol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8415
               TabIndex        =   162
               Top             =   120
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text127 
               DataSource      =   "data1"
               Height          =   285
               Left            =   7815
               TabIndex        =   378
               Top             =   390
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               _Version        =   327681
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text126 
               DataField       =   "espessorsol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6990
               TabIndex        =   377
               Top             =   405
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               _Version        =   327681
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text125 
               DataField       =   "fuellebocasol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6165
               TabIndex        =   376
               Top             =   405
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text124 
               DataField       =   "fuellebasesol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   5355
               TabIndex        =   375
               Top             =   405
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text123 
               DataField       =   "solapasol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4545
               TabIndex        =   374
               Top             =   405
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text122 
               DataField       =   "longitudsol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3735
               TabIndex        =   373
               Top             =   405
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text121 
               DataField       =   "amplesol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   2100
               TabIndex        =   371
               Top             =   405
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text120 
               DataField       =   "soldadora"
               DataSource      =   "data1"
               Height          =   285
               Left            =   960
               TabIndex        =   380
               Top             =   810
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text119 
               DataField       =   "ampleplegsol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   2910
               TabIndex        =   372
               Top             =   405
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text118 
               DataField       =   "cantitatsol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8790
               TabIndex        =   379
               Top             =   390
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text117 
               DataField       =   "numtaladros"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8100
               TabIndex        =   389
               Top             =   1695
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text116 
               DataField       =   "diametremm"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8985
               TabIndex        =   390
               Top             =   1710
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text115 
               DataField       =   "tac"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6810
               TabIndex        =   381
               Top             =   975
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox19 
               DataField       =   "unitatsxsac"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8580
               TabIndex        =   394
               Top             =   2340
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   503
               _Version        =   327681
               PromptChar      =   "_"
            End
            Begin VB.Label label1 
               Caption         =   "Cinta"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   177
               Left            =   1695
               TabIndex        =   485
               Top             =   1725
               Width           =   4350
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   270
               Index           =   173
               Left            =   30
               TabIndex        =   474
               ToolTipText     =   "Numero bossa soldadores"
               Top             =   195
               Width           =   1170
            End
            Begin VB.Label label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF00FF&
               Caption         =   "MICRO/MACROPERFORAT"
               DataSource      =   "data1"
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Index           =   164
               Left            =   6555
               TabIndex        =   434
               Top             =   -15
               Width           =   3210
            End
            Begin VB.Label label1 
               Caption         =   "MicroP:"
               DataSource      =   "data1"
               Height          =   210
               Index           =   112
               Left            =   8490
               TabIndex        =   193
               Top             =   765
               Width           =   630
            End
            Begin VB.Label label1 
               Caption         =   "C. Obert:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   113
               Left            =   9105
               TabIndex        =   192
               Top             =   765
               Width           =   750
            End
            Begin VB.Label label1 
               Caption         =   "Arxiu Soldadora:"
               DataSource      =   "clients"
               Height          =   255
               Index           =   103
               Left            =   7500
               TabIndex        =   191
               Top             =   2685
               Width           =   1290
            End
            Begin VB.Label label1 
               Caption         =   "Observacions del Client"
               DataSource      =   "data1"
               Height          =   255
               Index           =   104
               Left            =   75
               TabIndex        =   190
               Top             =   2655
               Width           =   1950
            End
            Begin VB.Label label1 
               Caption         =   "Observacions Soldadora"
               DataSource      =   "data1"
               Height          =   210
               Index           =   105
               Left            =   75
               TabIndex        =   189
               Top             =   2100
               Width           =   2520
            End
            Begin VB.Label label1 
               Caption         =   "Un. Paquet/Caixa/Sacs"
               DataSource      =   "data1"
               Height          =   270
               Index           =   106
               Left            =   7230
               TabIndex        =   188
               Top             =   2100
               Width           =   1890
            End
            Begin VB.Label label1 
               Caption         =   "TAC:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   108
               Left            =   6915
               TabIndex        =   187
               Top             =   765
               Width           =   480
            End
            Begin VB.Label label1 
               Caption         =   "Diam. m/m"
               DataSource      =   "data1"
               Height          =   270
               Index           =   109
               Left            =   8910
               TabIndex        =   186
               Top             =   1485
               Width           =   825
            End
            Begin VB.Label label1 
               Caption         =   "Nº Taladros:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   110
               Left            =   7995
               TabIndex        =   185
               Top             =   1470
               Width           =   930
            End
            Begin VB.Label label1 
               Caption         =   "Quantitat:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   111
               Left            =   8820
               TabIndex        =   184
               Top             =   165
               Width           =   825
            End
            Begin VB.Label label1 
               Caption         =   "Plegat:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   114
               Left            =   3060
               TabIndex        =   183
               Top             =   180
               Width           =   630
            End
            Begin VB.Label label1 
               Caption         =   "Soldadora:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   115
               Left            =   120
               TabIndex        =   182
               Top             =   885
               Width           =   1035
            End
            Begin VB.Label nomsoldadora 
               Caption         =   "nomsoldadora"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   1680
               TabIndex        =   181
               Top             =   900
               Width           =   4500
            End
            Begin VB.Label label1 
               Caption         =   "Simultaneitat:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   116
               Left            =   7470
               TabIndex        =   180
               Top             =   765
               Width           =   975
            End
            Begin VB.Label label1 
               Caption         =   "Ample:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   117
               Left            =   2250
               TabIndex        =   179
               Top             =   165
               Width           =   555
            End
            Begin VB.Label label1 
               Caption         =   "B/L/F/BB:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   118
               Left            =   1185
               TabIndex        =   178
               Top             =   150
               Width           =   1005
            End
            Begin VB.Label label1 
               Caption         =   "Longitud:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   119
               Left            =   3780
               TabIndex        =   177
               Top             =   180
               Width           =   750
            End
            Begin VB.Label label1 
               Caption         =   "Solapa:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   120
               Left            =   4695
               TabIndex        =   176
               Top             =   180
               Width           =   630
            End
            Begin VB.Label label1 
               Caption         =   "Fuelle Ba:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   121
               Left            =   5400
               TabIndex        =   175
               Top             =   180
               Width           =   810
            End
            Begin VB.Label label1 
               Caption         =   "Fuelle Bo:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   122
               Left            =   6225
               TabIndex        =   174
               Top             =   180
               Width           =   750
            End
            Begin VB.Label label1 
               Caption         =   "Espessor:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   123
               Left            =   7050
               TabIndex        =   173
               Top             =   180
               Width           =   750
            End
            Begin VB.Label label1 
               Caption         =   "Mesura:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   124
               Left            =   7935
               TabIndex        =   172
               Top             =   165
               Width           =   690
            End
            Begin VB.Label label1 
               Caption         =   "Tipus Soldadura:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   125
               Left            =   6390
               TabIndex        =   171
               Top             =   1470
               Width           =   1605
            End
            Begin VB.Label label1 
               Caption         =   "Un. Caixa"
               DataSource      =   "data1"
               Height          =   270
               Index           =   126
               Left            =   8085
               TabIndex        =   170
               Top             =   1710
               Width           =   780
            End
            Begin VB.Label ansa 
               Caption         =   "ansa"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   1710
               TabIndex        =   169
               Top             =   1485
               Width           =   4500
            End
            Begin VB.Label truquel 
               Caption         =   "Truquel"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   1710
               TabIndex        =   168
               Top             =   1185
               Width           =   4500
            End
            Begin VB.Label label1 
               Caption         =   "Cinta:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   107
               Left            =   105
               TabIndex        =   167
               Top             =   1785
               Width           =   705
            End
            Begin VB.Label label1 
               Caption         =   "Ansa:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   127
               Left            =   105
               TabIndex        =   166
               Top             =   1470
               Width           =   705
            End
            Begin VB.Label label1 
               Caption         =   "Troquel:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   128
               Left            =   105
               TabIndex        =   165
               Top             =   1185
               Width           =   705
            End
            Begin VB.Label solpes 
               BackStyle       =   0  'Transparent
               Caption         =   "Pes Soldadora."
               ForeColor       =   &H00FF0000&
               Height          =   225
               Left            =   3990
               TabIndex        =   164
               Top             =   2130
               Width           =   2880
            End
         End
         Begin VB.Label label1 
            Caption         =   "Altres"
            DataSource      =   "data1"
            Height          =   600
            Index           =   133
            Left            =   8670
            TabIndex        =   486
            Top             =   210
            Width           =   1545
         End
         Begin VB.Label label1 
            DataSource      =   "data1"
            Height          =   360
            Index           =   174
            Left            =   2655
            TabIndex        =   476
            Top             =   1755
            Width           =   1215
         End
         Begin VB.Label label1 
            Caption         =   "--------"
            DataSource      =   "data1"
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
            Index           =   17
            Left            =   1455
            TabIndex        =   344
            Top             =   3525
            Width           =   7065
         End
      End
   End
End
Attribute VB_Name = "formcomandes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim velclientvolPVPimpostinclos As Boolean
Dim vnodemanarcontrasenyapassaraimpresores As Boolean
Dim tipusimpresio As String
Dim reservaassignacioocompra As String
Dim vprimeraentradacomandes As Boolean
Sub comprovarsijahihaalgunalbaraentrataexpedicions()
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select grupdeclient from clients where codi=" + atrim(Data1.Recordset!client))
   If rst.EOF Then Exit Sub
   If atrim(rst!grupdeclient) = "ARDO" Then Exit Sub
   Set rst = dbtmp.OpenRecordset("select * from liniesalbara where lotinplacsa=" + Text1)
   If Not rst.EOF Then MsgBox "ATENCIÓ, AQUESTA COMANDA JA TÉ ALBARANS DE EXPEDICONS FETS." + vbNewLine + "S'HA DE BORRAR I TORNAR A FER L'ALBARÀ D'EXPEDICIONS PER CORREGIR EL PROBLEMA.", vbCritical, "ATENCIÓ"
   Set rst = Nothing
End Sub
Function recorregutcontrolspelvalor(v As String) As String
   Dim objecte As Object
   For Each objecte In formcomandes
      If TypeOf objecte Is MaskEdBox Or TypeOf objecte Is TextBox Or TypeOf objecte Is ComboBox Then
            If objecte.DataField = v Then recorregutcontrolspelvalor = objecte.Text: Exit Function
      End If
   Next objecte
End Function
Sub comandesamblamateixareferenciainplacsaIMPRESES(rst As Recordset)
   Dim rstc As Recordset
   Dim rstp As Recordset
   Dim vruta As String
   Dim camps As String
   Dim nomcamp As String
   Dim valorcamp As String
   Dim totsiguals As Boolean
   Dim comandesiguals As String
   Dim vreferencianova As Double
   Dim rstmat1 As Recordset
   Dim rstmat2 As Recordset
   Set rstp = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
   If rstp.EOF Then Exit Sub
   vruta = atrim(rstp!ruta)
   If InStr(1, vruta, "I") = 0 Then Exit Sub
   Set rstc = dbtmp.OpenRecordset("select * from comandes where numtreball=" + atrim(rst!numtreball) + " and tubolam='" + atrim(rst!tubolam) + "' and ampleesq=" + passaradecimalpunt(cadbl(rst!ampleesq)) + " and client=" + atrim(cadbl(rst!client)))
   If rstc.EOF Then Exit Sub
   Set rstmat1 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rst!materialex)))
   If rstmat1.EOF Then Exit Sub
   While Not rstc.EOF
      'he tret #Ltipusadhesiu
    camps = "#Etubolam#Eampleesq#Eplegatesq#Esolapa#Eespessor#Emicropex#Eoberturaex#Ematerialex#Inumtreball#Lampleutil#Lsimulteneitatlam#Rmigelaborat#Ramplereb#Rsimulteneitatreb"
    camps = camps + "#Smigelaboratsol#Samplesol#Sampleplegsol#Slongitudsol#Ssolapasol#Sfuellebasesol#Sfuellebocasol#Sespessorsol#Stroquel#Sansa#Scinta#"
    nomcamp = proximcamp(camps)
    totsiguals = True
    While nomcamp <> ""
      If InStr(1, vruta, Mid(nomcamp, 1, 1)) > 0 Then
       nomcamp = Mid(nomcamp, 2)
       valordelcamp = cvavalorcamp(rstc, nomcamp)
       If nomcamp = "numtreball" Then GoTo cont
       If nomcamp = "materialex" Then
           Set rstmat2 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstc!materialex)))
           If rstmat2.EOF Then GoTo proxim
           If rstmat1!familia <> rstmat2!familia Or rstmat1!subfamilia <> rstmat2!subfamilia Or rstmat1!familiacol <> rstmat2!familiacol Then
             totsiguals = False: GoTo proxim
           End If
           GoTo cont
       End If
       If atrim(valordelcamp) <> cvavalorcamp(rst, nomcamp) Then totsiguals = False: GoTo proxim
      End If
cont:
      nomcamp = proximcamp(camps)
    Wend
proxim:
   If totsiguals Then comandesiguals = comandesiguals + IIf(comandesiguals <> "", " ,", "") + atrim(rstc!comanda)
   'Me.Caption = atrim(rst!comanda)
   rstc.MoveNext
   Wend
   If comandesiguals = "" Then Exit Sub
   rstc.MoveFirst
   Set rstp = dbtmp.OpenRecordset("select * from comandes_extres where refinplacsa<>'' and refinplacsa<>null and comanda in (" + atrim(comandesiguals) + ")")
   If rstp.EOF Then
      vreferencianova = cadbl("9" + atrim(rstc!client) + atrim(proximnumeroreferencia))
        Else: vreferencianova = rstp!refinplacsa
   End If
   Set rstp = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda in (" + atrim(comandesiguals) + ")")
   While Not rstp.EOF
        dbtmp.Execute "update comandes_extres set refinplacsa='" + atrim(vreferencianova) + "' where comanda in (" + atrim(rstp!comanda) + ", " + atrim(rstp!linkcomanda1) + ", " + atrim(rstp!linkcomanda2) + ")"
        rstp.MoveNext
   Wend
   
End Sub
Function cvavalorcamp(rst As Recordset, nomcamp As String) As String
   Select Case nomcamp
      Case "ampleesq"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "plegatesq"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "solapa"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "espessor"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "micropex"
           cvavalorcamp = atrim((rst.Fields(nomcamp)))
           If cvavalorcamp = "" Then cvavalorcamp = "N"
      Case "oberturaex"
           cvavalorcamp = atrim((rst.Fields(nomcamp)))
           If cvavalorcamp = "" Then cvavalorcamp = "N"
      Case "ampleutil"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "simulteneitatlam"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "tipusadhesiu"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "migelaborat"
           cvavalorcamp = atrim(rst.Fields(nomcamp))
      Case "amplereb"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "simulteneitatreb"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
            Case "migelaboratsol"
           cvavalorcamp = atrim(rst.Fields(nomcamp))
      Case "amplesol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "ampleplegsol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "longitudsol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "solapasol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "fuellebasesol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "fuellebocasol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "troquel"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "ansa"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "cinta"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
       Case Else
           cvavalorcamp = atrim(rst.Fields(nomcamp))
   End Select
End Function
Function NUMERODEreferenciainplacsa(rst As Recordset, esimpresa As Boolean) As String
   Dim rstc As Recordset
   Dim rstp As Recordset
   Dim vruta As String
   Dim camps As String
   Dim nomcamp As String
   Dim valorcamp As String
   Dim totsiguals As Boolean
   Dim comandesiguals As String
   Dim vreferencianova As String
   Dim rstmat1 As Recordset
   Dim rstmat2 As Recordset
   If atrim(rst!producte) = "PC" Or atrim(rst!producte) = "PC2" Or atrim(rst!producte) = "PCP" Then Exit Function
   Set rstp = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
   If rstp.EOF Then Exit Function
   vruta = atrim(rstp!ruta)
   If esimpresa Then
        If InStr(1, vruta, "I") = 0 Then Exit Function
        Set rstc = dbtmp.OpenRecordset("SELECT comandesmesextres.*, InStr(1,[ruta],'I') AS Expr1 FROM comandesmesextres WHERE (((InStr(1,[ruta],'I'))>0) and producte='" + atrim(rst!producte) + "' and numtreball = " + atrim(cadbl(rst!numtreball)) + " And client = " + atrim(cadbl(rst!client)) + " AND mid(refinplacsa,1,2)<>'PR' and mid(refinplacsa,1,2)<>'FP') and comanda<>" + atrim(rst!comanda) + " order by comanda asc")
          Else:
            If InStr(1, vruta, "I") <> 0 Then Exit Function
            Set rstc = dbtmp.OpenRecordset("SELECT comandesmesextres.*, InStr(1,[ruta],'I') AS Expr1 FROM comandesmesextres WHERE (((InStr(1,[ruta],'I'))=0) and producte='" + atrim(rst!producte) + "' and client = " + atrim(cadbl(rst!client)) + " and tubolam='" + atrim(rst!tubolam) + "' and ampleesq=" + passaradecimalpunt(cadbl(rst!ampleesq)) + " AND mid(refinplacsa,1,2)<>'PR' and mid(refinplacsa,1,2)<>'FP') and comanda<>" + atrim(rst!comanda) + " order by comanda asc")
            'MsgBox "SELECT comandes.*, InStr(1,[ruta],'I') AS Expr1 FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi WHERE (((InStr(1,[ruta],'I'))=0) and client = " + atrim(cadbl(rst!client)) + " and tubolam='" + atrim(rst!tubolam) + "' and ampleesq=" + passaradecimalpunt(cadbl(rst!ampleesq)) + ")"
            'Set rstc = dbtmp.OpenRecordset("select * from comandes where tubolam='" + atrim(rst!tubolam) + "' and ampleesq=" + passaradecimalpunt(cadbl(rst!ampleesq)) + " and client=" + atrim(cadbl(rst!client)))
   End If
     
   If rstc.EOF Then Exit Function
   Set rstmat1 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rst!materialex)))
   If rstmat1.EOF Then Exit Function
   While Not rstc.EOF
     'he tret #Ltipusadhesiu
    camps = "#Etubolam#Eampleesq#Eplegatesq#Esolapa#Eespessor#Emicropex#Eoberturaex#Ematerialex#Inumtreball#Lampleutil#Lsimulteneitatlam#Rmigelaborat#Ramplereb#Rsimulteneitatreb"
    camps = camps + "#Smigelaboratsol#Samplesol#Sampleplegsol#Slongitudsol#Ssolapasol#Sfuellebasesol#Sfuellebocasol#Stroquel#Sansa#Scinta#"
    nomcamp = proximcamp(camps)
    totsiguals = True
'    MsgBox rstc!comanda
    While nomcamp <> ""
      If InStr(1, vruta, Mid(nomcamp, 1, 1)) > 0 Then
       nomcamp = Mid(nomcamp, 2)
       valordelcamp = cvavalorcamp(rstc, nomcamp)
       If nomcamp = "numtreball" Then GoTo cont
       If nomcamp = "materialex" Then
           Set rstmat2 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstc!materialex)))
           If rstmat2.EOF Then GoTo proxim
           If cadbl(rstmat1!familia) <> cadbl(rstmat2!familia) Or cadbl(rstmat1!subfamilia) <> cadbl(rstmat2!subfamilia) Or cadbl(rstmat1!familiacol) <> cadbl(rstmat2!familiacol) Then
             totsiguals = False: GoTo proxim
           End If
           GoTo cont
       End If
'       MsgBox nomcamp + "   =    " + atrim(valordelcamp) + "  ------  " + cvavalorcamp(rst, nomcamp)
       If atrim(valordelcamp) <> cvavalorcamp(rst, nomcamp) Then
         totsiguals = False: GoTo proxim
       End If
      End If
cont:
      nomcamp = proximcamp(camps)
    Wend
proxim:
   If totsiguals Then
      If complexesiguals(rst, rstc) Then comandesiguals = comandesiguals + IIf(comandesiguals <> "", " ,", "") + atrim(rstc!comanda)
   End If
   'Me.Caption = atrim(rst!comanda)
   rstc.MoveNext
   Wend
   If comandesiguals = "" Then NUMERODEreferenciainplacsa = rst!refinplacsa: Exit Function
   rstc.MoveFirst
   If InStr(1, comandesiguals, ",") = 0 And InStr(1, comandesiguals, atrim(rst!comanda)) > 0 And esimpresa Then     'que hi ha nomes el registre actiu igual
         Set rstp = dbtmp.OpenRecordset("select * from comandes_extres where comanda in (" + atrim(comandesiguals) + ") order by refinplacsa DESC")
         If Not rstp.EOF Then
            If InStr(1, atrim(rstp!refinplacsa), "C") > 0 Then
                  vreferencianova = atrim(rstp!refinplacsa)
                    Else: vreferencianova = novareferenciaimpresa(rst)
            End If
         End If
         GoTo gravar:
   End If
   Set rstp = dbtmp.OpenRecordset("select * from comandes_extres where (refinplacsa<>'' or refinplacsa<>null) and comanda in (" + atrim(comandesiguals) + ") order by refinplacsa DESC")
   If Not rstp.EOF Then rstp.FindFirst "refinplacsa like '*C*'"
   If Not esimpresa Then
        If rstp.EOF Or rstp.NoMatch Then
           vreferencianova = ""
             Else: vreferencianova = rstp!refinplacsa
        End If
          Else
            If rstp.EOF Or rstp.NoMatch Then
                vreferencianova = ""
                    Else:
                       vreferencianova = rstp!refinplacsa
            End If
          
   End If
gravar:
  NUMERODEreferenciainplacsa = vreferencianova
   
End Function
Sub comandesamblamateixareferenciainplacsa(rst As Recordset, esimpresa As Boolean)
   Dim rstc As Recordset
   Dim rstp As Recordset
   Dim vruta As String
   Dim camps As String
   Dim nomcamp As String
   Dim valorcamp As String
   Dim totsiguals As Boolean
   Dim comandesiguals As String
   Dim vreferencianova As String
   Dim vnoPR_RP_FP As String
   Dim rstmat1 As Recordset
   Dim rstmat2 As Recordset
   Dim vrefnova As Boolean
   Dim vreciclat As Boolean
   Dim vprefix As String
   
   vrefnova = False
   If atrim(rst!producte) = "PC" Or atrim(rst!producte) = "PC2" Or atrim(rst!producte) = "PCP" Then Exit Sub
   If cadbl(rst!materialex) < 501 Then MsgBox "El material seleccionat es inferior al codi 500 no es pot generar la REFERENCIA D'INPLACSA.", vbCritical, "Error": Exit Sub
   Set rstp = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
   If rstp.EOF Then Exit Sub
   vruta = atrim(rstp!ruta)
   vnoPR_RP_FP = "(mid(refinplacsa,1,2)&'   ')<>'PR' and (mid(refinplacsa,1,2)&'   ')<>'RP' and (mid(refinplacsa,1,2)&'   ')<>'FP'"
   If esimpresa Then
        If InStr(1, vruta, "I") = 0 Then Exit Sub
        'Set rstc = dbtmp.OpenRecordset("SELECT comandes.*, InStr(1,[ruta],'I') AS Expr1 FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi WHERE (((InStr(1,[ruta],'I'))>0) and producte='" + atrim(rst!producte) + "' and numtreball = " + atrim(cadbl(rst!numtreball)) + " And client = " + atrim(cadbl(rst!client)) + ")")
        Set rstc = dbtmp.OpenRecordset("SELECT comandesmesextres.*, InStr(1,[ruta],'I') AS Expr1 FROM comandesmesextres WHERE (" + vnoPR_RP_FP + ") and (((InStr(1,[ruta],'I'))>0) and producte='" + atrim(rst!producte) + "' and numtreball = " + atrim(cadbl(rst!numtreball)) + " And client = " + atrim(cadbl(rst!client)) + ")")
          Else:
            If InStr(1, vruta, "I") <> 0 Then Exit Sub
            'Set rstc = dbtmp.OpenRecordset("SELECT comandes.*, InStr(1,[ruta],'I') AS Expr1 FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi WHERE (((InStr(1,[ruta],'I'))=0) and producte='" + atrim(rst!producte) + "' and client = " + atrim(cadbl(rst!client)) + " and tubolam='" + atrim(rst!tubolam) + "' and ampleesq=" + passaradecimalpunt(cadbl(rst!ampleesq)) + ")")
            Set rstc = dbtmp.OpenRecordset("SELECT comandesmesextres.*, InStr(1,[ruta],'I') AS Expr1 FROM comandesmesextres WHERE (" + vnoPR_RP_FP + ") and (((InStr(1,[ruta],'I'))=0) and producte='" + atrim(rst!producte) + "' and client = " + atrim(cadbl(rst!client)) + " and tubolam='" + atrim(rst!tubolam) + "' and ampleesq=" + passaradecimalpunt(cadbl(rst!ampleesq)) + ")")
   End If
   '  Clipboard.Clear
   '  Clipboard.SetText "SELECT comandesmesextres.*, InStr(1,[ruta],'I') AS Expr1 FROM comandesmesextres WHERE (" + vnoPR_RP_FP + ") and (((InStr(1,[ruta],'I'))>0) and producte='" + atrim(rst!producte) + "' and numtreball = " + atrim(cadbl(rst!numtreball)) + " And client = " + atrim(cadbl(rst!client)) + ")"
     
   If rstc.EOF Then Exit Sub
   Set rstmat1 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rst!materialex)))
   If rstmat1.EOF Then Exit Sub
   vreciclat = False
   esmaterialreciclat rstmat1, vreciclat
   While Not rstc.EOF
     'he tret #Ltipusadhesiu
    camps = "#Etubolam#Eampleesq#Eplegatesq#Esolapa#Eespessor#Emicropex#Eoberturaex#Ematerialex#Inumtreball#Lampleutil#Lsimulteneitatlam#Rmigelaborat#Ramplereb#Rsimulteneitatreb"
    camps = camps + "#Smigelaboratsol#Samplesol#Sampleplegsol#Slongitudsol#Ssolapasol#Sfuellebasesol#Sfuellebocasol#Stroquel#Sansa#Scinta#"
    nomcamp = proximcamp(camps)
    totsiguals = True
    While nomcamp <> ""
      If InStr(1, vruta, Mid(nomcamp, 1, 1)) > 0 Then
       nomcamp = Mid(nomcamp, 2)
       valordelcamp = cvavalorcamp(rstc, nomcamp)
       If nomcamp = "numtreball" Then GoTo cont
       If nomcamp = "materialex" Then
           Set rstmat2 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstc!materialex)))
           If rstmat2.EOF Then GoTo proxim
           esmaterialreciclat rstmat2, vreciclat
           If cadbl(rstmat1!familia) <> cadbl(rstmat2!familia) Or cadbl(rstmat1!subfamilia) <> cadbl(rstmat2!subfamilia) Or cadbl(rstmat1!familiacol) <> cadbl(rstmat2!familiacol) Then
             totsiguals = False: GoTo proxim
           End If
           GoTo cont
       End If
'       MsgBox nomcamp + "   =    " + atrim(valordelcamp) + "  ------  " + cvavalorcamp(rst, nomcamp)
       If atrim(valordelcamp) <> cvavalorcamp(rst, nomcamp) Then
         totsiguals = False: GoTo proxim
       End If
      End If
cont:
      nomcamp = proximcamp(camps)
    Wend
proxim:
   If totsiguals And complexesiguals(rst, rstc) Then comandesiguals = comandesiguals + IIf(comandesiguals <> "", " ,", "") + atrim(rstc!comanda)
   'Me.Caption = atrim(rst!comanda)
   rstc.MoveNext
   Wend
   If comandesiguals = "" Then Exit Sub
   rstc.MoveFirst
   If InStr(1, comandesiguals, ",") = 0 And InStr(1, comandesiguals, atrim(rst!comanda)) > 0 And esimpresa Then     'que hi ha nomes el registre actiu igual
         Set rstp = dbtmp.OpenRecordset("select * from comandes_extres where comanda in (" + atrim(comandesiguals) + ") order by refinplacsa DESC")
         If Not rstp.EOF Then
            If InStr(1, atrim(rstp!refinplacsa), "C") > 0 Then
                  vreferencianova = atrim(rstp!refinplacsa)
                  vrefnova = False
                    Else: vreferencianova = novareferenciaimpresa(rst): vrefnova = True
            End If
         End If
         GoTo gravar:
   End If
   Set rstp = dbtmp.OpenRecordset("select * from comandes_extres where (refinplacsa<>'' or refinplacsa<>null) and comanda in (" + atrim(comandesiguals) + ") order by refinplacsa DESC")
   If Not rstp.EOF Then rstp.FindFirst "refinplacsa like '*C*'"
   If Not esimpresa Then
        If rstp.EOF Or rstp.NoMatch Then
           vreferencianova = "C" + atrim(rstc!client) + "A" + atrim(proximnumeroreferencia)
           vrefnova = True
             Else: vreferencianova = rstp!refinplacsa: vrefnova = False
        End If
          Else
            If rstp.EOF Or rstp.NoMatch Then
                vreferencianova = novareferenciaimpresa(rst) 'cadbl("801" + atrim(rstc!client) + atrim(rstc!numtreball))
                vrefnova = True
                    Else:
                       vreferencianova = rstp!refinplacsa
                       vrefnova = False
            End If
          
   End If
gravar:
   If atrim(vreferencianova) <> "" Then
        vprefix = ""
        If InStr(1, " FP RP PR ", Mid(vreferencianova, 1, 2)) = 0 And atrim(rst!impressio) <> "R" Then vprefix = DemanarPrefixRefInplacsa
        vreferencianova = vprefix + vreferencianova
        If vreciclat Then vreferencianova = vreferencianova + "_R"
        Set rstp = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2,client from comandes where comanda  in (" + atrim(comandesiguals) + ")")
        While Not rstp.EOF
             dbtmp.Execute "update comandes_extres set refinplacsa='" + atrim(vreferencianova) + "' where comandaimpresa=true and comanda in (" + atrim(rstp!comanda) + ", " + atrim(rstp!linkcomanda1) + ", " + atrim(rstp!linkcomanda2) + ")"
             rstp.MoveNext
        Wend
        If vrefnova Then
             dbtmp.Execute "update comandes_extres SET refinplacsa_nova=True, refinplacsa_validada=false where refinplacsa='" + atrim(vreferencianova) + "'"
              Else: dbtmp.Execute "update comandes_extres SET refinplacsa_nova=FALSE, refinplacsa_validada=TRUE where refinplacsa='" + atrim(vreferencianova) + "'"
             ' Else: dbtmp.Execute "update comandes_extres SET refinplacsa_validada=true where comanda=" + atrim(rst!comanda) + " or comanda=" + atrim(rst!linkcomanda1) + " or comanda=" + atrim(rst!linkcomanda2)
        End If
   End If
   
End Sub
Function DemanarPrefixRefInplacsa() As String
  ratoli "normal"
  Unload formseleccio
  Load formseleccio
  'formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "SELECT TOP 1 'NORMAL' AS Opcio From comandes Union All SELECT TOP 1 'PROVES' From comandes Union All SELECT TOP 1 'REPROCESATS' FROM comandes Union All SELECT TOP 1 'FINGER PRINTS' FROM comandes;"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 4000
  formseleccio.DBGrid2.Font.Name = "Arial"
  formseleccio.DBGrid2.Font.Size = 20
  formseleccio.Width = 5000
  formseleccio.Height = 3600
  formseleccio.Caption = "Escull Opció"
  formseleccio.Show 1
  If seleccioret = 1 Then
      If formseleccio.Data1.Recordset!Opcio = "NORMAL" Then DemanarPrefixRefInplacsa = ""
      If formseleccio.Data1.Recordset!Opcio = "PROVES" Then DemanarPrefixRefInplacsa = "PR"
      If formseleccio.Data1.Recordset!Opcio = "REPROCESATS" Then DemanarPrefixRefInplacsa = "RP"
      If formseleccio.Data1.Recordset!Opcio = "FINGER PRINTS" Then DemanarPrefixRefInplacsa = "FP"
  End If
  Unload formseleccio
  
End Function
Sub esmaterialreciclat(rstmat As Recordset, vreciclat As Boolean)
  If vreciclat Then Exit Sub
  If cadbl(rstmat!tanpercentimpostenvasos) < 100 And cadbl(rstmat!tanpercentimpostenvasos) > 0 Then
                 If Not familiapaper(rstmat!familia) Then vreciclat = True
  End If
End Sub
Function familiapaper(vcodifamilia As Double) As Boolean
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(vcodifamilia))
  If InStr(1, rst!descripcio, "PAPER ") > 0 Then familiapaper = True
  Set rst = Nothing
End Function
Function novareferenciaimpresa(rst As Recordset, Optional num As Byte) As String
    Dim rst2 As Recordset
    Dim vrefinplacsa As String
    Dim elgran As Byte
    elgran = 0
    If num > 0 Then elgran = num - 1: GoTo noref
    'Clipboard.Clear
    'Clipboard.SetText "select * from comandes_extres where refinplacsa<>null and comanda in (select comanda from comandes where comanda<>" + atrim(rst!comanda) + " and client=" + atrim(rst!client) + " and numtreball=" + atrim(cadbl(rst!numtreball)) + ") order by refinplacsa asc"
    
    Set rst2 = dbtmp.OpenRecordset("select * from comandes_extres where refinplacsa<>null and comanda in (select comanda from comandes where comanda<>" + atrim(rst!comanda) + " and client=" + atrim(rst!client) + " and numtreball=" + atrim(cadbl(rst!numtreball)) + ") order by refinplacsa asc")
    If rst2.EOF Then
       GoTo noref
         Else: vrefinplacsa = atrim(rst2!refinplacsa)
    End If
    While Not rst2.EOF
       If cadbl(Mid(rst2!refinplacsa, 1, 2)) > elgran Then
           elgran = cadbl(Mid(rst2!refinplacsa, 1, 2))
           vrefinplacsa = atrim(rst2!refinplacsa)
       End If
       rst2.MoveNext
    Wend
    'vrefinplacsa = atrim(rst2!refinplacsa)
noref:
    'If Len(vrefinplacsa) < 4 Then vrefinplacsa = ""
    'If rst2.RecordCount = 1 Or vrefinplacsa = "" Then
    novareferenciaimpresa = Format(elgran + 1, "00") + "C" + atrim(cadbl(rst!client)) + "I" + atrim(rst!numtreball)
    '   Else: novareferenciaimpresa = Format(cadbl(Mid(vrefinplacsa, 1, 2)) + 1, "00") + "C" + atrim(cadbl(rst!client)) + "I" + atrim(rst!numtreball)
    'End If
End Function
Function complexesiguals(rst As Recordset, rstc As Recordset) As Boolean
   Dim rstcpc As Recordset
   Dim rstpc As Recordset
   Dim rstmat1 As Recordset
   Dim rstmat2 As Recordset
   Dim totsiguals As Boolean
   Dim camps As String
   Dim nomcamp As String
   
   If cadbl(rst!linkcomanda1) = 0 And cadbl(rstc!linkcomanda1) = 0 Then totsiguals = True: GoTo fi
   Set rstcpc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rstc!linkcomanda1))
   Set rstpc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rst!linkcomanda1))
inici:
   Set rstmat1 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstpc!materialex)))
   If rstmat1.EOF Then totsiguals = False: GoTo fi
   If Not rstcpc.EOF Then
    camps = "#Etubolam#Eampleesq#Eplegatesq#Esolapa#Eespessor#Emicropex#Eoberturaex#Ematerialex#"
    nomcamp = proximcamp(camps)
    totsiguals = True
     While nomcamp <> ""
       nomcamp = Mid(nomcamp, 2)
       valordelcamp = cvavalorcamp(rstcpc, nomcamp)
       If nomcamp = "materialex" Then
           Set rstmat2 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstcpc!materialex)))
           If rstmat2.EOF Then GoTo proxim
           If cadbl(rstmat1!familia) <> cadbl(rstmat2!familia) Or cadbl(rstmat1!subfamilia) <> cadbl(rstmat2!subfamilia) Or cadbl(rstmat1!familiacol) <> cadbl(rstmat2!familiacol) Then
             totsiguals = False: GoTo proxim
           End If
           GoTo cont
       End If
'       MsgBox nomcamp + "   =    " + atrim(valordelcamp) + "  ------  " + cvavalorcamp(rst, nomcamp)
       If atrim(valordelcamp) <> cvavalorcamp(rstpc, nomcamp) Then
         totsiguals = False: GoTo proxim
       End If
cont:
      nomcamp = proximcamp(camps)
    Wend
proxim:
   End If
   If Not totsiguals Then GoTo fi
   If cadbl(rstpc!comanda) = cadbl(rst!linkcomanda2) Then GoTo fi
   If cadbl(rst!linkcomanda2) = 0 And cadbl(rstc!linkcomanda2) = 0 Then totsiguals = True: GoTo fi
   Set rstcpc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rstc!linkcomanda2))
   Set rstpc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rst!linkcomanda2))
   GoTo inici
   
fi:
   complexesiguals = totsiguals
   Set rstcpc = Nothing
   Set rstpc = Nothing
   Set rstmat1 = Nothing
   Set rstmat2 = Nothing
End Function
Function proximnumeroreferencia() As Long
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select numreferenciainplacsa from valorsgenerals")
   If rst.EOF Then
      proximnumeroreferencia = 1
        Else: proximnumeroreferencia = rst!numreferenciainplacsa + 1
   End If
   dbtmp.Execute "update valorsgenerals set numreferenciainplacsa=numreferenciainplacsa+1"
End Function
Function compleixelmateixcodidarticle(ruta As String) As Boolean
   Dim camps As String
   Dim nomcamp As String
   Dim valordelcamp As String
     'he tret #Ltipusadhesiu
   camps = "#Etubolam#Eampleesq#Eplegatesq#Esolapa#Eespessor#Emicropex#Eoberturaex#Ematerialex#Inumtreball#Lampleutil#Lsimulteneitatlam#Rmigelaborat#Ramplereb#Rsimulteneitatreb"
   camps = camps + "#Smigelaboratsol#Samplesol#Sampleplegsol#Slongitudsol#Ssolapasol#Sfuellebasesol#Sfuellebocasol#Sespessorsol#Stroquel#Sansa#Scinta#"
   nomcamp = proximcamp(camps)
   While nomcamp <> ""
      
      If InStr(1, ruta, Mid(nomcamp, 1, 1)) > 0 Then
        nomcamp = Mid(nomcamp, 2)
        valordelcamp = recorregutcontrolspelvalor(nomcamp)
        If nomcamp = "numtreball" Then valordelcamp = atrim(Data1.Recordset.Fields(nomcamp))
        If atrim(valordelcamp) <> atrim(Data1.Recordset.Fields(nomcamp)) Then MsgBox atrim(valordelcamp) + "  -  " + atrim(Data1.Recordset.Fields(nomcamp))
      End If
      nomcamp = proximcamp(camps)
   Wend
End Function
Function proximcamp(camps As String) As String
   If camps = "#" Then Exit Function
   proximcamp = Mid(camps, 2, InStr(2, camps, "#") - 2)
   camps = Mid(camps, InStr(2, camps, "#"))
End Function
Function posicioenlaruta(numc As Double, seccioactual As String, laruta As String) As String
  Dim rstp As Recordset
  'If InStr(1, "VPT", seccioactual) = 0 Then Exit Function
 
  Set rstp = dbbaixes.OpenRecordset("SELECT comandes.comanda, rebobinadorestot.acavada as acavadar, laminadorestot.acavada as acavadal, impressorestot.acavada as acavadai FROM ((comandes LEFT JOIN rebobinadorestot ON comandes.comanda = rebobinadorestot.comanda) LEFT JOIN laminadorestot ON comandes.comanda = laminadorestot.comanda) LEFT JOIN impressorestot ON comandes.comanda = impressorestot.comanda WHERE (((comandes.comanda)=" + atrim(numc) + "));")
  
  If Not rstp.EOF Then
     If InStr(1, laruta, "R") > 0 And cadblnull_1(rstp!acavadar) = 0 Then posicioenlaruta = "R"
     If InStr(1, laruta, "L") > 0 And cadblnull_1(rstp!acavadal) = 0 Then posicioenlaruta = "L"
     If InStr(1, laruta, "I") > 0 And cadblnull_1(rstp!acavadai) = 0 Then posicioenlaruta = "I"
  End If
  
  Set rstp = Nothing
End Function


Function cadblnull_1(acabada As Variant) As Double
   If IsNull(acabada) Then cadblnull_1 = -1: Exit Function
   cadblnull_1 = cadbl(acabada)
End Function

Private Sub camp1_Change()

End Sub

Private Sub adhesiu_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim vsql As String
 If KeyCode = 113 Then
   vsql = "SELECT adhesius.codi, adhesius.predeterminada, adhesius.color, adhesius.resina, adhesius.enduridor, familiescoles.descripcio AS descfamcola, subfamiliescoles.descripcio AS descsubfamcola FROM (adhesius LEFT JOIN familiescoles ON adhesius.idfamilia = familiescoles.codi) LEFT JOIN subfamiliescoles ON adhesius.idsubfamilia = subfamiliescoles.codi"
   triaralgu "Triar Adhesiu", vsql, vadhesiu, adhesiu, "resina"
 End If
  possar_noms_adhesius False
End Sub
Sub possar_noms_adhesius(Optional lookup As Boolean)
 Dim vsql As String
 vsql = "SELECT adhesius.*, familiescoles.descripcio AS descfamcola, subfamiliescoles.descripcio AS descsubfamcola FROM (adhesius LEFT JOIN familiescoles ON adhesius.idfamilia = familiescoles.codi) LEFT JOIN subfamiliescoles ON adhesius.idsubfamilia = subfamiliescoles.codi "
 Set rsttmp = dbtmp.OpenRecordset(vsql + " where adhesius.codi=" + atrim(cadbl(vadhesiu)))
 If Not rsttmp.EOF Then
    enduridor(0) = atrim(rsttmp!enduridor)
    enduridor(1) = IIf(atrim(rsttmp!descfamcola) <> "", atrim(rsttmp!descsubfamcola) + " - " + atrim(rsttmp!descfamcola), "")
    adhesiu = atrim(rsttmp!resina)
    grcm1(0) = cadbl(rsttmp!grmcm3_resina)
    grcm2(0) = cadbl(rsttmp!grmcm3_enduridor)
    ºC1(0) = cadbl(rsttmp!grausresina)
    ºC2(0) = cadbl(rsttmp!grausenduridor)
    If Not lookup Then
      pes1 = cadbl(rsttmp![%resina])
      pes2 = cadbl(rsttmp![%enduridor])
      If cadbl(grmt2) = 0 Then grmt2 = cadbl(rsttmp!aportcola)
    End If
    adhesiu.BackColor = possarcoloradhesiu(IIf(atrim(rsttmp!predeterminada) = "S", "BLANC", atrim(rsttmp!color)))
 End If
End Sub
Function possarcoloradhesiu(color As String) As String
  Dim codicolor As Double
  codicolor = QBColor(15)
  Select Case color
    Case "VERD"
       codicolor = QBColor(10)
    Case "TARONJA"
       codicolor = &H62B1F2
    Case "BLAU"
       codicolor = QBColor(9)
    Case "ROSA"
       codicolor = &HC78DFA
    Case "GROC"
       codicolor = QBColor(6)
    Case "VERMELL"
       codicolor = QBColor(12)
    Case "BLANC"
       codicolor = QBColor(15)
    Case Else
          codicolor = QBColor(15)
  End Select
  possarcoloradhesiu = codicolor
  fondocoloradhesiu.FillColor = codicolor
End Function
Private Sub alta_Click()
If comprovar_risc_comanda = 1 Then Exit Sub
If MsgBox("Segur que vols crear una comanda?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
alta_registre
Command23_Click
End Sub

Private Sub form1_AccessKeyPress(tecla As String)
  MsgBox tecla
End Sub

Private Sub arrastrar_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  If Data.GetFormat(vbCFFiles) Then
    Me.Caption = Data.Files(1)
      Else: Me.Caption = "E-Mail"
  End If
End Sub

Private Sub areadatos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim vvalor As String
  If Shift = 2 And Button = 2 Then
    vvalor = valordelcontrolalaposicioXY(x, y)
    If vvalor = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText vvalor
    'MsgBox vvalor + vbNewLine + "COPIAT AL PORTAPAPERS", vbInformation, "Valor copiat"
  End If
End Sub

Function valordelcontrolalaposicioXY(x As Single, y As Single) As String
  Dim Control As Control
  For Each Control In formcomandes
   If TypeOf Control Is MaskEdBox Or TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
      If Control.Visible Then
         If x > (Control.Left + 100) And x < (Control.Left + 100 + Control.Width) Then
             If y > (Control.Container.Top + Control.Top) And y < (Control.Container.Top + (Control.Top + Control.Height) + 100) Then
                 valordelcontrolalaposicioXY = Control.Text: Exit Function
             End If
         End If
      End If
   End If
  Next
End Function

Private Sub areadatos_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim ruta_documentacio_pressupostos As String
 Dim vnumc As Double
 Dim vclient As Double
 Dim vpdf As String
  'If Data.GetFormat(vbCFFiles) Then
  '  Me.Caption = Data.Files(1)
  '    Else:
  '      If State = 0 Then obrircarpetadragdrop
  'End If
  'If State = 0 Then obrircarpetadragdrop
  vpdf = Data.Files(1)
  If InStr(1, vpdf, ".pdf") Then
      vnumc = Data1.Recordset!comanda
      vclient = Data1.Recordset!client
      vnompressupost = InputBox("Entra el nom del pressupost que vols relacionar." + Chr(10) + "FES CANCELAR SI NO VOLS RELACIONAR AQUEST PDF AMB AQUESTA COMANDA.", "Numero de Pressupost", atrim(vnumc))
      If vnompressupost = "" Then Exit Sub
      'vnompressupost = vnompressupost
      'If MsgBox("Vols assignar aquest pressupost a aquesta comanda?", vbExclamation + vbDefaultButton2 + vbYesNo, "Assignar pressupost") = vbYes Then
      ruta_documentacio_pressupostos = llegir_ini("ruta", "ruta_documentacio_pressupostos", rutadelfitxer(cami) + "valorsprograma.ini")
      If Not existeix(ruta_documentacio_pressupostos + "\" + atrim(vclient)) Then MkDir ruta_documentacio_pressupostos + "\" + atrim(vclient)
      Copiar_Fitxer vpdf, ruta_documentacio_pressupostos + "\" + atrim(vclient) + "\" + atrim(vnompressupost) + "_" + atrim(vnumc) + ".pdf", 6
      dbtmp.Execute "update comandes set numpressupost='" + vnompressupost + "' where comanda=" + atrim(vnumc)
      'Kill Data.Files(1)
      Data1.Recordset.Move 0
      'End If
       Else: MsgBox "Aquest fitxer no es PDF i no puc linkar-lo.", vbCritical, "Error"
  End If
End Sub

Sub obrircarpetadragdrop()
 'With CommonDialog1
 ' .DialogTitle = "Deixar Fitxer comanda: " + atrim(data1.Recordset!comanda)
 ' .flags = cdlOFNExplorer
 ' .ShowOpen
 'End With
 Shell "explorer.exe c:\", vbNormalFocus
End Sub

Private Sub possarelpreu()
   Dim vpreu As String
   Dim vmoneda As String
   Dim vpressupost As String
   Dim rstcli As Recordset
   Dim vimpostinclos As Boolean
   
   If atrim(Data1.Recordset!producte) = "PC" Or atrim(Data1.Recordset!producte) = "PC2" Or atrim(Data1.Recordset!producte) = "PCP" Or atrim(Data1.Recordset!producte) = "PC3I" Then MsgBox "Només es pot possar preu a la comanda principal.", vbCritical, "Error": Exit Sub
   vimpostinclos = False
   vmoneda = buscarmonedadelclient(Data1.Recordset!client, cadbl(Text32(3).Tag))
   If velclientvolPVPimpostinclos Then If MsgBox("A T E N C I Ó" + vbNewLine + vbNewLine + "AQUEST CLIENT VOL L'IMPOST DEL PLÀSTIC INCLÒS AL PVP." + vbNewLine + " ASSEGUREU-VOS QUE EL PVP ES AMB L'IMPOST APLICAT." + vbNewLine + "EL PREU QUE ENTRARÀS ES AMB IMPOST?", vbExclamation + vbDefaultButton2 + vbYesNo, "A T E N C I Ó....") = vbYes Then vimpostinclos = True
   If Not comprovarrelaciomesuraPVPidemanada Then MsgBox "Les unitats de PVP i quantitat demanada han d'esser la mateixa", vbCritical, "Error": Exit Sub
   vpreu = InputBox("Entra el PVP en Euros d'aquesta comanda." + IIf(vimpostinclos = True, " (AMB IMPOST INCLÒS)", "") + Chr(10) + "Si vols preu zero escriu [zero]" + vbNewLine + "Si vols sense cost pel client escriu [sense cost]", "Preu Euros")
   If Command11.BackColor = &HC0C0FF Then
     MsgBox "Hi ha algú gravant, espera uns segons i torna-ho a provar.", vbCritical, "Error": Exit Sub
   End If
   If cadbl(vpreu) > 0 Or UCase(vpreu) = "ZERO" Or UCase(vpreu) = "SENSE COST" Then
      If UCase(vpreu) = "SENSE COST" Then vpreu = "-1"
      If UCase(vpreu) = "ZERO" Then vpreu = "0"
      dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','PVP','" + atrim(Data1.Recordset!pvp) + "','" + atrim(vpreu) + "')"
      dbtmp.Execute "update comandes set PVP=" + passaradecimalpunt(cadbl(vpreu)) + " where comanda=" + atrim(Data1.Recordset!comanda)
      dbtmp.Execute "update comandes_extres set PVPimpostinclos=" + IIf(vimpostinclos, "True", "False") + " where comanda=" + atrim(Data1.Recordset!comanda)
      If vpreu > 0 Then  'si es ARDO Posso una firma de PVP automàticament la segona es posarà automaticament a tot el grup
        Set rstcli = dbtmp.OpenRecordset("select grupdeclient from clients where codi=" + atrim(Data1.Recordset!client))
        If Not rstcli.EOF Then
            If atrim(rstcli!grupdeclient) = "ARDO" Then
              dbtmp.Execute "delete * from comandes_firmes where comanda=" + atrim(Data1.Recordset!comanda) + " and tipus='PVP'"
              dbtmp.Execute "insert into comandes_firmes (comanda,usuari,tipus,data) values (" + atrim(Data1.Recordset!comanda) + ",'ARDO_PVP1','PVP',now)"
            End If
        End If
      End If
   End If
   If vmoneda = "Dolars" Then
     vpreu = InputBox("Entra el PVP en Dolars d'aquesta comanda." + Chr(10) + "Si vols preu zero escriu [zero]", "Preu Dolars")
     If Command11.BackColor = &HC0C0FF Then MsgBox "Hi ha algú gravant, espera uns segons i torna-ho a provar.", vbCritical, "Error": Exit Sub
     If cadbl(vpreu) > 0 Or UCase(vpreu) = "ZERO" Then
       dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','PVPDOLAR','" + atrim(Data1.Recordset!pvpdolar) + "','" + atrim(vpreu) + "')"
       dbtmp.Execute "update comandes set pvpdolar=" + passaradecimalpunt(cadbl(vpreu)) + " where comanda=" + atrim(Data1.Recordset!comanda)
       dbtmp.Execute "update comandes_extres set PVPimpostinclos=" + IIf(vimpostinclos, "True", "False") + " where comanda=" + atrim(Data1.Recordset!comanda)
     End If
   End If
   If atrim(Data1.Recordset!numpressupost) = "" Then
        Set rstcli = dbtmp.OpenRecordset("select grupdeclient from clients where codi=" + atrim(Data1.Recordset!client))
        If Not rstcli.EOF Then
            If atrim(rstcli!grupdeclient) = "ARDO" Then
                vpressupost = UCase(InputBox("Entra el Codi d'agrupació per ARDO", "Codi agrupació ARDO"))
                If vpressupost <> "" Then
                   If Command11.BackColor = &HC0C0FF Then MsgBox "Hi ha algú gravant, espera uns segons i torna-ho a provar.", vbCritical, "Error": Exit Sub
                   dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','NUMPRESSUPOST','" + atrim(Data1.Recordset!numpressupost) + "','" + atrim(vnumpressupost) + "')"
                   dbtmp.Execute "update comandes set numpressupost='" + atrim(vpressupost) + "' where comanda=" + atrim(Data1.Recordset!comanda)
                End If
            End If
        End If
   End If
   Data1.Recordset.Move 0
End Sub
Function buscarmonedadelclient(vcodiclient As Double, vcodicomptable As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from clients_codiscomptables where codifabricacio=" + atrim(vcodiclient) + " and codicomptable=" + atrim(vcodicomptable))
   If Not rst.EOF Then
     buscarmonedadelclient = atrim(rst!moneda)
     If buscarmonedadelclient = "" Then buscarmonedadelclient = "Euros"
   End If
   Set rst = Nothing
End Function

Private Sub bpossarelpreu_Click()

End Sub

Private Sub bxl_Click()
 Dim vcontrol As Control
  Load formescullirlleixa
  formescullirlleixa.Top = formscrooll.Top + formcomandes.Top
  formescullirlleixa.Left = formscrooll.Left + formcomandes.Left
  formescullirlleixa.Show 1
  If seleccioret = 1 Then text77(27) = formescullirlleixa.valorescullit
  bxl.Visible = False
End Sub

Private Sub Check1_Click(Index As Integer)
   'dbtmp.Execute "update comandes_extres set pararaexpedicions=" + atrim(Check1(1).Value) + "  where comanda=" + atrim(cadbl(Text1.Text))
End Sub

Private Sub checkpassaraproduccio_Click()
   Dim vc As String
   On Error Resume Next
   vc = Screen.ActiveControl.Name
  'si es color verd el botó d'imprimir es que ja es impresa, avisar que no es pot canviar
   If Command9(0).BackColor = &HC0FFC0 Then MsgBox "Aquesta comanda ja està impresa amb paper per modificar-ho s'ha de fer desde Planificació StandBy.", vbInformation, "Atenció": Exit Sub
   If vc <> "" Then formcomandes.Controls(vc).SetFocus
   If vc <> "checkpassaraproduccio" Then Exit Sub
   If InStr(1, nomclient, "CROP´S") = 0 Then
         dbtmp.Execute "update comandes_extres set passaraimpresores=" + atrim(checkpassaraproduccio.Value) + "  where comanda=" + atrim(cadbl(Text1.Text))
           Else:
             dbtmp.Execute "update comandes_extres set passaraimpresores=0  where comanda=" + atrim(cadbl(Text1.Text))
             checkpassaraproduccio.Value = 0
             MsgBox "Aquest client es CROP'S i no es pot passar directament a producció." + Chr(10) + "S'ha d'activar desde planificació [StandBy]", vbInformation, "Atenció"
             
   End If
End Sub

Private Sub checkpassaraproduccio_GotFocus()
   If vnodemanarcontrasenyapassaraimpresores Then Exit Sub
   If UCase(InputBox("Entra la contrasenya per marcar comandes com a apunt per imprimir.", "Control de seguretat")) <> "INPLACSA" Then sortir.SetFocus: Exit Sub
   vnodemanarcontrasenyapassaraimpresores = True
End Sub

Private Sub cimpressio_Click()
  'If cimpressio = "Nova" Then MsgBox "Per escullir nova ho has de fer al moment de duplicar la comanda.", vbCritical, "Atenció": Exit Sub
  Text64.Text = Mid(cimpressio.Text, 1, 1)
 ' If cimpressio = "Modificada" Then
 '
 '
 '     If estatdelclixe(cadbl(data1.Recordset!numtreball), modificaciodeltreballmesgran(cadbl(data1.Recordset!numtreball))) <> "CLIXES ENTRATS" Then
 '        passardadestreballacomanda CInt(numtreball), numnovamodificacio, cadbl(Text1)
 '        data1.Recordset.Move 0
 '        posarmarcailinia numtreballdelacomanda(data1.Recordset!comanda)
 '              Else
 '                MsgBox "No hi ha cap versió oberta per aquest Nº de treball, no puc passar la comanda a modificada.", vbCritical, "Atenció"
 '                cimpressio = "Falta Autoritzar"
 '                Text64.Text = "F"
 '
 '     End If
 '
 ' End If
 ' If cimpressio = "Repetida" Then
 '    If estatdelclixe(cadbl(data1.Recordset!numtreball), cadbl(data1.Recordset!numordremodificacio)) <> "CLIXES ENTRATS" Then
 '        MsgBox "No pots passar aquesta comanda a repetida amb aquest treball i versió." + Chr(10) + "Els CLIXES NO ESTAN ENTRATS", vbCritical, "Atenció"
 '        cimpressio = "Falta Autoritzar"
  '       Text64.Text = "F"
  '   End If
  'End If
End Sub
Sub crearnovamodificacioaltreball()
  Dim numnovamodificacio As Integer
  Dim numtreball As Double
  numtreball = numtreballdelacomanda(Data1.Recordset!comanda)
  numnovamodificacio = fernovamodificacio(numtreball)
  If numnovamodificacio > 1 Then
     passardadestreballacomanda CInt(numtreball), numnovamodificacio, cadbl(Text1)
     Data1.Recordset.Move 0
     posarmarcailinia numtreballdelacomanda(Data1.Recordset!comanda)
  End If
  
End Sub
Function fernovamodificacio(numtreball As Double) As Integer
  Dim inici As Date
  
 Shell "\\serverprodu\dades\progcomandes\aplicacio\clixesnous.exe " + fitxerini + " " + atrim(numtreball) + " novamodificacio", vbNormalFocus
 inici = Now
 While llegir_ini("General", "creantmodificacio", "clixes.ini") = "si" And DateDiff("s", inici, Now) < 5
    DoEvents
 Wend
 If llegir_ini("General", "creantmodificacio", "clixes.ini") = "si" Then
    MsgBox "No s'ha pogut crear la modificacio nova per aquest treball"
    fernovamodificacio = 0
   Else: fernovamodificacio = cadbl(llegir_ini("General", "creantmodificacio", "clixes.ini"))
 End If
End Function
Private Sub cimpressio_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub





Private Sub cinta_Click(Index As Integer)

End Sub

Private Sub Combo1_Change(Index As Integer)
  'If Index = 0 Then If cadbl(Combo1(0).Text) > 4 Or cadbl(Combo1(0).Text) < 0 Then Combo1(0).Text = "0"
End Sub

Private Sub Combo1_Click(Index As Integer)
  If Index = 3 Then
     dbtmp.Execute "update comandes_extres set pararaexpedicions=" + atrim(Combo1(3).ItemData(Combo1(3).ListIndex)) + "  where comanda=" + atrim(cadbl(Text1.Text))
  End If
  If Index = 1 Then
      If Combo1(1) <> Data1.Recordset!marques Then
        If MsgBox("Segur que vols canviar el destí?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
             Data1.Recordset!marques = Combo1(1)
           Else
             Combo1(1) = Data1.Recordset!marques
        End If
          
      End If
  End If
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 3 Or Index = 4 Then
      KeyCode = 0
   End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 3 Or Index = 4 Then
    KeyAscii = 0
  End If
End Sub

Private Sub Combo10_LostFocus()
If Combo10 <> "B" And Combo10 <> "L" And Combo10 <> "F" And Combo10 <> "BB" Then
     Combo10 = ""
End If
End Sub

Private Sub Combo11_LostFocus()
If Combo11 <> "S" And Combo11 <> "N" Then
     Combo11 = "N"
End If
End Sub

Private Sub Combo12_Change()

End Sub

Private Sub Combo14_Change()
  If Chr$(KeyAscii) <> "N" And Chr$(KeyAscii) <> "C" And Chr$(KeyAscii) <> "1" And Chr$(KeyAscii) <> "2" Then
     KeyAscii = Asc("N")
   Else: Combo14.Text = ""
  End If

End Sub

Private Sub Combo15_LostFocus()
 If cadbl(Combo15.Text) > 4 Or cadbl(Combo15.Text) < 0 Then Combo15.Text = "0"
End Sub

Private Sub Combo2_Change()
 If cadbl(Combo2.Text) > 4 Or cadbl(Combo2.Text) < 0 Then Combo2.Text = "0"
 
End Sub

Private Sub Combo2_LostFocus()
calcular_ample_lam
Combo3.Text = passaradecimal(Combo2.Text)
End Sub

Private Sub Combo4_LostFocus()
If Combo4 <> "E" And Combo4 <> "I" Then
     Combo4 = ""
End If
End Sub



Private Sub Combo5_LostFocus(Index As Integer)
  If Index = 1 Then
    If Combo5(Index) <> "S" And Combo5(Index) <> "N" Then
         Combo5(Index) = "N"
    End If
  End If
  If Index = 0 Then
    If Combo5(Index) <> "C" And Combo5(Index) <> "F" And Combo5(Index) <> "N" Then
         Combo5(Index) = "N"
    End If
  End If
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Combo6_LostFocus()
'If Combo6 <> "S" And Combo6 <> "N" Then
'     Combo6 = "N"
'End If
End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub Combo7_LostFocus()
'If Combo7 <> "S" And Combo7 <> "N" Then
'     Combo7 = "N"
'End If
End Sub

Private Sub Combo8_LostFocus()
If Combo8 <> "T" And Combo8 <> "L" And Combo8 <> "ST" Then
     Combo8 = ""
End If
End Sub

Private Sub Combo9_LostFocus()
If Combo9 <> "T" And Combo9 <> "L" And Combo9 <> "ST" Then
     Combo9 = ""
End If

End Sub

Private Sub comboenvios_Click()
If comboenvios.ListIndex <> -1 Then
 Data1.Recordset!direnvio = comboenvios.ItemData(comboenvios.ListIndex)
 If comboenvios.Text = "" Then Data1.Recordset!direnvio = 0
 Label1(147).Caption = "Envio: " + comboenvios.Text
End If
 comboenvios.Visible = False
End Sub

Private Sub comboenvios_LostFocus()
comboenvios_Click
End Sub

Sub canvideclientoenvio()
  Dim clientnou As Double
  Dim direccionova As Double
  Dim direnviovell As Double
  Dim codicomptablevell As String
  Dim vimpresio As String
  Unload formduplicarcomanda
  Load formduplicarcomanda
  'formduplicarcomanda.Frame2.Height = 2800
  formduplicarcomanda.Frame1.Visible = False
  formduplicarcomanda.Height = 4000
  formduplicarcomanda.Caption = "Escullir Client"
  formduplicarcomanda.Command1.Visible = True
  formduplicarcomanda.Show 1

  codicomptablevell = Text32(3).Tag
  direnviovell = cadbl(Data1.Recordset.Fields!direnvio)
  If direnviovell = cadbl(formduplicarcomanda.direccioenvio.Tag) And cadbl(formduplicarcomanda.codicomptable.Tag) = cadbl(codicomptablevell) Then Exit Sub
  If formduplicarcomanda.Tag <> "sortir" Then
     enviaremailriscsuperat Data1.Recordset!comanda, formduplicarcomanda.codiclient.Tag
     Data1.Recordset.Fields!client = cadbl(formduplicarcomanda.codiclient)
     Data1.Recordset.Fields!direnvio = cadbl(formduplicarcomanda.direccioenvio.Tag)
     vimpresio = crearclientvinculat(Data1.Recordset)
     ' si hi ha canvi d'envio
     If cadbl(formduplicarcomanda.direccioenvio.Tag) <> direnviovell Then
           If Not (atrim(Data1.Recordset.Fields!impressio) = "N" And cadbl(Data1.Recordset!numtreball) = 0) Then
                Data1.Recordset.Fields!impressio = vimpresio
                Data1.Recordset.Fields!marques = IIf(vimpresio <> "N", "Si", "No")
           End If
           dbtmp.Execute "update comandes_Extres set aviscanvisambeltreball='Duplicada i canvi envio' where comanda=" + atrim(cadbl(Data1.Recordset!comanda))
     End If
     dbtmp.Execute "update comandes_Extres set codicomptable=" + atrim(cadbl(formduplicarcomanda.codicomptable.Tag)) + " where comanda=" + atrim(cadbl(Data1.Recordset!comanda))
     eliminar_dienviovell_sical direnviovell, cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), cadbl(Data1.Recordset!comanda)
     gravar_registre
  End If
  Unload formduplicarcomanda
End Sub
Sub eliminar_dienviovell_sical(dirvenviovell As Double, vntreball As Double, vordremodificacio As Double, comandaactual As Double)
   Dim rst As Recordset
   
   Set rst = dbtmp.OpenRecordset("select * from comandes where numtreball=" + atrim(vntreball) + " and numordremodificacio" + IIf(vordremodificacio = 0, "<2", "=" + atrim(vordremodificacio)) + " and direnvio=" + atrim(dirvenviovell) + " and comanda<>" + atrim(comandaactual))
   If rst.EOF Then dbclixes.Execute "delete * from clientsvinculats where not arxiuimp and id_treball=" + atrim(vntreball) + " and ordremodificacio=" + IIf(vordremodifiacio = 0, "1", atrim(vordremodificacio)) + " and direnvio=" + atrim(dirvenviovell)
   Set rst = Nothing
   
End Sub

Function espoteditarelcamp(nomcamp) As Boolean
   If InStr(1, llistadecampsvalids, "[" + nomcamp + "]") > 0 Or InStr(1, llistadecampsvalids, "[tots]") Then espoteditarelcamp = True
End Function
Sub vincularpressupost()
  Dim ruta_documentacio_pressupostos As String
  Dim vnomfitxer As String
  Dim vnompressupost As String
  Dim vcodiclient As Double
  Dim vnomp As String
  Dim vnumc As Double

  
  vnumc = cadbl(Data1.Recordset!comanda)
  If Data1.Recordset.EOF Then Exit Sub
  vcodiclient = cadbl(Data1.Recordset!client)
  ruta_documentacio_pressupostos = llegir_ini("ruta", "ruta_documentacio_pressupostos", rutadelfitxer(cami) + "valorsprograma.ini")
  vnomfitxer = ruta_documentacio_pressupostos + "\" + atrim(vcodiclient) + "\" + atrim(Text41) + IIf(InStr(1, Text41, "_") = 0, "_" + atrim(vnumc), "") + ".pdf"
  If existeix(vnomfitxer) Then
     If MsgBox("El pressupost ja està linkat, vols veure el PDF?", vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
        'obrir_document vnomfitxer
        Shell "cmd /c start chrome.exe " + vnomfitxer
        Exit Sub
          Else
            If MsgBox("Vols eliminar aquesta relació?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
            Data1.Database.Execute "update comandes set numpressupost='' where comanda=" + atrim(Data1.Recordset!comanda)
            Data1.Recordset.Move 0
            Exit Sub
     End If
  End If
  If existeix(ruta_documentacio_pressupostos + "\" + atrim(vcodiclient)) Then
    vnompressupost = buscarpressupostalspdfvinculats(ruta_documentacio_pressupostos + "\" + atrim(vcodiclient))
    If InStr(1, nomclient, "GLOBALINNOVA") > 0 Then
        If cadbl(Mid(vnompressupost, 1, 4)) <> Year(Now) And cadbl(Mid(vnompressupost, 1, 4)) <> (Year(Now) - 1) Then
             MsgBox "El nom d'aquest pressupost no sembla el format demanat per el client GLOBALINNOVA. " + vbNewLine + "Ex: 202408120913  ANYMESDIAHORAMINUT", vbCritical, "ATENCIÓ"
             Exit Sub
        End If
    End If
    If InStr(1, vnompressupost, ".pdf") > 0 Then
       vnomp = Mid(vnompressupost, 1, InStr(1, vnompressupost, ".pdf") - 1)
       If Data1.Recordset.EditMode > 0 Then
          Text41 = vnomp
           Else
              Copiar_Fitxer ruta_documentacio_pressupostos + "\" + atrim(vcodiclient) + "\" + vnompressupost, ruta_documentacio_pressupostos + "\" + atrim(vcodiclient) + "\" + vnomp + "_" + atrim(vnumc) + ".pdf"
              Data1.Database.Execute "update comandes set numpressupost='" + vnomp + "' where comanda=" + atrim(Data1.Recordset!comanda)
              Data1.Recordset.Move 0
       End If
    End If
      Else: MsgBox "No hi ha cap pressupost entrat per aquest client" + Chr(10) + "Primer entra el pdf a pressupostos", vbCritical, "Error"
  End If
End Sub
Function buscarpressupostalspdfvinculats(vdirectori As String)
CommonDialog1.CancelError = True
On Error Resume Next
With CommonDialog1
  .DialogTitle = "Seleccionar fitxer pressupost"
  .flags = cdlOFNExplorer
  .DefaultExt = ".pdf"
  .InitDir = vdirectori
  .ShowOpen
End With
If err.Number <> &H7FF3 Then buscarpressupostalspdfvinculats = CommonDialog1.FileTitle
End Function
Sub veure_documents_comanda(vnumc As Double)
   Dim ruta_documentacio As String
   ruta_documentacio = llegir_ini("ruta", "ruta_comandes_exportades ", rutadelfitxer(cami) + "valorsprograma.ini")
   ruta_documentacio = ruta_documentacio + "\Les_" + atrim(atrim(Int(cadbl(vnumc) / 1000)) + "000") + "\" + atrim(vnumc)
   ratoli "espera"
   If existeix(ruta_documentacio) Then
      Shell "explorer.exe " + ruta_documentacio, vbNormalFocus
      ratoli "normal"
       Else:
          MsgBox "No hi ha documentació encara d'aquesta comanda", vbInformation, "Atenció"
          MkDir ruta_documentacio
          Shell "explorer.exe " + ruta_documentacio, vbNormalFocus
   End If
   ratoli "normal"

End Sub
Sub buscar_tarifa_referencies(vcodiclient As String, vcoditarifa As String)
   Dim vruta As String
   Dim vdir As String
   Dim rst As Recordset
   Dim vcont As Byte
   Dim vrutacomplerta As String
   vruta = llegir_ini("General", "rutatarifesDRIVE", "comandes.ini")
   'If Not existeix(vruta) Then
   '   vruta = "G:\Unidades compartidas\Escandalls Tarifes"
   '   escriure_ini "General", "rutatarifesDRIVE", "comandes.ini"
   vcont = 0
   While (Not existeix(vruta) Or vruta = "") And vcont < 2
       vruta = InputBox("Entra la ruta del DRIVE on hi ha les tarifes" + Chr(10) + "No escriguis res per sortir.", "Atenció", "G:\Unidades compartidas\Escandalls Tarifes")
       If existeix(vruta) And vruta <> "" Then escriure_ini "General", "rutatarifesDRIVE", vruta, "comandes.ini"
       vcont = vcont + 1
   Wend
   If vcont > 9 Then Exit Sub
   Set rst = dbtmp.OpenRecordset("select grupdeclient from clients where codi=" + atrim(vcodiclient))
   If Not rst.EOF Then
      If atrim(rst!grupdeclient) <> "" Then vcodiclient = atrim(rst!grupdeclient)
       Set rst = Nothing
   End If
   If vruta = "" Then Exit Sub
   If cadbl(vcodiclient) > 0 Then vcodiclient = Format(vcodiclient, "000000")
   vdir = Dir(vruta + "\" + vcodiclient + "*", vbDirectory)
   While vdir <> ""
      If Mid(vdir + "                ", 1, Len(vcodiclient)) = vcodiclient Then GoTo cont
      vdir = Dir
   Wend
cont:
   If vdir <> "" Then
      vrutacomplerta = vruta + "\" + vdir
      vdir = Dir(vruta + "\" + vdir + "\TARIFA ACTIVA*", vbDirectory)
      If vdir <> "" Then
         vrutacomplerta = vrutacomplerta + "\" + vdir
          Else: GoTo fi
      End If
      vdir = Dir(vrutacomplerta + "\" + Format(cadbl(vcoditarifa), "000") + " -*.*", vbArchive)
      If vdir <> "" Then
'          Clipboard.Clear
'          Clipboard.SetText vrutacomplerta + "\" + vdir
          vrutacomplerta = vrutacomplerta + "\" + vdir
      End If
      
   End If
fi:
  
  If Len(vrutacomplerta) > 255 Then
       vrutacomplerta = rutadelfitxer(vrutacomplerta)
       MsgBox "Aquest arxiu té el nom massa llarg per obrir-lo desde el windows, t'obriré la carpeta on està ubicada la tarifa i has d'escullir manualment " + vbNewLine + "la tarifa Nº: " + atrim(vcoditarifa), vbInformation, "Atenció"
       Shell "c:\windows\system32\cmd.exe /c Explorer.exe " + """" + vrutacomplerta + """", vbMaximizedFocus
      Else
        If existeix(vrutacomplerta) Then
            Shell "c:\windows\system32\cmd.exe /c """ + vrutacomplerta + """, vbMaximizedFocus"
              Else: MsgBox "No trobo el fitxer especificat..." + vbNewLine + vrutacomplerta, vbCritical, "Error"
        End If
  End If
Set rst = Nothing
End Sub
Sub canvi_tarifa_referencia(vcodiclient As String, vrefcli As String, vcoditarifa As String)
    
    If vcodiclient = "" Or vrefcli = "" Or vcoditarifa = "" Then Exit Sub
    vcoditarifa = InputBox("Entra el NOU codi de tarifa de la referencia " + atrim(vrefcli) + Chr(10) + "Ex: 001" + Chr(10) + "PER ELIMINAR RELACIÓ ESCRIU [ELIMINAR]", "Codi tarifa", vcoditarifa)
    If UCase(vcoditarifa) = "ELIMINAR" Then dbtmp.Execute "delete * from tarifes_referencies where refinplacsa='" + atrim(vrefcli) + "'": GoTo fi
    If cadbl(vcoditarifa) = 0 Then MsgBox "Aquest codi no es correcte.", vbCritical, "Error": Exit Sub
    dbtmp.Execute "update tarifes_referencies set coditarifa='" + vcoditarifa + "'  where refinplacsa='" + atrim(vrefcli) + "'"
    dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','Codi_Tarifa','" + atrim(Label1(148)) + "','" + atrim(vcoditarifa) + "')"
fi:
    carregartarifesperreferencia vcodiclient, vrefcli
End Sub
Sub assignar_tarifa_referencia(vcodiclient As String, vrefinplacsa As String)
    Dim vcoditarifa As String
    Dim rst As Recordset
    If vrefinplacsa = "" Then MsgBox "NO ES POT ASSIGNAR TARIFA SENSE REFERENCIA INPLACSA.", vbCritical, "ERROR"
    If vcodiclient = "" Or vrefinplacsa = "" Then Exit Sub
    vcoditarifa = InputBox("Entra el codi de la referencia d'inplacsa " + atrim(vrefcli) + Chr(10) + "Ex: 001", "Codi tarifa")
    If cadbl(vcoditarifa) = 0 Then MsgBox "Aquest codi no es correcte.", vbCritical, "Error": Exit Sub
    Set rst = dbtmp.OpenRecordset("select * from tarifes_referencies where refinplacsa='" + vrefinplacsa + "'")
    If Not rst.EOF Then
       rst.Edit
       rst!coditarifa = vcoditarifa
       rst.Update
        Else: dbtmp.Execute "insert into tarifes_referencies (codiclient,refinplacsa,coditarifa) values ('" + atrim(vcodiclient) + "','" + atrim(vrefinplacsa) + "','" + atrim(vcoditarifa) + "')"
    End If
    'dbtmp.Execute "delete * from tarifes_referencies where refinplacsa='" + atrim(vrefinplacsa) + "'"
    dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','Codi_Tarifa','" + atrim(Label1(148)) + "','" + atrim(vcoditarifa) + "')"
    carregartarifesperreferencia vcodiclient, vrefinplacsa
End Sub

Private Sub Command1_Click(Index As Integer)
  Dim ruta As String
  Dim nomfitxer
  If Index = 10 Then
    escriure_ini "baixes", "imprimirpackinglist", "0", "comandes.ini"
    wait 1
    escriure_ini "baixes", "imprimirpackinglist", "1", "comandes.ini"
    Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe comandes.ini " + atrim(cadbl(Data1.Recordset!comanda) * -1), vbNormalFocus
  End If
  If Index = 8 Then veure_documents_comanda cadbl(Data1.Recordset!comanda)
  If Index = 5 Then canvideclientoenvio
  If Index = 9 Then
      If Command1(9).ToolTipText <> "" And atrim(Data1.Recordset!refclient) <> "" Then
           buscar_tarifa_referencies Command1(9).Tag, Mid(Command1(9).ToolTipText, Len("Codi tarifa: ") + 1)
             Else: assignar_tarifa_referencia Command1(9).Tag, atrim(Text32(5))
      End If
  End If
  If Index = 4 Then
    nomfitxer = Data1.Recordset!arxiuimpressora
    If cadbl(Mid(nomfitxer, 1, 6)) = 0 Then nomfitxer = numcarpetaclient + " " + Trim(nomfitxer)
    ruta = ruta_relativa_docs + "\" + nomfitxer ' + Chr$(34)
    If existeix(ruta) Then
     obrir_document ruta
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + ruta, vbCritical, "Error"
    End If
  End If
  If Index = 7 Then
    If Data1.Recordset.EditMode > 0 Then
       MsgBox "Estas editant la comanda, primer guarda els canvis abans de linkar el pressupost", vbCritical, "Linkar pressupost"
       Exit Sub
        Else: vincularpressupost
    End If
  End If
 If Index = 0 Then
   'If cadbl(modificaciotreball) = 0 Then
     'r = obre_fitxer(ruta_relativa_docs, 2)
     'Text78 = Mid(r, Len(ruta_relativa_docs) + 2)
     'Text78.SetFocus
    'Else
       obrir_pdf_treball cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio)
   'End If
 End If
 If Index = 1 And numtreballdelacomanda(Data1.Recordset!comanda) > 0 Then
   Shell "\\serverprodu\dades\progcomandes\aplicacio\clixesnous.exe " + fitxerini + " " + atrim(numtreballdelacomanda(Data1.Recordset!comanda)), vbNormalFocus
 End If
 If Index = 2 Then
  ' If cadbl(modificaciotreball) = 0 Then
  '     r = obre_fitxer(ruta_relativa_docs, 2)
  '     Text79 = Mid(r, Len(ruta_relativa_docs) + 2)
  '     Text79.SetFocus
  '    Else
          obrir_imp_treball cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), cadbl(Data1.Recordset!client), cadbl(Data1.Recordset!direnvio)
   'End If
 End If
 If Index = 3 Then
    If Not espoteditarelcamp("numtreball") Then MsgBox "No pots editar aquest camp, no tens permís per fer-ho", vbCritical, "Atenció": Exit Sub
    If atrim(Data1.Recordset!proximaseccio) = "T" Then
       If MsgBox("No pots modificar aquest valor si la comanda està entregada" + vbNewLine + "Vols fer-ho igualment?", vbCritical + vbDefaultButton2 + vbYesNo, "Error") = vbNo Then Exit Sub
    End If
    
    demanarnumerodetreballiordre
    
    
    If Not buscant And triarversiotreball.ntreball <> "" Then
      posarmarcailinia numtreballdelacomanda(Data1.Recordset!comanda)
      crearclientvinculat Data1.Recordset
      areadedatos False
    End If
 End If
End Sub
Function numtreballdelacomanda(numc As Double) As Double
   Dim rstc As Recordset
   Set rstc = dbtmp.OpenRecordset("select numtreball from comandes where comanda=" + atrim(cadbl(numc)))
   If Not rstc.EOF Then numtreballdelacomanda = cadbl(rstc!numtreball)
End Function
Sub demanarnumerodetreballiordre()
   Dim treball As Integer
   Dim ordremodificacio As Integer
   Unload triarversiotreball
   Load triarversiotreball
   triarversiotreball.ntreball = atrim(cadbl(Data1.Recordset!numtreball))
   triarversiotreball.carregarversions cadbl(Data1.Recordset!numtreball)
   triarversiotreball.Show 1
   
   If triarversiotreball.ntreball = "" Then Exit Sub
   treball = cadbl(triarversiotreball.ntreball)
   If treball = -1 Then GoTo guardartreball
   ordremodificacio = triarversiotreball.llistaversions.ItemData(triarversiotreball.llistaversions.ListIndex)
   If treball > 0 Then
      'ordremodificacio = cadbl(InputBox("Escriu el numero de modificacio que vols relacionar amb aquesta comanda." + Chr(10) + "La que surt per defecte es la mes gran que té aquest treball.", "Modificacio del treball", atrim(modificaciodeltreballmesgran(treball))))
      If ordremodificacio > 0 And Not buscant Then
guardartreball:
        If treball = -1 Then
           ordremodificacio = 0: treball = 0
           possardadesdeltreballazero Data1.Recordset
             Else: Text64.Text = Mid(triarversiotreball.llistaversions, 1, 1)
        End If
        
        gravar_registre
        dbtmp.Execute "update comandes_Extres set aviscanvisambeltreball='' where comanda=" + atrim(cadbl(Text1))
        'dbtmp.Execute "update comandes_Extres set refinplacsa=8" + atrim(Data1.Recordset!client) + atrim(treball) + " where comanda=" + atrim(cadbl(Text1)) + " or comanda=" + atrim(cadbl(text77(11))) + " or comanda=" + atrim(cadbl(text77(12))) + ""
        
        passardadestreballacomanda treball, ordremodificacio, cadbl(Text1)
        Data1.Recordset.Move 0
      End If
      If buscant Then
         Text103(3) = atrim(treball) + "/" + atrim(ordremodificacio)
         Text103(3).Tag = atrim(treball)
         Text103(3).WhatsThisHelpID = cadbl(ordremodificacio)
      End If
   End If
End Sub
Sub passardadestreballacomanda(treball As Integer, ordre As Integer, numc As Double)
  Dim vestatclixe As String
  Dim vultimacomandaimpresa As String
  If Data1.Recordset.EditMode > 0 Then Data1.Recordset.CancelUpdate ': areadedatos False
  If cadbl(Data1.Recordset!numtreball) <> cadbl(treball) Then
       dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(numc) + ",'" + nomordinador + "','numtreball','" + atrim(Data1.Recordset!numtreball) + "','" + atrim(treball) + "')"
  End If
  If cadbl(Data1.Recordset!numordremodificacio) <> cadbl(ordre) Then
       dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(numc) + ",'" + nomordinador + "','versiotreball','" + atrim(Data1.Recordset!numordremodificacio) + "','" + atrim(ordre) + "')"
  End If
       
  dbtmp.Execute "update comandes set numtreball=" + atrim(treball) + ",numordremodificacio=" + atrim(ordre) + " where comanda=" + atrim(numc)
  Text103(3) = atrim(treball) + "/" + atrim(ordre)
  mirardiferenciescomandaitreball numc
  If imprimirdiferenciescomandaitreball(numc, "P") Then
    If Text64.Text <> "N" Then If MsgBox("Segur que vols canviar les dades de la comanda per les del treball?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
    posardiferenciesacomandadeltreball numc
  End If
  If atrim(Data1.Recordset!proximaseccio) <> "T" Then
        If Text64.Text <> "N" Or ordre > 1 Then
              vestatclixe = estatdelclixe(treball, ordre)
              If InStr(1, vestatclixe, " - ") > 0 Then vestatclixe = Mid(vestatclixe, InStr(1, vestatclixe, " - ") + 3)
              If vestatclixe <> "CLIXES ENTRATS" And vestatclixe <> "REPOSICIÓ DEL CLIXE" Then
                'If Data1.Recordset!impressio = "R" Then
                   If vestatclixe <> "RETORNEM CLIXES" Then
                      MsgBox "Aquest treball no te els CLIXES ENTRATS. La passo a modificada.", vbCritical, "Atenció"
                      'cimpressio = "Modificada"
                      'Text64.Text = "M"
                      dbtmp.Execute "update comandes set impressio='M' where comanda=" + atrim(numc)
                      If atrim(Data1.Recordset!impressio) <> "M" Then
                        dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(numc) + ",'" + nomordinador + "','impressio_','" + atrim(Data1.Recordset!impressio) + "','M')"
                      End If
                      demanar_si_canvi_refclient numc
                   End If
                      
                'End If
                  Else
                    ' If Data1.Recordset!impressio = "M" Or Data1.Recordset!impressio = "F" Then
                      vultimacomandaimpresa = ultimacomandaimpresa(treball, ordre)
                      If vultimacomandaimpresa <> "" Then
                       MsgBox "Aquest treball ja te els CLIXES ENTRATS i comanda impresa. La passo a repetida." + Chr(10) + " Ultima comanda impresa " + vultimacomandaimpresa, vbCritical, "Atenció"
                       dbtmp.Execute "update comandes set impressio='R' where comanda=" + atrim(numc)
                       dbtmp.Execute "update comandes_Extres set clientvindraarevisarimpresio=false where comanda=" + atrim(numc)
                       If atrim(Data1.Recordset!impressio) <> "R" Then
                         dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(numc) + ",'" + nomordinador + "','impressio','" + atrim(Data1.Recordset!impressio) + "','R')"
                       End If
                         Else
                         If ordre > 1 Then
                            dbtmp.Execute "update comandes set impressio='M' where comanda=" + atrim(numc)
                            If atrim(Data1.Recordset!impressio) <> "M" Then
                              dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(numc) + ",'" + nomordinador + "','impressio','" + atrim(Data1.Recordset!impressio) + "','M')"
                            End If
                            demanar_si_canvi_refclient numc
                         End If
                      End If
                      'cimpressio = "Repetida"
                      'Text64.Text = "R"
                     'End If
                     
              End If
         End If
   End If
'   MaskEdBox18.SetFocus
End Sub
Sub demanar_si_canvi_refclient(vnumc As Double)
     Dim rst As Recordset
     Dim vrefnova As String
     

     
     Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
     If rst.EOF Then Exit Sub
     vrefnova = InputBox("S'ha passat la comanda a MODIFICADA vols canviar la referència del Client?", "Canvi de referència?", atrim(rst!refclient))
     If vrefnova <> atrim(rst!refclient) Then
        dbtmp.Execute "update comandes set refclient='" + atrim(vrefnova) + "' where comanda=" + atrim(rst!comanda) + IIf(cadbl(rst!linkcomanda1) > 0, " or comanda=" + atrim(rst!linkcomanda1), "") + IIf(cadbl(rst!linkcomanda2) > 0, " or comanda=" + atrim(rst!linkcomanda2), "")
        dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(vnumc) + ",'" + nomordinador + "','refclient','" + atrim(rst!refclient) + "','" + vrefnova + "')"
     End If
     
     Set rst = Nothing
End Sub
Function ultimacomandaimpresa(vtreball As Integer, vordre As Integer) As String
    Dim rst As Recordset
    Set rst = dbtmp.OpenRecordset("select * from comandes where proximaseccio<>'E' and proximaseccio<>'I' and numtreball=" + atrim(vtreball) + " and numordremodificacio=" + atrim(vordre) + " order by comanda desc")
    If Not rst.EOF Then ultimacomandaimpresa = atrim(rst!comanda)
    Set rst = Nothing
End Function
Function modificaciodeltreballmesgran(treball As Integer) As Integer
   Dim rst As Recordset
   modificaciodeltreballmesgran = 0
   Set rst = dbclixesnous.OpenRecordset("select ordre from modificacions where id_Treball=" + atrim(treball) + " order by ordre desc")
   If Not rst.EOF Then modificaciodeltreballmesgran = cadbl(rst!ordre)
    Set rst = Nothing
End Function

Sub obrir_imp_treball(treball As Double, modificacio As Double, codiclient As Double, direnvio As Double, Optional vnomfitxer As String)
   Dim generarfitxer_imp As String
   Dim vnoavisar As Boolean
   If vnomfitxer = "noavisar" Then vnoavisar = True
   If modificacio = 0 Then modificacio = 1
   generarfitxer_imp = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\IMP" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "-" + Format(codiclient, "000000") + "_" + atrim(direnvio) + ".doc"
   vnomfitxer = generarfitxer_imp
   If Not existeix(vnomfitxer) Then vnomfitxer = vnomfitxer + "x"
   If existeix(vnomfitxer) Then
     obrir_document generarfitxer_imp
    Else:
      If Not vnoavisar Then MsgBox "No he trobat el fitxer" + Chr(10) + vnomfitxer, vbCritical, "Error"
  End If
End Sub
Function existeix_imp_treball(treball As Double, modificacio As Double, codiclient As Double, direnvio As Double) As Boolean
   Dim generarfitxer_imp As String
   If modificacio = 0 Then modificacio = 1
   generarfitxer_imp = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\IMP" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "-" + Format(codiclient, "000000") + "_" + atrim(direnvio) + ".doc"
   If existeix(generarfitxer_imp) Then
     existeix_imp_treball = True
    Else: existeix_imp_treball = False
  End If
End Function

Sub obrir_pdf_treball(treball As Double, modificacio As Double)
   Dim generarfitxer_pdf As String
   If modificacio = 0 Then modificacio = 1
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + ".pdf"
   If existeix(generarfitxer_pdf) Then
     obrir_document generarfitxer_pdf
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + generarfitxer_pdf, vbCritical, "Error"
  End If
End Sub
Function existeix_pdf_treball(treball As Double, modificacio As Double) As Boolean
   Dim generarfitxer_pdf As String
   modificacio = 1
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + ".pdf"
   If existeix(generarfitxer_pdf) Then
     existeix_pdf_treball = True
    Else: existeix_pdf_treball = False
  End If
End Function

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Index = 9 And Button = 2 Then
       canvi_tarifa_referencia Command1(9).Tag, atrim(Text32(5)), Mid(Command1(9).ToolTipText, Len("Codi tarifa: ") + 1)
   End If
End Sub

Private Sub Command10_Click()
  escriure_ini "General", "anarbaixasec", "V", fitxerini
  escriure_ini "General", "anarbaixacom", atrim(cadbl(Data1.Recordset!comanda)), fitxerini
  obrir_baixes

End Sub

Private Sub Command11_Click()
   usuari_guarda_registre
End Sub

Private Sub Command12_Click()
  escriure_ini "General", "anarbaixasec", "E", fitxerini
  escriure_ini "General", "anarbaixacom", atrim(cadbl(Data1.Recordset!comanda)), fitxerini
  obrir_baixes
End Sub

Private Sub Command13_Click()
  recordsourcetotals = "select comanda,hcanvi,havaria,hfuncio,tkilos,kiloshora,extrussora from extrussorestot "
  querywhere = " where comanda=" + atrim(Data1.Recordset!comanda)
  formtotals.Show
End Sub

Private Sub Command14_Click()
escriure_ini "General", "anarbaixasec", "I", fitxerini
  escriure_ini "General", "anarbaixacom", atrim(cadbl(Data1.Recordset!comanda)), fitxerini
  obrir_baixes
End Sub

Private Sub Command15_Click()
  recordsourcetotals = "select comanda,hclixe,hmaquina,hajusts,hfuncio,tmetres,metresmin,impressora,dataimpressio  from impressorestot "
    querywhere = " where comanda=" + atrim(Data1.Recordset!comanda)
  formtotals.Show
  formtotals.Tag = "7"
End Sub

Private Sub Command16_Click()
escriure_ini "General", "anarbaixasec", "R", fitxerini
  escriure_ini "General", "anarbaixacom", atrim(cadbl(Data1.Recordset!comanda)), fitxerini
  obrir_baixes
End Sub

Private Sub Command17_Click()
recordsourcetotals = "select comanda,hcanvi,havaria,hfuncio,tmetres,mtrsmin,simulteneitat,rebobinadora,datarebobinat from rebobinadorestot "
  querywhere = " where comanda=" + atrim(Data1.Recordset!comanda)
  formtotals.Show
  formtotals.Tag = "9"
End Sub

Private Sub Command18_Click()
  escriure_ini "General", "anarbaixasec", "L", fitxerini
  escriure_ini "General", "anarbaixacom", atrim(cadbl(Data1.Recordset!comanda)), fitxerini
  obrir_baixes
End Sub

Private Sub Command19_Click()
recordsourcetotals = "select comanda,hcanvi,havaria,hfuncio,tmetres,metresmin,laminadora,datalaminacio  from laminadorestot "
  querywhere = " where comanda=" + atrim(Data1.Recordset!comanda)
  formtotals.Show
End Sub



Private Sub Command2_Click()
  
End Sub

Private Sub Command20_Click()
  recordsourcetotals = "select comanda,hcanvi,havaria,hparada,hfuncio,tunitats,unitatshora,simultaneitat,soldadora,datasoldadora  from soldadorestot "
    querywhere = " where comanda=" + atrim(Data1.Recordset!comanda)
  formtotals.Show
  formtotals.Tag = "9"
End Sub

Private Sub Command21_Click()
escriure_ini "General", "anarbaixasec", "S", fitxerini
  escriure_ini "General", "anarbaixacom", atrim(cadbl(Data1.Recordset!comanda)), fitxerini
  obrir_baixes
End Sub

Private Sub Command22_Click()

End Sub

Private Sub Command23_Click()
  If cadbl(Text2) > 0 Then
   Unload Avisos
   r = ruta_relativa_docs + "\" + carpeta_del_client + "\avisos"
   If existeix(r) Then
    fitxers.path = r
    fitxers.Refresh
    If fitxers.ListCount > 0 Then Avisos.Show 1, Me: Unload Avisos
   End If
  End If
End Sub

Private Sub Command24_Click()
  dataactivacio = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Command25_Click()
MaskEdBox21 = Format(Now, "dd/mm/yyyy")
End Sub



Private Sub Command27_Click()
  Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe", vbNormalFocus
End Sub
Sub preguntar_verificacio_referencia()
     If UCase(InputBox("Aquesta referencia nova no està verificada encara." + vbNewLine + "Escriu [VERIFICADA] per passar-la a verificada.", "VERIFICAR REFERENCIA")) = "VERIFICADA" Then
         dbtmp.Execute "update comandes_extres set refinplacsa_validada=true where comanda=" + atrim(Data1.Recordset!comanda) + " or comanda=" + atrim(Data1.Recordset!linkcomanda1) + " or comanda=" + atrim(Data1.Recordset!linkcomanda2)
         dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','VerificacioRefInplacsa','" + atrim(Text32(5)) + "','" + atrim(Text32(5)) + "')"
         possar_boto_refinplacsavalida True
     End If
End Sub

Sub obrir_disposicio_materials()
   If Text32(5) = "" Then Exit Sub
   Unload Formdisposiciomaterialscomanda
   Load Formdisposiciomaterialscomanda
   Formdisposiciomaterialscomanda.etrefinplacsa.Tag = Text32(5)
  ' Formdisposiciomaterialscomanda.etrefinplacsa.Tag = "01C5979I7356"

   Formdisposiciomaterialscomanda.Show 1
   Unload Formdisposiciomaterialscomanda
End Sub
Private Sub Command26_Click(Index As Integer)
 Dim vcodiean13 As String
 Dim v As String
'If Index = 0 Then MaskEdBox22 = Format(Now, "dd/mm/yyyy")
 If Index = 12 Then obrir_disposicio_materials
If Index = 4 Then
    vcodiean13 = "8402320" + Format(cadbl(Command26(4).Tag), "00000")
    vcodiean13 = vcodiean13 + atrim(EAN13_Control(vcodiean13))
    If cadbl(Command26(4).Tag) > 0 Then MsgBox "El codi GTIN assignat a aquesta referència es:" + Chr(10) + vcodiean13, vbInformation, "GTIN"
    Exit Sub
End If
If Index = 1 Then
  comprovar_sihihabobinesetiquetades
  ensenyar_refaltervatives
End If
If Index = 3 Then
  canviar_lestatdelacomanda
End If
If Index = 5 Then
   If Command26(5).Tag = "F" Then
        preguntar_verificacio_referencia
   End If
   posarframeopcionsrefinplacsa
End If
If Index = 6 Then
   canviar_observacioPVP
End If
If Index = 7 Then
   If hiharelacions(cadbl(Data1.Recordset!comanda), cadbl(Data1.Recordset!linkcomanda1), cadbl(Data1.Recordset!linkcomanda2)) Then
      v = UCase(InputBox("Aquesta comanda encara te RESERVES, ASSIGNACIONS O COMPRES assignades no es pot fer canvis fins que totes estiguin eliminades." + vbNewLine + "SI VOLS RECALCULAR EL CODI ESCRIU [RECALCULAR]", "Atenció"))
      If v = "RECALCULAR" Then
            dbtmp.Execute "update comandes_extres set refinplacsa='' where comanda=" + atrim(cadbl(Data1.Recordset!comanda)) + " or comanda=" + atrim(cadbl(Data1.Recordset!linkcomanda1)) + " or comanda=" + atrim(cadbl(Data1.Recordset!linkcomanda2))
            wait 1
            generarrefinplacsadefinitiu cadbl(Data1.Recordset!comanda)
            wait 1
            Data1.Recordset.Move 0
      End If
      ratoli "normal"
      Exit Sub
       Else: canviar_refinplacsa atrim(Text32(5))
   End If
End If
If Index = 8 Then
   If MsgBox("Segur que vols canviar l'estat a ACTIU d'aquesta referencia?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
     canviar_estatActiuInactiu "A"
   End If
End If
If Index = 9 Then
   If MsgBox("Segur que vols canviar l'estat a INACTIU d'aquesta referencia?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
     canviar_estatActiuInactiu "I"
   End If
   'donar_referenciacomavalida
End If
If Index = 10 Then
   comprovar_refinplacsa
End If
If Index = 11 Then
   ferbusquedapertreball
End If
End Sub
Sub ferbusquedapertreball()
   Unload subbusqueda
   subbusqueda.Show
   subbusqueda.Text1 = formcomandes.Data1.Recordset!client
   subbusqueda.nomclient = nomclient
   subbusqueda.cnumtreball.Tag = "1a"
   vtreballbuscatsubbusqueda = Text32(5)
   If InStr(1, vtreballbuscatsubbusqueda, "I") Then
      subbusqueda.cnumtreball = Mid(vtreballbuscatsubbusqueda, InStr(1, vtreballbuscatsubbusqueda, "I"))
        Else: subbusqueda.cnumtreball = "ANÒNIM"
   End If
   subbusqueda.SetFocus
   subbusqueda.ferlabusquedaperresum
End Sub
Sub donar_referenciacomavalida()
   Dim rst As Recordset
   If Data1.Recordset.EOF Then Exit Sub
   dbtmp.Execute "update comandes_extres set refinplacsa_valida=not refinplacsa_valida where comanda=" + atrim(Data1.Recordset!comanda)
   If cadbl(Data1.Recordset!linkcomanda1) > 0 Then dbtmp.Execute "update comandes_extres set refinplacsa_valida=not refinplacsa_valida where comanda=" + atrim(Data1.Recordset!linkcomanda1)
   If cadbl(Data1.Recordset!linkcomanda2) > 0 Then dbtmp.Execute "update comandes_extres set refinplacsa_valida=not refinplacsa_valida where comanda=" + atrim(Data1.Recordset!linkcomanda2)
   Set rst = dbtmp.OpenRecordset("select refinplacsa_valida from comandes_extres where comanda=" + atrim(Data1.Recordset!comanda))
   'If Not rst.EOF Then possar_boto_refinplacsavalida rst!refinplacsa_valida
   Set rst = Nothing
End Sub
Sub canviar_estatActiuInactiu(vestat As String)
    Dim rst As Recordset
    If Data1.Recordset.EOF Then Exit Sub
    Set rst = dbtmp.OpenRecordset("select * from tarifes_referencies where refinplacsa='" + Text32(5) + "'")
    If Not rst.EOF Then
       rst.Edit
       rst!inactiva = True
       If vestat = "A" Then rst!inactiva = False
       rst.Update
    End If
    carregartarifesperreferencia Data1.Recordset!client, Text32(5)
End Sub
Sub posarframeopcionsrefinplacsa()
   Frame1(1).Visible = Not Frame1(1).Visible
End Sub

Sub canviar_observacioPVP()
    Dim v As String
    Dim vcomandesafectades As String
    Dim vliniaalbara As String
    Dim vextracost As String
    Dim rst As Recordset
    Set rst = dbtmp.OpenRecordset("select * from comandeS_observaciopvp where comanda=" + atrim(Text1))
    If Not rst.EOF Then
        v = atrim(rst!observacio)
        vextracost = atrim(cadbl(rst!extracost))
        vliniaalbara = atrim(rst!liniaalbara)
        vcomandesafectades = atrim(rst!comandesafectades)
         Else
          Set rst = dbtmp.OpenRecordset("select * from comandes_observacioPVP where comandesafectades like '*" + atrim(cadbl(Text1)) + "*'")
          If Not rst.EOF Then MsgBox " Extracost compartit amb: " + atrim(rst!comanda) + " i " + atrim(rst!comandesafectades), vbInformation, "Informació"
    End If
    v = InputBox("Escriu la observació d'aquest PVP." + vbNewLine + "Equesta observació surtirà al moment de firmar PVP.", "Observació PVP", v)
    If StrPtr(v) = 0 Then Exit Sub
    vextracost = InputBox("Si hi ha extracost de transport escriu l'import aqui." + vbNewLine + "Equesta observació surtirà al moment de firmar PVP.", "Observació PVP i Extracost", vextracost)
    dbtmp.Execute "delete * from comandes_observacioPVP where comanda=" + atrim(cadbl(Text1))
    If v <> "" Or cadbl(vextracost) > 0 Then
        If cadbl(vextracost) > 0 Then
            vliniaalbara = InputBox("Escriu la descripció que vols que surti a l'albarà:", "Descripció linia albarà", vliniaalbara)
            vliniaalbara = Mid(vliniaalbara, 1, 100)
            If vliniaalbara <> "" Then
                 vcomandesafectades = InputBox("Escriu les comandes afectades separades per comes:", "Comandes afectades", vcomandesafectades)
                 If InStr(1, vcomandesafectades, Text1) = 0 Then vcomandesafectades = Text1 + IIf(vcomandesafectades <> "", "," + vcomandesafectades, "")
            End If
        End If
        dbtmp.Execute "insert into comandes_observacioPVP (comanda,observacio,extracost,liniaalbara,comandesafectades) values (" + atrim(Text1) + ",'" + treure_apostruf(atrim(v)) + "'," + Trim(cadbl(vextracost)) + ",'" + atrim(treure_apostruf(vliniaalbara)) + "','" + Trim(vcomandesafectades) + "')"
    End If
    carregar_observacioPVP
    Set rst = Nothing
End Sub
Sub canviar_refinplacsa(vrefantiga As String)
  Dim v As String
  If LCase(InputBoxEx("Escriu la contrasenya per modificar la referencia d'inplacsa.", "Contrasenya", , , , , , SPassword)) <> "inplacsa" Then Exit Sub
  If Data1.Recordset.EditMode > 0 Then Data1.Recordset.CancelUpdate
  v = InputBox("Entra la nova referencia d'inplacsa." + vbNewLine + "SI VOLS ELIMINAR-LA ESCRIU UN ESPAI", "Nova referencia", vrefantiga)
  If v <> "" And v <> vrefantiga Then
      dbtmp.Execute "update comandes_extres set refinplacsa='" + atrim(v) + "' where comanda=" + atrim(Data1.Recordset!comanda)
      If Data1.Recordset!linkcomanda1 <> 0 Then dbtmp.Execute "update comandes_extres set refinplacsa='" + v + "' where comanda=" + atrim(Data1.Recordset!linkcomanda1)
      If Data1.Recordset!linkcomanda2 <> 0 Then dbtmp.Execute "update comandes_extres set refinplacsa='" + v + "' where comanda=" + atrim(Data1.Recordset!linkcomanda2)
      dbtmp.Execute "insert into tarifes_referencies (codiclient,refinplacsa) values (" + atrim(Data1.Recordset!client) + ",'" + v + "')"
      dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','RefInplacsa','" + atrim(vrefantiga) + "','" + atrim(v) + "')"
      Data1.Recordset.Move 0
      hihaalgugravant 2
      activaronocampsimpresio False
      enabled_campscontrolcodiinplacsa True
  End If
  
End Sub
Sub canviar_lestatdelacomanda()
  Dim vnouestat As String
  Dim vnumc As String
  vnumc = cadbl(Data1.Recordset!comanda)
  vnouestat = atrim(UCase(InputBox("Entra el nou estat de la comanda." + Chr(10) + "Si vols passar a no muntada i no impresa escriu 'M'.", "Canvi d'estat", atrim(Data1.Recordset!proximaseccio))))
  If vnouestat = "" Then Exit Sub
  If InStr(1, "EMILRSVPT", vnouestat) <> 0 And Len(vnouestat) = 1 Then
      If vnouestat = "M" Then
         dbbaixes.Execute "update muntadoratot set acabada=false where comanda=" + atrim(vnumc)
         vnouestat = "I"
      End If
      Data1.Recordset!proximaseccio = vnouestat
      usuari_guarda_registre
       Else: MsgBox "Aquest estat de comanda no existeix", vbCritical, "Error"
  End If
End Sub
Sub ensenyar_refaltervatives()
  Dim rstref As Recordset
  Dim rr As String
  Dim pos As Byte
  Dim nomtaulatmp As String
  nomtaulatmp = "tmp_refclialt" ' + nomordinador
  If Not existeixlataula(cami, nomtaulatmp) Then
    dbtmp.Execute "create table " + nomtaulatmp + " (referencia text(255),nomusuari text(50),estat text(10))"
   Else: dbtmp.Execute "delete * from " + nomtaulatmp + " where nomusuari='" + nomordinador + "'"
  End If
  r = Text32(1)
  If InStr(1, r, "|") = 0 And r <> "" Then r = r + "|"
  If InStr(1, r, "|") = 0 Then r = "": rr = "no"
  While r <> ""
    pos = InStr(1, r, "|")
    If pos = 0 Then pos = 254
    rr = Mid(r, 1, pos - 1)
    If atrim(rr) > "" And atrim(rr) <> Text44 Then dbtmp.Execute "insert into " + nomtaulatmp + " (referencia,nomusuari,estat) values ('" + atrim(rr) + "','" + nomordinador + "','INACTIVA')"
    r = Mid(r, pos + 1)
  Wend
  If Text44 <> "" Then dbtmp.Execute "insert into " + nomtaulatmp + " (referencia,nomusuari,estat) values ('" + atrim(Text44) + "','" + nomordinador + "','INACTIVA')"
  If rr <> "no" Then
   mantenimentreferencies nomtaulatmp
   Set rstref = dbtmp.OpenRecordset("select * from " + nomtaulatmp + " where nomusuari='" + nomordinador + "'")
   While Not rstref.EOF
    If r <> "" Then r = r + " | "
    r = r + atrim(rstref!referencia)
    rstref.MoveNext
   Wend
   Text32(1) = r
  End If
  Set rstref = Nothing
  DoEvents
  'On Error Resume Next
  'dbtmp.Execute "drop table " + nomtaulatmp
  dbtmp.Execute "delete * from " + nomtaulatmp + " where nomusuari='" + nomordinador + "'"

End Sub

Sub mantenimentreferencies(nomtaulatmp As String)
  Dim vrefanterior As String
  vrefanterior = Text44
  Load formaltarep
  formaltarep.Caption = "Manteniment de Referències"
  formaltarep.Tag = ""
  formaltarep.lvariable = nomordinador
  formaltarep.alta.Tag = "nomusuari"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from " + nomtaulatmp + " where nomusuari='" + nomordinador + "' order by estat"
  formaltarep.Command1.Visible = True
  formaltarep.bactiva.Visible = True
 ' formaltarep.Width = formaltarep.Width * 2
  'formaltarep.DBGrid1.Width = formaltarep.DBGrid1.Width * 2
  DoEvents
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(0).Width = 2000
  formaltarep.DBGrid1.Columns(1).Visible = False
  formaltarep.DBGrid1.Columns(2).Visible = False
  formaltarep.DBGrid1.Columns(2).Width = 1000
'  formaltarep.DBGrid1.Columns(3).Width = 200
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.AllowUpdate = True
  'formaltarep.DBGrid1.Columns(2).Button = True
  actualitzar_tarifes_referencies formaltarep.Data1.Recordset, False
  formaltarep.Data1.Recordset.Sort = "estat"
  formaltarep.Data1.Recordset.FindFirst "referencia='" + Text4 + "'"
  formaltarep.Show 1
  If seleccioret = 1 Then
     If Not formaltarep.Data1.Recordset.EOF Then
      If atrim(formaltarep.Data1.Recordset!referencia) <> "" Then
        If Len(atrim(formaltarep.Data1.Recordset!referencia)) < 64 Then
            Text44 = formaltarep.Data1.Recordset!referencia
            If vrefanterior <> Text44 Then
                'If MsgBox("Vols passar totes les referencies anterior com a inactives?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
                '    formaltarep.Data1.Recordset.MoveFirst
                '    While Not formaltarep.Data1.Recordset.EOF
                '      If atrim(formaltarep.Data1.Recordset!referencia) <> Text44 Then
                '          'desactivar les referencies
                '           formaltarep.Data1.Recordset.Edit
                '           formaltarep.Data1.Recordset!estat = "INACTIVA"
                '           formaltarep.Data1.Recordset.Update
                '      End If
                '      formaltarep.Data1.Recordset.MoveNext
                '    Wend
                'End If
            End If
           Else: MsgBox "Aquesta referencia es massa gran per aquet camp", vbCritical, "Atenció"
        End If
      End If
     End If
  End If
  If seleccioret > 0 Then
     actualitzar_tarifes_referencies formaltarep.Data1.Recordset, True
  End If
  Unload formaltarep
End Sub
Sub actualitzar_tarifes_referencies(rst As Recordset, vtaulareferencies As Boolean)
  Dim rst2 As Recordset
  Dim vinactivarst As Boolean
  If rst.EOF Then Exit Sub
  rst.MoveFirst
  If vtaulareferencies Then
    While Not rst.EOF
        dbtmp.Execute "update tarifes_referencies set inactiva=" + IIf(rst!estat = "INACTIVA", "True", "False") + " where refclient='" + atrim(formaltarep.Data1.Recordset!referencia) + "' and codiclient='" + Command1(9).Tag + "'"
        rst.MoveNext
    Wend
     Else
      Set rst2 = dbtmp.OpenRecordset("select * from tarifes_referencies " + " where codiclient='" + Command1(9).Tag + "'")
      rst.MoveFirst
      While Not rst.EOF
          rst2.FindFirst "refclient='" + atrim(rst!referencia) + "'"
          If Not rst2.NoMatch Then
             vinactivarst = IIf(rst!estat = "INACTIVA", True, False)
             If vinactivarst <> rst2!inactiva Then
              rst.Edit
              rst!estat = IIf(rst2!inactiva, "INACTIVA", "ACTIVA")
              rst.Update
             End If
              Else
                dbtmp.Execute "insert into tarifes_referencies (codiclient,refclient,inactiva) values ('" + Command1(9).Tag + "','" + atrim(rst!referencia) + "',false)"
          End If
          rst.MoveNext
      Wend
  End If
  Set rst2 = Nothing
End Sub
Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
 r = obre_fitxer(ruta_relativa_docs, 2)
 Text97 = Mid(r, Len(ruta_relativa_docs) + 2)
 Text97.SetFocus
End Sub

Private Sub Command5_Click()
 r = obre_fitxer(ruta_relativa_docs, 2)
 Text109 = Mid(r, Len(ruta_relativa_docs) + 2)
 Text109.SetFocus
End Sub

Private Sub Command6_Click()
'  MsgBox atrim(data1.Recordset!obsimp1)
'MsgBox
  If Menu.nomusuari <> "Usr_JM" Then
   r = InputBox("Entra la contrasenya per entrar:", "Atenció")
   If LCase(r) <> "inplacsa" Then MsgBox "Has d'entrar la contrasenya correcte per poder veure l'etiqueta.": Exit Sub
  End If
  numcomanda = Data1.Recordset!comanda
  Unload comprovaretrebo
  Load comprovaretrebo
  comprovaretrebo.linia.BackColor = Command6.BackColor
  comprovaretrebo.Show 1
  If atrim(comprovaretrebo.Tag) <> "" Then
     Data1.Recordset.Edit
     Data1.Recordset!etrebvistiplau = IIf(comprovaretrebo.Tag = "OK", True, False)
     Data1.Recordset.Update
  End If
End Sub

Private Sub Command7_Click()
 r = obre_fitxer(ruta_relativa_docs, 2)
 Text111 = Mid(r, Len(ruta_relativa_docs) + 2)
 Text111.SetFocus
End Sub
Function comprovar_risc_comanda() As Byte
comprovar_risc_comanda = 0
If Label1(17).ForeColor = QBColor(9) Then
  If MsgBox("Aquest no se li ha demanat RISC , Vols fer-la igualment?", vbCritical + vbYesNo + vbDefaultButton2, "A T E N C I O") = vbNo Then comprovar_risc_comanda = 1
End If
If Label1(17).ForeColor <> &H80000012 And comprovar_risc_comanda <> 1 Then
     If MsgBox("Aquest Client ha superat el RISC o està IMPAGAT, NO ES POT FER COMANDA, Vols fer-la igualment?", vbCritical + vbYesNo + vbDefaultButton2, "A T E N C I O") = vbNo Then comprovar_risc_comanda = 1
  End If
  
End Function
Private Sub Command8_Click()
  Dim rsttmpdup As Recordset
  Dim rsttmpdup2 As Recordset
  Dim rsttmpdup3 As Recordset
  Dim rsttmpd As Recordset
  Dim esimpresa As Boolean
  Dim ultim As String
  Dim sistema As Byte
  Dim dies As Byte
  Dim tres As Double
  Dim dos As Double
  Dim un As Double
  Dim comandessel As String
  Dim i As Double
  un = 0: dos = 0: tres = 0: sistema = 1
  If Data1.Recordset!producte = "PC" Or Data1.Recordset!producte = "PC2" Then MsgBox "No pots duplicar un PC o PC2 has de sel.lecionar el material imprès", vbCritical, "Atenció": Exit Sub
  If Text32(5).BackColor <> &H6BEBB1 Then MsgBox "ATENCIÓ AQUESTA REFERENCIA ESTÀ MARCADA COM A NO ACTIVA." + vbNewLine + "ASSEGURA'T QUE VOLS DUPLICAR AQUESTA COMANDA.", vbCritical, "ATENCIÓ"
  'comprovo que els numeros de comanda siguin correlatius
  If cadbl(Data1.Recordset!linkcomanda1) > 0 And (Data1.Recordset!comanda + 1) <> cadbl(Data1.Recordset!linkcomanda1) Then MsgBox "Els numeros de comanda complexes no son correlatius no es pot duplicar.", vbCritical, "Error": Exit Sub
  If cadbl(Data1.Recordset!linkcomanda2) > 0 And (Data1.Recordset!comanda + 2) <> cadbl(Data1.Recordset!linkcomanda2) Then MsgBox "Els numeros de comanda complexes no son correlatius no es pot duplicar.", vbCritical, "Error": Exit Sub
  '------------------------
  If comprovar_risc_comanda = 1 Then Exit Sub
  If atrim(Data1.Recordset!proximaseccio) <> "T" Then
    If MsgBox("Aquesta comanda no està entregada. Voleu duplicar-la igualment?", vbCritical + vbYesNo, "Atenció") <> vbYes Then Exit Sub
  End If
  'If MsgBox("Segur que vols Duplicar aquesta commanda?", 64 + 4, "Atenció") = vbNo Then Exit Sub
  If InStr(1, ruta, "I") > 0 Then esimpresa = True
  'If esimpresa Then
    Unload formduplicarcomanda
    formduplicarcomanda.Show 1
    DoEvents
    If formduplicarcomanda.Tag = "sortir" Then Exit Sub
  'End If
  
  'areadatos.Visible = False
  duplicant = True
  If hihaalgugravant Then MsgBox "Hi ha algú gravant ara mateix, espera un moment i torna-ho a provar.", vbCritical, "Atenció": GoTo fi
  hihaalgugravant 1
  ensenyarframementreduplica True
  'ratoli "espera"
  carregar_controlscampsalicia True
  'valors per possar els recordset a no trobat
  Set rsttmpdup2 = dbtmp.OpenRecordset("select * from comandes where comanda=-99999")
  Set rsttmpdup3 = dbtmp.OpenRecordset("select * from comandes where comanda=-99999")
  '-----
  
  Set rsttmpdup = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(Data1.Recordset!comanda)))
  If cadbl(Data1.Recordset!linkcomanda1) > 0 Then Set rsttmpdup2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(Data1.Recordset!linkcomanda1)))
  If cadbl(Data1.Recordset!linkcomanda2) > 0 Then
    Set rsttmpdup3 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(Data1.Recordset!linkcomanda2)))
    If Not rsttmpdup3.EOF Then
       If cadbl(rsttmpdup3!lotmatdesb1) = cadbl(Data1.Recordset!linkcomanda1) Or cadbl(rsttmpdup3!lotmatdesb2) = cadbl(Data1.Recordset!linkcomanda1) Then
           sistema = 2
       End If
    End If
  End If
  
  'sistema = 1 'posso aixó perque el sistema al final serà sempre l'1
  If Not rsttmpdup.EOF Then
     ensenyarframementreduplica True
     passar_dades_registre_nou rsttmpdup, ultim, esimpresa
     'wait 2
     un = cadbl(ultim)
     'passo email si hi havia risc superat al crear la comanda nova
     enviaremailriscsuperat un, formduplicarcomanda.codiclient.Tag
     
     i = 0
     Do
      Set rsttmpdup = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(un)))
      i = i + 1
     Loop Until Not rsttmpdup.EOF Or i > 200000
     If i > 2000000 Then MsgBox "Error duplicant la comanda1, la base de dades ha tardat massa a crear el registre." + Chr(10) + "BORREU LA COMANDA CREADA I TORNEU A PROVAR EL PROCES.", vbCritical, "ERROR": GoTo fi
     comandessel = "comanda=" + ultim
     ensenyarframementreduplica True
     If Not rsttmpdup2.EOF Then
        passar_dades_registre_nou rsttmpdup2, ultim, esimpresa
      '  wait 3
        dos = cadbl(ultim)
        If un + 1 <> dos Then MsgBox "Error duplicant el producte complexe PC" + Chr(10) + " El numero de comanda no era correlatiu.", vbCritical, "Error": GoTo fi
        i = 0
        Do
         Set rsttmpdup2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(dos)))
         i = i + 1
        Loop Until Not rsttmpdup2.EOF Or i > 200000
        If i > 2000000 Then MsgBox "Error duplicant la comanda2, la base de dades ha tardat massa a crear el registre." + Chr(10) + "BORREU LA COMANDA CREADA I TORNEU A PROVAR EL PROCES.", vbCritical, "ERROR": GoTo fi
        
        comandessel = comandessel + " or comanda=" + ultim
     End If
     ensenyarframementreduplica True
     If Not rsttmpdup3.EOF Then
        passar_dades_registre_nou rsttmpdup3, ultim, esimpresa
       ' wait 3
        tres = cadbl(ultim)
        If un + 2 <> tres Then MsgBox "Error duplicant el producte complexe PC2" + Chr(10) + " El numero de comanda no era correlatiu.", vbCritical, "Error": GoTo fi
        i = 0
        'While rsttmpdup.EOF And i < 100000
        Do
          Set rsttmpdup3 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(tres)))
          i = i + 1
        Loop Until Not rsttmpdup3.EOF Or i > 200000
        If i > 2000000 Then MsgBox "Error duplicant la comanda3, la base de dades ha tardat massa a crear el registre." + Chr(10) + "BORREU LA COMANDA CREADA I TORNEU A PROVAR EL PROCES.", vbCritical, "ERROR": GoTo fi
        comandessel = comandessel + " or comanda=" + ultim
       End If
     ensenyarframementreduplica True
     'If Not rsttmpdup.EOF Then rsttmpdup.Edit: rsttmpdup!linkcomanda1 = dos: rsttmpdup!linkcomanda2 = tres: rsttmpdup!lotmatdesb1 = un: rsttmpdup!lotmatdesb2 = dos: rsttmpdup.Update
     'If Not rsttmpdup2.EOF Then rsttmpdup2.Edit: rsttmpdup2!linkcomanda1 = un: rsttmpdup2!linkcomanda2 = tres: rsttmpdup2!lotmatdesb1 = un: rsttmpdup2.Update
     'If Not rsttmpdup3.EOF Then rsttmpdup3.Edit: rsttmpdup3!linkcomanda1 = un: rsttmpdup3!linkcomanda2 = dos: rsttmpdup3!lotmatdesb1 = dos: rsttmpdup3!lotmatdesb2 = tres: rsttmpdup3.Update
     rsttmpdup.Close: rsttmpdup2.Close: rsttmpdup3.Close
     Set rsttmpdup = Nothing
     Set rsttmpdup2 = Nothing
     Set rsttmpdup3 = Nothing
     actualitzar_linkscomplexes un, dos, tres, sistema
     'wait 1
     Data1.RecordSource = "select * from comandes where " + comandessel + " order by comanda ASC"
     refrescar
     DoEvents
     'wait 2
     'Data1.Recordset.FindFirst ("comanda=" + ultim)
     modificar_Click
     ensenyarframementreduplica True
     carregar_controlscampsalicia True
     'wait (1)
     
     DoEvents
     If InStr(1, ruta, "I") > 0 Then Command1_Click 3
     Command23_Click
  End If
fi:
  areadatos.Visible = True
  Frame1(0).Enabled = True
   duplicant = False
   ensenyarframementreduplica False
   formcomandes.SetFocus
   ratoli "normal"
   Unload formduplicarcomanda
   hihaalgugravant 2
   Set rsttmpdup = Nothing
   Set rsttmpdup2 = Nothing
   Set rsttmpdup3 = Nothing
   Set rsttmpd = Nothing
End Sub
Sub enviaremailriscsuperat(vnumc As Double, vmsg As String)
  If atrim(vmsg) = "" Then Exit Sub
  enviaremailgeneric "jmiralles@inplacsa.com;eruscalleda@inplacsa.com", "Comanda creada amb crèdit superat  " + IIf(vnumc > 0, atrim(vnumc), ""), vmsg
End Sub
Sub actualitzar_linkscomplexes(un As Double, dos As Double, tres As Double, sistema As Byte)
 Dim s As String
  If sistema = 1 Then
   If un > 0 Then
    s = "linkcomanda1 = " + atrim(cadbl(dos)) + ",linkcomanda2 = " + atrim(cadbl(tres)) + ", lotmatdesb1 = " + atrim(cadbl(un)) + ",lotmatdesb2 = " + atrim(cadbl(dos))
    dbtmp.Execute "update comandes set " + s + " where comanda=" + atrim(cadbl(un))
   End If
   If dos > 0 Then
    s = "linkcomanda1 = " + atrim(cadbl(un)) + ",linkcomanda2 = " + atrim(cadbl(tres)) + ", lotmatdesb1 = " + atrim(cadbl(un)) + ",lotmatdesb2 = " + atrim(cadbl(un))
    dbtmp.Execute "update comandes set " + s + " where comanda=" + atrim(cadbl(dos))
   End If
   
   If tres > 0 Then
    s = "linkcomanda1 = " + atrim(cadbl(un)) + ",linkcomanda2 = " + atrim(cadbl(dos)) + ", lotmatdesb1 = " + atrim(cadbl(un)) + ",lotmatdesb2 = " + atrim(cadbl(tres))
    dbtmp.Execute "update comandes set " + s + " where comanda=" + atrim(cadbl(tres))
   End If
 End If
 If sistema = 2 Then
   If un > 0 Then
    s = "linkcomanda1 = " + atrim(cadbl(dos)) + ",linkcomanda2 = " + atrim(cadbl(tres)) + ", lotmatdesb1 = " + atrim(cadbl(un)) + ",lotmatdesb2 = " + atrim(cadbl(tres))
    dbtmp.Execute "update comandes set " + s + " where comanda=" + atrim(cadbl(un))
   End If
   If dos > 0 Then
    s = "linkcomanda1 = " + atrim(cadbl(un)) + ",linkcomanda2 = " + atrim(cadbl(tres)) + ", lotmatdesb1 = " + atrim(cadbl(un)) + ",lotmatdesb2 = " + atrim(cadbl(un))
    dbtmp.Execute "update comandes set " + s + " where comanda=" + atrim(cadbl(dos))
   End If
   
   If tres > 0 Then
    s = "linkcomanda1 = " + atrim(cadbl(un)) + ",linkcomanda2 = " + atrim(cadbl(dos)) + ", lotmatdesb1 = " + atrim(cadbl(dos)) + ",lotmatdesb2 = " + atrim(cadbl(tres))
    dbtmp.Execute "update comandes set " + s + " where comanda=" + atrim(cadbl(tres))
   End If
 End If
End Sub
Sub ensenyarframementreduplica(ensenyar As Boolean)
'  formscrooll.Visible = Not ensenyar
  If ensenyar Then
   Load avis
   avis.missatge = "Duplicant la comanda espera fins que acabi la duplicació..."
   avis.Caption = "Duplicant..."
   avis.Command1.Visible = False
   
   avis.Show
     Else: Unload avis
  End If
   
   
End Sub
Sub passar_dades_registre_nou(rsttmpdup As Recordset, ultim As String, seccioimpresores As Boolean)
     Dim rsttmpd As Recordset
     Dim rutaproducte As String
     Dim rstextra As Recordset
     Dim rstenvio As Recordset
     Dim vobservacionsalbaraperclient As String
     Command8.Tag = "duplicant"
     alta_registre
     For i = 1 To Data1.Recordset.Fields.Count
      If Data1.Recordset.Fields(i - 1).Name <> "comanda" Then
        Data1.Recordset.Fields(i - 1) = Null
      End If
      DoEvents
     Next i
     'canvio els camps que vull que siguin diferents al registre duplicat
     Set rsttmpd = dbtmp.OpenRecordset("select dies,ruta from productes where codi='" + atrim(rsttmpdup.Fields!producte) + "'", , dbReadOnly)
     If Not rsttmpd.EOF Then
        dies = rsttmpd!dies
        rutaproducte = rsttmpd!ruta
      Else: dies = 0
     End If
     dies = 0 'hem tret el calcul de dies perque no es cumplien mai
     For i = 1 To rsttmpdup.Fields.Count
      If rsttmpdup.Fields(i - 1).Name <> "comanda" Then
       Data1.Recordset.Fields(i - 1) = rsttmpdup.Fields(i - 1)
      End If
     Next i
     'Data1.Recordset.Update
'     Data1.Recordset.Edit
     'miro la comisió del client
     Data1.Recordset.Fields!com_representant = Null
     Set rsttmpd = dbtmp.OpenRecordset("select com_representant,fix_com_rep,clientvindraarevisarimpresio from clients where codi=" + atrim(rsttmpdup.Fields!client) + "", , dbReadOnly)
     If Not rsttmpd.EOF Then
         If rsttmpd!fix_com_rep Then Data1.Recordset.Fields!com_representant = cadbl(rsttmpd!com_representant)
     End If
     'poso les dades del client de formduplicarcomanda
     Set rstenvio = dbtmp.OpenRecordset("select observacionscomandaalalbara from clients_envios where id=" + atrim(cadbl(formduplicarcomanda.direccioenvio.Tag)))
     Set rstextra = dbtmp.OpenRecordset("select * from comandes_extres where comanda=" + atrim(rsttmpdup!comanda), , dbReadOnly)
     Data1.Recordset.Fields!client = cadbl(formduplicarcomanda.codiclient)
     Data1.Recordset.Fields!direnvio = cadbl(formduplicarcomanda.direccioenvio.Tag)
     If Not rstenvio.EOF Then vobservacionsalbaraperclient = atrim(rstenvio!observacionscomandaalalbara)
     dbtmp.Execute "update comandes_extres set observacionsalbara='" + vobservacionsalbaraperclient + "'" + " where comanda=" + atrim(cadbl(Text1.Text))
     dbtmp.Execute "update comandes_Extres set materialexacte=" + IIf(rstextra!materialexacte, "True", "False") + " where comanda=" + atrim(cadbl(Text1.Text))
     dbtmp.Execute "update comandes_Extres set codicomptable=" + atrim(cadbl(formduplicarcomanda.codicomptable.Tag)) + " where comanda=" + atrim(cadbl(Text1.Text))
     dbtmp.Execute "update comandes_Extres set carametall='" + atrim(rstextra!carametall) + "' where comanda=" + atrim(cadbl(Text1.Text))
     dbtmp.Execute "update comandes_Extres set tipusmaterialcanutureb='" + atrim(rstextra!tipusmaterialcanutureb) + "' where comanda=" + atrim(cadbl(Text1.Text))
     dbtmp.Execute "update comandes_Extres set noplanificable=false where comanda=" + atrim(cadbl(Text1.Text))
     dbtmp.Execute "update comandes_Extres set comandaduplicadade=" + atrim(cadbl(rsttmpdup!comanda)) + " where comanda=" + atrim(cadbl(Text1.Text))
     
    ' If seccioimpresores And InStr(1, rutaproducte, "I") > 0 Then

          
          'trec el numero de treball de la comanda si es nou producte
          
          If formduplicarcomanda.canviproducte.Tag = "1" Then
              Data1.Recordset.Fields!numtreball = Null
              Data1.Recordset.Fields!numordremodificacio = Null
              Data1.Recordset.Fields!refclialt = Null
              Data1.Recordset.Fields!refclient = Null
              Data1.Recordset.Fields!impressio = "N"
              Data1.Recordset.Fields!marques = "No"
              Data1.Recordset.Fields!texteimpressio = ""
              Data1.Recordset!arxiuimpressora = ""
              Data1.Recordset!arxiupdf = ""
              dbtmp.Execute "update comandes_Extres set clientvindraarevisarimpresio=" + atrim(rsttmpd!clientvindraarevisarimpresio) + " where comanda=" + atrim(cadbl(Text1.Text))
              netejarcampsblaus Data1.Recordset
          End If
          If formduplicarcomanda.repetir.Tag = "1" Then
            'si es repetida
            Data1.Recordset.Fields!impressio = "R"
            Data1.Recordset.Fields!marques = "No"
            
          ' si hi ha canvi d'envio
           If cadbl(formduplicarcomanda.direccioenvio.Tag) <> cadbl(rsttmpdup!direnvio) Then
             Data1.Recordset.Fields!impressio = "M"
             Data1.Recordset.Fields!marques = "Si"
             dbtmp.Execute "update comandes_Extres set aviscanvisambeltreball='Duplicada i canvi envio' where comanda=" + atrim(cadbl(Data1.Recordset!comanda))
           End If
           'miro si es modificada
           If cadbl(rsttmpdup!numordremodificacio) > 0 And (cadbl(rsttmpdup!numordremodificacio) < modificaciodeltreballmesgran(cadbl(rsttmpdup!numtreball))) Then
             Data1.Recordset.Fields!impressio = "M"
             Data1.Recordset.Fields!numordremodificacio = modificaciodeltreballmesgran(cadbl(rsttmpdup!numtreball))
             dbtmp.Execute "update comandes_Extres set aviscanvisambeltreball='Modificada al duplicar' where comanda=" + atrim(cadbl(Data1.Recordset!comanda))
             dbtmp.Execute "update comandes_Extres set clientvindraarevisarimpresio=" + atrim(cabool(rsttmpd!clientvindraarevisarimpresio)) + " where comanda=" + atrim(cadbl(Text1.Text))
           End If
          End If
          crearclientvinculat Data1.Recordset
   '  End If
     Data1.Recordset.Fields!datacomanda = Date
     Data1.Recordset.Fields!dataentrega = Null
     Data1.Recordset.Fields!dataactivacio = Null
     Data1.Recordset.Fields!datapreu = Null
     Data1.Recordset.Fields!datamaterial = Null
     Data1.Recordset.Fields!linkcomanda1 = Null
     Data1.Recordset.Fields!linkcomanda2 = Null
     Data1.Recordset.Fields!proximaseccio = "E"
     Data1.Recordset.Fields!seccioactual = " "
     Data1.Recordset.Fields!obspedgen1 = Null
     'data1.Recordset.Fields!obsext1 = Null
     'data1.Recordset.Fields!obsimp1 = Null
     Data1.Recordset.Fields!obsimp2 = Null
     'data1.Recordset.Fields!obslam1 = Null
     Data1.Recordset.Fields!obslam2 = Null
     'data1.Recordset.Fields!obsreb1 = Null
     'data1.Recordset.Fields!obsreb2 = Null
     'data1.Recordset.Fields!obssol1 = Null
     Data1.Recordset.Fields!obssol2 = Null
     Data1.Recordset.Fields!mesuracantex = Null
     Data1.Recordset.Fields!cantitatex = Null
     Data1.Recordset.Fields!tubbaseext = Null
     Data1.Recordset.Fields!cantitatsol = Null
     Data1.Recordset.Fields!pvp = 0
     Data1.Recordset.Fields!pvpdolar = 0
     Data1.Recordset.Fields!pes1000mtrs = 0
     Data1.Recordset.Fields!numpressupost = Null
     'data1.Recordset.Fields!kghora = 0
     Data1.Recordset!etrebvistiplau = False
    ' data1.Recordset.Update
     'ultim = data1.Recordset!comanda
     ultim = Text1.Text
     'If InStr(1, rutaproducte, "I") = 0 Then
     '   data1.Recordset!numtreball = 0
     '   data1.Recordset!numordretreball = 0
     '   data1.Recordset!texteimpresio = ""
     'End If
     Data1.Recordset!proximaseccio = "E"
     Data1.Recordset!seccioactual = " "

     Data1.Recordset.Update
     passar_accessoris_soldadores Data1.Recordset!comanda, rsttmpdup!comanda
     Data1.Recordset.Move 0
     Data1.Recordset.Edit
     gravar_registre
     

End Sub
Sub passar_accessoris_soldadores(vnumc_desti As Double, vnumc_original As Double)
   Dim rst As Recordset
   Dim rstaccessori As Recordset
   Dim rstdesti As Recordset
   dbbaixes.Execute "delete * from soldadores_accessorisutilitzats where comanda=" + atrim(vnumc_desti)
   Set rstdesti = dbbaixes.OpenRecordset("select * from soldadores_accessorisutilitzats")
   Set rstaccessori = dbtmp.OpenRecordset("select * from accessoris_soldadora")
   Set rst = dbbaixes.OpenRecordset("select * from soldadores_accessorisutilitzats where comanda=" + atrim(vnumc_original))
   While Not rst.EOF
      rstdesti.AddNew
      rstdesti!nomaccessori = rst!nomaccessori
      rstdesti!idaccessori = rst!idaccessori
      rstaccessori.FindFirst "numaccessori=" + atrim(rst!idaccessori)
      If Not rstaccessori.NoMatch Then If cabool(rstaccessori!control_traçabilitat) Then rstdesti!lottraçabilitat = "-"
      rstdesti!comanda = vnumc_desti
      rstdesti.Update
      rst.MoveNext
   Wend
   Set rst = Nothing
   Set rstdesti = Nothing
End Sub
Function cabool(valor As Variant) As String
'On Error Resume Next
   If atrim(valor) = "Verdadero" Then valor = True
   If atrim(valor) = "Falso" Then valor = False
   If IsNull(valor) Or atrim(valor) = "" Then
      cabool = "False"
        Else: cabool = IIf(valor, "True", "False")
   End If
End Function
Function possarelcodimuntadora(id_treball As Double) As String
   Dim vmuntadora As String
   If id_treball = 0 Then Exit Function
   vmuntadora = buscarcodimuntadora(cadbl(id_treball))
   If vmuntadora = "" Then vmuntadora = Format(id_treball, "00000000")
   dbclixes.Execute "update clientsvinculats set codimuntadora='" + atrim(vmuntadora) + "' where id_Treball=" + atrim(id_treball)
   possarelcodimuntadora = vmuntadora
End Function
Function buscarcodimuntadora(idtreball As Double) As String
    Dim rst As Recordset
    Set rst = dbclixes.OpenRecordset("select codimuntadora from clientsvinculats where id_Treball=" + atrim(idtreball) + " order by codimuntadora Desc")
    If rst.EOF Then Exit Function
    buscarcodimuntadora = atrim(rst!codimuntadora)
End Function
Function crearclientvinculat(rstc As Recordset) As String
   Dim rst As Recordset
   Dim rstcli As Recordset
   Dim numordre As Integer
   Dim vimpresio As String
   Dim vmuntadora As String
   Dim numc As Double
   numordre = cadbl(rstc!numordremodificacio)
   If numordre = 0 Then numordre = 1
   Set rst = dbclixes.OpenRecordset("select * from clientsvinculats where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordremodificacio=" + atrim(numordre) + " and direnvio=" + atrim(cadbl(rstc!direnvio)))
   If (cadbl(rstc!numtreball) > 0 And cadbl(rstc!numordremodificacio) > 1) Then vimpresio = "M"
   If rst.EOF Then
        Set rstcli = dbtmp.OpenRecordset("SELECT Clients_envios.id, clients.nom, Clients_envios.poblacioe FROM Clients_envios INNER JOIN clients ON Clients_envios.codi = clients.codi WHERE Clients_envios.id=" + atrim(cadbl(rstc!direnvio)) + ";")
        If rstcli.EOF Then Exit Function 'sinotrobo el client també surtu
        rst.AddNew
        rst!id_treball = cadbl(rstc!numtreball)
        rst!ordremodificacio = numordre
        rst!codiclient = rstc!client
        rst!direnvio = rstc!direnvio
        rst!nomclient = atrim(rstcli!nom)
        rst!nomdirenvio = atrim(rstcli!poblacioe)
        rst!codimuntadora = atrim(rstc!arxiumontadora)
        rst!refclient = atrim(rstc!refclient)
        rst!refclientalternatives = atrim(rstc!refclialt)
        rst!arxiuimp = False
        rst.Update
          Else
             If cadbl(rstc!numtreball) = 0 Or cadbl(rstc!numordremodificacio) = 1 Then vimpresio = "N"
             If Not rst.EOF Then
                If atrim(rst!arxiuimp) <> "" Then
                    vimpresio = "R"
                   Else: vimpresio = "M"
                End If
             End If
             
   End If
   If vimpresio = "" Then vimpresio = atrim(rstc!impressio)
   vmuntadora = possarelcodimuntadora(cadbl(rstc!numtreball))
   If atrim(rstc!arxiumontadora) = "" Then
      'dbtmp.Execute "update comandes set arxiumontadora='" + vmuntadora + "' where comanda=" + atrim(rstc!comanda)
      If rstc.EditMode > 0 Then
          rstc!arxiumontadora = vmuntadora
           Else
            numc = rstc!comanda
            rstc.Edit
            rstc!arxiumontadora = vmuntadora
            rstc.Update
            rstc.FindFirst "comanda=" + atrim(numc)
      End If
   End If
   Set rst = Nothing
   Set rstcli = Nothing
   crearclientvinculat = vimpresio
   
End Function
Sub llistademodificacions()
  Unload formseleccio
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from comandes_controlcanvis where comanda=" + atrim(Data1.Recordset!comanda) + " order by data "
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(1).Width = 1000
  formseleccio.DBGrid2.Columns(2).Width = 1200
  formseleccio.DBGrid2.Columns(3).Width = 1700
  formseleccio.DBGrid2.Columns(4).Width = 1700
  formseleccio.DBGrid2.Columns(5).Width = 1500
  formseleccio.DBGrid2.Columns(6).Width = 1500
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(2).NumberFormat = "dd/mm/yy hh:nn"
  formseleccio.Width = 10000
  If formseleccio.Data1.Recordset.EOF Then Exit Sub
  formseleccio.Show 1
  Unload formseleccio
End Sub
Sub netejarcampsblaus(taula As Recordset)
 For Each objecte In formcomandes
      If TypeOf objecte Is MaskEdBox Or TypeOf objecte Is TextBox Or TypeOf objecte Is ComboBox Then
          If objecte.HelpContextID = 99 Then
              If objecte.DataField <> "" And objecte.DataField <> "refinplacsa" Then taula.Fields(objecte.DataField) = Null
          End If
      End If
   Next objecte
End Sub

Sub modificar_comandaclient_crops()
   formentradacomandacrops.Show 1
End Sub
Sub comprovarcomandaitreball()
  Dim numc As Double
  numc = cadbl(Data1.Recordset!comanda)
  mirardiferenciescomandaitreball numc
  imprimirdiferenciescomandaitreball numc, "P"
End Sub
Sub possar_numero_bossasoldadores(Optional vforçarimprimirVQ As Boolean)
   Dim rst As Recordset
   Dim vnumbossa As String
   Set rst = dbtmp.OpenRecordset("select * from comandesmesextres where comanda=" + atrim(Text1))
   If Not rst.EOF Then
       If atrim(rst!numerobossasoldadores) = "" Then
          'si comença per C es perquè es anonim, si es imprés es farà desde clixesnous a l'EVA quan passi el treball a clixes entrats
            If Mid(rst!refinplacsa + " ", 1, 1) = "C" Then
             vnumbossa = atrim(rst!refinplacsa)
               Else
                  'si els clixes ja estan entrats possaré el numero de treball
                  If InStr(1, estatdelclixe(cadbl(rst!numtreball), cadbl(rst!numordremodificacio)), "CLIXES ENTRATS") > 0 Then
                    vnumbossa = atrim(rst!numtreball)
                     Else: GoTo fi
                  End If
            End If
            rst.Edit
            rst!numerobossasoldadores = rst!refinplacsa
            rst.Update
            Label1(173) = vnumbossa
            'com que no tenia posat el numero vol dir que ha d'imprimir també la VQ per soldadores
            imprimir_VQ_soldadores rst
          Else: If vforçarimprimirVQ Then imprimir_VQ_soldadores rst
       End If
   End If
fi:
   Set rst = Nothing
End Sub
Sub imprimir_VQ_soldadores(rst As Recordset)
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.report
  Dim vnumtreballiversio As String
  Dim vtexteimpresio As String
  If Mid(atrim(rst!numerobossasoldadores) + " ", 1, 1) <> "C" Then
       vnumtreballiversio = atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio)
       vtexteimpresio = atrim(rst!marcailinia)
  End If
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "verificacioqualitatVQsoldadores.rpt", 1)
  oreport.FormulaFields.GetItemByName("numbossa").Text = "'NºBossa: " + atrim(rst!numerobossasoldadores) + "'"
  oreport.FormulaFields.GetItemByName("nomclient").Text = "'" + treure_apostruf(atrim(rst!nomclient)) + "'"
  oreport.FormulaFields.GetItemByName("treballversio").Text = "'" + atrim(vnumtreballiversio) + "'"
  oreport.FormulaFields.GetItemByName("texteimpresio").Text = "'" + atrim(vtexteimpresio) + "'"
  oreport.FormulaFields.GetItemByName("refclient").Text = "'" + atrim(rst!refclient) + "'"
  oreport.FormulaFields.GetItemByName("comanda").Text = "'" + atrim(rst!comanda) + "'"
   oreport.DiscardSavedData
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   If Not existeix("c:\ordprog.ini") Then
          veurereport.CRViewer.PrintReport
          oreport.PrintOut False
         Else:
           veurereport.CRViewer.ViewReport
           veurereport.WindowState = 2
          veurereport.Show 1
   End If
   
End Sub
Sub imprimirbossasoldadores(Optional vforçarimprimirBossa As Boolean)
    Dim rst As Recordset
    If Not larutahiha(atrim(Data1.Recordset!producte), "S") Then Exit Sub
    possar_numero_bossasoldadores
    If Label1(173) <> "" Then
       Set rst = dbtmp.OpenRecordset("select numerobossasoldadores from comandes_extres where numerobossasoldadores='" + atrim(Label1(173)) + "' and comanda<>" + atrim(Data1.Recordset!comanda))
       If rst.EOF Or vforçarimprimirBossa Then
           Shell llegir_ini("General", "rutallistats", "comandes.ini") + "ClixesNous.exe comandes.ini 0 imprimirbossasoldadores " + atrim(Text1), vbHide
       End If
       Set rst = Nothing
    End If
End Sub

Private Sub Command9_Click(Index As Integer)
 Static dins As Boolean
 Dim visexp As Byte
 Dim comanda0 As Double
 Dim comanda1 As Double
 Dim comanda2 As Double
 Dim contador As Byte
 Dim vobrirdisposiciodematerials As Boolean
 Dim vcomandaimpresa As Boolean
 Dim vpreguntarferVQ As Boolean
 Dim rste As Recordset
 If Index = 9 Then FormExtresSoldadores.Show 1: Exit Sub
 If Index = 8 Then
   If isloaded("formfirmes") Then Unload formfirmes
   formfirmes.Show
   carregar_firmes
   Exit Sub
 End If
 
 If Command9(0).BackColor = &HC0FFC0 Then vcomandaimpresa = True
 If Index = 7 Then
  If Data1.Recordset.EditMode > 0 Then MsgBox "No es pot imprimir editant la comanda.", vbCritical, "Error": Exit Sub
  If Label1(173) <> "" Then vpreguntarferVQ = True
  imprimirbossasoldadores True
  If vpreguntarferVQ Then
     If MsgBox("Vols imprimir també la VQ de Soldadores?", vbInformation + vbDefaultButton2 + vbYesNo, "VQ?") = vbYes Then possar_numero_bossasoldadores True
  End If
  Exit Sub
 End If
 If Index = 3 Then possarelpreu:    Exit Sub
 If Index = 4 Then formdesactivades.Show 1: Exit Sub
 If Index = 6 Then formstopped.Show 1: Exit Sub
 If Index = 5 Then modificar_comandaclient_crops: Exit Sub
 comanda0 = cadbl(Data1.Recordset!comanda)
 comanda1 = cadbl(Data1.Recordset!linkcomanda1)
 comanda2 = cadbl(Data1.Recordset!linkcomanda2)
 If Not dins Then
     dins = True
    Else: Exit Sub
  End If
 If Index = 1 Then GoTo dia
 If Index = 2 Then llistademodificacions: GoTo fi
 visexp = 1
 If llegir_ini("General", "exportant", fitxerini) = "1" Then GoTo exportant
 'comprovo que si es complexa hi hagi el linkcomanda1 o linkcomanda2
   If InStr(1, Text3, "PC") Or InStr(1, ruta, "L") Then
      If cadbl(text77(11)) = 0 And cadbl(text77(12)) = 0 Then
         MsgBox "No pots imprimir un complexa sense possar la comanda que fa referencia..."
         GoTo fi
      End If
   End If
 '.....
 'comprovo que abans d'imprimir la fulla complexa generin la referencia d'inplacsa
 If InStr(1, Text3, "PC") And atrim(Text32(5)) = "" Then
     MsgBox "Abans d'imprimir la fulla complexa has d'imprimir la principal per generar la referència d'inplacsa", vbCritical, "Atenció": GoTo fi
 End If
 '....
 'miro si el treball i la comanda estan igual
 comprovarcomandaitreball
 ' trec aquesta verificació perque ningu sap que fa... 22/07/25
      'If comprovarquelaverificacioanteriorestafeta = False Then dins = False: Exit Sub
 seleccioimpresio.Show 1
 If seleccioimpresio.Tag = "" Then dins = False: Exit Sub
 
 'If (cadbl(llegir_ini("General", "programador", fitxerini)) = 1 Or previprint = 1) And Index <> 1 Then
 '    If MsgBox("Vols visualitzar el full d'expedicions?", vbInformation + vbYesNo, "Atenció") <> vbYes Then visexp = 0
 'End If
  ratoli "espera"
  
  If Index = 0 Then
    If seleccioimpresio.imprimir.Item(0) Then
        If cadbl(Text33) = 0 Then MsgBox "Compte el camp de pes 1000 peces està a zero.", vbCritical, "Error"
exportant:
        r = ""
        If seleccioimpresio.Checkimps.Value = 1 Then contador = imprimirtotselsarxius: wait (2)
        llistar_comanda False
        contador = contador + 1
        If llegir_ini("General", "exportant", fitxerini) <> "1" Then
          wait (2)
          If Label1(173) = "" Then
            imprimirbossasoldadores
          End If
          llistar_comanda True
        End If
        
        If cadbl(comanda1) > 0 Then
         carregant = True
         Data1.RecordSource = "select * from comandes where comanda=" + atrim(comanda1)
         Data1.Refresh
         esperarlacarrega 'wait (2)
         
         llistar_comanda False, atrim(comanda1)
         contador = contador + 1
        End If
        If cadbl(comanda2) > 0 Then
         carregant = True
         Data1.RecordSource = "select * from comandes where comanda=" + atrim(comanda2)
         Data1.Refresh
         esperarlacarrega 'wait (2)
         llistar_comanda False, atrim(comanda2)
         contador = contador + 1
        End If
        If comanda1 > 0 Then
         carregant = True
         Data1.RecordSource = "select * from comandes where comanda=" + atrim(comanda0)
         Data1.Refresh
         esperarlacarrega
        End If
        If llegir_ini("General", "exportant", fitxerini) = "1" Then
           esperarqueeltotaldepdf contador
           ratoli "normal": Exit Sub
             Else: vobrirdisposiciodematerials = True
        End If
    End If
    If seleccioimpresio.imprimir.Item(1) Or seleccioimpresio.imprimir.Item(2) Then
        llistar_comanda False
    End If
    If seleccioimpresio.imprimir.Item(3) Then
        llistar_comandes_pendentsdepassaraimpresores
    End If
    wait 2
  End If
dia:
  If Index = 1 Then
      ratoli "espera"
      imprimir_diaadia
  End If
 
fi:
  Set rste = Data1.Database.OpenRecordset("select * from comandes_extres where comanda=" + atrim(cadbl(Text1)))
  If Not rste.EOF Then If Not vcomandaimpresa And rste!comandaimpresa Then enviarmissatgesdelacomanda cadbl(Text1)
  dins = False
  ratoli "normal"
  If vobrirdisposiciodematerials Then
      Set rste = dbtmp.OpenRecordset("select * from referencies_disposiciomaterials where refinplacsa='" + Text32(5) + "'")
      If rste.EOF Then
          obrir_disposicio_materials
      End If
  End If
  Set rste = Nothing
End Sub
Sub enviarmissatgesdelacomanda(vnumc As Double)
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim vmsg As String
   Dim vdescclient As String
   Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
   If rst.EOF Then GoTo fi
   Set rstc = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(rst!client))
   If Not rstc.EOF Then vdescclient = atrim(rst!client) + " - " + atrim(rstc!nom)
   vmsg = "Comanda: " + atrim(vnumc) + "   " + vdescclient + Chr(13) + Chr(10) + "Referència del client: " + atrim(rst!refclient) + Chr(13) + Chr(10) + "Text impresió: " + atrim(rst!marcailinia) + Chr(13) + Chr(10) + "Missatge: " + atrim(rst!obsreb2)
   'avismissatgeclientarebobinadores
   If atrim(rst!obsreb2) <> "" Then enviaremailgeneric "avismissatgeclientarebobinadores", "Comanda Nova Ref:" + atrim(rst!refclient) + " --> Avís per a rebobinadores.", vmsg
fi:
   Set rst = Nothing
End Sub
Function comprovarquelaverificacioanteriorestafeta() As Boolean
 Dim vsql As String
 Dim vtreball As String
 Dim vmodificacio As String
 Dim rst As Recordset
 Set rst = dbtmp.OpenRecordset("SELECT comandes.*, InStr(1,[ruta],'I') AS hihaimpresora FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi WHERE comanda=" + atrim(Data1.Recordset!comanda))
 If rst.EOF Then comprovarquelaverificacioanteriorestafeta = True: Exit Function
 If rst!hihaimpresora = 0 Then comprovarquelaverificacioanteriorestafeta = True: Exit Function
 vtreball = atrim(cadbl(Data1.Recordset!numtreball))
 vmodificacio = atrim(cadbl(Data1.Recordset!numordremodificacio))
 vsql = "SELECT comandes.numtreball, comandes.numordremodificacio, impresores_aniloxos.comanda, impresores_aniloxos.okcanvi "
 vsql = vsql + " FROM impresores_aniloxos INNER JOIN comandes ON impresores_aniloxos.comanda = comandes.comanda "
 vsql = vsql + " WHERE (((comandes.numtreball)=" + vtreball + ") AND ((comandes.numordremodificacio)=" + vmodificacio + ") AND ((impresores_aniloxos.okcanvi)=1));"
 Set rst = dbbaixes.OpenRecordset(vsql)
 'Clipboard.Clear
 'Clipboard.SetText vsql
 If Not rst.EOF Then
      vtreball = UCase(InputBox("El treball d'aquesta comanda encara està pendent de verificar per l'impresor de la comanda " + atrim(rst!comanda) + Chr(10) + "Aviseu-lo que ho faci per evitar ERRORS A FABRICACIÓ." + Chr(10) + "Escriu [OK] per acceptar i continuar la impresió.", "ATENCIÓ"))
      If vtreball = "OK" Then comprovarquelaverificacioanteriorestafeta = True
       Else: comprovarquelaverificacioanteriorestafeta = True
 End If
 
 Set rst = Nothing
End Function
Sub llistar_comandes_pendentsdepassaraimpresores()
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim rsttintes As Recordset
  Dim dbenvio As Database
  Set dbenvio = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
  Set rst = dbtmp.OpenRecordset("select * from comandes_extres where passaraimpresores=1 order by comanda")
  While Not rst.EOF
     'imprimir la primera fulla de la comanda
     llistar_comanda False, atrim(rst!comanda), True
     Set rst2 = dbtmp.OpenRecordset("SELECT comandes.comanda, clients.nom, comandes.marcailinia,comandes.numtreball,comandes.numordremodificacio,comandes.rebmtrs,comandes.cantitatex FROM (comandes INNER JOIN clients ON comandes.client = clients.codi) INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda WHERE (((comandes.comanda)=" + atrim(rst!comanda) + "));")
     If Not rst2.EOF Then
        vmetres = cadbl(rst2!rebmtrs)
        If vmetres = 0 Then cadbl (rst2!cantitatex)
        msg = justificar(atrim(rst2!comanda), 8) + justificar(Format(vmetres, "#,##0") + "mtrs", 15) + justificar(atrim(rst2![nom]), 40) + "   " + atrim(rst2!marcailinia)
        If atrim(msg) <> "" Then dbenvio.Execute "insert into envios_mails_linies (id_envio,descripcio) values (0,'" + treure_apostruf(msg) + "')"
        Set rsttintes = dbclixesnous.OpenRecordset("select * from tintes where id_treball=" + atrim(cadbl(rst2!numtreball)) + " and (ordremodificacio=" + atrim(cadbl(rst2!numordremodificacio)) + " or ordremodificacio=" + atrim(cadbl(rst2!numordremodificacio) * -1) + ")")
        While Not rsttintes.EOF
            If InStr(1, atrim(rsttintes!color), "PRIMAR") = 0 And atrim(rsttintes!color) <> "" Then
               dbenvio.Execute "insert into envios_mails_linies (id_envio,descripcio) values (0,'" + Chr(9) + treure_apostruf(atrim(rsttintes!color)) + "')"
            End If
            rsttintes.MoveNext
        Wend
        dbenvio.Execute "insert into envios_mails_linies (id_envio,descripcio) values (0,'_________________________________________________________')"
        dbenvio.Execute "insert into envios_mails_linies (id_envio,descripcio) values (0,'                                                         ')"
     End If
     dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(rst!comanda) + ",'" + nomordinador + "','ImprimirFullaImpresores','','')"
     rst.MoveNext
  Wend
  'If msg <> "" Then enviaremailgeneric "miquel.inplacsa@gmail.com", "Noves comandes a punt per imprimir.", "Relació de Comandes confirmades per imprimir."
  If msg <> "" Then enviaremailgeneric "tintes@inplacsa.com;impresores@inplacsa.com", "Noves comandes a punt per imprimir.", "Relació de Comandes confirmades per imprimir."
  dbtmp.Execute "update comandes_extres set passaraimpresores=2 where passaraimpresores=1"
  Set rst = Nothing
  Set rsttintes = Nothing
End Sub

Sub borrartaulestemp()
  On Error Resume Next
  Kill Environ("temp") + "\~llistdad*.*"
End Sub
Sub esperarlacarrega()
   While carregant
     DoEvents
   Wend
End Sub
Sub esperarqueeltotaldepdf(c As Byte)
   Dim cintern As Byte
   Dim d As String
   Dim horainici As Date
   horainici = Now
   While c <> cintern And DateDiff("n", horainici, Now) < 2
     d = Dir("c:\temp\exportar\*.pdf")
     cintern = 0
     While d <> ""
       cintern = cintern + 1
       d = Dir
     Wend
     DoEvents
   Wend
End Sub
Sub imprimir_diaadia()
Dim v As String
 Dim dbtmp As Database
 Dim taulatemp As String
 Dim camps As String
 Dim dies As Double
 Dim inici As Date
 Dim fi As Date
 Dim where As String
 Dim db As Database
 Dim rst As Recordset
 Dim rstc As Recordset
 Dim xx As Variant
 Dim subconsulta As String
 Dim tipusllistat As String
 Dim vmetres As Double
 Dim vkilos As Double
 Dim k As Double
 Dim e As Double
 Dim eurus1 As Double
 Dim eurus2 As Double
 Dim eurus3 As Double
 Dim agrupades As Boolean
 Dim vprofunditat As String
 'where = querywhere
 'where = Mid(where, 1, 6)
 ratoli "espera"
 borrartaulestemp
 taulatemp = Environ("temp") + "\~llistdad" + Format(Now, "ddmmhhnnss") + ".mdb"
 'On Error Resume Next
 v = InputBox("Entra la data comanda d'inici: ", "Llistat Dia a Dia", Format(DateAdd("d", -7, Now), "dd/mm/yy"))
 If v = "" Then Exit Sub
 inici = v
 v = InputBox("Entra la data comanda d'acavament: ", "Llistat Dia a Dia", Format(Now, "dd/mm/yy"))
 If v = "" Then Exit Sub
 fi = v
demanarllistat:
 tipusllistat = InputBox("Vols llistat de comandes (R)RECEPCIONADES O (E)ENTREGADES.", "TIPUS DE LLSITAT", "R")
 If UCase(tipusllistat) <> "R" And UCase(tipusllistat) <> "E" Then MsgBox "No s'ha escullit cap llistat.", vbCritical, "Error": Exit Sub
 If tipusllistat = "e" Then
   If MsgBox("Has possat una e minúscula es correcte?", vbCritical + vbYesNo + vbDefaultButton1, "Atenció") = vbNo Then GoTo demanarllistat
 End If
 'If UCase(tipusllistat) = "R" Then
   If MsgBox("Vols veure els kilos agrupats per clients?", vbInformation + vbYesNo, "Tipus llistat") = vbYes Then agrupades = True
 'End If
 'On Error GoTo 0
 If inici = 0 Or fi = 0 Then Exit Sub
 If Not IsDate(inici) Or Not IsDate(fi) Then MsgBox "Comandes mal entrades": Exit Sub
 Set dbtmp = OpenDatabase(cami)
 If existeix(taulatemp) Then Kill taulatemp
 DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
 
 'subconsulta = "SELECT comandes.comanda FROM comandes INNER JOIN bobinesent ON comandes.comanda = bobinesent.comanda WHERE (((bobinesent.data) Between #" + Format(inici, "mm/dd/yy") + "# And #" + Format(fi, "mm/dd/yy") + "#)) GROUP BY comandes.comanda;"
 subconsulta = "SELECT bobinesent.comanda into tmp_diaadia From bobinesent WHERE (((bobinesent.data) Between #" + Format(inici, "mm/dd/yy") + "# And #" + Format(fi, "mm/dd/yy") + "#)) GROUP BY bobinesent.comanda;"
 camps = "datacomanda,comandes.comanda,client,' ' as nomclient,0.0 as kilos,0 as metres,0.0 as kilosmat,0.0 as eurusmat,0.0 as eurokg,' ' as mesuraquant, producte,' ' as descproducte,cantitatsol,cantitatex,mesuracantex,rebmtrs,rebkilos "
 
 If tipusllistat = "R" Then
 '(linkcomanda1=null or laminadora<>null) and
     where = " (datacomanda between #" & (Format(inici, "mm/dd/yy")) & "# and #" & (Format(fi, "mm/dd/yy")) & "#) and producte<>'PC' and producte<>'PC2' and producte<>'PCP' " ' and rebmtrs>0"
     Set rst = dbtmp.OpenRecordset("select " + camps + " from comandes  where" + where, , dbReadOnly)
    Else:
      crear_tmp_diaadia subconsulta
'      where = " comanda in (" + subconsulta + ")"
      Set rst = dbtmp.OpenRecordset("select " + camps + " from tmp_diaadia LEFT JOIN comandes ON tmp_diaadia.comanda = comandes.comanda", , dbReadOnly)
 End If
 
 
 'Set rst = dbtmp.OpenRecordset("select " + camps + " from comandes  where" + where, , dbReadOnly)
 dbtmp.Execute ("select " + camps + " into temporal in '" + taulatemp + "' from comandes  where comanda=0") '  + where)
 'dbtmp.Execute ("select " + camps + " into temporal in '" + taulatemp + "' from comandes  where" + where)
 Me.Caption = "seleccio feta"
 Set db = OpenDatabase(taulatemp)
 Set rsttmp = db.OpenRecordset("temporal")
 Me.Caption = "obro els palets"
 Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb", , True)
 While Not rst.EOF
'    MsgBox "Gravant " + atrim(rst.PercentPosition) + " %"
    Me.Caption = "1   " + atrim(rst!comanda) + "   --->   " + atrim(rst.PercentPosition)
    rsttmp.AddNew
    For i = 0 To rst.Fields.Count - 1
     rsttmp.Fields(i) = rst.Fields(i)
    Next i
    rsttmp.Update
    rst.MoveNext
    DoEvents
 Wend
  Me.Caption = "obro temporal"
 Set rsttmp = db.OpenRecordset("temporal")
 While Not rsttmp.EOF
    DoEvents
    Me.Caption = "2    " + atrim(rsttmp!comanda) + "   --->   " + atrim(rsttmp.PercentPosition)
    rsttmp.Edit
    'nom producte
   Set rst = dbtmp.OpenRecordset("select descripcio from productes where codi='" + atrim((rsttmp!producte)) + "'")
   If Not rst.EOF Then rsttmp!descproducte = rst!descripcio
    ' nom client
    Set rst = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(rsttmp!client) + "")
   If Not rst.EOF Then rsttmp!nomclient = rst!nom
   If cadbl(rsttmp!rebmtrs) > 0 Then
       rsttmp!metres = rsttmp!rebmtrs
   End If
   If cadbl(rsttmp!rebkilos) > 0 Then
      rsttmp!kilos = rsttmp!rebkilos
   End If
  
   vprofunditat = UCase(llegir_ini("General", "profunditatllistatdiaadia", fitxerini))
   If vprofunditat = "{[}]" Then escriure_ini "General", "profunditatllistatdiaadia", IIf(Menu.nomusuari = "Usr_JM", "total", "parcial"), fitxerini
  DoEvents
   If UCase(tipusllistat) = "E" Then
     If tipusllistat = "e" Then
       If impresafora(rsttmp!comanda) Or laminatafora(rsttmp!comanda) Or soldatafora(rsttmp!comanda) Then rsttmp.Update: rsttmp.Delete: GoTo proxima
     End If
     vmetres = rsttmp!metres
     vkilos = rsttmp!kilos
     possarelsmetresikilos vmetres, vkilos, rsttmp!comanda, inici, fi
     rsttmp!metres = vmetres
     rsttmp!kilos = vkilos
     
     If vprofunditat = "TOTAL" Then
       Set rstc = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(rsttmp!comanda))
       k = 0
       e = 0
       eurus1 = 0
       eurus2 = 0
       eurus3 = 0
       If Not rstc.EOF Then
         calculareurokilodematerial rstc!comanda, k, e, eurus1
         calculareurokilodematerial rstc!linkcomanda1, k, e, eurus2
         calculareurokilodematerial rstc!linkcomanda2, k, e, eurus3
         rsttmp!eurusmat = eurus1 + eurus2 + eurus3
         rsttmp!kilosmat = k
         If k > 0 Then rsttmp!eurokg = rsttmp!eurusmat / k
       End If
       Set rstc = Nothing
     End If
   End If
   rsttmp.Update
proxima:
   rsttmp.MoveNext
 Wend
 Me.Caption = "Fora"
 subconsulta = "SELECT distinct bobinesent.data as ladata FROM comandes INNER JOIN bobinesent ON comandes.comanda = bobinesent.comanda GROUP BY bobinesent.data,comandes.comanda HAVING (((Last(bobinesent.data)) Between #" + Format(inici, "mm/dd/yy") + "# And #" + Format(fi, "mm/dd/yy") + "#)) order by bobinesent.data;"
 If tipusllistat <> "R" Then
    Set rst = dbtmp.OpenRecordset(subconsulta)
   Else: Set rst = db.OpenRecordset("select distinct datacomanda from temporal")
 End If
 
 db.Execute "delete * from temporal where comanda=0"
 If Not rst.EOF Then
    rst.MoveLast
    dies = rst.RecordCount
    Set rst = Nothing
 End If
 For i = 0 To 80
     llistat.Formulas(i) = ""
 Next i
 llistat.Formulas(0) = "inicifi=' Inici:" + atrim(inici) + " -- fi:" + atrim(fi) + "'"
 llistat.Formulas(1) = "totaldies=" + atrim(dies) '+ atrim(Format(DateDiff("d", inici, fi), "#,##0"))
 llistat.Formulas(2) = "titol='" + IIf(tipusllistat = "R", "RECEPCIONADES", "ENTREGADES") + "'"
 
 If vprofunditat = "TOTAL" Then
    llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatcomandesvalorades.rpt"
   Else: llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatcomandes.rpt"
 End If
 If agrupades Then
    If UCase(tipusllistat) = "R" Then llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatcomandesrecepcionadesagrupades.rpt"
    If UCase(tipusllistat) = "E" Then llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatcomandesentregadesagrupades.rpt"
 End If
 llistat.DataFiles(0) = taulatemp
 llistat.DiscardSavedData = True
 llistat.Destination = crptToWindow
 wait (2)
 llistat.Action = 1
 Set db = Nothing
 Set rsttmp = Nothing
 'Set dbstocks = Nothing
 Set dbtmp = Nothing
 Set rst = Nothing
 ratoli "normal"
End Sub
Sub crear_tmp_diaadia(v As String)
  On Error Resume Next
  dbtmp.Execute "drop table tmp_diaadia"
  dbtmp.Execute v
  On Error GoTo 0
End Sub
Function impresafora(numc As Double) As Boolean
 Dim rst As Recordset
 Dim rstm As Recordset
 Set rst = dbbaixes.OpenRecordset("select numeromaquina from impressores where comanda=" + atrim(numc))
 If Not rst.EOF Then
     Set rstm = dbtmp.OpenRecordset("select descripcio from maquines where maquina='I' and codi=" + atrim(rst!numeromaquina))
     If Not rstm.EOF Then
        If InStr(1, rstm!descripcio, "#") > 0 Then impresafora = True
     End If
 End If
 Set rst = Nothing
 Set rstm = Nothing
End Function
Function laminatafora(numc As Double) As Boolean
 Dim rst As Recordset
 Dim rstm As Recordset
 Set rst = dbbaixes.OpenRecordset("select numeromaquina from laminadores where comanda=" + atrim(numc))
 If Not rst.EOF Then
     Set rstm = dbtmp.OpenRecordset("select descripcio from maquines where maquina='L' and codi=" + atrim(rst!numeromaquina))
     If Not rstm.EOF Then
        If InStr(1, rstm!descripcio, "#") > 0 Then laminatafora = True
     End If
 End If
 Set rst = Nothing
 Set rstm = Nothing
End Function
Function soldatafora(numc As Double) As Boolean
 Dim rst As Recordset
 Dim rstm As Recordset
 Set rst = dbbaixes.OpenRecordset("select numeromaquina from soldadores where comanda=" + atrim(numc))
 If Not rst.EOF Then
     Set rstm = dbtmp.OpenRecordset("select descripcio from maquines where maquina='S' and codi=" + atrim(rst!numeromaquina))
     If Not rstm.EOF Then
        If InStr(1, rstm!descripcio, "#") > 0 Then soldatafora = True
     End If
 End If
 Set rst = Nothing
 Set rstm = Nothing
End Function
Sub calculareurokilodematerial(numc As Double, ByRef kilos As Double, ByRef eurus As Double, eurusc As Double)
   Dim rstp As Recordset
   Dim k As Double
   Dim e As Double
   eurusc = 0
   If numc < 100000 Then Exit Sub
   Set rstp = dbstocks.OpenRecordset("SELECT avg(Palets.preucompra) AS tpreucompra,count(*) as registres, First(Bobines.Mts) AS metresbobina, First(Bobines.kilos) AS kilosbobina, Parcials.comanda, Sum(Parcials.metres) AS metresgastats FROM Parcials INNER JOIN (Palets INNER JOIN Bobines ON Palets.Idpalet = Bobines.Idpalet) ON (Bobines.Idbobina = Parcials.idbobina) AND (Parcials.idpalet = Bobines.Idpalet) GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + atrim(numc) + "'));", , dbReadOnly)
   If Not rstp.EOF Then
      If cadbl(rstp!metresbobina) > 0 Then k = cadbl(rstp!metresgastats) * cadbl(rstp!kilosbobina) / cadbl(rstp!metresbobina)
      eurusc = cadbl(rstp!tpreucompra) * k
      e = cadbl(rstp!tpreucompra)
   End If
   kilos = kilos + k
   eurus = eurus + e
   Set rstp = Nothing
End Sub
Sub possarelsmetresikilos(ByRef metres As Double, ByRef kilos As Double, numc As Double, inici As Date, fi As Date)
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select sum(metresisacs) as metres,sum(kilosiunitats) as kilos from bobinesent where data between #" + Format(inici, "mm/dd/yy") + "# and #" + Format(fi, "mm/dd/yy") + "# and entregat='S' and comanda=" + atrim(numc))
   If Not rst.EOF Then
       If cadbl(rst!kilos) > 0 Then
          metres = cadbl(rst!metres)
          kilos = cadbl(rst!kilos)
         Else
           metres = cadbl(rst!metres)
            Set rst = dbtmp.OpenRecordset("SELECT comandes.*,COMANDES_EXTRES.solpesgrmcm2 FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda where comandes.comanda = " + atrim(numc))
            If Not rst.EOF Then
              kilos = calcularpesxrpeça(rst, cadbl(rst!solpesgrmcm2)) * metres
            End If
       End If
   End If
   Set rst = Nothing
End Sub
Function imprimirtotselsarxius() As Byte
   Dim generarfitxer_imp As String
   Dim ordremodificacio As Double
   Dim contador As Byte
    'exportarcomandes = True
    r = ""
    Set rsttmp = dbtmp.OpenRecordset("select arxiuult,arxiuexp from clients_envios where id=" + atrim(cadbl(Data1.Recordset!direnvio)))
    If Not rsttmp.EOF Then
      If Len(Trim(rsttmp!arxiuult)) > 5 Then
        imprimir_word3 r + ruta_relativa_docs + "\" + carpeta(rsttmp!arxiuult, cadbl(Text2.Text)), True
        
        contador = contador + 1
      End If
      If Len(Trim(rsttmp!arxiuexp)) > 5 Then
         imprimir_word3 r + ruta_relativa_docs + "\" + carpeta(rsttmp!arxiuexp, cadbl(Text2.Text)), True
         
         contador = contador + 1
      End If
    End If
    ordremodificacio = cadbl(Data1.Recordset!numordremodificacio)
    If ordremodificacio = 0 Then ordremodificacio = 1
    generarfitxer_imp = ruta_documentacio_clixes + "\" + Format(cadbl(Data1.Recordset!numtreball), "00000") + "\IMP" + Format(cadbl(Data1.Recordset!numtreball), "00000") + "-" + Format(ordremodificacio, "000") + "-" + Format(cadbl(Data1.Recordset!client), "000000") + "_" + atrim(cadbl(Data1.Recordset!direnvio)) + ".doc"
    If Not existeix(generarfitxer_imp) Then generarfitxer_imp = generarfitxer_imp + "x"
    'If Text42.Container.Visible And Len(Trim(Text42)) > 5 Then imprimir_word3 r + ruta_relativa_docs + "\" + carpeta(Text42, cadbl(Text2.Text)), True: contador = contador + 1
    'If Len(Trim(Text79)) > 5 Then imprimir_word3 r + ruta_relativa_docs + "\" + carpeta(Text79, cadbl(Text2.Text)), True
    If larutahiha(atrim(Data1.Recordset!producte), "I") Then
      If atrim(Data1.Recordset!impressio) = "R" Or (estatdelclixe(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio)) = "CLIXES ENTRATS") Then
        If existeix(generarfitxer_imp) Then
           'MsgBox "Intentant imprimir el IMP " + vbNewLine + generarfitxer_imp
           'poso una espera de 2 segons perquè a en Ricard hi ha vegades que no li imprimeix l'IMP
           'AMB EL MSGBOX D'ABANS HA FUNCIONAT PERO AMB L'ESPERA A VEURE SI N'HI HA PROU
            wait 2
           imprimir_word3 generarfitxer_imp, True: contador = contador + 1
        End If
      End If
    End If
    If Text97.Container.Visible And Len(Trim(Text97)) > 5 Then imprimir_word3 r + ruta_relativa_docs + "\" + carpeta(Text97, cadbl(Text2.Text)), True: contador = contador + 1
    If Text111.Container.Visible And Len(Trim(Text111)) > 5 Then imprimir_word3 r + ruta_relativa_docs + "\" + carpeta(Text111, cadbl(Text2.Text)), True: contador = contador + 1
    If Text109.Container.Visible And Len(Trim(Text109)) > 5 Then imprimir_word3 r + ruta_relativa_docs + "\" + carpeta(Text109, cadbl(Text2.Text)), True: contador = contador + 1
    'If Text35.Container.Name And Len(Trim(Text35)) > 5 Then imprimir_word3 r + ruta_relativa_docs + "\" + carpeta(Text35, cadbl(Text2.Text)), True: contador = contador + 1
    imprimirtotselsarxius = contador
End Function
Sub imprimir_word3(vnomfitxer As String, Optional vVeurel As Boolean)
    Dim vimp As String
    Dim v As String
   ' vnomfitxer = nomfitxer
    If InStr(1, LCase(vnomfitxer), ".lnk") > 0 Then MsgBox "Hi ha relacionat un archiu d'acces directe enlloc d'assignar un WORD, arregla aquesta direcció d'enviament i torna-ho a imprimir.", vbCritical, "Atenció"
    If llegir_ini("General", "exportant", fitxerini) <> "1" Then
     If InStr(1, LCase(vnomfitxer), ".doc") > 0 And InStr(1, LCase(vnomfitxer), ".docx") = 0 And Not existeix(vnomfitxer + "x") Then
       '  If Not existeix("c:\temp\docx") Then MkDir "c:\temp\docx"
       '  v = "c:\temp\docx\" + Format(Now, "yymmddhhnnss") + ".doc"
       '  FileCopy vnomfitxer, v
       '  guardar_doc_a_docx v
       '  If existeix(v) Then vnomfitxer = v
       MsgBox "El fitxer word es de la versió anterior primer s'ha de convertir", vbCritical, "Atenció"
       obrir_document vnomfitxer
       MsgBox "FES ACCEPTAR QUAN HAGIS CANVIAT EL FORMAT DEL FITXER.", vbExclamation, "ATENCIÓ"
       If Not existeix(vnomfitxer + "x") Then
             MsgBox "No s'ha trobat el fitxer convertit.", vbCritical, "Error": Exit Sub
              Else: Kill vnomfitxer
       End If
     End If
    End If
    
    vimp = llegir_ini("General", "segonaimpresoradecomandes", fitxerini)
    If vimp = "{[}]" Then vimp = ""
    If Not existeix(vnomfitxer) Then vnomfitxer = vnomfitxer + "x"
    PrintAnyDocument vnomfitxer, vimp
End Sub
Public Sub PrintAnyDocument(ByVal strPathFile As String, Optional vImpresoraonimprimir As String)
    Dim TargetFolder
    Dim FileName
    Dim ObjShell As Object
    Dim ObjFolder As Object
    Dim ObjItem As Object
    Dim ColItems As Object
    Dim vimpresoraactual As String
    Dim printerx As Printer
    
    vimpresoraactual = Printer.DeviceName
    If vImpresoraonimprimir <> "" Then Establecer_Impresora vImpresoraonimprimir
    If InStr(1, strPathFile, "\") <> 0 Then
        TargetFolder = rutadelfitxer(strPathFile)
        FileName = Right(strPathFile, Len(strPathFile) - Len(TargetFolder))
    End If
    Set ObjShell = CreateObject("Shell.Application")
    Set ObjFolder = ObjShell.NameSpace(TargetFolder)
    Set ColItems = ObjFolder.Items
    For Each ObjItem In ColItems
        If ObjItem.Name = FileName Then
            ObjItem.InvokeVerbEx ("Print")
            Exit For
        End If
    Next
    If vImpresoraonimprimir <> "" Then Establecer_Impresora vimpresoraactual
    Set ObjItem = Nothing
    Set ColItems = Nothing
    Set ObjFolder = Nothing
    Set ObjShell = Nothing
End Sub
Private Function Establecer_Impresora(ByVal NamePrinter As String) As Boolean
On Error GoTo errSub
      
    'Variable de referencia
    Dim obj_Impresora As Object
      
    'Creamos la referencia
    Set obj_Impresora = CreateObject("WScript.Network")
        obj_Impresora.setdefaultprinter NamePrinter
      
    Set obj_Impresora = Nothing
          
        'La función devuelve true y se cambió con éxito
        Establecer_Impresora = True
        'MsgBox "La impresora se cambió correctamente", vbInformation
    Exit Function
      
      
'Error al cambiar la impresora
errSub:
If err.Number = 0 Then Exit Function
   Establecer_Impresora = False
   MsgBox "error: " & err.Number & Chr(13) & "Description: " & err.Description
   On Error GoTo 0
End Function

Function carpeta(ruta, client) As String
  If cadbl(Mid(ruta, 1, 6)) = 0 Then ruta = numcarpetaclient + " " + Trim(ruta)
  carpeta = treure_apostruf(ruta)
End Function
Function nomfitxertemporal() As String
     nomfitxertemporal = "c:\temp\" + Format(Now, "ddmmhhnnss") + ".mdb"
End Function
Sub actualitzarcampstintes(rstc As Recordset, ByRef rsttemp2 As Recordset, ByRef rsttemp As Recordset)
   Dim i As Byte
   Dim rstclixe As Recordset
   Dim rstmodificacio As Recordset
   Dim rsttintes As Recordset
   Dim rstobstintes As Recordset
   Dim rstlink As Recordset
   Dim rstfoto As Recordset
   Dim treball As Integer
   Dim modificacio As Integer
   Dim dbtintes As Database
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb", , True)
   treball = cadbl(rstc!numtreball): modificacio = cadbl(rstc!numordremodificacio)
   If modificacio = 0 Then modificacio = 1
   Set rstclixe = dbclixesnous.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(rstc!numtreball)))
   If rstclixe.EOF Then Exit Sub
   Set rstmodificacio = dbclixesnous.OpenRecordset("select * from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(cadbl(modificacio)))
   If rstmodificacio.EOF Then Exit Sub
   Set rstfoto = dbclixes.OpenRecordset("select nomfotogravador from fotogravadors where codi=" + atrim(cadbl(rstmodificacio!fotograbador)))
   Set rsttintes = dbclixesnous.OpenRecordset("select * from tintes where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordremodificacio=" + atrim(cadbl(modificacio)) + " order by ordretinter")
   'rsttemp2.Edit
  ' rsttemp.Edit
   If Not rstfoto.EOF Then rsttemp2!nomfotograbador = atrim(rstfoto!nomfotogravador)
   rsttemp2!treballsistemaimpresio = rstmodificacio!sistemadimpresio
   While Not rsttintes.EOF
     Set rstlink = dbclixesnous.OpenRecordset("select * from tintes where id_tinter=" + IIf(cadbl(rsttintes!tinterlinkambid_treball) > 0, atrim(rsttintes!tinterlinkambid_treball), atrim(rsttintes!id_tinter)))
     If Not rstlink.EOF Then
        rsttemp2.Fields("lintreball" + atrim(cadbl(rsttintes!ordretinter))) = cadbl(rstlink.Fields!aniloxclixe)
        rsttemp2.Fields("densitattinta" + atrim(cadbl(rsttintes!ordretinter))) = cadbl(rstlink.Fields!densitatutilitzada)
        rsttemp2.Fields("volumtinta" + atrim(cadbl(rsttintes!ordretinter))) = cadbl(rstlink.Fields!volum)
        rsttemp2.Fields("detalltinter" + atrim(cadbl(rsttintes!ordretinter))) = atrim(rstlink.Fields!detalltinter) + IIf(atrim(rstlink.Fields!clixeosleeve) <> "" And atrim(rstlink.Fields!clixeosleeve) <> "Clixé", "[" + atrim(rstlink.Fields!clixeosleeve) + "]", "")
        rsttemp2.Fields("teextensiotinter" + atrim(cadbl(rsttintes!ordretinter))) = mirarsiteextensiofeta(cadbl(treball), CByte(modificacio), cadbl(rstlink.Fields!coditinta))
        rsttemp2.Fields("colortinter" + atrim(cadbl(rsttintes!ordretinter))) = QBColor(buscarcolorrecuadre(cadbl(rstlink!coditinta), dbtintes))
     End If
     rsttintes.MoveNext
   Wend
   
   'posso el reprint
   Set rsttintes = dbclixesnous.OpenRecordset("select * from tintes where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordremodificacio=" + atrim(cadbl(modificacio) * -1) + " order by ordretinter")
   rsttemp2!formaimpresioreprint = IIf(atrim(rstmodificacio!reprintformaimpres) = "N", "NORMAL", IIf(atrim(rstmodificacio!reprintformaimpres) = "T", "TRANSPARENT", ""))
   While Not rsttintes.EOF
       Set rstlink = dbclixesnous.OpenRecordset("select * from tintes where id_tinter=" + IIf(cadbl(rsttintes!tinterlinkambid_treball) > 0, atrim(rsttintes!tinterlinkambid_treball), atrim(rsttintes!id_tinter)))
       If Not rstlink.EOF Then
          rsttemp2.Fields("nomtintareprint" + atrim(cadbl(rsttintes!ordretinter))) = atrim(rstlink!color) + IIf(atrim(rstlink!detalltinter) <> "", "(" + atrim(rstlink!detalltinter) + ")", "")
          rsttemp2.Fields("desarrollreprint" + atrim(cadbl(rsttintes!ordretinter))) = atrim(rstlink!anilox)
          rsttemp2.Fields("densitattintareprint" + atrim(cadbl(rsttintes!ordretinter))) = cadbl(rstlink.Fields!densitatutilitzada)
          rsttemp2.Fields("volumtintareprint" + atrim(cadbl(rsttintes!ordretinter))) = cadbl(rstlink.Fields!volum)
          rsttemp2.Fields("lintreballreprint" + atrim(cadbl(rsttintes!ordretinter))) = cadbl(rstlink.Fields!aniloxclixe)
       End If
       rsttintes.MoveNext
   Wend
   
   '---------------------------------------------------
   
   
   
   Set rstobstintes = dbtmp.OpenRecordset("select * from comandes_observacionstintes where comanda=" + atrim(rstc!comanda) + " order by id")
   i = 1
   While Not rstobstintes.EOF And i < 3
     rsttemp2.Fields("obstinta" + atrim(i)) = atrim(rstobstintes!observacio)
     i = i + 1
     rstobstintes.MoveNext
   Wend
   possar_observaciotintesoperaris rsttemp2, rstc!numtreball
   
   Set rstclixe = Nothing
   Set rstmodificacio = Nothing
   Set rsttintes = Nothing
   Set rstfoto = Nothing
   Set rstobstintes = Nothing
   Set dbtintes = Nothing
End Sub
Function buscarcolorrecuadre(vcoditinta As Double, dbtintes As Database) As Double
   Dim vsql As String
   Dim rst As Recordset
   buscarcolorrecuadre = 15
   vsql = "SELECT subfamiliestintes.color, colorsetiquetes.codicolor FROM (tintes LEFT JOIN subfamiliestintes ON tintes.idsubfamilia = subfamiliestintes.codi) LEFT JOIN colorsetiquetes ON subfamiliestintes.color = colorsetiquetes.nomcolor where tintes.codi='" + atrim(vcoditinta) + "'"
   Set rst = dbtintes.OpenRecordset(vsql)
   If Not rst.EOF Then buscarcolorrecuadre = IIf(cadbl(rst!codicolor) = 0, 15, rst!codicolor)

End Function
Function mirarsiteextensiofeta(vnumtreball As Double, vnumordre As Byte, vcoditinta As Double) As String
  Dim rsttintes As Recordset
  Dim rstext As Recordset
  Set rsttintes = dbbaixes.OpenRecordset("select codiextensio from extensions_treballsrelacionats where numtreball=" + atrim(vnumtreball) + " and numordremodificacio=" + atrim(vnumordre) + " and coditinta=" + atrim(vcoditinta))
   If rsttintes.EOF Then
      mirarsiteextensiofeta = ""
       Else
        Set rstext = dbbaixes.OpenRecordset("select anilox,volum from extensions where codiextensio='" + atrim(rsttintes!codiextensio) + "'")
        mirarsiteextensiofeta = atrim(rsttintes!codiextensio)
        If Not rstext.EOF Then
           If cadbl(rstext!volum) > 0 Then mirarsiteextensiofeta = mirarsiteextensiofeta + "/" + atrim(cadbl(rstext!anilox)) + "-" + atrim(cadbl(rstext!volum))
        End If
   End If
  Set rsttintes = Nothing
End Function
Function mirarsihihaCingularReal(vnumtreball As Double, vordremodificacio As Double) As Boolean
   Dim vurl As String
   Dim generarfitxer_pdf As String
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(vnumtreball, "00000") + "\pdf" + Format(vnumtreball, "00000") + "-" + Format(vordremodificacio, "000") + "_CR.pdf"
   If existeix(generarfitxer_pdf) Then
      mirarsihihaCingularReal = True
   End If
End Function
Function convertiratexte(vcampmemo As String) As String
   Dim i As Integer
   Dim v As String
   Dim valorant As String
   For i = 1 To Len(vcampmemo)
     If Asc(Mid(vcampmemo, i, 1)) > 32 Then
       'If Mid(vcampmemo, i, 1) = "-" And valorant <> "-" Then
        v = v + Mid(vcampmemo, i, 1)
       ' valorant = Mid(vcampmemo, i, 1)
       'End If
     End If
   Next i
   convertiratexte = v
End Function
Sub possar_observaciotintesoperaris(rsttemp2 As Recordset, vnumtreball As Double)
   Dim rsto As Recordset
   Dim i As Byte
   Dim vtext As String
   Set rsto = dbbaixes.OpenRecordset("select * from idstreball where id=" + atrim(vnumtreball))
   If rsto.EOF Then Exit Sub
   vtext = convertiratexte(rsto!obsidtreball)
   vcar_repetit = buscar_car_repetit(vtext)
   While atrim(vcar_repetit) <> ""
      While InStr(1, vtext, vcar_repetit + vcar_repetit)
       substituir vtext, vcar_repetit + vcar_repetit, ""
      Wend
      vcar_repetit = buscar_car_repetit(vtext)
   Wend
   substituirtots vtext, Chr(13), " "
   substituirtots vtext, Chr(10), " "
   rsttemp2!obstintabaixes1 = Mid(atrim(vtext), 2, 80)
   rsttemp2!obstintabaixes2 = (Mid(atrim(vtext), 81, 80))
End Sub
Function buscar_car_repetit(vtext As String) As String
   Dim i As Integer
   Dim vultimcar As String
   For i = 1 To Len(vtext)
      If vultimcar = Mid(vtext, i, 1) And Not IsNumeric(vultimcar) Then buscar_car_repetit = vultimcar: Exit Function
      vultimcar = Mid(vtext, i, 1)
   Next i
   buscar_car_repetit = ""
End Function
Sub generarrefinplacsadefinitiu(numc As Double)
  Dim rst As Recordset

  Dim vclient As Double
  
  Set rst = dbtmp.OpenRecordset("SELECT comandes.*, InStr(1,[ruta],'I') AS hihaimpresora FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi WHERE comanda=" + atrim(numc))
  Set rst2 = dbtmp.OpenRecordset("select * from comandes_extres where comanda=" + atrim(numc))
  If rst.EOF Then Exit Sub
  If rst2.EOF Then Exit Sub
  If atrim(rst2!refinplacsa) <> "" Then Exit Sub
  vclient = cadbl(rst!client)
  comandesamblamateixareferenciainplacsa rst, IIf(rst!hihaimpresora > 0, True, False)
  Set rst = dbtmp.OpenRecordset("select refinplacsa from comandes_extres where comanda=" + atrim(rst!comanda))
  If Not rst.EOF Then gravar_refinplacsaSICAL atrim(rst!refinplacsa), vclient
  comprovarsihihaalgunacomandasemblantaordredimpresio numc
  comprovar_dades_client_ventes_3anys numc
  Set rst = Nothing
End Sub
Sub comprovar_dades_client_ventes_3anys(vnumc As Double)
Dim rst As Recordset
  Dim vcos As String
  Set rst = dbtmp.OpenRecordset("SELECT comandes.client from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
    Set rst = dbtmp.OpenRecordset("SELECT comandesmesextres.datacomanda, comandesmesextres.nomclient, comandesmesextres.client, comandesmesextres.codicomptable From comandesmesextres where (((comandesmesextres.client) = " + atrim(rst!client) + ")) ORDER BY comandesmesextres.datacomanda DESC;")
    If Not rst.EOF Then
        If DateDiff("yyyy", rst!datacomanda, Now) >= 3 Then
            vcos = "DADES DE LA COMANDA CREADA Nº: " + atrim(vnumc) + vbNewLine + vbNewLine
            vcos = vcos + "Client: " + atrim(rst!client) + "-" + atrim(rst!nomclient) + vbNewLine
            vcos = vcos + "Codi comptable: " + atrim(rst!codicomptable) + vbNewLine
            vcos = vcos + "Data ultima comanda: " + atrim(rst!datacomanda) + vbNewLine + vbNewLine
            vcos = vcos + "Revisar les dades d'aquest client el mes ràpid possible."
            enviaremailgeneric "recepcion@inplacsa.com;rgirones@inplacsa.com", "Nova comanda amb un client que fa mes de 3 anys que no compra. " + atrim(vnumc), vcos
        End If
    End If
  End If
fi:
  Set rst = Nothing
End Sub
Sub comprovarsihihaalgunacomandasemblantaordredimpresio(vnumc As Double)
  Dim rst As Recordset
  Dim vcos As String
  
     'la onana no en fa cas i ho trec
  Exit Sub
  Set rst = dbtmp.OpenRecordset("SELECT comandes.client,Modificacions.codidelinia, Modificacions.codideliniav FROM comandes LEFT JOIN Modificacions ON (comandes.numtreball = Modificacions.id_treball) AND (comandes.numordremodificacio = Modificacions.ordre) where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
    If rst!client = 6841 Or cadbl(rst!codidelinia) = 0 Then GoTo fi
    Set rst = dbtmp.OpenRecordset("SELECT impresores_ordreimpresio.comanda, Modificacions.codidelinia, Modificacions.codideliniav FROM (impresores_ordreimpresio LEFT JOIN comandes ON impresores_ordreimpresio.comanda = comandes.comanda) LEFT JOIN Modificacions ON (comandes.numordremodificacio = Modificacions.ordre) AND (comandes.numtreball = Modificacions.id_treball) where Modificacions.codidelinia=" + atrim(cadbl(rst!codidelinia)))
    If Not rst.EOF Then
        vcos = "DADES DE LA COMANDA AMB CdL A IMPRESORES." + vbNewLine + vbNewLine
        vcos = vcos + "Nova comanda. " + atrim(vnumc) + vbNewLine
        vcos = vcos + "Comanda a impresores. " + atrim(rst!comanda) + vbNewLine
        vcos = vcos + "Numero de CdL. " + atrim(Format(rst!codidelinia, "000") + "#" + atrim(rst!codideliniav)) + vbNewLine
        enviaremailgeneric "controlCdLnovescomandes", "Nova comanda amb un CdL igual a l'ordre d'impresió. " + Format(Now, "dd/mm/yy hh.nn"), vcos
    End If
  End If
fi:
  Set rst = Nothing
End Sub
Sub seleccionar_refinplacsa_activa(vrefinplacsa As String)
  Dim vrefactiva As String
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select refinplacsa,inactiva from tarifes_referencies where refinplacsa like '??" + Trim(Mid(vrefinplacsa + "   ", 3)) + "*'"
  formseleccio.refrescar
  formseleccio.Data1.Recordset.MoveLast
  If formseleccio.Data1.Recordset.RecordCount > 1 Then
     MsgBox "Escull la referencia d'Inplacsa que vols deixar com ACTIVA, les altres passaran a no actives.", vbExclamation, "Atenció"
       Else: GoTo fi
  End If
  formseleccio.DBGrid2.Columns(1).Visible = False
  formseleccio.DBGrid2.Font.Size = 14
  formseleccio.Show 1
  If seleccioret = 1 Then
     vrefactiva = atrim(formseleccio.Data1.Recordset!refinplacsa)
     formseleccio.Data1.Recordset.MoveFirst
     While Not formseleccio.Data1.Recordset.EOF
        If formseleccio.Data1.Recordset!refinplacsa = vrefactiva Then
             formseleccio.Data1.Recordset.Edit: formseleccio.Data1.Recordset!inactiva = False: formseleccio.Data1.Recordset.Update
           Else: formseleccio.Data1.Recordset.Edit: formseleccio.Data1.Recordset!inactiva = True: formseleccio.Data1.Recordset.Update
        End If
        formseleccio.Data1.Recordset.MoveNext
     Wend
  End If
fi:
  Unload formseleccio
End Sub
Sub gravar_refinplacsaSICAL(vrefinplacsa As String, vclient As Double)
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from tarifes_referencies where refinplacsa='" + atrim(vrefinplacsa) + "'")
   If rst.EOF Then
      dbtmp.Execute "insert into tarifes_referencies (codiclient,refinplacsa,coditarifa) values ('" + atrim(vclient) + "','" + atrim(vrefinplacsa) + "',null)"
      If llegir_ini("General", "exportant", fitxerini) <> "1" Then
          seleccionar_refinplacsa_activa vrefinplacsa
      End If
      dbtmp.Execute "update comandes_extres SET refinplacsa_nova=true , refinplacsa_validada=False where refinplacsa='" + vrefinplacsa + "'"
   End If
   Set rst = Nothing
End Sub
Sub llistar_comanda(Optional expedicions As Boolean, Optional numerocomanda As String, Optional imprimirnomespagina1 As Boolean)
  Dim vcomandaimpresa As Boolean
  Dim vcolorenvio As Double
  Dim taulatemp As String
  Dim rsttmp22 As Recordset
  Dim rsttmp2 As Recordset
  Dim rsttmp3 As Recordset
  Dim rstfirmes As Recordset
  Dim nomcontable As String
  Dim rutallistat As String
  Dim cont As Integer
  Dim ultimasec As String
  Dim colorenvio As Variant
  Dim coloradhesiu As Variant
  
  colorenvio = Array("crSilver", "crMaroon", "crGreen", "crOlive", "crNavy", "crPurple", "crTeal", "crGray", "crFuchsia", "crRed", "crLime", "crYellow", "crBlue", "crBlack", "crSilver", "crMaroon", "crGreen", "crOlive", "crNavy", "crPurple", "crTeal", "crGray", "crFuchsia", "crRed", "crLime", "crYellow", "crBlue", "crBlack")

  'Dim numerocomanda As String
  Dim descultimasec As String
  'numerocomanda = "117704"
  
  
  If numerocomanda = "" Then numerocomanda = atrim(Data1.Recordset!comanda)
  If Not Data1.Recordset.EOF Then
   If Data1.Recordset!comanda <> numerocomanda Then
     Data1.RecordSource = "select * from comandes where comanda=" + atrim(numerocomanda)
     Data1.Refresh
   End If
     Else
        Data1.RecordSource = "select * from comandes where comanda=" + atrim(numerocomanda)
        Data1.Refresh
  End If
  On Error Resume Next
  MkDir "c:\temp"
  On Error GoTo 0
  '"c:\temporal.mdb"
  taulatemp = nomfitxertemporal
 'la linia del kill hauria d'anar al final del llistat xo no afecta al resusltat final
  If existeix(taulatemp) Then Kill taulatemp
  DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  Set bdllistat = DBEngine.OpenDatabase(taulatemp)
  dbtmp.Execute ("select * into temporal in '" + taulatemp + "' from comandes where comanda=" + numerocomanda)
  bdllistat.Execute ("create table temporal2 (id double);")
 ' afegeixo els camps que em calen pel llistat
  crear_campsextres
  Set rsttmp = bdllistat.OpenRecordset("select * from temporal")
  Set rsttmp22 = bdllistat.OpenRecordset("select * from temporal2")
  
  If Not seleccioimpresio.imprimir(2) Then
    'passar comanda a impresa
    Command9(0).Tag = "Siimpresa"
    rsttmp.FindFirst "comanda=" + atrim(numerocomanda)
    If Not rsttmp.NoMatch Then
      dbtmp.Execute "update comandes_Extres set comandaimpresa=true where comanda=" + atrim(cadbl(rsttmp!comanda))
      dbtmp.Execute "update comandes_Extres set comandaimpresa=true where comanda=" + atrim(cadbl(rsttmp!linkcomanda1))
      dbtmp.Execute "update comandes_Extres set comandaimpresa=true where comanda=" + atrim(cadbl(rsttmp!linkcomanda2))
      'TAMBÉ FIRMO LA COMANDA AMB EL NOM DEL QUE HA IMPRES LA COMANDA (EN RICARD EN PRINCIPI)
      If llegir_ini("General", "exportant", fitxerini) <> "1" And (atrim(Data1.Recordset!proximaseccio) <> "T" And atrim(Data1.Recordset!proximaseccio) <> "V" And atrim(Data1.Recordset!proximaseccio) <> "P") Then 'SI NO ESTÀ EXPORTANT A PDF EL SERVIDOR
          Set rstfirmes = dbtmp.OpenRecordset("select * from comandes_firmes where comanda=" + atrim(numerocomanda) + " and tipus='INI'")
          If rstfirmes.EOF Then
                'dbtmp.Execute "delete * from comandes_firmes where comanda=" + atrim(numerocomanda) + " and usuari='" + nomordinador + "' and tipus='INI'"
              dbtmp.Execute "insert into comandes_firmes (comanda,usuari,tipus,data) values (" + atrim(numerocomanda) + ",'" + nomordinador + "','INI',now)"
          End If
          Set rstfirmes = Nothing
      End If
      'SI ES DE COPS ES PASSA A PRODUCCIÓ DIRECTAMENT
      If checkpassaraproduccio.Value = 0 And InStr(1, nomclient, "CROP´S") = 0 Then dbtmp.Execute "update comandes_extres set passaraimpresores=1  where comanda=" + atrim(cadbl(rsttmp!comanda))
      'SI ES DE DARTA I ES NOVA O MODIFICADDA ES PASSA UN EMAIL
      If InStr(1, nomclient, "D´ARTA") > 0 And (atrim(Data1.Recordset!impressio) = "N" Or atrim(Data1.Recordset!impressio) = "M") Then enviaremailsiesdedARTAinovaomodificada cadbl(rsttmp!comanda)
    End If
    'generar numero de referencia inplacsa
    generarrefinplacsadefinitiu cadbl(numerocomanda)
    
  End If
  
  
  crear_taulafullexpedicions
  possar_taulafullexpedicions numerocomanda
  
  
  While Not rsttmp.EOF
   ' busco la ruta del producte
   Set rsttmp2 = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim((rsttmp!producte)) + "'")
   If Not rsttmp2.EOF Then
      rutallistat = atrim(rsttmp2!ruta)
     Else: rutallistat = ""
  End If
   'possem els camps extres de la taula temporal
   rsttmp.Edit
   rsttmp!arxiupdf = ""
   rsttmp!arxiuimpressora = ""
   rsttmp22.AddNew
   rsttmp22!peuimprenta = Label1(153).Caption
   rsttmp22!sistemaimpresio = tipusimpresio
   rsttmp22!dataiusuari = nomordinador + " - " + Format(Now, "dd/mm/yy hh:nn")
   rsttmp22!cext = "F"
   rsttmp22!cimp = "F"
   rsttmp22!clam = "F"
   rsttmp22!creb = "F"
   rsttmp22!csol = "F"
   rsttmp22!cultimasec = IIf(InStr(1, Data1.Recordset!producte, "PC") > 0, "F", "T")
   rsttmp22!cexpedsec = "F"
   
    If InStr(1, rutallistat, "E") Then rsttmp22!cext = "T"
    If InStr(1, rutallistat, "I") Then rsttmp22!cimp = "T"
    If InStr(1, rutallistat, "L") Then rsttmp22!clam = "T"
    If InStr(1, rutallistat, "R") Then rsttmp22!creb = "T"
    If InStr(1, rutallistat, "S") Then rsttmp22!csol = "T"
    
'sabela ultima seccio
    If rsttmp22!csol = "T" Then
          ultimasec = "S"
            Else
             If rsttmp22!creb = "T" Then
                ultimasec = "R"
                  Else
                    If rsttmp22!clam = "T" Then
                      ultimasec = "L"
                     Else
                       If rsttmp22!cimp = "T" Then ultimasec = "I"
                    End If
             End If
      End If
'fins aqui ultima seccio

   If expedicions Then
      rsttmp22!cultimasec = "F"
      rsttmp22!cexpedsec = "T"
      If rsttmp22!csol = "T" Then
          rsttmp22!creb = "F": rsttmp22!clam = "F": rsttmp22!cimp = "F": rsttmp22!cext = "F"
            Else
             If rsttmp22!creb = "T" Then
                rsttmp22csol = "F": rsttmp22!clam = "F": rsttmp22!cimp = "F": rsttmp22!cext = "F"
                  Else
                    If rsttmp22!clam = "T" Then
                      rsttmp22!csol = "F": rsttmp22!creb = "F": rsttmp22!cimp = "F": rsttmp22!cext = "F"
                     Else
                       If rsttmp22!cimp = "T" Then
                            rsttmp22!csol = "F": rsttmp22!clam = "F": rsttmp22!creb = "F": rsttmp22!cext = "F"
                       End If
                    End If
             End If
      End If
                
   End If
   carregar_lookups_llistat rsttmp, rsttmp22
   rsttmp22!refclient1 = Mid(rsttmp!refclient, 1, 30)
   rsttmp22!refclient2 = Mid(rsttmp!refclient, 31)
   r = ""
   If cadbl(rsttmp!simulteneitat) > 0 Then
       r = atrim(rsttmp!simulteneitat) + " BANDES"
   End If
   rsttmp22!bandesobertmicro = Trim(r)
   r = ""
   
   If cadbl(rsttmp!oberturaex) > 0 Then
            r = "OBERT " + atrim(rsttmp!oberturaex) + " COSTATS --"
   End If
   rsttmp22!obert2 = r
   If atrim(rsttmp!micropex) = "S" Then
      r = "MICROPERFORAT"
   End If
   rsttmp22!micro = r
   
   'carrego les dades del client
   'Set rsttmp2 = dbtmp.OpenRecordset("select * from clients where codi=" + atrim((rsttmp!client)) + "")
   Set rsttmp2 = dbtmp.OpenRecordset("SELECT comandes_extres.refinplacsa,comandes_extres.refinplacsa_nova,COMANDES_EXTRES.comandaimpresa,comandes_extres.numerobossasoldadores,comandes_extres.tipusmaterialcanutureb,clients.*, Clients_codiscomptables.nomclient,Clients_codiscomptables.codicomptable FROM (comandes_extres RIGHT JOIN (clients LEFT JOIN comandes ON clients.codi = comandes.client) ON comandes_extres.comanda = comandes.comanda) LEFT JOIN Clients_codiscomptables ON comandes_extres.codicomptable = Clients_codiscomptables.codicomptable WHERE (((comandes.comanda)=" + atrim(rsttmp!comanda) + "));")
   If Not rsttmp2.EOF Then
      If atrim(rsttmp2!refinplacsa) = "" Then MsgBox "ERROR AMB LA REFERENCIA D'INPLACSA, PER ALGUN MOTIU NO S'HA CREAT CORRECTAMENT." + vbNewLine + "REVISA-HO I TORNA A PROVAR D'IMPRIMIR", vbCritical, "ERROR": GoTo fi
      vcomandaimpresa = rsttmp2!comandaimpresa
      vtipusmaterialcanutureb = atrim(rsttmp2!tipusmaterialcanutureb)
      nomcontable = atrim(rsttmp2!nomclient)
       
      If InStr(1, atrim(rsttmp2!codicomptable), atrim((rsttmp!client))) = 0 Then rsttmp22!nomfacturacio = atrim(rsttmp2!codicomptable) + " - " + atrim(nomcontable)
    '  rsttmp22!arxiuexp = atrim(rsttmp2!arxiuexp)
      rsttmp!refinplacsa = atrim(rsttmp2!refinplacsa)
      rsttmp!refinplacsanova = IIf(atrim(rsttmp2!refinplacsa_nova), "S", "N")
      rsttmp22!nomclient = rsttmp2!nom
      rsttmp22!nomdestinatari = ""
      rsttmp22!adreça = rsttmp2!domicili
      rsttmp22!poblacio = rsttmp2!poblacio
      rsttmp22!provincia = rsttmp2!provincia
      rsttmp22!codipostal = rsttmp2!codipostal
      rsttmp22!horarientrega = rsttmp2!horaridesc
      rsttmp22!numproveidor = rsttmp2!numproveidor
      rsttmp22!numerobossasoldadores = atrim(rsttmp2!numerobossasoldadores)
      obsultimaclient = atrim(rsttmp2!obsultima)
      descultimasec = IIf(cadbl(rsttmp2!paperfrontal), "PAPER FRONTAL PALET ", "") + IIf(cadbl(rsttmp2!anonim), "-EMBALATGE ANONIM", "")
      descultimasec = descultimasec + IIf(cadbl(rsttmp2!guardarmostra), "-GUARDAR MOSTRA CLIENT", "") + IIf(cadbl(rsttmp2!europeu), "-PALET EUROPEU", "")
      descultimasec = descultimasec + IIf(cadbl(rsttmp2!refpalet), "-1 REF x PALET ", "") + IIf(atrim(rsttmp2!arxiuult) <> "", "       ARXIU: " + (UCase(agafarnomfitxer(atrim(rsttmp2!arxiuult)))), "")
      
      If cadbl(rsttmp2!albvalorat) <> 0 Then
          rsttmp22!albvaloraticertqualitat = "ALBARA VALORAT"
         Else: rsttmp22!albvaloraticertqualitat = "           "
      End If
      If cadbl(rsttmp2!certqualitat) <> 0 Then
          rsttmp22!albvaloraticertqualitat = rsttmp22!albvaloraticertqualitat + "   CERTIFICAT QUALITAT"
         Else: rsttmp22!albvaloraticertqualitat = rsttmp22!albvaloraticertqualitat
      End If
      'carrega la forma de pagament
      Set rsttmp3 = dbtmp.OpenRecordset("select descripcio from [formes de pagament] where codi='" + atrim((rsttmp2!formapag)) + "'")
      If Not rsttmp3.EOF Then
        r = rsttmp3!descripcio
         Else: r = ""
       End If
      rsttmp22!formapagament = r
   End If
   
   'posso el nom del client d'envio
   If rsttmp!direnvio > 0 Then
      Set rsttmp2 = dbtmp.OpenRecordset("select * from clients_envios where id=" + atrim(cadbl(rsttmp!direnvio)))
      If Not rsttmp2.EOF Then
       rsttmp22!nomdestinatari = atrim(rsttmp2!ID) + "-" + atrim(rsttmp2!nome)
       rsttmp22!adreça = rsttmp2!domicilie
       rsttmp22!poblacio = rsttmp2!poblacioe
       rsttmp22!provincia = rsttmp2!provinciae
       rsttmp22!codipostal = rsttmp2!codipostale
      End If
   End If
   'comprovo els tractats
   rsttmp!tractatex = IIf(cadbl(rsttmp!tractatex) = 0, "N", rsttmp!tractatex)
   rsttmp!oberturaex = IIf(cadbl(rsttmp!oberturaex) = 0, "N", rsttmp!oberturaex)
   rsttmp!obert = IIf(cadbl(rsttmp!obert) = 0, "N", rsttmp!obert)
   'poso el texte dimpresio
   posarmarcailinia cadbl(rsttmp!numtreball)
    rsttmp!texteimpressio = atrim(Mid(Text103(4), 1, 98))
   'actualitzar camps de les tintes
   actualitzarcampstintes Data1.Recordset, rsttmp22, rsttmp
   'possodetalls de l'etiqueta de rebobinadora
   posardetallsetiquetareb Data1.Recordset!direnvio, rsttmp
   'coloco la descripcio de la ultima seccio
   i = 95
   If ruta <> "" Then
      ultimaseccio = Mid(ruta, Len(ruta), 1)
   End If
   For i = 0 To 80
     llistat.Formulas(i) = ""
   Next i
   If InStr(1, ruta, "R") = 0 Then rebpes = 0
  'If data1.Recordset!producte = "PC" Or data1.Recordset!producte = "PC2" Then rebpes = 0
   f = 0
   i = 80
   llistat.Formulas(f) = "cingularreal2='" + IIf(mirarsihihaCingularReal(cadbl(rsttmp!numtreball), cadbl(rsttmp!numordremodificacio)), "Si", "No") + "'"
   f = f + 1
   llistat.Formulas(f) = "obsultimasec='" + treure_apostruf(Mid(UCase(descultimasec), 1, i)) + "'"
   f = f + 1
   llistat.Formulas(f) = "obsultimasec2='" + treure_apostruf(Mid(UCase(descultimasec), i + 1)) + "'"
   f = f + 1
   llistat.Formulas(f) = "obsultimaclient='" + treure_apostruf(Mid(UCase(obsultimaclient), 1, i)) + "'"
   f = f + 1
      
   llistat.Formulas(f) = "arxiuexp='" + UCase(agafarnomfitxer(atrim(rsttmp22!arxiuexp))) + "'"
   f = f + 1
   
   llistat.Formulas(f) = "arxiuext='" + (UCase(agafarnomfitxer(atrim(Data1.Recordset!arxiuext)))) + "'"
   f = f + 1
   llistat.Formulas(f) = "arxiupdf='" + IIf(existeix_pdf_treball(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio)), "Sí", "No") + "'"
   f = f + 1
   llistat.Formulas(f) = "arxiuimp='" + IIf(existeix_imp_treball(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), cadbl(Data1.Recordset!client), cadbl(Data1.Recordset!direnvio)), "Sí", "No") + "'"
   f = f + 1
   llistat.Formulas(f) = "arxiulam='" + UCase(agafarnomfitxer(atrim(Data1.Recordset!arxiulaminadora))) + "'"
   f = f + 1
   llistat.Formulas(f) = "arxiureb='" + UCase(agafarnomfitxer(atrim(Data1.Recordset!arxiureb))) + "'"
   f = f + 1
   llistat.Formulas(f) = "arxiusol='" + UCase(agafarnomfitxer(atrim(Data1.Recordset!arxiusol))) + "'"
   f = f + 1
   llistat.Formulas(f) = "avismodificionstreball='" + UCase(buscaravismodificacionstreball(atrim(Data1.Recordset!comanda))) + "'"
   f = f + 1

   llistat.Formulas(f) = "reducciocilindrefw='" + UCase(buscarreducciocilindretreball(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), "fw")) + "'"
   f = f + 1

   llistat.Formulas(f) = "reducciocilindref2='" + UCase(buscarreducciocilindretreball(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), "f2")) + "'"
   f = f + 1
   vubicacio = atrim(buscarubicacioeneltreball(cadbl(Data1.Recordset!numtreball), atrim(Data1.Recordset!impressio)))
   llistat.Formulas(f) = "ubicacioclixe='" + vubicacio + "'"
   f = f + 1

   llistat.Formulas(f) = "linialam='" + atrim(Data1.Recordset!arxiuext) + "'"
   f = f + 1
   llistat.Formulas(f) = "clixescomparteixenamb='" + atrim(mirarsicompartit(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), cadbl(Data1.Recordset!arxiu))) + "'"
   f = f + 1
   llistat.Formulas(f) = "reprint='" + atrim(mirarsireprint(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio))) + "'"
   f = f + 1

   On Error Resume Next
   If rebpes = "" Or rebpes = "0" Or InStr(1, rutallistat, "S") > 0 Then
       If cadbl(Data1.Recordset!cantitatex) > 0 Then
           'rebpes = atrim(Redondejar(cadbl(data1.Recordset!pes1000mtrs) * (data1.Recordset!cantitatex / 1000), 1)) + "+"
           rebpes = atrim(cadbl(solpes.Tag)) + "+"
           If cadbl(solpes.Tag) = 0 Then rebpes = atrim(Redondejar(cadbl(Data1.Recordset!pes1000mtrs) * (Data1.Recordset!cantitatex / 1000), 1)) + "+"
       End If
   End If
   llistat.Formulas(f) = "totalpesreb='" + atrim(Format(cadbl(Mid(rebpes, 1, InStr(1, rebpes, "+") - 1)), "#,##0")) + "'"
   f = f + 1
   llistat.Formulas(f) = "totalmetresreb='" + atrim(Format(cadbl(rebmetres), "#,##0")) + "'"
   'llistat.Formulas(f) = "totalmetresreb='" + atrim(Format(cadbl(Mid(rebmetres, 1, InStr(1, rebmetres, "(") - 1)), "#,##0")) + "'"
   On Error GoTo 0
   f = f + 1
   llistat.Formulas(f) = "totalpeces='" + rebpcs + "'"
   f = f + 1
   llistat.Formulas(f) = "ultimaseccio='" + ultimasec + "'"
   f = f + 1
   vcolorenvio = IIf(cadbl(Label1(147).Tag) - 1 > 28, 1, 28 - cadbl(Label1(147).Tag) - 1)
   llistat.Formulas(f) = "enviodiferent='" + colorenvio(vcolorenvio) + "'"
   f = f + 1
   llistat.Formulas(f) = "hihaafectatspelcanvi=" + mirarsihihaclixesafectatspelcanvi(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio))
   f = f + 1
   llistat.Formulas(f) = "colormaterial=" + atrim(IIf(cmarcmaterial.BorderColor < 0, 0, cmarcmaterial.BorderColor)) + ""
   f = f + 1
   llistat.Formulas(f) = "materialtubbasereb='" + IIf(vtipusmaterialcanutureb = "C", "CARTRÓ", IIf(vtipusmaterialcanutureb = "P", "PVC", "")) + "'"
   f = f + 1
   llistat.Formulas(f) = "velocitatimpresio='" + descripcio_metresminut(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio)) + "'"
   f = f + 1
   
   If Not expedicions Then
     k = 1
    'For k = 1 To 10
      llistat.Formulas(f) = "lareb" + atrim(k) + "='" + IIf(InStr(1, rutallistat, "R") = 4, ".", "") + "'"
      f = f + 1
   ' Next k
   ' For k = 1 To 10
      llistat.Formulas(f) = "lasol" + atrim(k) + "='" + IIf(InStr(1, rutallistat, "S") = 4, ".", "") + "'"
      f = f + 1
   ' Next k
   End If
   
   
   r = ""
   'fi de la col.locacio de la ultima seccio
   
   'fi ns aqui
   rsttmp.Update
   rsttmp22.Update
   rsttmp.MoveNext
  Wend
  If Not rsttmp.EOF And Not rsttmp.BOF Then rsttmp.MoveFirst
  If rsttmp22.EOF And rsttmp22.BOF Then Exit Sub
  rsttmp22.MoveFirst
  llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "comandes_tintesnoves.rpt" '   _tintesnoves
  'If expedicions Then llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "copia comandes.rpt"
  llistat.Destination = crptToWindow
  llistat.WindowState = crptMaximized
  
  For i = 1 To 10
   llistat.SectionFormat(i) = ""
  Next i
  
  comprovosihaigdocultarseccions rsttmp22
  llistat.SectionFormat(0) = "GH2;" + rsttmp22!cext + ";X;X;X;X;X;X"
  llistat.SectionFormat(1) = "GH3;" + rsttmp22!cimp + ";X;X;X;X;X;X"
  llistat.SectionFormat(2) = "GH4;" + rsttmp22!clam + ";" + IIf(InStr(1, rutallistat, "L") = 3 And Not expedicions, "T", "F") + ";X;X;X;X;X"                '   possar =3 a tots dos per novestintes
  llistat.SectionFormat(3) = "GH5;" + rsttmp22!creb + ";" + IIf(InStr(1, rutallistat, "R") = 3 And Not expedicions, "T", "F") + ";X;X;X;X;X"
  llistat.SectionFormat(4) = "GH6;" + rsttmp22!csol + ";" + IIf(InStr(1, rutallistat, "S") = 3 And Not expedicions, "T", "F") + ";X;X;X;X;X;X"
 ' llistat.SectionFormat(3) = "GH5;" + rsttmp22!creb + ";F;X;X;X;X;X"
 ' llistat.SectionFormat(4) = "GH6;" + rsttmp22!csol + ";F;X;X;X;X;X"

  llistat.SectionFormat(5) = "GH7;" + rsttmp22!cultimasec + ";X;X;X;X;X;X"
  llistat.SectionFormat(6) = "GH8;" + rsttmp22!cexpedsec + ";X;X;X;X;X;X"
  llistat.DataFiles(0) = taulatemp
'  llistat.DiscardSavedData = True
  'llistat.RetrieveDataFiles
    'MsgBox llistat.ReportFileName
    
  If Not seleccioimpresio.imprimir(2) And Not existeix("c:\ordprog.ini") Then llistat.Destination = crptToPrinter
  If llegir_ini("baixes", "imprimircomandanomesimp", "comandes.ini") = "S" Or imprimirnomespagina1 Then
      llistat.Destination = crptToPrinter
      seleccionar_segonaimpresora
      llistat.PrinterStartPage = 1
      llistat.PrinterStopPage = 1
      llistat.PageLast
      escriure_ini "baixes", "imprimircomandanomesimp", "N", "comandes.ini"
        Else
          llistat.PrinterStartPage = 1
          llistat.PrinterStopPage = 9999
  End If
  seleccionar_segonaimpresora
  wait (2)
  llistat.WindowTitle = "Imprimint comanda"
  llistat.Action = 1
  If Not expedicions Then
    dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(numerocomanda) + ",'" + nomordinador + "','ImprimirComanda','','')"
    generar_gtin14 Data1.Recordset!comanda
  End If
fi:
  Set rsttmp3 = Nothing
  Set rsttmp2 = Nothing
  Set rsttmp = Nothing
  Set rstfirmes = Nothing
  bdllistat.Close
  Set bdllistat = Nothing
  Data1.Recordset.Move 0
End Sub
Sub posardetallsetiquetareb(vid_envio As Double, rst As Recordset)
   Dim rstetenvio As Recordset
   Set rstetenvio = dbtmp.OpenRecordset("select * from clients_etbobina where id_envio=" + atrim(vid_envio))
   If Not rstetenvio.EOF Then
       rst!etiqintcanutu = Mid("Et:" + atrim(rstetenvio!etinteriorbob), 1, 14)
   End If
   Set rstetenvio = Nothing
End Sub
Sub enviaremailsiesdedARTAinovaomodificada(vnumc As String)
   Dim rst As Recordset
   Dim cos As String
   Set rst = dbtmp.OpenRecordset("SELECT  comandes.comanda, clients.codi, clients.nom, comandes.refclient, comandes.marcailinia FROM clients INNER JOIN comandes ON clients.codi = comandes.client where comanda=" + atrim(vnumc))
   cos = "ACTUALITZAR BASE DE DADES TARIFA PREUS CLIENTS CSV" + Chr(13) + Chr(10)
   cos = cos + Chr(13) + Chr(10) + "Codi Client: " + atrim(rst!codi) + " - " + atrim(rst!nom) + Chr(13) + Chr(10) + "Ref.Client: " + atrim(rst!refclient) + Chr(10) + Chr(13) + "Texte Imp.: " + atrim(rst!marcailinia)
   enviaremailgeneric "avisarNOVAMODIFICADAdeARTA", "Comanda " + vnumc + " -  Nova/Modificada de d'ARTA.", cos
   Set rst = Nothing
End Sub
Sub generar_gtin14(vnumc As Double)
   Dim vsql As String
   Dim rst As Recordset
   Dim vcodigtin As Double
   vsql = "SELECT comandes.comanda, clients.nom,clients.codi, comandes.refclient, clients.ultimcodiarticle, codis_gtin14.codigtin, Clients_envios.estilfrontal FROM ((comandes LEFT JOIN codis_gtin14 ON (comandes.client = codis_gtin14.codiclient) AND (comandes.refclient = codis_gtin14.refclient)) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN Clients_envios ON comandes.direnvio = Clients_envios.id "
   vsql = vsql + " where comanda=" + atrim(vnumc)
   Set rst = dbtmp.OpenRecordset(vsql)
   If Not rst.EOF Then
     If atrim(rst!estilfrontal) <> "Estil UCC128" Or atrim(rst!refclient) = "" Then GoTo fi
     If cadbl(rst!codigtin) > 0 Then
         vcodigtin = cadbl(rst!codigtin)
          Else
           vcodigtin = cadbl(rst!ultimcodiarticle) + 1
           dbtmp.Execute "update clients set ultimcodiarticle=" + atrim(vcodigtin) + " where codi=" + atrim(cadbl(rst!codi))
           dbtmp.Execute "insert into codis_gtin14 (codiclient,refclient,codigtin) values (" + atrim(rst!codi) + ",'" + atrim(rst!refclient) + "'," + atrim(vcodigtin) + ")"
     End If
     dbtmp.Execute "update comandes_extres set gtin14=" + atrim(vcodigtin) + " where comanda=" + atrim(vnumc)
   End If
fi:
   Set rst = Nothing
End Sub
Function descripcio_metresminut(vnumtreball As Long, vnumordremodificacio As Long) As String
   Dim v1fw As Double
   Dim v2fw As Double
   Dim v1f2 As Double
   Dim v2f2 As Double
   If vnumtreball = 0 Then Exit Function
   calcular_mtrsminut vnumtreball, vnumordremodificacio, v1fw, v2fw, v1f2, v2f2
   descripcio_metresminut = "" + IIf(v1fw > 0 Or v2fw > 0, "FW:" + atrim(v1fw) + "~" + atrim(v2fw), "") + IIf(v1f2 > 0 Or v2f2 > 0, " F2:" + atrim(v1f2) + "~" + atrim(v2f2), "")
End Function
Function mirarsihihaclixesafectatspelcanvi(numtreball As Double, ordre As Double) As String
   Dim rst As Recordset
   mirarsihihaclixesafectatspelcanvi = "True"
   Set rst = dbclixes.OpenRecordset("select afectatspelcanvi from tintes where id_treball=" + atrim(numtreball) + " and ordremodificacio=" + atrim(ordre) + " and afectatspelcanvi=true")
   If rst.EOF Then mirarsihihaclixesafectatspelcanvi = "False"
   Set rst = Nothing
End Function
Sub seleccionar_segonaimpresora()
  Dim x As Printer
  For Each x In Printers
     If x.DeviceName = llegir_ini("General", "segonaimpresoradecomandes", fitxerini) Then GoTo canviar_impresora
  Next x
  Exit Sub
canviar_impresora:
  llistat.PrinterName = x.DeviceName
  llistat.PrinterDriver = x.DriverName
  llistat.PrinterPort = x.Port
End Sub
Function mirarsireprint(numtreball As Double, ordre As Double) As String
  Dim rst As Recordset
  mirarsireprint = "N"
  Set rst = dbclixes.OpenRecordset("select reimpres from modificacions where id_treball=" + atrim(numtreball) + " and ordre=" + atrim(ordre))
  If Not rst.EOF Then
     If rst!reimpres Then mirarsireprint = "S"
  End If
End Function
Function mirarsicompartit(numc As Double, numordre As Double, varxiu As String) As String
  Dim rst As Recordset
  Dim rstclixe As Recordset
  Dim vconsulta As String
  Dim vconsultafinal As String
  Dim varxiuv As String
  mirarsicompartit = ""
  'Set rst = dbclixes.OpenRecordset("select * from tintes where tinterlinkambid_treball,id_treball=" + atrim(numc))
  vconsulta = "id_tinter in(select tinterlinkambid_treball from tintes where tinterlinkambid_treball and id_treball=" + atrim(numc) + "and ordremodificacio=" + atrim(numordre) + ")"
  GoTo cont
  Set rst = dbclixes.OpenRecordset("select distinct id_treball from tintes where " + vconsulta)
  While Not rst.EOF
      Set rstclixe = dbclixes.OpenRecordset("select arxiu from clixes where id_treball=" + atrim(rst!id_treball))
      If Not rstclixe.EOF Then
        If atrim(rstclixe!arxiu) <> varxiu Then
         mirarsicompartit = mirarsicompartit + atrim(rst!id_treball) + atrim(rstclixe!arxiu)
        End If
      End If
      rst.MoveNext
  Wend
cont:
  vconsultafinal = "select distinct id_treball from tintes where tinterlinkambid_treball in (select id_tinter from tintes where comparteix and id_treball=" + atrim(numc) + " and ordremodificacio=" + atrim(numordre) + ") or (" + vconsulta + ")"
  Set rst = dbclixes.OpenRecordset("select id_treball,arxiu from clixes where id_treball in (" + vconsultafinal + ") order by arxiu")
  If Not rst.EOF Then
     varxiuv = atrim(rst!arxiu)
     mirarsicompartit = atrim(rst!arxiu) + "(" + atrim(rst!id_treball)
  End If
  While Not rst.EOF
      'Set rstclixe = dbclixes.OpenRecordset("select arxiu from clixes where id_treball=" + atrim(rst!id_treball))
       If varxiuv = atrim(rst!arxiu) Then
         mirarsicompartit = mirarsicompartit + "," + atrim(rst!id_treball)
          Else
            varxiuv = atrim(rst!arxiu)
            mirarsicompartit = mirarsicompartit + ") " + atrim(rst!arxiu) + "(" + atrim(rst!id_treball)
       End If
      rst.MoveNext
  Wend
  If mirarsicompartit <> "" Then mirarsicompartit = "Compartit_amb:_" + mirarsicompartit + ")"
  Set rst = Nothing
  Set rstclixe = Nothing
End Function

Sub comprovosihaigdocultarseccions(rsttmp22 As Recordset)
    If llegir_ini("baixes", "imprimircomandanomesimp", "comandes.ini") = "S" Then
       rsttmp22.Edit
        rsttmp22!clam = "F"
        rsttmp22!creb = "F"
        rsttmp22!csol = "F"
        rsttmp22!cultimasec = "F"
        rsttmp22.Update
    End If
End Sub
Function buscarubicacioeneltreball(treball As Double, impressio As String) As String
  Dim rst As Recordset
  
    Set rst = dbclixes.OpenRecordset("select * from clixes where id_Treball=" + atrim(treball))
    If Not rst.EOF Then
            If atrim(rst!ubicacio) <> "" Then buscarubicacioeneltreball = "Ubicació: " + atrim(rst!ubicacio)
            If Mid(atrim(rst!ubicacio), 1, 2) <> "P-" And Not (impressio = "M" Or impressio = "N") Then buscarubicacioeneltreball = ""
    End If
  Set rst = Nothing
  
End Function

Function buscarreducciocilindretreball(treball As Double, ordre As Double, fwof2 As String, Optional msgcurt As Boolean) As String
  Dim rst As Recordset
  
  Set rst = dbclixes.OpenRecordset("select * from clixes where id_Treball=" + atrim(treball))
  If Not rst.EOF Then
    If cadbl(rst!reduccioxmetre) <> 0 Then
     If fwof2 = "fw" Then
          buscarreducciocilindretreball = IIf(msgcurt, "Distorsió: ", "Distorsió mtr/lin: ") + atrim(cadbl(rst!reduccioxmetre)) + IIf(msgcurt, "", " mm") + " FW: "
          buscarreducciocilindretreball = buscarreducciocilindretreball + atrim(Redondejar(cadbl(rst!redcilindrefw), 2)) + IIf(msgcurt, "", " ")
     End If
     If fwof2 = "f2" Then
         buscarreducciocilindretreball = " F2: " + atrim(Redondejar(cadbl(rst!redcilindref2), 2)) + IIf(msgcurt, "", " ")
     End If
    End If
  End If
  
  
End Function
Function buscaravismodificacionstreball(numc As Double) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from comandes_Extres where comanda=" + atrim(numc))
  If Not rst.EOF Then buscaravismodificacionstreball = treure_apostruf(atrim(rst!aviscanvisambeltreball))
End Function
Sub borrarcarpetatemporal()
   On Error Resume Next
   Kill "c:\temp\exportar\*.*"
End Sub
Sub exportartotalacomanda(numc As Double)

   Data1.RecordSource = "select * from comandes where comanda=" + atrim(numc)
   Data1.Refresh
   borrarcarpetatemporal
   wait 3
   Command9_Click 0
   
 
     
End Sub
Sub imprimirperpantalla(numc As Double)
   
   Load seleccioimpresio
   seleccioimpresio.imprimir(2).Value = True
   llistar_comanda False, atrim(numc)
   
   'If llegir_ini("baixes", "imprimircomanda", "comandes.ini") <> "" Then escriure_ini "baixes", "imprimircomanda", "", "comandes.ini"
   'formcomandes.Hide
End Sub
Function agafarnomfitxer(vnomf As String) As String
  Dim Tm As String
  Tm = vnomf
  While InStr(1, Tm, "\") > 0
    Tm = Mid(Tm, InStr(1, Tm, "\") + 1)
  Wend
  If Len(Tm) < 4 Then Tm = Tm + "         "
  agafarnomfitxer = treure_apostrof(Trim(Mid(Tm, 1, Len(Tm) - 4)))
End Function
Sub crear_taulafullexpedicions()
 Dim camps As String
 Dim camps2 As String
 camps = " monedafact text, unitatxproducte text, valorat text, codibarres text, datafab text, detallbobalb text, detallbobfrontal text "
 camps2 = " pesnet text, alcadamax text, tipuspalet text, guardarmostres text, certqualitat text, albarapalet text, packingpalet text "
 camps3 = " protecciob text, protecciop text, proteccios text, emb_anonim text,  conosprotec text "
 camps4 = " pfpesnet text, pfdatafab text, pfpacking text, pfcodibarres text, pesmaxpalet text ,bobinesmaxpalet text, pfpaperfrontal text,okenvio text "
 bdllistat.Execute ("create table seccioexpedicions (id double);")
 GoSub crearcamps
 camps = camps2: GoSub crearcamps
 camps = camps3: GoSub crearcamps
 camps = camps4: GoSub crearcamps
 
 Exit Sub
crearcamps:
 While Trim(camps) <> ""
    If InStr(1, camps, ",") > 0 Then
       r = Mid(camps, 1, InStr(1, camps, ",") - 1)
      Else: r = camps: camps = ""
    End If
    bdllistat.Execute ("alter table seccioexpedicions  add column " + r)
    camps = Mid(camps, InStr(1, camps, ",") + 1)
 Wend
 Return
End Sub

Sub possar_taulafullexpedicions(numerodecomanda As String)
  Dim rste As Recordset
  Dim rstll As Recordset
  Dim rstenvio As Recordset
  Dim rstclient As Recordset
  Dim codienvio As Long
  Set rstll = bdllistat.OpenRecordset("seccioexpedicions")
  If rsttmp.EOF Then Exit Sub
  Set rstclient = dbtmp.OpenRecordset("select * from clients where codi=" + atrim(cadbl(rsttmp!client)))
  
'  !!!!!!! PENSAR-HI !!!!!!   SI ES FA UN CANVI AMB MES O MENYS OPCIONS TAMBÉ S'HA D'APLICAR A LA BAIXA DE ENFLAJAR BOBINES
  
  If cadbl(rsttmp!direnvio) > 0 Then
      Set rstenvio = dbtmp.OpenRecordset("select * from clients_envios where id=" + atrim(cadbl(rsttmp!direnvio)))
      codienvio = cadbl(rsttmp!direnvio)
     Else
        Set rstenvio = dbtmp.OpenRecordset("select * from clients where codi=" + atrim(cadbl(rsttmp!client)))
        codienvio = cadbl(rsttmp!client) * -1
  End If
  
  If rstenvio.EOF Then
    'no hi ha dades a direccions d'envio
      Exit Sub
  End If
  
  rstll.AddNew
  llistatlookupde "mesures", atrim(rsttmp!mesurapvp), inventat
  rstll!monedafact = inventat
  rstll!unitatxproducte = ""
  r = imp_unitatsxproducte(codienvio, numerodecomanda)
  If atrim(r) <> "" Then
     rstll!unitatxproducte = "" + r
    Else: rstll!unitatxproducte = ""
  End If
  If cadbl(rstenvio!albaravalorat) Then rstll!valorat = "VALORAT"
  If cadbl(rstenvio!codibarres) Then rstll!codibarres = "CODI DE BARRES"
  If cadbl(rstenvio!datafabricacio) Then rstll!datafab = "DATA DE FABRICACIÓ"
  If cadbl(rstenvio!detallbobalpalet) Then rstll!detallbobalb = "DETALL BOBINES AL PALET"
  If cadbl(rstenvio!detallbobalfrontal) Then rstll!detallbobfrontal = "DETALL BOBINES AL FRONTAL"
  If cadbl(rstenvio!pesnetbrut) Then rstll!pesnet = "PES NET"
  'If atrim(rstenvio!bobinesmaxpalet) <> "" Then rstll!bobinesmaxpalet = rsttmp!bobinesmaxpalet
  If cadbl(rstenvio!alcadapalet) Then
    llistatlookupde "alcadespalets", atrim(rstenvio!alcadapalet), inventat
    rstll!alcadamax = inventat + " CM" + IIf(cadbl(rstenvio!pesmaxpalet) > 0, "  Pes Màx. Palet: " + atrim(cadbl(rstenvio!pesmaxpalet)) + " Kg", "")
  End If
  If cadbl(rstenvio!tipuspalet) Then
    llistatlookupde "tipuspalets", atrim(rstenvio!tipuspalet), inventat
    rstll!tipuspalet = inventat
  End If
  If cadbl(rstenvio!guardarmostres) Then
    llistatlookupde "guardarmostres", atrim(rstenvio!guardarmostres), inventat
    rstll!guardarmostres = inventat
  End If
  If cadbl(rstenvio!cert_qualitat) Then
    llistatlookupde "cert_qualitat", atrim(rstenvio!cert_qualitat), inventat
    rstll!certqualitat = inventat
  End If
  If cadbl(rstenvio!albaraalpalet) Then rstll!albarapalet = "ALBARÀ AL PALET"
  If cadbl(rstenvio!packingalpalet) Then rstll!packingpalet = "PACKING-LIST"
  If cadbl(rstenvio!tipusprotecciob) Then
    llistatlookupde "tipusproteccions", atrim(rstenvio!tipusprotecciob), inventat
    rstll!protecciob = inventat
  End If
  If cadbl(rstenvio!tipusprotecciop) Then
    llistatlookupde "tipusproteccions", atrim(rstenvio!tipusprotecciop), inventat
    rstll!protecciop = inventat
  End If
  If cadbl(rstenvio!tipusprotecciospr) Then
    llistatlookupde "tipusproteccions", atrim(rstenvio!tipusprotecciospr), inventat
    rstll!proteccios = inventat
  End If
  If cadbl(rstenvio!emb_anonim) Then
    llistatlookupde "embalatgesanonims", atrim(rstenvio!emb_anonim), inventat
    rstll!emb_anonim = inventat
  End If
   If cadbl(rstenvio!guardarmostres) Then
    llistatlookupde "guardarmostres", atrim(rstenvio!guardarmostres), inventat
    rstll!guardarmostres = inventat
  End If
  If cadbl(rstenvio!conosprotectors) Then
    llistatlookupde "conosprotectors", atrim(rstenvio!conosprotectors), inventat
    rstll!conosprotec = inventat
  End If
  If cadbl(rstenvio!conosprotectors) Then
    llistatlookupde "conosprotectors", atrim(rstenvio!conosprotectors), inventat
    rstll!conosprotec = inventat
  End If
  If cadbl(rstenvio!okenvio) Then rstll!okenvio = "DEMANAR OK PER ENVIAR"
  If cadbl(rstenvio!pfpaperfrontal) Then
    llistatlookupde "tipuspaperfrontal", atrim(rstenvio!pfpaperfrontal), inventat
    rstll!pfpaperfrontal = inventat
  End If
  If atrim(rstll!pfpaperfrontal) <> "" Then
     If cadbl(rstenvio!pfpesnet) Then rstll!pfpesnet = "PES NET"
     If cadbl(rstenvio!pfdatafab) Then rstll!pfdatafab = "DATA FABRICACIO"
     If cadbl(rstenvio!pfpacking) Then rstll!pfpacking = "PACKING-LIST"
     If cadbl(rstenvio!pfcodibarres) Then rstll!pfcodibarres = "CODI DE BARRES" + IIf(atrim(rstenvio!estilfrontal) <> "", " (" + atrim(rstenvio!estilfrontal) + ")", "")
  End If
  
  rstll.Update
End Sub
Function imp_unitatsxproducte(idenvio As Long, numerodecomanda As String)
  Dim rstuxp As Recordset
  Dim f As String
  Dim familia As String
  Set rstuxp = dbtmp.OpenRecordset("select producte from comandes where comanda=" + atrim(numerodecomanda))
  If rstuxp.EOF Then Exit Function
  Set rstuxp = dbtmp.OpenRecordset("select familia from productes where codi='" + atrim(rstuxp!producte) + "'")
  If rstuxp.EOF Then Exit Function
  familia = Mid(atrim(rstuxp!familia), 1, 1)
  Set rstuxp = dbtmp.OpenRecordset("select * from unitatsxproducte where idenvio=" + atrim(cadbl(idenvio)))
  While Not rstuxp.EOF
   If Mid(rstuxp!idproducte, 1, 1) = familia Then
    If rstuxp!kg Then f = f + " KG"
    If rstuxp!mtrs Then f = f + " METRES"
    If rstuxp!unts Then f = f + " UNITATS"
    If rstuxp!pcs Then f = f + " PECES"
    If rstuxp!mt2 Then f = f + " MTRS²"
    If rstuxp!km Then f = f + " KILOMETRES"
    If rstuxp!emiler Then f = f + " €/MILER"
   End If
   rstuxp.MoveNext
  Wend
  imp_unitatsxproducte = f
End Function
Sub crear_campsextres()
Dim i As Byte
bdllistat.Execute ("alter table temporal drop column obsimpgen2")
bdllistat.Execute ("alter table temporal drop column obsextgen2")
bdllistat.Execute ("alter table temporal drop column obsext2")
bdllistat.Execute ("alter table temporal drop column obsimp2")
bdllistat.Execute ("alter table temporal drop column obslamgen2")
bdllistat.Execute ("alter table temporal drop column obslam2")
bdllistat.Execute ("alter table temporal drop column obssol2")
bdllistat.Execute ("alter table temporal drop column obssolgen2")
bdllistat.Execute ("alter table temporal drop column texteimpressio")
bdllistat.Execute ("alter table temporal add column texteimpressio text(100)")
bdllistat.Execute ("alter table temporal add column refinplacsa text(20)")
bdllistat.Execute ("alter table temporal add column refinplacsanova text(1)")

 bdllistat.Execute ("alter table temporal drop column arxiuexp")
 ' bdllistat.Execute ("alter table temporal2 add column desccomlaminar text")
 bdllistat.Execute ("alter table temporal2 add column arxiuexp text(125)")
 bdllistat.Execute ("alter table temporal2 add column cext text(1)")
 bdllistat.Execute ("alter table temporal2 add column cimp text(1)")
 bdllistat.Execute ("alter table temporal2 add column clam text(1)")
 bdllistat.Execute ("alter table temporal2 add column creb text(1)")
 bdllistat.Execute ("alter table temporal2 add column csol text(1)")
 bdllistat.Execute ("alter table temporal2 add column cultimasec text(1)")
 bdllistat.Execute ("alter table temporal2 add column cexpedsec text(1)")
 bdllistat.Execute ("alter table temporal2 add column nomclient text")
 bdllistat.Execute ("alter table temporal2 add column ruta text(10)")
 bdllistat.Execute ("alter table temporal2 add column descripcioproducte text(50)")
 bdllistat.Execute ("alter table temporal2 add column tipussoldadura text(50) ")
 bdllistat.Execute ("alter table temporal2 add column dtipusentrega text(50) ")
 bdllistat.Execute ("alter table temporal2 add column dmesures text(70) ")
 bdllistat.Execute ("alter table temporal2 add column dmesureslineals1 text(70) ")  'text(70)
 bdllistat.Execute ("alter table temporal2 add column dmesureslineals2 text(70)")
 bdllistat.Execute ("alter table temporal2 add column dcolorant text(70)")
 bdllistat.Execute ("alter table temporal2 add column dmaterial text(70)")
 bdllistat.Execute ("alter table temporal2 add column daditiu text(70)")
 bdllistat.Execute ("alter table temporal2 add column dextrusora text(30)")
 bdllistat.Execute ("alter table temporal2 add column dimpressora Text(30) ")
 bdllistat.Execute ("alter table temporal2 add column dlaminadora text(30)")
 bdllistat.Execute ("alter table temporal2 add column dsoldadora text(30)")
 bdllistat.Execute ("alter table temporal2 add column drebobinadora text(30)")
 bdllistat.Execute ("alter table temporal2 add column dcinta text(50)")
 bdllistat.Execute ("alter table temporal2 add column dansa text(50)")
 bdllistat.Execute ("alter table temporal2 add column dtruquel text(50)")
 bdllistat.Execute ("alter table temporal2 add column dmesurasoldadora text(30)")
 bdllistat.Execute ("alter table temporal2 add column dlotmatdesb1 text(50)")
 bdllistat.Execute ("alter table temporal2 add column dlotcolorantdesb1 text(50)")
 bdllistat.Execute ("alter table temporal2 add column dlotgalgadesb1 text(30)")
 bdllistat.Execute ("alter table temporal2 add column dlotimpressora text(50)")
 bdllistat.Execute ("alter table temporal2 add column dgalgaespsol text(30)")
 
 
 
 bdllistat.Execute ("alter table temporal2 add column dlotmatdesb2 text(50)")
 bdllistat.Execute ("alter table temporal2 add column dlotcolorantdesb2 text(50)")
 bdllistat.Execute ("alter table temporal2 add column dlotgalgadesb2 text(50)")
 
 bdllistat.Execute ("alter table temporal2 add column dmatintbob text(50)")
 
 bdllistat.Execute ("alter table temporal2 add column denduridor text(50)")
 bdllistat.Execute ("alter table temporal2 add column dresina text(50)")
 bdllistat.Execute ("alter table temporal2 add column descfamiliescoles text(100)")
 bdllistat.Execute ("alter table temporal2 add column dgrmcm2resina text(50)")
 bdllistat.Execute ("alter table temporal2 add column dgrmcm2enduridor text(50)")
 bdllistat.Execute ("alter table temporal2 add column dgrausresina text(50)")
 bdllistat.Execute ("alter table temporal2 add column dgrausenduridor text(50)")
 bdllistat.Execute ("alter table temporal2 add column dxresina text(50)")
 bdllistat.Execute ("alter table temporal2 add column dxenduridor text(50)")
 bdllistat.Execute ("alter table temporal2 add column dltsresina text(50)")
 bdllistat.Execute ("alter table temporal2 add column dltsenduridor text(50)")
 bdllistat.Execute ("alter table temporal2 add column dcoloradhesiu double")
 bdllistat.Execute ("alter table temporal2 add column daportcola text(50)")
 bdllistat.Execute ("alter table temporal2 add column horarientrega text(50)")
 bdllistat.Execute ("alter table temporal2 add column nomdestinatari text(50)")
 bdllistat.Execute ("alter table temporal2 add column adreça text(50)")
 bdllistat.Execute ("alter table temporal2 add column poblacio text(50)")
 bdllistat.Execute ("alter table temporal2 add column codipostal text(15)")
 bdllistat.Execute ("alter table temporal2 add column provincia text(30)")
 bdllistat.Execute ("alter table temporal2 add column telfclient text(30)")
 bdllistat.Execute ("alter table temporal2 add column faxclient text(20)")
 bdllistat.Execute ("alter table temporal2 add column numproveidor text(20)")
 bdllistat.Execute ("alter table temporal2 add column formapagament text(50)")
 
 bdllistat.Execute ("alter table temporal2 add column albvaloraticertqualitat text(40)")
 bdllistat.Execute ("alter table temporal2 add column refclient1 text(30)")
 bdllistat.Execute ("alter table temporal2 add column refclient2 text(30)")
 bdllistat.Execute ("alter table temporal2 add column bandesobertmicro text(20)")
 bdllistat.Execute ("alter table temporal2 add column obert2 text(20)")
 bdllistat.Execute ("alter table temporal2 add column micro text(20)")
 bdllistat.Execute ("alter table temporal2 add column dbobinesembolicades text(50)")
 bdllistat.Execute ("alter table temporal2 add column dtipusetiquetes text(50)")
 bdllistat.Execute ("alter table temporal2 add column peuimprenta text(50)")
 bdllistat.Execute ("alter table temporal2 add column sistemaimpresio text(15)")
  bdllistat.Execute ("alter table temporal2 add column nomfacturacio text(80)")
  bdllistat.Execute ("alter table temporal2 add column nomfotograbador text(30)")
  bdllistat.Execute ("alter table temporal2 add column treballsistemaimpresio text(15)")
  bdllistat.Execute ("alter table temporal2 add column altramaterial text(50)")
  bdllistat.Execute ("alter table temporal2 add column dataiusuari text(80)")
  bdllistat.Execute ("alter table temporal2 add column numerobossasoldadores text(20)")
  
  'formaimpresio del reprint
  bdllistat.Execute ("alter table temporal2 add column formaimpresioreprint text(15)")
 
  For i = 1 To 8
    bdllistat.Execute ("alter table temporal2 add column densitattinta" + atrim(i) + " double")
    bdllistat.Execute ("alter table temporal2 add column lintreball" + atrim(i) + " double")
    bdllistat.Execute ("alter table temporal2 add column detalltinter" + atrim(i) + " text(20)")
    bdllistat.Execute ("alter table temporal2 add column volumtinta" + atrim(i) + " double")
    bdllistat.Execute ("alter table temporal2 add column teextensiotinter" + atrim(i) + " text(50)")
    bdllistat.Execute ("alter table temporal2 add column colortinter" + atrim(i) + " double")
    
    'tintes reprint
    bdllistat.Execute ("alter table temporal2 add column nomtintareprint" + atrim(i) + " text(50)")
    bdllistat.Execute ("alter table temporal2 add column desarrollreprint" + atrim(i) + " double")
    bdllistat.Execute ("alter table temporal2 add column densitattintareprint" + atrim(i) + " double")
    bdllistat.Execute ("alter table temporal2 add column volumtintareprint" + atrim(i) + " double")
    bdllistat.Execute ("alter table temporal2 add column lintreballreprint" + atrim(i) + " double")
    
    
  Next i
  For i = 1 To 2
     bdllistat.Execute ("alter table temporal2 add column obstinta" + atrim(i) + " text(100)")
     bdllistat.Execute ("alter table temporal2 add column obstintabaixes" + atrim(i) + " text(100)")
  Next i
  
 
 'crear els camps pels consums laminadora
 bdllistat.Execute ("create table temporalconsums (id double);")
 For i = 1 To 16
  bdllistat.Execute ("alter table temporalconsums add column consuma" + Trim(i) + " double")
  bdllistat.Execute ("alter table temporalconsums add column consumb" + Trim(i) + " double")
 Next i

     
End Sub
Sub carregar_lookups_llistat(rsttmp As Recordset, rsttmp22 As Recordset)
  Dim rsttmp2 As Recordset
  Dim valtramat As String
  Dim invcol As String
  Dim invgalga As String
  Dim rstmat As Recordset
  
   llistatlookupde "clients", rsttmp!client, inventat, "nom"
   rsttmp22!nomclient = inventat
   llistatlookupde "clients", rsttmp!client, inventat, "telefon1"
   rsttmp22!telfclient = Mid(inventat, 1, 30)
   llistatlookupde "clients", rsttmp!client, inventat, "fax1"
   rsttmp22!faxclient = Mid(inventat, 1, 20)
   
  
  'LOOKUP DE producte
  Set rsttmp2 = dbtmp.OpenRecordset("select descripcio,ruta from productes where codi='" + atrim((rsttmp!producte)) + "'")
  If Not rsttmp2.EOF Then
     rsttmp22!descripcioproducte = atrim(rsttmp2!descripcio)
     rsttmp22!ruta = atrim(rsttmp2!ruta)
  End If
  'lookup de tipussoldadura
  Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from tipussoldadura where codi='" + atrim((rsttmp!tipusoldadura)) + "'")
  If Not rsttmp2.EOF Then
     rsttmp22!tipussoldadura = atrim(rsttmp2!descripcio)
  End If

llistatlookupde "tipusentregues", atrim(rsttmp!tipoentrega), inventat
rsttmp22!dtipusentrega = inventat
llistatlookupde "bobinesembolicades", atrim(rsttmp!rebidbobinesembolicades), inventat
rsttmp22!dbobinesembolicades = inventat
llistatlookupde "tipusetiquetes", atrim(rsttmp!rebidtipusetiqueta), inventat
rsttmp22!dtipusetiquetes = inventat

llistatlookupde "mesures", atrim(rsttmp!mesurapvp), inventat
rsttmp22!dmesures = inventat
llistatlookupde "mesureslineals", atrim(rsttmp!mesuraesp), inventat
rsttmp22!dmesureslineals1 = inventat
llistatlookupde "mesureslineals", atrim(rsttmp!mesuracantex), inventat
rsttmp22!dmesureslineals2 = inventat
If cadbl(rsttmp!materialex) < 500 Then
  llistatlookupde "colorants", atrim(rsttmp!colorex), inventat
  rsttmp22!dcolorant = inventat
  llistatlookupde "materials", atrim(rsttmp!materialex), inventat
  rsttmp22!dmaterial = inventat
  'llistatlookupde "aditius", atrim(rsttmp!aditiuex), inventat
  'rsttmp22!daditiu = inventat
    Else
       Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(rsttmp!materialex))
       If Not rstmat.EOF Then
          rsttmp22!dmaterial = Mid(descripciomaterial(rstmat) + " ", 1, 70)
          If nomcolor(23) <> "" Then rsttmp22!daditiu = atrim(rstmat!proveidor) + "-" + atrim(rstmat!refproducte)
       End If
       Set rstmat = Nothing
End If
llistatlookupde "select descripcio from maquines where maquina='E' and codi=" + atrim(cadbl((rsttmp!extrusora))), , inventat
rsttmp22!dextrusora = Mid(inventat, 1, 30)
llistatlookupde "select descripcio from maquines where maquina='I' and codi=" + atrim(cadbl((rsttmp!impressora))), , inventat
rsttmp22!dimpressora = Mid(inventat, 1, 30)
llistatlookupde "select descripcio from maquines where maquina='L' and codi=" + atrim(cadbl((rsttmp!laminadora))), , inventat
rsttmp22!dlaminadora = Mid(inventat, 1, 30)
llistatlookupde "select descripcio from maquines where maquina='S' and codi=" + atrim(cadbl((rsttmp!soldadora))), , inventat
rsttmp22!dsoldadora = Mid(inventat, 1, 30)
llistatlookupde "select descripcio from maquines where maquina='R' and codi=" + atrim(cadbl((rsttmp!rebobinadora))), , inventat
rsttmp22!drebobinadora = Mid(inventat, 1, 30)
llistatlookupde "accessoris", atrim(rsttmp!cinta), inventat
rsttmp22!dcinta = Mid(inventat, 1, 50)
llistatlookupde "accessoris", atrim(rsttmp!ansa), inventat
rsttmp22!dansa = Mid(inventat, 1, 50)
llistatlookupde "accessoris", atrim(rsttmp!troquel), inventat
rsttmp22!dtruquel = Mid(inventat, 1, 50)
impr = 0
llistat_possar_desc_lot atrim(rsttmp!lotmatdesb1), inventat, invcol, invgalga, impr
rsttmp22!dlotmatdesb1 = inventat
rsttmp22!dlotcolorantdesb1 = invcol
rsttmp22!dlotgalgadesb1 = invgalga + " " + r
rsttmp22!dlotimpressora = impr
invcol = ""
invgalga = ""
llistat_possar_desc_lot atrim(rsttmp!lotmatdesb2), inventat, invcol, invgalga, impr
rsttmp22!dlotmatdesb2 = inventat
rsttmp22!dlotcolorantdesb2 = invcol
rsttmp22!dlotgalgadesb2 = invgalga + " " + r
valtramat = descripcioaltramaterial(cadbl(rsttmp!lotmatdesb1), cadbl(rsttmp!lotmatdesb2), cadbl(Data1.Recordset!comanda), cadbl(Data1.Recordset!linkcomanda1), cadbl(Data1.Recordset!linkcomanda2))
If Len(valtramat) > 50 Then valtramat = Mid(valtramat, 1, 50)
rsttmp22!altramaterial = valtramat

'miro si hi ha alguna amb seccio impressora
If impr = 1 Then impr = 2
If rsttmp22!dlotimpressora = 0 Then rsttmp22!dlotimpressora = impr
'POSSO ELS CONSUMS
Set rsttmp2 = bdllistat.OpenRecordset("select * from temporalconsums")
rsttmp2.AddNew
For i = 0 To 15
  reixaconsums.row = 1
  reixaconsums.col = i
  rsttmp2.Fields("consuma" + Trim(i + 1)) = reixaconsums.Text
  reixaconsums.row = 2
  rsttmp2.Fields("consumb" + Trim(i + 1)) = reixaconsums.Text
Next i
rsttmp2.Update
inventat = ""
invcol = ""
If cadbl(rsttmp!matintbob) > 0 Then
  llistat_possar_desc_lot atrim(rsttmp!matintbob), inventat, invcol, invgalga
  rsttmp22!dmatintbob = inventat + " " + invcol
 Else: rsttmp22!dmatintbob = ""
End If
llistatlookupde "mesureslineals", atrim(rsttmp!unitatespsol), inventat
rsttmp22!dmesurasoldadora = inventat
llistat_possar_noms_adhesius True, rsttmp22

'mesures de soldadora
Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp!unitatespsol)))
If Not rsttmp2.EOF Then rsttmp22!dgalgaespsol = atrim(rsttmp2!descripcio)
'possarconsums
Set rsttmp2 = Nothing
End Sub
Function descripcioaltramaterial(desb1 As Double, desb2 As Double, lot1 As Double, lot2 As Double, lot3 As Double) As String
  Dim invcol As String
  Dim invgalga As String
  Dim lotmatmig As Double
  Dim lotmatmin As Double
  Dim lotmatmax As Double
  If (lot1 + lot2 + lot3) = (lot1 - lot2 + lot3) Then Exit Function
  If lot1 = 0 Or lot2 = 0 Or lot3 = 0 Then lotmatmig = IIf(lot1 > lot2, lot1, lot2): GoTo fi
  lotmatmax = IIf(lot1 > lot2, lot1, lot2)
  lotmatmax = IIf(lotmatmax > lot3, lotmatmax, lot3)
  lotmatmin = IIf(lot1 > lot2, lot2, lot1)
  lotmatmin = IIf(lotmatmin > lot3, lot3, lotmatmin)
  lotmatmig = (lot1 + lot2 + lot3) - (lotmatmin + lotmatmax)
fi:
  llistat_possar_desc_lot atrim(lotmatmig), inventat, invcol, invgalga, impr
  
  'If lot1 = desb1 Or lot1 = desb2 Then lot1 = 0
  'If lot2 = desb1 Or lot2 = desb2 Then lot2 = 0
  'If lot3 = desb1 Or lot3 = desb2 Then lot3 = 0
  'If lot1 > 0 Then llistat_possar_desc_lot atrim(lot1), inventat, invcol, invgalga, impr
  'If lot2 > 0 Then llistat_possar_desc_lot atrim(lot2), inventat, invcol, invgalga, impr
  'If lot3 > 0 Then llistat_possar_desc_lot atrim(lot3), inventat, invcol, invgalga, impr
  descripcioaltramaterial = inventat + " - " + invcol

End Function
Sub possarreduccionscilindre()
  Dim msg As String
   Label1(31) = UCase(buscarreducciocilindretreball(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), "fw", True))
   Label1(31) = Label1(31) + UCase(buscarreducciocilindretreball(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), "f2"))
End Sub
Sub llistat_possar_noms_adhesius(Optional lookup As Boolean, Optional rsttmp22 As Recordset)
    Dim rsttmp2 As Recordset
    Dim vsql As String
    vsql = "SELECT adhesius.*, familiescoles.descripcio AS descfamcola, subfamiliescoles.descripcio AS descsubfamcola FROM (adhesius LEFT JOIN familiescoles ON adhesius.idfamilia = familiescoles.codi) LEFT JOIN subfamiliescoles ON adhesius.idsubfamilia = subfamiliescoles.codi "
    Set rsttmp2 = dbtmp.OpenRecordset(vsql + " where adhesius.codi=" + atrim(cadbl(rsttmp!tipusadhesiu)))
    If Not rsttmp2.EOF Then
      rsttmp22!denduridor = atrim(rsttmp2!enduridor)
      rsttmp22!descfamiliescoles = IIf(atrim(rsttmp2!descfamcola) <> "", UCase(atrim(rsttmp2!descsubfamcola) + " - " + atrim(rsttmp2!descfamcola)), "")
      rsttmp22!dresina = atrim(rsttmp2!resina)
      rsttmp22!dgrmcm2resina = cadbl(rsttmp2!grmcm3_resina)
      rsttmp22!dgrmcm2enduridor = cadbl(rsttmp2!grmcm3_enduridor)
      rsttmp22!dgrausresina = cadbl(rsttmp2!grausresina)
      rsttmp22!dgrausenduridor = cadbl(rsttmp2!grausenduridor)
      rsttmp22![dxresina] = cadbl(pes1) 'cadbl(rsttmp2![%resina])
      rsttmp22![dxenduridor] = cadbl(pes2) 'cadbl(rsttmp2![%enduridor])
      rsttmp22!daportcola = cadbl(grmt2)
      rsttmp22!dltsresina = cadbl(litres1(1))
      rsttmp22!dltsenduridor = cadbl(litres2(2))
      rsttmp22!dcoloradhesiu = possarcoloradhesiu(IIf(atrim(rsttmp2!predeterminada) = "S", "BLANC", atrim(rsttmp2!color)))
        Else: rsttmp22!dcoloradhesiu = possarcoloradhesiu("BLANC")
    End If
    Set rsttmp2 = Nothing
End Sub

Sub llistat_possar_desc_lot(numlot As String, desclotx As Control, descolorant As String, galga As String, Optional impr)
  Dim desctmp As String
  Dim rsttmp2 As Recordset
  Dim rsttmp3 As Recordset
  desctmp = ""
  desclotx = desctmp
  If cadbl(numlot) < 1 Then Exit Sub
  Set rsttmp2 = dbtmp.OpenRecordset("select tubolam,producte,materialex,colorex,espessor,mesuraesp from comandes where comanda=" + atrim(cadbl(numlot)))
  If Not rsttmp2.EOF Then
     Set rsttmp3 = dbtmp.OpenRecordset("select descripcio,ruta from productes where codi='" + atrim(rsttmp2!producte) + "'")
     If Not rsttmp3.EOF Then If InStr(1, UCase(ruta), "I") <> 0 Then impr = 1
     Set rsttmp3 = dbtmp.OpenRecordset("select descripcio,familiacol from materials where codi=" + atrim(cadbl(rsttmp2!materialex)))
     If Not rsttmp3.EOF Then
       desctmp = rsttmp3!descripcio
       Set rsttmp3 = dbtmp.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(rsttmp3!familiacol)))
     End If
     If Not rsttmp3.EOF Then descolorant = atrim(rsttmp3!descripcio)
     Set rsttmp3 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp2!mesuraesp)))
     r = ""
     If Not rsttmp3.EOF Then
        galga = atrim(rsttmp2!espessor) + " " + rsttmp3!descripcio
        If rsttmp3!descripcio = "GALGUES" Then
            If rsttmp2!tubolam = "T" Then
                 r = Format(rsttmp2!espessor / 4, "#,##0") + " Mic"
                  Else: r = Format(rsttmp2!espessor / 2, "#,##0") + " Mic"
            End If
        End If
     End If
  End If
  desclotx = desctmp
  Set rsttmp2 = Nothing
  Set rsttmp3 = Nothing
End Sub



Private Sub consultar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim v As String
  If Button = 2 And Shift = 2 Then
     v = InputBox("Entra el valor de la variable campsvalids: ", "Atenció", "[tots]")
     If atrim(v) <> "" Then
         escriure_ini "General", "llistadecampsvalids", v, "comandes.ini"
     End If
  End If
End Sub

Private Sub ctipusimp_Click()
  Text65.Text = Mid(ctipusimp.Text, 1, 1)
End Sub

Private Sub ctipusimp_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Data1_Reposition()
  Dim c As Double
  'Set rsttmp = dbtmp.OpenRecordset("select comanda from comandes where client=0 or cantitatex=null")
  'If rsttmp.EOF Then
  ' clients0.Visible = False
  '  Else: clients0.Visible = True
  'End If
 ' possarvalordcamps
  Data1.BackColor = QBColor(12)
  DoEvents
  
  Data1.Recordset.LockEdits = False
  dataentrega2.Text = ""
  Frame1(1).Visible = False
  Text5.Enabled = True
  Data1.BackColor = QBColor(15)
  MaskEdBox12.BackColor = QBColor(15)
  ensenyar_etiqueta_vandamme False
  activaronocampsimpresio True
  carregant = True
  If formcomandes.Tag <> "100" Then carregar_lookups
  
  If Data1.Recordset.EditMode = 0 And formcomandes.Tag <> "100" Then situar_seccions
  Data1.Caption = "" + atrim(Data1.Recordset.AbsolutePosition + 1) + "/" + atrim(Data1.Recordset.RecordCount)
  If Not Data1.Recordset.EOF Then
    vdirenvio = ""
    If InStr(1, Label1(147), ":") > 0 Then vdirenvio = Trim(Mid(Label1(147), InStr(1, Label1(147) + " ", ":") + 1))
    formcomandes.Caption = "Comanda:    " + Text2 + " - " + nomclient + " (Envio a: " + vdirenvio + ")" + "      | Lot: " + Text1
  End If
  If cadbl(Data1.Recordset.RecordCount) = 0 Then Data1.Caption = "Cap Resultat": Data1.BackColor = QBColor(12)
  'On Error Resume Next
 ' If atrim(Data1.Recordset!proximaseccio) = "" And Not Data1.Recordset.EOF Then
 '   If Data1.Recordset.EditMode = 0 Then i = 1: Data1.Recordset.Edit
 '   Data1.Recordset!proximaseccio = "E"
 '   If i = 1 Then Data1.Recordset.Update
 ' End If
  If atrim(dataactivacio.Text) = "" Then
          dataactivacio.BackColor = QBColor(12)
      Else: dataactivacio.BackColor = QBColor(15)
  End If
  'coloco les variables de linkcomanda
  c = cadbl(Data1.Recordset!comanda)
  'If c <> vlink1 And c <> vlink2 And c <> vlink3 Then
    vlink1 = c
    vlink2 = cadbl(Data1.Recordset!linkcomanda1)
    vlink3 = cadbl(Data1.Recordset!linkcomanda2)
    If vlink2 < vlink1 Then vlink1 = vlink2: vlink2 = c
    If vlink3 > 0 Then
     If vlink3 < vlink1 Then c = vlink1: vlink1 = vlink3: vlink3 = c
     If vlink3 < vlink2 Then c = vlink2: vlink2 = vlink3: vlink3 = c
    End If
  'End If
  ' fins aqui
  Data1.BackColor = QBColor(15)
  Text32(12).Visible = False
  Text32(15).Visible = False
  'comandesamblamateixareferenciainplacsa data1.Recordset
  carregant = False
  'trec lo de comprovar que estigui impres perquè ens dona problemes 20/03/22
  'comprovarsinoshaimprespassaraE
  If isloaded("formfirmes") Then formfirmes.refrescar_firmes
  'If isloaded("formcomandes") Then f
  'ACTIVO AQUEST VARIABLE PER TANCAR EL FORM D'IMPRIMIR EL PACKINGLIST PER PANTALLA (PER SI DE CAS)
  escriure_ini "baixes", "imprimirpackinglist", "0", "comandes.ini"
End Sub
Sub comprovarsinoshaimprespassaraE()
  If Not Data1.Recordset.EOF Then
    If Data1.Recordset!comanda > 0 And Data1.Recordset!proximaseccio <> "E" And Command9(0).BackColor <> &HC0FFC0 Then
      If UCase(InputBox("Aquesta comanda no s'ha llistat per impresora encara i la seccio no es Extrusora, la canviaré automàticament a 'E'." + vbNewLine + "ESCRIU [EXTRUSORES] PER PASSARLA.", vbCritical, "Error")) <> "EXTRUSSORES" Then GoTo fi
      Data1.Recordset.Edit
      Data1.Recordset!proximaseccio = "E"
      Data1.Recordset.Update
      If cadbl(Data1.Recordset!linkcomanda1) <> 0 Then dbtmp.Execute "update comandes set proximaseccio='E' where comanda=" + atrim(Data1.Recordset!linkcomanda1)
      If cadbl(Data1.Recordset!linkcomanda2) <> 0 Then dbtmp.Execute "update comandes set proximaseccio='E' where comanda=" + atrim(Data1.Recordset!linkcomanda2)
    End If
  End If
fi:
End Sub
Sub situar_seccions()
  Dim sec(9, 2)
  Dim ultimapos As Double
  Dim marge As Double
  marge = 150
  
  'amago tots els botons i seccions
  ext.Visible = False
  imp1.Visible = False
  lam1.Visible = False
  sol.Visible = False
  reb.Visible = False
  Command10.Visible = True: Command12.Visible = False: Command13.Visible = False
  Command14.Visible = False: Command15.Visible = False: Command18.Visible = False
  Command19.Visible = False: Command16.Visible = False: Command17.Visible = False
  Command20.Visible = False: Command21.Visible = False: Command1(0).Visible = False
  Command9(7).Visible = False
  
  If ruta = "" Then Exit Sub
  
  ultimapos = formscrooll.Top
  ultimapos = cap.Height + cap.Top
  For i = 1 To 10
    taulapos(i) = taulapos(0) + ultimapos
    If Mid(ruta, i, 1) = "" Then taulapos(i) = 0
    Select Case Mid(ruta, i, 1)
      Case "E"
         ext.Visible = True
         ext.Top = ultimapos + marge
         ultimapos = ultimapos + marge + ext.Height
         Command12.Visible = True: Command13.Visible = True
         Command12.Top = ext.Top + ext.Height - 570
         Command13.Top = Command12.Top + Command12.Height
         ext.Tag = taulapos(0) + ultimapos
      Case "I"
         imp1.Visible = True
         imp1.Top = ultimapos + marge
         ultimapos = ultimapos + marge + imp1.Height
         Command14.Visible = True: Command15.Visible = True
         Command1(0).Visible = True
         Command14.Top = imp1.Top + imp1.Height - 570
         Command15.Top = Command14.Top + Command14.Height
         Command1(0).Top = imp1.Top + Command1(4).Top
         imp1.Tag = taulapos(0) + ultimapos
      Case "L"
         lam1.Visible = True
         lam1.Top = ultimapos + marge
         ultimapos = ultimapos + marge + lam1.Height
         Command18.Visible = True: Command19.Visible = True
         Command18.Top = lam1.Top + lam1.Height - 570
         Command19.Top = Command18.Top + Command18.Height
         lam1.Tag = taulapos(0) + ultimapos
      Case "R"
         reb.Visible = True
         reb.Top = ultimapos + marge
         ultimapos = ultimapos + marge + reb.Height
         Command16.Visible = True: Command17.Visible = True
         Command16.Top = reb.Top + reb.Height - 570
         Command17.Top = Command16.Top + Command16.Height
      Case "S"
         sol.Visible = True
         sol.Top = ultimapos + marge
         ultimapos = ultimapos + marge + sol.Height
         Command20.Visible = True: Command21.Visible = True
         Command21.Top = sol.Top + sol.Height - 570
         Command20.Top = Command21.Top + Command21.Height
         Command9(7).Visible = True
         Command9(7).Top = sol.Top + 450
    End Select
    
  Next i
End Sub
Sub triarclient()
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from clients"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text2.Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   Data1.Recordset!client = Text2.Text
   nomclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
   If Not buscant Then possar_direccio_envio
  End If
  Unload formseleccio
  
End Sub
Sub possar_direccio_envio(Optional nomissatge As Boolean)
  If buscant Then nomissatge = True
  If Data1.Recordset.EditMode = 0 Then Exit Sub
   Set rsttmp = dbtmp.OpenRecordset("select * from clients_envios where codi=" + atrim(cadbl(Text2.Text)) + " and id=" + atrim(cadbl(Data1.Recordset!direnvio)))
   If Not rsttmp.EOF Then Exit Sub
   Data1.Recordset!direnvio = 0
   Data1.Recordset!client = cadbl(Text2.Text)
   Set rsttmp = dbtmp.OpenRecordset("select * from clients_envios where codi=" + atrim(cadbl(Text2.Text)))
   If Not rsttmp.EOF Then
      rsttmp.MoveLast
      If rsttmp.RecordCount = 1 Then
         Data1.Recordset!direnvio = rsttmp!ID
        Else: If Not nomissatge Then MsgBox "Hi ha mes d'una direcció d'enviament has d'escullir una"
      End If
  End If
End Sub
Sub triarmesura()
  Load formseleccio
  formseleccio.Caption = "Triar Unitat Mesura"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from mesures"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 1200
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text7.Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   Text16.Text = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub

Sub triarmesuraquantitat()
  Load formseleccio
  formseleccio.Caption = "Triar Unitat Mesura"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from mesureslineals"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 1200
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text31.Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   Text30.Text = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub
Sub triarmesuraquantitatdesitjada()
  Load formseleccio
  formseleccio.Caption = "Triar Unitat Mesura"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from mesureslineals"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 1200
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text32(8).Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   Text32(7).Text = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub

Sub triarmesuraespesor()
  Load formseleccio
  formseleccio.Caption = "Triar Unitat Mesura"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from mesureslineals where codi=10 or codi=11"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 1200
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text23.Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   Text22.Text = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub



'Sub triarextrussora()
'  Load formseleccio
'  formseleccio.Caption = "Triar Màquina Extrussora"
'  formseleccio.data1.DatabaseName = data1.DatabaseName
'  formseleccio.data1.RecordSource = "select * from maquines where maquina='E' order by codi"
'  formseleccio.refrescar
 ' formseleccio.Show 1
 ' If seleccioret = 1 Then
 '  Text27.Text = atrim(formseleccio.data1.Recordset!codi)
 ''  nomextrussora(0).Caption = atrim(formseleccio.data1.Recordset!descripcio)
 ' End If
 ' Unload formseleccio
  
'End Sub

Sub triaralgu(titol As String, taula As String, control1 As Control, control2 As Control, Optional Camp As String, Optional anularcolsel As Byte)
  If atrim(Camp) = "" Then Camp = "descripcio"
  Load formseleccio
  If cadbl(anularcolsel) > 0 Then formseleccio.Tag = "1"
  formseleccio.Caption = titol
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = IIf(Len(taula) < 10, "select * from " + taula, taula)
  formseleccio.refrescar
  formseleccio.Command3.Tag = "filtre"
  If r = "triaraccessoris" Then
   formseleccio.DBGrid2.Columns(2).Width = 4000
   formseleccio.Text1.Tag = "2"
   r = 0
  End If
  If Camp = "resina" Then
    formseleccio.Width = 11500
    
    formseleccio.DBGrid2.Columns(2).Width = 900
    formseleccio.DBGrid2.Columns(3).Width = 2200
    formseleccio.DBGrid2.Columns(4).Width = 2200
    formseleccio.DBGrid2.Columns(5).Width = 2200
    formseleccio.DBGrid2.Columns(6).Width = 2200
'    formseleccio.DBGrid2.Columns(7).Width = 2000
    formseleccio.Data1.Recordset.FindFirst "predeterminada='S'"
    formseleccio.Caption = "Triar Adhesiu"
  End If
  formseleccio.Show 1
  If seleccioret = 1 Then
'   On Error Resume Next
   control1 = atrim(formseleccio.Data1.Recordset!codi)
   control2 = atrim(formseleccio.Data1.Recordset.Fields(Camp))
  End If
  Unload formseleccio
  
End Sub

Sub triarproducte()
  Load formseleccio
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from productes"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text3.Text = atrim(formseleccio.Data1.Recordset!codi)
   nomproducte.Caption = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub


Sub triartipusentrega()
  Load formseleccio
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from tipusentregues"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text11.Text = atrim(formseleccio.Data1.Recordset!codi)
   Label3.Caption = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub




Private Sub consultar_Click()
  botoconsultar
End Sub
Sub botoconsultar()
     
 If Data1.Recordset.EditMode > 0 Then MsgBox "Primer finalitza l'edició actual": Exit Sub
  Data1.RecordSource = "select * from comandes where comanda=0"
  'Data1.Refresh
  vtreballbuscatsubbusqueda = ""
  Unload formbusquedahabitual
  formbusquedahabitual.Show 1
  If formbusquedahabitual.bproximabusqueda.Tag <> "1" Then Unload formbusquedahabitual: Exit Sub
  Unload formbusquedahabitual
  buscant = True
  i = 0
  activaronocampsimpresio True
'  On Error GoTo carregaform
  subbusqueda.Visible = True
 ' On Error Resume Next
  'If subbusqueda.Tag = "" Then subbusqueda.Show 1
  While subbusqueda.Visible
    DoEvents
  Wend
  If r = "sortir" Then r = "": buscant = False: Exit Sub
  querywhere = ""
  queryorder = ""
  
  If i = 1 Then
   busqueda_xr_formulari
     Else
       'executar la consulta de la variable r
     querywhere = r
     finalitzarbusqueda 1
  End If
  'ext.Visible = True
  'imp1.Visible = True
  'lam1.Visible = True
  'sol.Visible = True
  'reb.Visible = True
  Exit Sub
End Sub
Sub busqueda_xr_formulari()
   alta_registre
   deixartotblanc
   ruta = "EILRS"
   situar_seccions
   cimpressio.Clear
   cimpressio.AddItem "Nova"
   cimpressio.AddItem "Modificada"
   cimpressio.AddItem "Repetida"
   cimpressio.AddItem "Falta Autoritzar"
   On Error Resume Next
   Text1.SetFocus
   activaronocampsimpresio True
End Sub
Private Sub Data1_Validate(Action As Integer, Save As Integer)
 Dim Control As Control
 
Save = False

'If Not buscant And Data1.Recordset.RecordCount > 0 Then

 'Data1.Recordset.Edit
'  For Each control In formcomandes
'       If telapropietatdatachanged(control) Then
'        If control.DataChanged And control.DataField <> "" Then
'            If Data1.Recordset.Fields(control.DataField).Type = 8 And control = "" Then
'                Data1.Recordset.Fields(control.DataField) = Null
'                  Else: Data1.Recordset.Fields(control.DataField) = control
'             End If
'
'
'            control.DataChanged = False
'
'        End If
'       End If
'   Next
'  Data1.Recordset.Update
'End If
   If Action = 2 Or Action = 3 Then
'     data1.Recordset!descquantitat = quantitatdelacomanda
     areadedatos False
   End If
End Sub
Function telapropietatdatachanged(contrl As Control) As Boolean
   On Error GoTo err
   If contrl.DataChanged <> "a" Then telapropietatdatachanged = True
   telapropietatdatachanged = True
   On Error GoTo 0
   Exit Function
err:
   telapropietatdatachanged = False
   On Error GoTo 0
End Function
Function quantitatdelacomanda() As String
  Dim ultima As String
  ultima = Mid(ruta, Len(ruta), 1)
  If ultima = "I" Then ultima = e
  If ultima = "E" Then
      quantitatdelacomanda = atrim(Data1.Recordset!cantitatex) + Text30
  End If
  If ultima = "R" Then
      quantitatdelacomanda = rebpes + " Kgs."
  End If
  If ultima = "S" Then
      quantitatdelacomanda = Data1.Recordset!cantitatsol + " Unitats"
  End If
End Function

Private Sub data1tmp_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub dataactivacio_LostFocus()
  If dataactivacio = "" And Data1.Recordset.EditMode > 0 Then
     Data1.Recordset!dataactivacio = Null
     dataactivacio.DataChanged = False
  End If
End Sub
Function hiharelacions(numc As Double, numc2 As Double, numc3 As Double) As Boolean
   If numc > 0 Then hiharelacions = hiharelacionsunacomanda(numc)
   If Not hiharelacions And numc2 > 0 Then hiharelacions = hiharelacionsunacomanda(numc2)
   If Not hiharelacions And numc3 > 0 Then hiharelacions = hiharelacionsunacomanda(numc3)
End Function

Function hiharelacionsunacomanda(numc As Double) As Boolean
  Dim dbt As Database
  Dim rstt As Recordset

  Set dbt = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
  reservaassignacioocompra = "COMPRES"
  Set rstt = dbt.OpenRecordset("select * from comandesxlinia where numcomanda=" + atrim(numc))
  If Not rstt.EOF Then hiharelacionsunacomanda = True: GoTo fi
  
  Set dbt = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
    reservaassignacioocompra = "ASSIGNACIO"
  Set rstt = dbt.OpenRecordset("select * from parcials where comanda='" + atrim(numc) + "'")
  If Not rstt.EOF Then hiharelacionsunacomanda = True: GoTo fi
  
  Set rstt = dbt.OpenRecordset("select * from percomandaoclient where numcomanda=" + atrim(numc))
    reservaassignacioocompra = "RESERVA"
  If Not rstt.EOF Then hiharelacionsunacomanda = True: GoTo fi
  hiharelacionsunacomanda = False
fi:
  Set dbt = Nothing
  Set rstt = Nothing

End Function
Private Sub dataentrega2_DblClick()
  Dim d As String
  d = InputBox("Entra la segona data d'entrega.", "Entrada de data")
  If IsDate(d) Then
     dataentrega2 = Format(d, "dd/mm/yy")
     dbplanificacio.Execute "update  planificacioimp set imp_data2=#" + Format(d, "mm/dd/yy") + "# where comanda=" + atrim(Text1)
  End If
End Sub

Private Sub eliminar_Click()
Dim possicioreg As Variant
Dim un As Double
Dim dos As Double
Dim tres As Double
 On Error GoTo err
 ratoli "espera"
 possicioreg = Data1.Recordset.Bookmark
 
 un = cadbl(Data1.Recordset!comanda)
 dos = cadbl(Data1.Recordset!linkcomanda1)
 tres = cadbl(Data1.Recordset!linkcomanda2)
 If Data1.Recordset!producte = "PC" Or Data1.Recordset!producte = "PC2" Then MsgBox "No pots eliminar un PC o PC2 has de sel.lecionar la comanda principal.", vbCritical, "Atenció": ratoli "normal": Exit Sub
 If dos > 0 And (un + 1) <> dos Then MsgBox "Els numeros de comanda complexes no son correlatius no es pot duplicar.", vbCritical, "Error": Exit Sub
 If tres > 0 And (un + 2) <> tres Then MsgBox "Els numeros de comanda complexes no son correlatius no es pot duplicar.", vbCritical, "Error": Exit Sub
 
 If hiharelacions(cadbl(Data1.Recordset!comanda), cadbl(Data1.Recordset!linkcomanda1), cadbl(Data1.Recordset!linkcomanda2)) Then MsgBox "Aquesta comanda encara te RESERVES, ASSIGNACIONS O COMPRES assignades no es pot eliminar fins que totes estiguin eliminades.", vbCritical + vbOKOnly, "Atenció": ratoli "normal": Exit Sub
 ratoli "normal"
 If InputBox("Segur que vols donar de baixa aquesta comanda?" + Chr$(13) + "Escriu... [eliminar] ...per eliminar la comanda" + Chr(10) + Chr(13) + IIf(dos > 0, "ES UNA COMANDA COMPLEXE, S'ELIMINARAN LES RELACIONADES TAMBÉ", ""), "Atenció Borrant...") = "eliminar" Then
 ratoli "espera"
   Frame1(0).Enabled = False
   'areadatos.Enabled = False
   ratoli "espera"
   If un > 0 Then
      'si hi ha impresora avisar a tintes que la comanda s'ha anulat si ja s'ha passat a impresa BOTO VERD IMPRIMIR
      
    If Command9(0).Tag = "Siimpresa" Then
      If larutahiha(Data1.Recordset!producte, "I") Then
        enviaremailgeneric "tintes@inplacsa.com", "La comanda " + atrim(un) + " s´ha eliminat de producció.", "Revisa si hi ha tinta preparada per aquesta comanda, el client ha anulat la comanda " + atrim(un) + "." + Chr(10) + Chr(13) + "Client: " + atrim(nomclient) + Chr(10) + Chr(13) + "Texte Impresió: " + atrim(Data1.Recordset!marcailinia)
      End If
    End If
   ' gravar_avis_alicia "COMANDA: " + atrim(data1.Recordset!comanda) + "   ELIMINADA"
    Data1.Recordset.Delete
    dbtmp.Execute "delete * from comandes_extres where comanda=" + atrim(un), 512
    Data1.RecordSource = "select * from comandes where comanda=" + atrim(un)
    refrescar
    wait (2)
   End If
   If dos > 0 Then
    'gravar_avis_alicia "COMANDA: " + atrim(dos) + "   ELIMINADA"
    dbtmp.Execute "delete * from comandes where comanda=" + atrim(dos), 512
    dbtmp.Execute "delete * from comandes_extres where comanda=" + atrim(dos), 512
   End If

   If tres > 0 Then
    'gravar_avis_alicia "COMANDA: " + atrim(tres) + "   ELIMINADA"
    dbtmp.Execute "delete * from comandes where comanda=" + atrim(tres), 512
    dbtmp.Execute "delete * from comandes_extres where comanda=" + atrim(tres), 512
   End If
   dbtmp.Execute "delete * from comandes_controlcanvis where comanda in (" + atrim(un) + "," + atrim(dos) + "," + atrim(tres) + ")"
   wait 2
   Data1.RecordSource = "select * from comandes order by comanda Desc"
   refrescar
  End If
  Frame1(0).Enabled = True
   'areadatos.Enabled = False
   ratoli "normal"
 Exit Sub
err:
  MsgBox "No s'ha pogut eliminar possiblement perque tingui registres relacionats. O bé no hi ha res per eliminar."
  Frame1(0).Enabled = True
 ' Resume Next
   'areadatos.Enabled = False
   ratoli "normal"
End Sub



Private Sub etrebmetres_Click()

End Sub
Sub possardadesdeltreballazero(rstc As Recordset)
   Dim i As Byte
   rstc.Edit
   For i = 1 To 8
          rstc.Fields("tinta" + atrim(i) + "a") = ""
          rstc.Fields("lin" + atrim(i)) = "0"
          rstc.Fields("tinta" + atrim(i) + "b") = ""
          rstc!continu = "N"
          rstc!cilindres = 0
          rstc!dessarroll = 0
          
   Next i
   rstc!arxiumontadora = ""
   rstc!arxiu = ""
   rstc!formaimp = ""
   rstc!gruixpol = 0
   rstc!codibarras = ""
   rstc!cmaquina = ""
   rstc!marcailinia = ""
   rstc.Update
End Sub
Sub posardiferenciesacomandadeltreball(numc As Double)
   Dim i As Byte
   Dim cilindre As Double
   Dim desarroll As Double
   Dim continu As String
   Dim rstc As Recordset
   Dim rstclixe As Recordset
   Dim rstmodificacio As Recordset
   Dim rstlink As Recordset
   Dim rsttintes As Recordset
   Dim treball As Integer
   Dim arxiumontadora As String
   Dim modificacio As Integer
   Dim rsttintesllaunes As Recordset
   Dim vtinters As Byte
   Dim vnomtinta As String
   
   Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
   If rstc.EOF Then GoTo noinfo
   treball = cadbl(rstc!numtreball): modificacio = cadbl(rstc!numordremodificacio)
   'If treball = 0 Then possardadesdeltreballazero rstc: GoTo fi
   If modificacio = 0 Then modificacio = 1
   Set rstclixe = dbclixesnous.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(rstc!numtreball)))
   If rstclixe.EOF Then GoTo noinfo
   Set rstmodificacio = dbclixesnous.OpenRecordset("select * from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(cadbl(modificacio)))
   If rstmodificacio.EOF Then GoTo noinfo
   rstc.Edit
   Set rsttintes = dbclixesnous.OpenRecordset("select * from tintes where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordremodificacio=" + atrim(cadbl(modificacio)) + " order by ordretinter")
   Set rsttintesllaunes = dbclixesnous.OpenRecordset("select codi,descripcio from tintes_llaunes", , dbonly)
   'comprovo tintes
   continu = ""
   If Not rsttintes.EOF Then
      rstc!dessarroll = 0
      For i = 1 To 8
        rsttintes.FindFirst "ordretinter = " + atrim(i)
        If Not rsttintes.NoMatch Then
          If atrim(rsttintes!color) <> "" Or cadbl(rsttintes!tinterlinkambid_treball) > 0 Then vtinters = vtinters + 1
          Set rstlink = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(IIf(rsttintes!tinterlinkambid_treball > 0, rsttintes!tinterlinkambid_treball, rsttintes!id_tinter)))
          rsttintesllaunes.FindFirst "codi='" + atrim(rstlink!coditinta) + "'"
          If rsttintesllaunes.NoMatch And atrim(rsttintes!color) = "" Then posartinterdecomandaazero rstc, i: GoTo proxima
          vnomtinta = ""
          If atrim(rsttintes!color) <> "" And cadbl(rsttintes!tinterlinkambid_treball) = 0 Then
              vnomtinta = atrim(rsttintes!color)
                Else: vnomtinta = atrim(rsttintesllaunes!descripcio)
          End If
          rstc.Fields("tinta" + atrim(i) + "a") = vnomtinta
          rstc.Fields("lin" + atrim(i)) = atrim(rstlink!anilox)
          'rstc.Fields("tinta" + atrim(i) + "b") = atrim(rstlink!observacions)
          rstc.Fields("detalltinter" + atrim(i)) = atrim(rstlink!detalltinter)
          If atrim(rstlink!color) <> "" Or i = 1 Then
            If continu <> "S" Then
               rstc!continu = IIf(Not rstlink!continuu, "N", "S")
                Else: continu = rstc!continu
            End If
            If cadbl(rstlink!cilindre) <> 0 Then rstc!cilindres = cadbl(rstlink!cilindre)
            If cadbl(rstlink!desarroll) > cadbl(rstc!dessarroll) Then rstc!dessarroll = cadbl(rstlink!desarroll)
              Else: If atrim(rstlink!color) <> "" Then rstc!dessarroll = 0
          End If
           Else: posartinterdecomandaazero rstc, i
        End If
proxima:
      Next i
      copiarobservacionstreballacomanda treball, modificacio, numc
   End If
   arxiumontadora = buscararxiumontadora(cadbl(rstmodificacio!id_treball), cadbl(rstmodificacio!ordre), cadbl(rstc!client), cadbl(rstc!direnvio))
   rstc!arxiumontadora = arxiumontadora
   If Not comandaimpresa(Data1.Recordset!comanda) Then If larutahiha(rstc!producte, "R") Then rstc!amplereb = rstmodificacio!amplelamina
   If vtinters <> cadbl(rstc!numerotintes) Then rstc!numerotintes = vtinters
   rstc!arxiu = atrim(rstclixe!arxiu)
   rstc!formaimp = atrim(rstmodificacio!formaimpresio)
   rstc!gruixpol = cadbl(rstmodificacio!gruixpolimer)
   If Len(atrim(rstclixe!codidebarres)) > 15 Then
           rstc!codibarras = Mid(atrim(rstclixe!codidebarres), Len(atrim(rstclixe!codidebarres)) - 14)
            Else: rstc!codibarras = atrim(rstclixe!codidebarres)
   End If
   rstc!cmaquina = IIf(atrim(rstclixe!redcilindrefw) <> "", atrim(rstclixe!reduccioxmetre), atrim(rstc!cmaquina))
   rstc!marcailinia = atrim(rstclixe!marca) + " - " + atrim(rstclixe!linia)
   rstc.Update
   Exit Sub
fi:
   Set rstc = Nothing
   Set rsttintes = Nothing
   Set rstlink = Nothing
   Set rstclixe = Nothing
   Set rsttintesllaunes = Nothing
noinfo:
   MsgBox "Falta informació de comparació amb el CLIXE per aquesta comanda.", vbCritical, "Atenció"
End Sub
Sub posartinterdecomandaazero(rstc As Recordset, i As Byte)
   rstc.Fields("tinta" + atrim(i) + "a") = ""
   rstc.Fields("lin" + atrim(i)) = 0
   rstc.Fields("detalltinter" + atrim(i)) = ""
End Sub
Function larutahiha(producte As String, seccio As String) As Boolean
   Dim rstp As Recordset
   Set rstp = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(producte) + "'")
   If rstp.EOF Then Exit Function
   If InStr(1, rstp!ruta, seccio) > 0 Then
       larutahiha = True
      Else: larutahiha = False
   End If
   Set rstp = Nothing
End Function
Sub mirardiferenciescomandaitreball(numc As Double)
   Dim i As Byte
   Dim cilindre As Double
   Dim desarroll As Double
   Dim continu As String
   Dim rstc As Recordset
   Dim rstclixe As Recordset
   Dim rstlink As Recordset
   Dim rstmodificacio As Recordset
   Dim rsttintes As Recordset
   Dim rsttintesllaunes As Recordset
   Dim treball As Integer
   Dim modificacio As Integer
   Dim arxiumontadora As String
   Dim vtinters As Byte
   Dim vnomtinta As String
   
   dbclixesnous.Execute "delete * from diferenciescomandaitreball where comanda=" + atrim(numc)
   Set rstc = dbtmp.OpenRecordset("SELECT comandes.*, productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi where comanda = " + atrim(numc))
   If rstc.EOF Then GoTo noinfo
   If InStr(1, rstc!ruta, "I") = 0 Then Exit Sub
   treball = cadbl(rstc!numtreball): modificacio = cadbl(rstc!numordremodificacio)
   If modificacio = 0 Then modificacio = 1
   Set rstclixe = dbclixesnous.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(rstc!numtreball)))
   If rstclixe.EOF Then GoTo noinfo
   Set rstmodificacio = dbclixesnous.OpenRecordset("select * from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(cadbl(modificacio)))
   If rstmodificacio.EOF Then GoTo noinfo
   
   Set rsttintes = dbclixesnous.OpenRecordset("select * from tintes where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordremodificacio=" + atrim(cadbl(modificacio)) + " order by ordretinter", , dbonly)
   Set rsttintesllaunes = dbclixesnous.OpenRecordset("select codi,descripcio from tintes_llaunes", , dbonly)
   'If rsttintes.EOF Then GoTo noinfo
      'comprovo tintes
      If Not rsttintes.EOF Then
        
        For i = 1 To 8
          rsttintes.FindFirst "ordretinter = " + atrim(i)
          If rsttintes.NoMatch Then
              If atrim(rstc.Fields("tinta" + atrim(i) + "a")) <> "" Then
                posardiferencia "Tinter Nº " + atrim(i), atrim(rstc.Fields("tinta" + atrim(i) + "a")), "<Sense Tinta>", treball, modificacio, numc
              End If
              GoTo proxima
          End If
          If atrim(rsttintes!color) <> "" Or cadbl(rsttintes!tinterlinkambid_treball) > 0 Then vtinters = vtinters + 1
          Set rstlink = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(IIf(rsttintes!tinterlinkambid_treball > 0, rsttintes!tinterlinkambid_treball, rsttintes!id_tinter)))
          If rstlink.EOF Then GoTo proxima
          rsttintesllaunes.FindFirst "codi='" + atrim(rstlink!coditinta) + "'"
          If rsttintesllaunes.NoMatch And atrim(rsttintes!color) = "" Then
            If atrim(rstc.Fields("tinta" + atrim(i) + "a")) <> "" Then
              posardiferencia "Tinter Nº " + atrim(i), atrim(rstc.Fields("tinta" + atrim(i) + "a")), "<Sense Tinta>", treball, modificacio, numc
            End If
            GoTo proxima
          End If
          vnomtinta = IIf(Not rsttintesllaunes.NoMatch, atrim(rsttintesllaunes!descripcio), atrim(rsttintes!color))
          If atrim(rstc.Fields("tinta" + atrim(i) + "a")) <> vnomtinta Then posardiferencia "Tinter Nº " + atrim(i), atrim(rstc.Fields("tinta" + atrim(i) + "a")), vnomtinta, treball, modificacio, numc
          If atrim(rstc.Fields("detalltinter" + atrim(i))) <> atrim(rstlink!detalltinter) Then posardiferencia "Detalltinter Nº " + atrim(i), atrim(rstc.Fields("detalltinter" + atrim(i))), atrim(rstlink!detalltinter), treball, modificacio, numc
          If cadbl(atrim(rstc.Fields("lin" + atrim(i)))) <> atrim(rstlink!anilox) Then posardiferencia "Anilox Nº " + atrim(i), atrim(rstc.Fields("lin" + atrim(i))), atrim(rstlink!anilox), treball, modificacio, numc
            'If atrim(rstc.Fields("tinta" + atrim(i) + "b")) <> atrim(rstlink!observacions) Then posardiferencia "Observacions Nº " + atrim(i), atrim(rstc.Fields("tinta" + atrim(i) + "b")), atrim(rstlink!observacions), treball, modificacio, numc
          If continu = "" Then
            If cadbl(rstlink!cilindre) <> 0 Then cilindre = cadbl(rstlink!cilindre)
            If cadbl(rstlink!desarroll) <> 0 Then desarroll = cadbl(rstlink!desarroll)
          End If
          If continu <> "S" Then continu = IIf(Not rstlink!continuu, "N", "S")
proxima:
        Next i
      End If
      If diferenciesobservacionstreballicomanda(treball, modificacio, numc) Then posardiferencia "Observacions diferents", "Diferencies", "Diferencies", treball, modificacio, numc
      If vtinters <> cadbl(rstc!numerotintes) Then posardiferencia "NºTinters", atrim(rstc!numerotintes), atrim(vtinters), treball, modificacio, numc
      If IIf(atrim(rstc!continu) = "", "N", atrim(rstc!continu)) <> continu Then posardiferencia "Continuu", atrim(rstc!continu), continu, treball, modificacio, numc
      If cadbl(rstc!cilindres) <> cilindre Then posardiferencia "Cilindre", cadbl(rstc!cilindres), atrim(cilindre), treball, modificacio, numc
      If cadbl(rstc!dessarroll) <> desarroll Then posardiferencia "Desarroll", cadbl(rstc!dessarroll), atrim(desarroll), treball, modificacio, numc
   
   If atrim(rstc!formaimp) <> atrim(rstmodificacio!formaimpresio) Then posardiferencia "Forma Impresio", atrim(rstc!formaimp), atrim(rstmodificacio!formaimpresio), treball, modificacio, numc
   If cadbl(rstc!gruixpol) <> cadbl(rstmodificacio!gruixpolimer) Then posardiferencia "Gruix Pol.", atrim(rstc!gruixpol), atrim(rstmodificacio!gruixpolimer), treball, modificacio, numc
   If atrim(rstc!codibarras) <> atrim(rstclixe!codidebarres) Then
      If Len(rstclixe!codidebarres) > 15 Then
       If Mid(rstclixe!codidebarres, Len(rstclixe!codidebarres) - 14) <> atrim(rstc!codibarras) Then
           posardiferencia "Codi de Barres", atrim(rstc!codibarras), atrim(rstclixe!codidebarres), treball, modificacio, numc
       End If
         Else: posardiferencia "Codi de Barres", atrim(rstc!codibarras), atrim(rstclixe!codidebarres), treball, modificacio, numc
      End If
   End If
   If atrim(rstc!cmaquina) <> atrim(rstclixe!reduccioxmetre) And atrim(rstclixe!redcilindrefw) <> "" Then posardiferencia "Reduccio per metre", atrim(rstc!cmaquina), atrim(rstclixe!reduccioxmetre), treball, modificacio, numc
   If larutahiha(rstc!producte, "R") And atrim(rstc!amplereb) <> atrim(rstmodificacio!amplelamina) Then
     posardiferencia "Ample lamina", atrim(rstc!amplereb), atrim(rstmodificacio!amplelamina), treball, modificacio, numc
   End If
   If alltrim(rstc!marcailinia) <> (alltrim(atrim(rstclixe!marca)) + "-" + alltrim(atrim(rstclixe!linia))) Then posardiferencia "Marca i Linia", atrim(rstc!marcailinia), (atrim(rstclixe!marca) + " - " + atrim(rstclixe!linia)), treball, modificacio, numc
   If atrim(rstc!arxiu) <> atrim(rstclixe!arxiu) Then posardiferencia "Arxiu Clixe", atrim(rstc!arxiu), atrim(rstclixe!arxiu), treball, modificacio, numc
   arxiumontadora = buscararxiumontadora(cadbl(rstmodificacio!id_treball), cadbl(rstmodificacio!ordre), cadbl(rstc!client), cadbl(rstc!direnvio))
   If atrim(rstc!arxiumontadora) <> arxiumontadora Then posardiferencia "Arxiu Muntadora", atrim(rstc!arxiumontadora), atrim(arxiumontadora), treball, modificacio, numc

   Set rsttintes = Nothing
   Set rsttintesllaunes = Nothing
   
noinfo:
  ' MsgBox "No hi ha informació de comparació amb el CLIXE per aquesta comanda.", vbCritical, "Atenció"
End Sub
Function copiarobservacionstreballacomanda(idtreball As Integer, ordremodificacio As Integer, numc As Double) As Boolean
   Dim vc1 As String
   Dim vc2 As String
   Dim vt1 As String
   Dim vt2 As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(idtreball) + " and ordre=" + atrim(ordremodificacio) + " order by id")
   If Not rst.EOF Then vt1 = treure_apostruf(atrim(rst!observacio)): rst.MoveNext
   If Not rst.EOF Then vt2 = treure_apostruf(atrim(rst!observacio)): rst.MoveNext
   dbtmp.Execute "delete * from comandes_observacionstintes where comanda=" + atrim(numc)
   If vt1 <> "" Then
    dbtmp.Execute "insert into comandes_observacionstintes (comanda,observacio) values (" + atrim(numc) + ",'" + treure_apostruf(vt1) + "')"
    If vt2 <> "" Then dbtmp.Execute "insert into comandes_observacionstintes (comanda,observacio) values (" + atrim(numc) + ",'" + treure_apostruf(vt2) + "')"
   End If
   Set rst = Nothing
End Function


Function diferenciesobservacionstreballicomanda(idtreball As Integer, ordremodificacio As Integer, numc As Double) As Boolean
   Dim vc1 As String
   Dim vc2 As String
   Dim vt1 As String
   Dim vt2 As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(idtreball) + " and ordre=" + atrim(ordremodificacio) + " order by id")
   If Not rst.EOF Then vt1 = atrim(rst!observacio): rst.MoveNext
   If Not rst.EOF Then vt2 = atrim(rst!observacio): rst.MoveNext
   
   Set rst = dbtmp.OpenRecordset("select * from comandes_observacionstintes where comanda=" + atrim(numc) + " order by id")
   If Not rst.EOF Then vc1 = atrim(rst!observacio): rst.MoveNext
   If Not rst.EOF Then vc2 = atrim(rst!observacio): rst.MoveNext
   
   If treure_apostruf(vt1) <> vc1 Then diferenciesobservacionstreballicomanda = True
   If treure_apostruf(vt2) <> vc2 Then diferenciesobservacionstreballicomanda = True
   
   Set rst = Nothing
End Function

Function alltrim(v As String) As String
  
  While InStr(1, v, " ") > 0
     v = Mid(v, 1, InStr(1, v, " ") - 1) + Mid(v, InStr(1, v, " ") + 1)
  Wend
  alltrim = v
End Function
Function buscararxiumontadora(treball As Double, modificacio As Double, client As Double, direnvio As Double) As String
  Dim rst As Recordset
  Set rst = dbclixes.OpenRecordset("select * from clientsvinculats where id_treball=" + atrim(treball) + " and ordremodificacio=" + atrim(modificacio) + " and codiclient=" + atrim(client) + " and direnvio=" + atrim(direnvio))
  If Not rst.EOF Then buscararxiumontadora = rst!codimuntadora
  Set rst = Nothing
End Function


Sub copiarmontadoraireferenciesaltreball(treball As Double, modificacio As Double, client As Double, direnvio As Double)
  Dim rst As Recordset
  Set rst = dbclixes.OpenRecordset("select * from clientsvinculats where id_treball=" + atrim(treball) + " and ordremodificacio=" + atrim(modificacio) + " and codiclient=" + atrim(client) + " and direnvio=" + atrim(direnvio))
  If Not rst.EOF And Not Data1.Recordset.EOF Then
       rst.Edit
       If atrim(rst!codimuntadora) = "" Then rst!codimuntadora = Data1.Recordset!arxiumontadora
       rst!refclient = atrim(Data1.Recordset!refclient)
       rst!refclientalternatives = atrim(Data1.Recordset!refclialt)
       rst.Update
  End If
  Set rst = Nothing
End Sub


Sub posardiferencia(Camp As String, valorcomanda As String, valortreball As String, treball As Integer, modificacio As Integer, numc As Double)
   Dim valors As String
   valors = atrim(treball) + "," + atrim(modificacio) + "," + atrim(numc) + ",'" + treure_apostruf(Camp) + "','" + treure_apostruf(valorcomanda) + "','" + treure_apostruf(valortreball) + "'"
   dbclixesnous.Execute "insert into diferenciescomandaitreball (id_treball,ordremodificacio,comanda,camp,valorcomanda,valortreball) values (" + valors + ")"
End Sub
Sub gravantcomandaordinador(rst As Recordset, gravantcomanda As Variant, gravantnomordinador As Variant)
   On Error GoTo error
   rst.Edit
   rst!gravantcomanda = IIf(IsNull(gravantcomanda), Null, CVDate(Format(Now, "dd/mm/yy hh:nn:ss")))
   rst!gravantnomordinador = gravantnomordinador
   rst.Update
   'dbtmp.Execute "update valorsgenerals set gravantcomanda=#" + Format(Now, "mm/dd/yy hh:nn:ss") + "#"
      'dbtmp.Execute "update valorsgenerals set gravantnomordinador='" + atrim(nomordinador) + "'"
   Exit Sub
error:
    MsgBox err.Description
    If rst.EditMode > 0 Then rst.CancelUpdate
End Sub
Function hihaalgugravant(Optional gravant As Byte) As Boolean
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from valorsgenerals")
   hihaalgugravant = False
   Command11.Tag = ""
   If rst.EOF Then Exit Function
   If atrim(rst!gravantnomordinador) <> "" Then
       If DateDiff("s", rst!gravantcomanda, Now) > 15 Or IsNull(rst!gravantcomanda) Then
           hihaalgugravant = False
           gravantcomandaordinador rst, Null, Null
          Else:
            If atrim(rst!gravantnomordinador) <> atrim(nomordinador) Then
                hihaalgugravant = True
                Command11.Tag = atrim(rst!gravantnomordinador)
            End If
       End If
   End If
   If gravant = 1 And Not hihaalgugravant Then
       gravantcomandaordinador rst, Now, atrim(nomordinador)
      'dbtmp.Execute "update valorsgenerals set gravantcomanda=#" + Format(Now, "mm/dd/yy hh:nn:ss") + "#"
      'dbtmp.Execute "update valorsgenerals set gravantnomordinador='" + atrim(nomordinador) + "'"
   End If
   If gravant = 2 And Not hihaalgugravant Then
      gravantcomandaordinador rst, Null, Null
      'dbtmp.Execute "update valorsgenerals set gravantcomanda=null"
      'dbtmp.Execute "update valorsgenerals set gravantnomordinador=null"
   End If
   Set rst = Nothing
End Function
Function comprovarrelaciomesuraPVPidemanada() As Boolean
   'relaciomesureslineals
   Dim rst As Recordset
   If Mid(Text3, 1, 2) = "PC" Then comprovarrelaciomesuraPVPidemanada = True: Exit Function
   comprovarrelaciomesuraPVPidemanada = True
   Set rst = dbtmp.OpenRecordset("select * from mesures where codi=" + atrim(cadbl(Text7)))
   If Not rst.EOF Then
       If cadbl(Text32(8)) <> cadbl(rst!relaciomesureslineals) Then
          comprovarrelaciomesuraPVPidemanada = False
           Else: If cadbl(rst!relaciomesureslineals) = 0 Then comprovarrelaciomesuraPVPidemanada = True
       End If
   End If
   If cadbl(Text7) = 0 Then comprovarrelaciomesuraPVPidemanada = True
End Function
Function soniguals_simulteneitatTreballiRebobinadora() As Boolean
   Dim rst As Recordset
   Dim vt As String
   Dim vo As String
   soniguals_simulteneitatTreballiRebobinadora = True
   If InStr(1, ruta, "I") = 0 Then Exit Function
   vt = atrim(cadbl(Data1.Recordset!numtreball)): vo = atrim(cadbl(Data1.Recordset!numordremodificacio))
   Set rst = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + vt + " and ordre=" + vo)
   If Not rst.EOF Then
       If cadbl(Combo3) <> cadbl(rst!bandes) Then soniguals_simulteneitatTreballiRebobinadora = False
   End If
   Set rst = Nothing
End Function
Sub usuari_guarda_registre()
  'If buscant Then
  If Text13.Enabled Then Text13.SetFocus    ' canvio de camp per fer saltar el lostfocus del registre actual abans de guardar els canvis
  'If atrim(Text16) <> atrim(Text32(7)) And atrim(Text32(7)) <> "" Then MsgBox "Les unitats de PVP i quantitat demanada han d'esser la mateixa", vbCritical, "Error": Exit Sub
  'relaciomesureslineals
  'If Not comprovarrelaciomesuraPVPidemanada And cadbl(Text6) > 0 Then MsgBox "Les unitats de PVP i quantitat demanada han d'esser la mateixa", vbCritical, "Error": Exit Sub
  If Not soniguals_simulteneitatTreballiRebobinadora Then MsgBox "Atenció la simulteneitat del treball i la rebobinadora son diferents.", vbCritical, "Atenció"
  If Not comprovarrelaciomesuraPVPidemanada Then MsgBox "Les unitats de PVP i quantitat demanada han d'esser la mateixa", vbCritical, "Error": Exit Sub
  If cadbl(Text6) > 0 And cadbl(MaskEdBox6) = 0 Then
     MsgBox "Si poses PVP a la comanda també has de possar la quantitat demanada.", vbCritical, "Atenció"
     MaskEdBox6.SetFocus
     Exit Sub
  End If
  If MaskEdBox6.Tag = "obligatquantitatdemanada" And cadbl(MaskEdBox6) <= 0 Then MsgBox "No has entrat la quanitat demanda pel client, per aquest client es necessari possar-ho.", vbCritical, "Quantitat demanada": Exit Sub
  If Not hihaalgugravant Then
     gravar_registre
       Else: MsgBox "Hi ha un altra usuari gravant un registre, espera uns segons i torna-ho a provar", vbExclamation, "Atenció"
  End If
   '  Else
 '      gravar_registre
 ' End If
  hihaalgugravant 2
  activaronocampsimpresio False
  enabled_campscontrolcodiinplacsa True
End Sub

Private Sub Form_Activate()
   
'  DemanarPrefixRefInplacsa
  carregar_firmes
'  passar_accessoris_soldadores 300000, 226912
 'importar_tarifes_referenciesinplacsa
 '  comprovar_refinplacsaigualalsPC

  'comprovar_refinplacsa


'  comprovar_firmesrepetides
  'comprovar_referencies
  If Not imprimircomandes Then
     assignardecimalipunt
     If vprimeraentradacomandes = False Then vprimeraentradacomandes = True: consultar_Click
  End If
  colorrisc = cap.BackColor
  Label1(146).ForeColor = colorrisc
  
   
End Sub
Sub importar_tarifes_referenciesinplacsa()
  Dim rst As Recordset
  Dim rst2 As Recordset
  If MsgBox("importar tarifes referencies", vbCritical + vbDefaultButton2 + vbYesNo, "atencio") = vbNo Then Exit Sub
  Set rst = dbtmp.OpenRecordset("select * from tarifes_referencies")
  Set rst2 = dbtmp.OpenRecordset("select * from tarifes_referencies_prova_BARREJATS")
  While Not rst2.EOF
    rst.FindFirst "refinplacsa='" + atrim(rst2!refinplacsa) + "'"
    If Not rst.NoMatch Then
        rst.Edit
        'rst!coditarifa = rst2!coditarifa
        rst!inactiva = rst2!inactiva
        rst.Update
          Else
            rst.AddNew
            rst!refinplacsa = rst2!refinplacsa
         '   rst!coditarifa = rst2!coditarifa
            rst!inactiva = rst2!inactiva
            rst!codiclient = rst2!codiclient
            rst.Update
    End If
    rst2.MoveNext
  Wend
  MsgBox "acabat"
  Set rst = Nothing
  Set rst2 = Nothing
End Sub
Sub comprovar_refinplacsaigualalsPC()
  Dim rst As Recordset
  Dim rst2 As Recordset
  Set rst = dbtmp.OpenRecordset("select * from comandesmesextres where producte<>'PC' and producte<>'PC2' and producte <>'PCP' and producte<>'PCI3'")
  Set rst2 = dbtmp.OpenRecordset("select * from comandes_extres")
  While Not rst.EOF
   If cadbl(rst!linkcomanda1) <> 0 Then
    rst2.FindFirst "comanda=" + atrim(rst!linkcomanda1)
    If Not rst2.NoMatch Then
       If rst!refinplacsa <> rst2!refinplacsa Then
         rst2.Edit: rst2!refinplacsa = rst!refinplacsa: rst2.Update
       End If
    End If
   End If
   If cadbl(rst!linkcomanda2) <> 0 Then
    rst2.FindFirst "comanda=" + atrim(rst!linkcomanda2)
    If Not rst2.NoMatch Then
       If rst!refinplacsa <> rst2!refinplacsa Then
         rst2.Edit: rst2!refinplacsa = rst!refinplacsa: rst2.Update
       End If
    End If
   End If
    
   rst.MoveNext
   Me.Caption = rst!comanda
   DoEvents
  Wend
  
  
  Set rst = Nothing
  Set rst2 = Nothing
  
End Sub
Sub comprovar_firmesrepetides()
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim v As String
  
  Set rst = dbtmp.OpenRecordset("select distinct comanda from comandes_firmes where dataanulacio=null")
  While Not rst.EOF
    Set rst2 = dbtmp.OpenRecordset("select * from comandes_firmes where comanda=" + atrim(rst!comanda) + " order by tipus,usuari,data desc")
    v = ""
    While Not rst2.EOF
      If v = rst2!tipus + rst2!usuari Then
         If rst2!tipus <> "PVP" Then rst2.Delete
        Else: v = rst2!tipus + rst2!usuari
      End If
      rst2.MoveNext
    Wend
    rst.MoveNext
  Wend
  Set rst = Nothing
  Set rst2 = Nothing
  
End Sub
Sub comprovar_refinplacsa()
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim vref As String
  Dim vsql As String
  Dim vcont As Double
  Dim vcodiclient As Double
  Dim vrefi As String
  Dim vdiferencia As String
  Static esticdins As Boolean
  If esticdins Then Exit Sub
  esticdins = True
  vsql = InputBox("Escriu el codi de client que vols revisar o res per tots." + vbNewLine + "POTS POSSAR UNA REFERENCIA DIRECTAMENT TAMBÉ.", "Escull Client")
  If StrPtr(vsql) = 0 Then GoTo fi
  If Len(vsql) > 5 Then
     vrefi = vsql
       Else: vcodiclient = cadbl(vsql)
  End If
  vsql = "SELECT  min(comandesmesextres.comanda) as numcomanda From comandesmesextres where refinplacsa<>'' and (((comandesmesextres.producte)<> 'PC' and (comandesmesextres.producte) <> 'PCP' and (comandesmesextres.producte)<> 'PC2') and (comandesmesextres.producte)<> 'PCI3')"
  vsql = vsql + " GROUP BY comandesmesextres.refinplacsa HAVING (((Count(comandesmesextres.client))>1));"
  If vcodiclient > 0 Then
     vsql = "SELECT  min(comandesmesextres.comanda) as numcomanda From comandesmesextres where refinplacsa<>'' and (((comandesmesextres.producte)<> 'PC' and (comandesmesextres.producte) <> 'PCP' and (comandesmesextres.producte)<> 'PC2') and (comandesmesextres.producte)<> 'PCI3')"
     vsql = vsql + " and client=" + atrim(vcodiclient) + " GROUP BY comandesmesextres.refinplacsa HAVING (((Count(comandesmesextres.client))>1));"
  End If
  If vrefi <> "" Then
      vsql = "SELECT  min(comandesmesextres.comanda) as numcomanda From comandesmesextres where refinplacsa<>'' and (((comandesmesextres.producte)<> 'PC' and (comandesmesextres.producte) <> 'PCP' and (comandesmesextres.producte)<> 'PC2') and (comandesmesextres.producte)<> 'PCI3')"
      vsql = vsql + " and refinplacsa='" + atrim(vrefi) + "' GROUP BY comandesmesextres.refinplacsa HAVING (((Count(comandesmesextres.client))>1));"
  End If
'  Clipboard.Clear
'  Clipboard.SetText vsql
  MsgBox "COMENÇA EL LLISTAT"
  Open "c:\temp\llistat.csv" For Output As #1
 ' vsql = "select comanda as numcomanda from comandes_extres where refinplacsa='02C7035I5394'"
  Set rst = dbtmp.OpenRecordset(vsql)
  If rst.EOF Then MsgBox "No hi ha resultats de la busqueda.": GoTo fi
  rst.MoveLast
  rst.MoveFirst
  While Not rst.EOF
   vdiferencia = ""
   Set rst2 = dbtmp.OpenRecordset("select * from comandesmesextres where comanda=" + atrim(rst!numcomanda) + " order by comanda asc")
   If Not rst2.EOF Then
    If Mid(rst2!refinplacsa, 1, 2) <> "PR" And Mid(rst2!refinplacsa, 1, 2) <> "RP" And Mid(rst2!refinplacsa, 1, 2) <> "FP" Then
         vdiferencia = hihadiferenciesenaquestaRefInplacsa(rst2)
         If vdiferencia <> "" Then
              If noesdelesmarcadescomacorrectes(rst2!refinplacsa) Then
                Print #1, atrim(rst2!comanda) + ";" + rst2!refinplacsa + ";" + atrim(rst2!client) + ";" + IIf(InStr(1, rst2!ruta, "I"), atrim(rst2!numtreball), "") + ";" + vdiferencia
                Me.Caption = atrim(rst2!comanda) + ";" + rst2!refinplacsa + ";" + vdiferencia
                vcont = vcont + 1
              End If
              DoEvents
         End If
    End If
    
   End If
   DoEvents
   Me.Caption = atrim(rst!numcomanda) + " -> " + atrim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount) + "  Cont: " + atrim(vcont)
   rst.MoveNext
  Wend
  Close #1
  Set rst = Nothing
  If existeix("c:\temp\llistat.csv") Then
    obrir_document "c:\temp\llistat.csv"
     Else:: MsgBox "Procès acabat."
  End If
fi:
  esticdins = False
End Sub
Function noesdelesmarcadescomacorrectes(vref As String) As Boolean
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select refinplacsa from comandes_extres where refinplacsa='" + vref + "' and refinplacsa_valida")
   
   If rst.EOF Then
      noesdelesmarcadescomacorrectes = True
       Else: noesdelesmarcadescomacorrectes = False
   End If
   Set rst = Nothing
End Function
Function hihadiferenciesenaquestaRefInplacsa(rst As Recordset) As String

   Dim rstc As Recordset
   Dim rstp As Recordset
   Dim vruta As String
   Dim camps As String
   Dim nomcamp As String
   Dim valorcamp As String
   Dim totsiguals As Boolean
   Dim comandesiguals As String
   Dim vreferencianova As String
   Dim rstmat1 As Recordset
   Dim rstmat2 As Recordset
   Dim vdiferencia As String
   
   If atrim(rst!producte) = "PC" Or atrim(rst!producte) = "PC2" Or atrim(rst!producte) = "PCP" Then Exit Function
   Set rstp = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
   If rstp.EOF Then Exit Function
   vruta = atrim(rstp!ruta)
 
   If InStr(1, vruta, "I") = 0 Then
      Set rstc = dbtmp.OpenRecordset("SELECT comandesmesextres.*, InStr(1,[ruta],'I') AS Expr1 FROM comandesmesextres WHERE producte<>'PC' and producte<>'PCP' and producte<>'PC2' and producte<>'PCI3' and refinplacsa='" + rst!refinplacsa + "' and comanda<>" + atrim(rst!comanda) + " order by comanda asc")
          Else:
            Set rstc = dbtmp.OpenRecordset("SELECT comandesmesextres.*, InStr(1,[ruta],'I') AS Expr1 FROM comandesmesextres WHERE producte<>'PC' and producte<>'PCP' and producte<>'PC2' and producte<>'PCI3' and refinplacsa='" + rst!refinplacsa + "' and comanda<>" + atrim(rst!comanda) + " order by comanda asc")
            'MsgBox "SELECT comandes.*, InStr(1,[ruta],'I') AS Expr1 FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi WHERE (((InStr(1,[ruta],'I'))=0) and client = " + atrim(cadbl(rst!client)) + " and tubolam='" + atrim(rst!tubolam) + "' and ampleesq=" + passaradecimalpunt(cadbl(rst!ampleesq)) + ")"
            'Set rstc = dbtmp.OpenRecordset("select * from comandes where tubolam='" + atrim(rst!tubolam) + "' and ampleesq=" + passaradecimalpunt(cadbl(rst!ampleesq)) + " and client=" + atrim(cadbl(rst!client)))
   End If
   If rstc.EOF Then Exit Function
   Set rstmat1 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rst!materialex)))
   If rstmat1.EOF Then Exit Function
   While Not rstc.EOF
    If rstc!refinplacsa_valida Then totsiguals = False: GoTo proxim
    ''camps = "#Etubolam#Eampleesq#Eplegatesq#Esolapa#Eespessor#Emicropex#Eoberturaex#Ematerialex#Inumtreball#Lampleutil#Lsimulteneitatlam#Ltipusadhesiu#Rmigelaborat#Ramplereb#Rsimulteneitatreb"
    camps = "#Etubolam#Eampleesq#Eplegatesq#Esolapa#Eespessor#Emicropex#Eoberturaex#Ematerialex#Inumtreball#Lampleutil#Lsimulteneitatlam#Rmigelaborat#Ramplereb#Rsimulteneitatreb"
    camps = camps + "#Smigelaboratsol#Samplesol#Sampleplegsol#Slongitudsol#Ssolapasol#Sfuellebasesol#Sfuellebocasol#Stroquel#Sansa#Scinta#"
    nomcamp = proximcamp(camps)
    totsiguals = True
'    MsgBox rstc!comanda
    While nomcamp <> ""
      If InStr(1, vruta, Mid(nomcamp, 1, 1)) > 0 Then
       nomcamp = Mid(nomcamp, 2)
       valordelcamp = cvavalorcamp(rstc, nomcamp)
       If nomcamp = "numtreball" Then GoTo cont
       If nomcamp = "materialex" Then
           Set rstmat2 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstc!materialex)))
           If rstmat2.EOF Then GoTo proxim
           If cadbl(rstmat1!familia) <> cadbl(rstmat2!familia) Or cadbl(rstmat1!subfamilia) <> cadbl(rstmat2!subfamilia) Or cadbl(rstmat1!familiacol) <> cadbl(rstmat2!familiacol) Then
             vdiferencia = "families diferents"
             totsiguals = False: GoTo proxim
             
           End If
           GoTo cont
       End If
'       MsgBox nomcamp + "   =    " + atrim(valordelcamp) + "  ------  " + cvavalorcamp(rst, nomcamp)
       If atrim(valordelcamp) <> cvavalorcamp(rst, nomcamp) Then
         vdiferencia = nomcamp + " = " + atrim(valordelcamp) + "<>" + cvavalorcamp(rst, nomcamp)
         totsiguals = False: GoTo proxim
       End If
      End If
cont:
      nomcamp = proximcamp(camps)
    Wend
proxim:
   If totsiguals Then
      If Not complexesiguals(rst, rstc) Then hihadiferenciesenaquestaRefInplacsa = "Diferencia en els Complexes"
       Else: hihadiferenciesenaquestaRefInplacsa = vdiferencia
   End If
   'Me.Caption = atrim(rst!comanda)
   If hihadiferenciesenaquestaRefInplacsa <> "" Then GoTo fi
   rstc.MoveNext
   Wend
fi:
   Set rstc = Nothing
   Set rstmat1 = Nothing
   Set rstmat2 = Nothing
End Function
Sub comprovar_referencies()
  Dim rst As Recordset
  Dim rstc As Recordset
  Dim vref As String
  Dim vref2 As String
  
  Set rst = dbtmp.OpenRecordset("select * from tarifes_referencies where  refclient not like '*TEST*' ")
  While Not rst.EOF
     vref = ""
     vref2 = ""
     Set rstc = dbtmp.OpenRecordset("select distinct refinplacsa from comandesmesextres where refinplacsa<>'' and refclient='" + atrim(rst!refclient) + "'")
     If Not rstc.EOF Then vref = rstc!refinplacsa
     While Not rstc.EOF
       vref2 = vref2 + " " + rstc!refinplacsa
       rstc.MoveNext
     Wend
     rst.Edit
     rst!refinplacsa = vref
     If atrim(vref) <> atrim(vref2) Then rst!altresrefinplacsa = vref2
     rst.Update
     rst.MoveNext
  Wend
  Set rst = Nothing
  Set rst2 = Nothing
End Sub
Function tepackinglist(numc As Double) As Boolean
   Dim rstt As Recordset
   tepackinglist = False
   Set rstt = dbstocks.OpenRecordset("select * from parcials where  comanda='" + atrim(numc) + "'")
   If Not rstt.EOF Then tepackinglist = True
   Set rstt = dbtmp.OpenRecordset("select assignarstock from comandes_extres where comanda=" + atrim(numc))
   If Not rstt.EOF Then
      If rstt!assignarstock Then tepackinglist = True
   End If
   Set rstt = Nothing
   'Set dbstocks = Nothing
End Function

Sub comprovarcomandesambEsenseestarimpreses()
  Dim rst As Recordset
  Dim vmsg As String
  Dim vt As Date
  'a dia 8/5/23 l'alicia diu que tregui aquest missatge... no s'utilitza i a vegades molesta.
    Exit Sub
  vsql = "SELECT comandes.obspedido1, COMANDES.cantitatex,comandes.refclient,comandes_extres.assignarstock,comandes.comanda, comandes.proximaseccio, comandes_extres.comandaimpresa, comandes.dataactivacio, comandes.client "
  vsql = vsql + " FROM comandes_extres INNER JOIN comandes ON comandes_extres.comanda = comandes.comanda "
  vsql = vsql + " WHERE  (comandes.proximaseccio='I' or comandes.proximaseccio='R' or comandes.proximaseccio='L' or comandes.proximaseccio='R' or comandes.proximaseccio='S');"
  Set rst = dbtmp.OpenRecordset(vsql)
  While Not rst.EOF
   If cadbl(rst!cantitatex) > 0 And InStr(1, UCase(rst!refclient), "TEST") = 0 And InStr(1, UCase(rst!refclient), "PROVA") = 0 Then
    If Not rst!assignarstock Then
     If Not tepackinglist(rst!comanda) Then
          '8/9/22
          'UNA EXCEPCIÓ QUE M'HA FET POSAR L'ALICIA PER QUAN FABRIQUEN FORA D 'INPLACSA
       If InStr(1, UCase(atrim(rst!obspedido1)), " EXTERN ") = 0 Then
         vmsg = vmsg + atrim(rst!comanda) + " "
       End If
     End If
    End If
   End If
   rst.MoveNext
  Wend
  If vmsg <> "" Then MsgBox "He trobat les seguents comandes que no estan en estat 'E' i encara no tenen packinglist." + vbNewLine + "Reviseu-les sisplau." + vbNewLine + vmsg, vbCritical, "Atenció"
  
  Set rst = Nothing
End Sub
Sub possarordrelaminadora(numc As Double)
   Dim rst As Recordset
   
   Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
   If Not rst.EOF Then
      If cadbl(rst!linkcomanda2) < cadbl(rst!linkcomanda1) Or cadbl(rst!linkcomanda1) < cadbl(rst!comanda) Then Exit Sub
      If rst!refilatd = 1 Then
          'possar ordre normal
             rst.Edit
             rst!lotmatdesb1 = rst!comanda
             rst!lotmatdesb2 = rst!linkcomanda1
             rst.Update
             dbtmp.Execute "update comandes set lotmatdesb1=" + atrim(cadbl(rst!comanda)) + ", lotmatdesb2=" + atrim(cadbl(rst!linkcomanda2)) + " where comanda=" + atrim(cadbl(rst!linkcomanda2))
           Else
                'possar odre invertit
                rst.Edit
                rst!lotmatdesb1 = rst!comanda
                rst!lotmatdesb2 = rst!linkcomanda2
                rst.Update
                dbtmp.Execute "update comandes set lotmatdesb1=" + atrim(cadbl(rst!linkcomanda1)) + ", lotmatdesb2=" + atrim(cadbl(rst!linkcomanda2)) + " where comanda=" + atrim(cadbl(rst!linkcomanda2))
             
      End If
   End If
End Sub
Sub comprovarsihihacomandaxrobrirdeplanificacio()
  Dim comanda As Double
  comanda = cadbl(llegir_ini("Planificacio", "comandaxrobrir", "comandes.ini"))
  escriure_ini "Planificacio", "comandaxrobrir", "0", "comandes.ini"
  If comanda > 0 Then
     If Data1.Recordset.EditMode > 0 Then
       cancelar_registre
       r = "sortir"
       subbusqueda.Hide
     End If
     Data1.RecordSource = "select * from comandes where comanda=" + atrim(comanda)
     Data1.Refresh
     vprimeraentradacomandes = True
  End If
End Sub

Function imprimirdiferenciescomandaitreball(numc As Double, Optional PoI As String) As Boolean
 ' Dim rst As Recordset
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.report
 imprimirdiferenciescomandaitreball = True
  Set rst = dbclixesnous.OpenRecordset("select * from diferenciescomandaitreball where comanda=" + atrim(numc))
  If rst.EOF Then imprimirdiferenciescomandaitreball = False: Exit Function
  
 ' llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "incidenciescomandaitreball.rpt"
 ' If Not existeix("c:\ordprog.ini") Then llistat.Destination = crptToPrinter
 ' llistat.DataFiles(0) = rutadelfitxer(cami) + "clixesnous.mdb"
 ' llistat.SelectionFormula = "{diferenciescomandaitreball.comanda}=" + atrim(numc)
 ' llistat.Action = 1
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "incidenciescomandaitreball.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
  oreport.RecordSelectionFormula = "{diferenciescomandaitreball.comanda}=" + atrim(numc)
  
  oreport.DiscardSavedData
   
  'If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   
   If PoI = "I" Then
      dbtmp.Execute "update comandes_Extres set aviscanvisambeltreball='Imprès full diferències treball' where comanda=" + atrim(numc)
      If Not existeix("c:\ordprog.ini") Then
          veurereport.CRViewer.PrintReport
         Else: veurereport.CRViewer.ViewReport
      End If
        Else
           veurereport.CRViewer.ViewReport
  End If
  
   veurereport.WindowState = 2
   veurereport.Show 1
   ' Else
   '   oreport.PrintOut False, 1
 ' End If
  
End Function

Function no_imprimirdiferenciescomandaitreball(numc As Double, Optional PoI As String) As Boolean
  Dim rst As Recordset
  Dim i As Byte
  no_imprimirdiferenciescomandaitreball = True
  Set rst = dbclixesnous.OpenRecordset("select * from diferenciescomandaitreball where comanda=" + atrim(numc))
  If rst.EOF Then no_imprimirdiferenciescomandaitreball = False: Exit Function
  
  llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "incidenciescomandaitreball.rpt"
  
  If PoI = "I" Then
      llistat.Destination = crptToPrinter
      dbtmp.Execute "update comandes_Extres set aviscanvisambeltreball='Imprès full diferències treball' where comanda=" + atrim(numc)
     Else
       If PoI = "P" Then llistat.Destination = crptToWindow
  End If
  
  If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
  For i = 0 To 50
    llistat.Formulas(i) = ""
  Next i
  
  llistat.DataFiles(0) = dbclixesnous.Name
  llistat.SelectionFormula = "{diferenciescomandaitreball.comanda}=" + atrim(numc)
  llistat.WindowState = crptMaximized
  llistat.Action = 1
End Function

Sub possartoteslesmarcailinia()
   Dim rst As Recordset
   Dim rstcl As Recordset
   Set rst = dbtmp.OpenRecordset("select distinct numtreball from comandes")
   If Not rst.EOF Then rst.MoveLast: rst.MoveFirst
   While Not rst.EOF
      Set rstcl = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(cadbl(rst!numtreball)))
      If Not rstcl.EOF Then
         dbtmp.Execute "update comandes set marcailinia='" + treure_apostruf(atrim(rstcl!marca)) + "-" + treure_apostruf(atrim(rstcl!linia)) + "' where numtreball=" + atrim(rst!numtreball)
      End If
      rst.MoveNext
      Me.Caption = rst.AbsolutePosition
      DoEvents
   Wend
   Set rst = Nothing
   Set rstcl = Nothing
End Sub

Sub possarreferenciainplacsa()
  Dim rst As Recordset
  Dim rstextra As Recordset
  Dim rstp As Recordset
  dbtmp.Execute "update comandeS_extres set refinplacsa=null"
  Set rst = dbtmp.OpenRecordset("SELECT * FROM comandes where materialex>499 order by comanda")
  rst.MoveLast
  rst.MoveFirst
  While Not rst.EOF
     Set rstextra = dbtmp.OpenRecordset("SELECT * FROM comandes_extres where comanda=" + atrim(rst!comanda))
     Set rstp = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
     If rstp.EOF Then GoTo cont
     Me.Caption = atrim(rst![comanda]) + "   " + atrim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount)
     DoEvents
     If rstextra.EOF Then crearcampsextra rst, rstextra
     If Not rstextra!comandaimpresa Then GoTo cont
     If atrim(rstextra!refinplacsa) <> "" Then GoTo cont
     If InStr(1, rstp!ruta, "I") > 0 Then
         If cadbl(rst!numtreball) > 0 Then
             comandesamblamateixareferenciainplacsa rst, True
             'dbtmp.Execute "update comandes_extres set refinplacsa=" + "8" + atrim(cadbl(rst!client)) + atrim(rst!numtreball) + " where comanda in (" + atrim(cadbl(rst![comanda])) + ", " + atrim(cadbl(rst!linkcomanda1)) + ", " + atrim(cadbl(rst!linkcomanda2)) + ")"
         End If
           Else: comandesamblamateixareferenciainplacsa rst, False
     End If
cont:
     rst.MoveNext
  Wend
End Sub
Sub crearcampsextra(rst As Recordset, rstextra As Recordset)
   Dim vdata As Date
   
   If IsNull(rst!datacomanda) Then
          vdata = "01/01/2000"
            Else: vdata = rst!datacomanda
    End If
   dbtmp.Execute "insert into comandes_extres (comanda,data) values (" + atrim(rst!comanda) + ",#" + Format(vdata, "mm/dd/yy") + "#)"
   Set rstextra = dbtmp.OpenRecordset("select * from comandes_extres where comanda=" + atrim(rst!comanda))
End Sub
Sub passarobstintesaliniesobs()
  Dim rst As Recordset
  Dim vtext As String
  Dim v1 As String
  Dim v2 As String
  Static jahisoc As Boolean
  If jahisoc Then Exit Sub
  jahisoc = True
  Set rst = dbtmp.OpenRecordset("SELECT comandes.*, InStr(1,[ruta],'I') AS vruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi order by comanda;")
  
  While Not rst.EOF
    If rst!vruta = 0 Then GoTo proxim
    vtext = "": v1 = "": v2 = ""
    For i = 1 To 8
       vtext = atrim(vtext) + IIf(atrim(rst.Fields("tinta" + atrim(i) + "b")) <> "", " " + atrim(i) + "# " + atrim(rst.Fields("tinta" + atrim(i) + "b")), "")
    Next i
   ' While Len(vtext) > 160
   '    vtext = InputBox("Modifica per fer 160 caracters", "atenció", vtext)
   ' Wend
    v1 = Mid(vtext, 1, 80)
    v2 = Mid(vtext, 81, 160)
    If v1 <> "" Then dbtmp.Execute "insert into comandes_observacionstintes (comanda,observacio) values (" + atrim(rst!comanda) + ",'" + treure_apostruf(v1) + "')"
    If v2 <> "" Then
      dbtmp.Execute "insert into comandes_observacionstintes (comanda,observacio) values (" + atrim(rst!comanda) + ",'" + treure_apostruf(v2) + "')"
    End If
proxim:
    rst.MoveNext
    Me.Caption = atrim(rst!comanda)
    DoEvents
  Wend
  Set rst = Nothing
End Sub

Private Sub Form_Click()

'seleccionar_refinplacsa_activa "03C7035I5240"
 ' Dim rst As Recordset
  'Set rst = dbtmp.OpenRecordset("SELECT comandes_extres.comanda, comandes_extres.refinplacsa as crefinplacsa, tarifes_referencies.refinplacsa, comandes.client FROM (comandes_extres LEFT JOIN tarifes_referencies ON comandes_extres.refinplacsa = tarifes_referencies.refinplacsa) INNER JOIN comandes ON comandes_extres.comanda = comandes.comanda WHERE (((comandes_extres.refinplacsa)<>'') AND ((tarifes_referencies.refinplacsa) Is Null));")
  'While Not rst.EOF
  '  If rst!client > 0 Then
  '    dbtmp.Execute "insert into tarifes_referencies (codiclient,refinplacsa) values (" + atrim(rst!client) + ",'" + rst!crefinplacsa + "')"
  '  End If
  '  rst.MoveNext
  'Wend
  'Set rst = Nothing
End Sub

Private Sub Form_DblClick()
'calcular_resumcomanda
'llistar_comanda True
' recalcular_comandescomplexes
End Sub
Sub calcular_resumcomanda()
  Dim rstcom1 As Recordset
  Dim rstcom2 As Recordset
  Dim rstcom3 As Recordset
  Dim rstresum As Recordset
  Dim numcom As String
  Dim instancia As Double
  Randomize
  instancia = Int((1000 * Rnd) + 1) * -1
  Set rstresum = dbtmp.OpenRecordset("resumcomanda")
  rstresum.AddNew
  rstresum!ID = instancia
  numcom = cadbl(Data1.Recordset!comanda)
  Set rstcom1 = dbtmp.OpenRecordset("select * from comandes where comanda=" + numcom)
  If rstcom1.EOF Then Exit Sub
  Set rstcom2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(rstcom1!linkcomanda1)))
  Set rstcom3 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(rstcom1!linkcomanda2)))
  rstresum!codiproducte = atrim(rstcom1!producte)
  rstresum!amplada = cadbl(rstcom1!amplereb)
  rstresum!desarroll = cadbl(rstcom1!dessarroll)
  

  'mat 1
  r = atrim(cadbl(rstcom1!materialex))
  Set rsttmp = dbtmp.OpenRecordset("select familia from materials where codi=" + r)
  If Not rsttmp.EOF Then rstresum!familiamat1 = rsttmp!familia
  rstresum!espesormat1 = micresmaterial(cadbl(rstcom1!mesuraesp), cadbl(rstcom1!espessor), atrim(rstcom1!tubolam))
  'mat2
  r = atrim(cadbl(rstcom2!materialex))
  Set rsttmp = dbtmp.OpenRecordset("select familia from materials where codi=" + r)
  If Not rsttmp.EOF Then rstresum!familiamat2 = rsttmp!familia
  rstresum!espesormat2 = micresmaterial(cadbl(rstcom2!mesuraesp), cadbl(rstcom2!espessor), atrim(rstcom2!tubolam))
  'mat3
  r = atrim(cadbl(rstcom3!materialex))
  Set rsttmp = dbtmp.OpenRecordset("select familia from materials where codi=" + r)
  If Not rsttmp.EOF Then rstresum!familiamat3 = rsttmp!familia
  rstresum!espesormat3 = micresmaterial(cadbl(rstcom3!mesuraesp), cadbl(rstcom3!espessor), atrim(rstcom3!tubolam))
  
  rstresum!espesor = rstresum!espesormat1 + rstresum!espesormat2 + rstresum!espesormat3
  
  'col1
  r = atrim(cadbl(rstcom1!colorex))
  Set rsttmp = dbtmp.OpenRecordset("select familia from colorants where codi=" + r)
  If Not rsttmp.EOF Then rstresum!familiacol1 = rsttmp!familia
  'col2
  r = atrim(cadbl(rstcom2!colorex))
  Set rsttmp = dbtmp.OpenRecordset("select familia from colorants where codi=" + r)
  If Not rsttmp.EOF Then rstresum!familiacol2 = rsttmp!familia
  'col3
  r = atrim(cadbl(rstcom3!colorex))
  Set rsttmp = dbtmp.OpenRecordset("select familia from colorants where codi=" + r)
  If Not rsttmp.EOF Then rstresum!familiacol3 = rsttmp!familia
  
  'id treball
  rstresum!codibarres = rstcom1!codibarras
  
  rstresum.Update
  Set rstresum = dbtmp.OpenRecordset("select * from resumcomanda where id=" + atrim(cadbl(instancia)))
  With rstresum
  r = !codiproducte + atrim(!amplada) + atrim(!desarroll) + atrim(desc_fam(!familiamat1)) + atrim(desc_fam(!familiacol1, 2)) + atrim(!espesormat1)
  If !familiamat2 > 0 Then r = r + atrim(desc_fam(!familiamat2)) + atrim(desc_fam(!familiacol2, 2)) + atrim(!espesormat2)
  If !familiamat3 > 0 Then r = r + atrim(desc_fam(!familiamat3)) + atrim(desc_fam(!familiacol3, 2)) + atrim(!espesormat3)
  r = r + atrim(!codibarres)
  MsgBox r
  End With
  Set rsttmp = dbtmp.OpenRecordset("select max(id) from resumcomanda ")
  If Not rsttmp.EOF Then proximid = cadbl(rsttmp!ID) + 1
  
  dbtmp.Execute "delete * from resumcomanda where id=" + atrim(cadbl(instancia))
  
  Set rstresum = Nothing
  Set rstcom1 = Nothing
  Set rstcom2 = Nothing
  Set rstcom3 = Nothing
  Set rsttmp = Nothing
End Sub
Function desc_fam(numfam As Long, Optional taula As Byte)
   Dim rstt As Recordset
   Dim nomtaula As String
   nomtaula = "familiesmaterials"
   If taula = 2 Then nomtaula = "familiescolorants"
   Set rstt = dbtmp.OpenRecordset("select descripcio from " + nomtaula + " where codi=" + atrim(numfam))
   If Not rstt.EOF Then desc_fam = atrim(rstt!descripcio)
   If Len(desc_fam) < 2 Then desc_fam = ""
End Function
Function micresmaterial(codimesuralineal As Byte, espesor As Double, tubolam As String) As Double
  Dim rstmesural As Recordset
  Set rstmesural = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(codimesuralineal), dbOpenSnapshot, dbReadOnly)
  If rstmesural.EOF Then Exit Function
  r = espesor
  If rstmesural!descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = Format(espesor / 4, "#,##0")
                  Else: r = Format(espesor / 2, "#,##0")
            End If
  End If
  If InStr(1, rstmesural!descripcio, "GR/") > 0 Then
    r = espesor * -1
  End If
  micresmaterial = r
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 123 Then buscar_tarifa_corresponent
  If duplicant Then KeyCode = 0: Exit Sub
 If KeyCode = 38 Or KeyCode = 40 Then KeyCode = 0
If Data1.Recordset.EditMode = 0 Then
 'If KeyCode = 65 Then alta_registre: KeyCode = 0
 If KeyCode = 69 Then buscar_registre
End If
 If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then
  usuari_guarda_registre
End If

If KeyCode = 13 Or KeyCode = 40 Then canviardecamp: KeyCode = 0
On Error GoTo cont
  If Not buscant And llistadecampsvalids <> "[tots]" And KeyCode > 100 And InStr(1, llistadecampsvalids, "[" + atrim(ActiveControl.DataField) + "]") = 0 Then KeyCode = 0: MsgBox "No pots canviar aquest camp", vbCritical, "Seguretat": Exit Sub
  On Error GoTo 0
cont:
If Shift = 2 And KeyCode = 49 And cadbl(vlink1) > 0 Then
     buscant = True
     queryorder = ""
     querywhere = "comanda=" + atrim(cadbl(vlink1))
     finalitzarbusqueda 1
End If
If Shift = 2 And KeyCode = 50 And cadbl(vlink2) > 0 Then
     buscant = True
     queryorder = ""
     querywhere = "comanda=" + atrim(cadbl(vlink2))
     finalitzarbusqueda 1
End If

If Shift = 2 And KeyCode = 51 And cadbl(vlink3) > 0 Then
     buscant = True
     queryorder = ""
     querywhere = "comanda=" + atrim(cadbl(vlink3))
     finalitzarbusqueda 1
End If
If Shift = 2 And KeyCode = 48 Then
     consultar_Click
'     buscant = True
'     queryorder = " comanda DESC"
'     querywhere = ""
'     busqueda_xr_formulari
End If


'34 pag avall    33 pag amunt
'esquerra 37  amunt 38 dreta 39 avall 40
'control    shift=2

'amb aquesta linia controlo els camps que canvient quan busco
If buscant And KeyCode <> 9 And KeyCode <> 0 And Shift = 0 Then
           Screen.ActiveControl.Tag = "9"
End If

If Shift = 2 And Data1.Recordset.EditMode = 0 Then
   If KeyCode = 38 Then
     If Data1.Recordset.AbsolutePosition > 0 Then Data1.Recordset.MovePrevious
   End If
   If KeyCode = 40 Then If Data1.Recordset.AbsolutePosition < Data1.Recordset.RecordCount - 1 Then Data1.Recordset.MoveNext
   If KeyCode = 37 Then If Not Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
   If KeyCode = 39 Then If Not Data1.Recordset.EOF Then Data1.Recordset.MoveLast
End If
If KeyCode = 34 Then
   pujar_formseccions
End If

If KeyCode = 33 Then
   abaixar_formseccions
End If


End Sub
Sub pujar_formseccions()
llocform = llocform + 1
   'If llocform > topeform Then llocform = topeform
   If taulapos(llocform) <> 0 Then
    formscrooll.SetValues formscrooll.Values.HorzValue, taulapos(llocform)
     Else: llocform = llocform - 1
   End If
   DoEvents
End Sub
Sub abaixar_formseccions()
If llocform <> 0 Then llocform = llocform - 1
    formscrooll.SetValues formscrooll.Values.HorzValue, IIf(llocform = 0, formscrooll.Values.VertMin, taulapos(llocform))
   DoEvents
End Sub
Sub canviardecamp(Optional enviotab As Boolean)
Call keybd_event(&H9, 0, 0, 0)
Call keybd_event(&H9, 0, KEYEVENTF_KEYUP, 0)
DoEvents
End Sub
Sub controlarseccions()
 
On Error Resume Next
'amb aixó reposiciono la secció a pantalla
'If formscrooll.Tag = "1" Or Screen.ActiveControl.Name = "formscrooll"  Then
'   formscrooll.SetFocus
'   formscrooll.Tag = ""
'   Exit Sub
'End If

If cadbl(formscrooll.Tag) < 10 And cadbl(formscrooll.Tag) > 0 Then
    formscrooll.Tag = cadbl(formscrooll.Tag) + 1
    Exit Sub
End If
contador = 0

If estattaula.Caption = "" Or InStr(1, estattaula, "Buscant") <> 0 Then Exit Sub
If Screen.ActiveForm.Name = "formcomandes" Then
 
 If Mid(Screen.ActiveControl.Container, 1, 1) <> "" Then
    llocform = InStr(1, ruta, Mid(Screen.ActiveControl.Container, 1, 1)) - 1
    If taulapos(llocform) <> 0 Then
       formscrooll.Tag = "20"
       formscrooll.SetValues formscrooll.Values.HorzValue, taulapos(llocform)
       formscrooll.Tag = ""
    End If
     Else:
      llocform = 0
      'formscrooll.SetValues formscrooll.Values.HorzValue, formscrooll.Values.HorzMax
 End If
End If
 '*************************+

End Sub

Sub buscar_registre()
consultar_Click
End Sub
Sub alta_registre()
 If cap.Enabled = False Then
      areadedatos True
      On Error GoTo cont
      Data1.Recordset.AddNew
      On Error GoTo 0
      possarvalordcamps 255
      DoEvents
      Text1.Enabled = True
      'busco el mes gran i el poso a codi +1
      If Not buscant Then
        Set rsttmp = dbtmp.OpenRecordset("select max(comanda) as [grancodi] from comandes", , dbReadOnly)
        If Not rsttmp.EOF Then
          Text1 = atrim(cadbl(rsttmp!grancodi) + 1)
         Else: Text1 = "1"
        End If
        Data1.Recordset!comanda = cadbl(Text1)
      End If
      
      gravar_camps_extres_Data
     ' Text1.SetFocus
 End If
 On Error GoTo 0
 Exit Sub
cont:
 Resume
End Sub
Sub gravar_camps_extres_Data(Optional numc As String)
   If cadbl(numc) = 0 Then numc = Text1.Text
   On Error Resume Next
   dbtmp.Execute "insert into comandes_extres (comanda,data) values (" + numc + ",now)"
End Sub
Function sihihacompres(numc As Double) As Boolean
 If Not IsDate(alta.Tag) Then Exit Function
  If DateDiff("n", CVDate(alta.Tag), Now) > 10 Then
   If modificar.Tag = "campsalicia" And Data1.Recordset!proximaseccio <> "T" Then sihihacompres = controlarcampsavisaralicia: modificar.Tag = ""
'   If sihihacompres Then
'      relacions = hiharelacions(cadbl(Data1.Recordset!comanda), cadbl(Data1.Recordset!linkcomanda1), cadbl(Data1.Recordset!linkcomanda2))
'      If Not relacions Then sihihacompres = False
'   End If
  End If
   
End Function
Function elmaterialdelmighauriadesertractatdoscares() As Boolean
  Dim rst As Recordset
  Dim numc1 As String
  Dim numc2 As String
  Dim numc3 As String
  Dim codimat As Double
  Dim rstm As Recordset
  elmaterialdelmighauriadesertractatdoscares = False
  numc1 = cadbl(Data1.Recordset!comanda)
  numc2 = cadbl(Data1.Recordset!linkcomanda1)
  numc3 = cadbl(Data1.Recordset!linkcomanda2)
  If (numc2 > 0 And numc3 > 0) Then
    Set rst = dbtmp.OpenRecordset("Select * from comandes where producte='PC' and (comanda=" + atrim(numc1) + " or comanda=" + atrim(numc2) + " or comanda=" + atrim(numc3) + ")", , dbReadOnly)
    If Not rst.EOF Then
       If numc1 = rst!comanda Then
           codimat = cadbl(Text25)
             Else: codimat = cadbl(rst!materialex)
       End If
       Set rstm = dbtmp.OpenRecordset("select material2cares from materials where codi=" + atrim(codimat), , dbReadOnly)
       If Not rstm.EOF Then If Not rstm!material2cares Then elmaterialdelmighauriadesertractatdoscares = True
    End If
   End If
   Set rst = Nothing
   Set rstm = Nothing
End Function
Sub aviscomprovarcomplexamesimpresionormal(numc As Double)
 Dim rstc As Recordset
 Dim rstclixes As Recordset
 Set rstc = dbtmp.OpenRecordset("Select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(numc))
 If rstc.EOF Then GoTo fi
 If rstc!linkcomanda1 = 0 And rstc!linkcomanda2 = 0 Then GoTo fi
 Set rstc = dbtmp.OpenRecordset("Select producte,numtreball,numordremodificacio from comandes where comanda>0 and (comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + ") order by comanda")
 If InStr(1, rstc!producte, "PC") = 0 Then
   
   Set rstclixes = dbclixes.OpenRecordset("select formaimpresio from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(cadbl(rstc!numordremodificacio)))
   If rstclixes.EOF Then GoTo fi
   If atrim(rstclixes!formaimpresio) = "N" Then MsgBox "MATERIAL COMPLEXA I IMPRESIÓ NORMAL" + Chr(10) + "VIGILAR MATERIAL TRACTAT DOS CARES", vbExclamation, "Atenció"
 End If
fi:
 Set rstc = Nothing
 Set rstclixes = Nothing

End Sub
Sub gravar_registre()
Dim com As Double
Dim haentratbuscant As Boolean
Dim hihaerror As Boolean
Dim relacions As Boolean
Static complexes As Boolean
Frame1(0).Enabled = False
hihaerror = False
hihaalgugravant 1
ratoli "espera"
If Data1.Recordset.EditMode = 0 Then ratoli "normal": GoTo fi
  wait (1)
  com = cadbl(Data1.Recordset!comanda)
' On Error Resume Next
' Screen.ActiveControl.Text = passaradecimal(Screen.ActiveControl.Text)
' On Error GoTo 0
 If (areadatos.Enabled And Not buscant) Then
    If sihihacompres(com) Then cancelar_registre: GoTo fi
    If Data1.Recordset!direnvio = 0 And Not duplicant Then MsgBox "Has d'escullir direcció d'enviament": ratoli "normal": GoTo fi
    'If atrim(Data1.Recordset!producte) = "PC" And (cadbl(Data1.Recordset!linkcomanda1) > 0 And cadbl(Data1.Recordset!linkcomanda2) > 0) Then     MaskEdBox17 = "2"
    If (atrim(MaskEdBox17) <> "2" And elmaterialdelmighauriadesertractatdoscares) And Not duplicant Then
       MsgBox "Fulla amb dos procesos de laminació la fulla del mig ha de tenir el material tractat a dos cares." + Chr(10) + " El material marcat com Tractat 2 Cares i la comanda amb material tractat 2 cares", vbCritical, "Atenció"
       'If atrim(Data1.Recordset!producte = "PC") Then
       '  ratoli "normal"
       '  GoTo fi
       'End If
    End If
    If InStr(1, ruta, "R") And atrim(Combo9) = "" Then MsgBox "Has d'escullir un migelaborat de la secció de Rebobinadora.": ratoli "normal": GoTo fi
    If InStr(1, ruta, "I") > 0 Then
       posarmarcailinia numtreballdelacomanda(Data1.Recordset!comanda)
         Else:: Data1.Recordset!numtreball = 0: Data1.Recordset!numordremodificacio = 0: Data1.Recordset!impressora = 0
    End If
    guardar_Calloff
    Combo1(0) = Combo1(0) + "   "
    dbtmp.Execute "update comandes_extres set carametall='" + atrim(Mid(Combo1(0), 1, 1)) + "' where comanda=" + atrim(Data1.Recordset!comanda)
    dbtmp.Execute "update comandes_extres set desarrollclient=" + atrim(cadbl(Text32(9))) + " where comanda=" + atrim(Data1.Recordset!comanda)
    dbtmp.Execute "update comandes_extres set observacionsalbara='" + atrim(Text32(10)) + "' where comanda=" + atrim(Data1.Recordset!comanda)
    dbtmp.Execute "update comandes_extres set est_o_past='" + atrim(Combo1(4)) + "' where comanda=" + atrim(Data1.Recordset!comanda)
    possarprimerprocesalpc2
    possarordrelaminadora vlink1
    text77(18) = generarlinialaminacio(Data1.Recordset!comanda)
    generarlinialaminacio cadbl(Data1.Recordset!linkcomanda1)
    generarlinialaminacio cadbl(Data1.Recordset!linkcomanda2)
    Text63 = calcular_tinters
  ' desactivo el calcul de pes al guardar per provar problemes de valors mal guardats     02/07/2024
    Text33 = calcular_pes1000kg
    'p100ossarpes1000metresalescomplexes cadbl(Text1)
    'calcular_pesmtr2imetresrebipesreb
    Text1.Enabled = False
    
    DoEvents
    
     formcomandes.Tag = "100"
    
     On Error GoTo rutinaerror
      Data1.Recordset.Update
      If Data1.Recordset.EditMode > 0 Then Data1.Recordset.Update
      If Data1.Recordset.EditMode > 0 And hihaerror Then hihaerror = False: Data1.Recordset.Update
      If Data1.Recordset.EditMode > 0 And hihaerror Then hihaerror = False: Data1.Recordset.Update
      If Data1.Recordset.EditMode > 0 And hihaerror Then hihaerror = False: Data1.Recordset.Update
      If Data1.Recordset.EditMode > 0 And hihaerror Then Data1.Recordset.Update
      possarordrelaminadora vlink1
'      If modificar.Tag = "campsalicia" And Data1.Recordset!proximaseccio <> "T" Then relacions = hiharelacions(cadbl(Data1.Recordset!comanda), cadbl(Data1.Recordset!linkcomanda1), cadbl(Data1.Recordset!linkcomanda2))
'      If modificar.Tag = "campsalicia" And relacions And Data1.Recordset!proximaseccio <> "T" Then controlarcampsavisaralicia: modificar.Tag = ""
      If cadbl(Data1.Recordset!numtreball) = 0 And InStr(1, ruta, "I") > 0 And InStr(1, "TPV", Data1.Recordset!proximaseccio) = 0 Then passaravisevasinoteidtreball cadbl(Data1.Recordset!comanda)
      If cadbl(Data1.Recordset!numtreball) > 0 And InStr(1, ruta, "I") = 0 Then Data1.Recordset!numtreball = 0
      areadedatos False
      formcomandes.Tag = ""
      Data1.RecordSource = "select * from comandes where comanda=" + atrim(com)
      Data1.Refresh
      control_campsqueafectencanvidepreu
      If Command9(0).BackColor = &HC0FFC0 Then enviaremailsihihauncanvienlescampsqueefectenlestarifes
      control_de_modificacions
   ' End If
 End If
 If buscant Then
    haentratbuscant = True
    finalitzarbusqueda
   Else
     If Not complexes Then
       If cadbl(Data1.Recordset!linkcomanda1) > 0 Or cadbl(Data1.Recordset!linkcomanda2) > 0 Then
        actualitzar_comuns_complexes cadbl(Data1.Recordset!linkcomanda1), cadbl(Data1.Recordset!linkcomanda2)
        complexes = False
       End If
     End If
 End If
 possarvalordcamps
 'actualitzo els camps comuns en les fulles complexes
 
 'control de impresio normal amb complexes
 aviscomprovarcomplexamesimpresionormal com
  
 'formscrooll.SetValues formscrooll.Values.HorzValue, formscrooll.Values.HorzMin    'taulapos(llocform)
 ratoli "normal"
fi:
  
  If Not haentratbuscant Then
     diferenciescomandaitreball
     copiarmontadoraireferenciesaltreball cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), cadbl(Data1.Recordset!client), cadbl(Data1.Recordset!direnvio)
     modificar_refinplacsa_sical cadbl(Data1.Recordset!comanda), Data1.Recordset!proximaseccio
  End If
   activaronocampsimpresio False
   Frame1(0).Enabled = True
   formcomandes.SetFocus
  hihaalgugravant 2
 Exit Sub
rutinaerror:
hihaerror = True

Select Case err.Number  ' Evalúa el número de error.
        Case 3260 ' Error "Archivo ya está abierto".
           MsgBox "Aquest registre està bloquejat." + Chr$(10) + err.Description

       ' Case 3197 ' Error "Archivo ya está abierto".
       '    MsgBox "Aquest registre està bloquejat per algú altra." + Chr$(10) + err.Description
        Case Else
            'MsgBox "Ni ha hagut un error " + atrim(err.Description)
        ' Puede incluir aquí otras situaciones...
    End Select
    Resume Next ' Continuar ejecución en la línea que
                ' causó el error.
formcomandes.SetFocus
End Sub
Sub guardar_Calloff()
    Dim vnumc As Double
    Dim vcalloffanterior As String
    vnumc = cadbl(Data1.Recordset!comanda)
      If Combo1(2).ListCount = 0 Then
         vcalloffanterior = buscarcalloffgeneric(vnumc)
         buscarcalloffgeneric vnumc, Combo1(2)
         If vcalloffanterior <> Combo1(2) Then
            dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(vnumc) + ",'" + nomordinador + "','NumCalloff','" + vcalloffanterior + "','" + atrim(Combo1(2)) + "')"
         End If
      End If
End Sub
Sub modificar_refinplacsa_sical(numcomanda As Double, proximaseccio As String)
  Dim rst As Recordset
  If proximaseccio = "T" Then Exit Sub
  Set rst = dbtmp.OpenRecordset("select comandaimpresa from comandes_extres where comanda=" + atrim(numcomanda))
  If Not rst.EOF Then
        If rst!comandaimpresa Then generarrefinplacsadefinitiu cadbl(numcomanda)
  End If
  Set rst = Nothing
End Sub
Function comandaimpresa(numc As Double) As Boolean
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select comandaimpresa from comandes_extres where comanda=" + atrim(numc))
  If Not rst.EOF Then comandaimpresa = rst!comandaimpresa
  Set rst = Nothing
End Function
Sub avisar_a_tintes_canvidemetres(vnumc As Double, vmetresabans As Double, vmetresara As Double)
   enviaremailgeneric "tintes@inplacsa.com", "La comanda " + atrim(vnumc) + " s'ha modificat els metres.", "La comanda " + atrim(vnumc) + " tenia " + atrim(vmetresabans) + " metres i ara té " + atrim(vmetresara) + " metres."
   
End Sub
Sub control_campsqueafectencanvidepreu()
  Dim i As Long
   If rstcontrolcanvis.EOF Or duplicant Then Exit Sub
   If Not comandaimpresa(Data1.Recordset!comanda) Then GoTo fi
   With rstcontrolcanvis
   If cadbl(!cantitatex) <> cadbl(Data1.Recordset!cantitatex) Then avisar_a_tintes_canvidemetres Data1.Recordset!comanda, cadbl(!cantitatex), cadbl(Data1.Recordset!cantitatex)
   If cadbl(rstcontrolcanvis!pvp) = 0 Then GoTo fi
   If cadbl(!cantitatex) <> cadbl(Data1.Recordset!cantitatex) Or cadbl(!materialex) <> cadbl(Data1.Recordset!materialex) Or cadbl(!espessor) <> cadbl(Data1.Recordset!espessor) Then
        MsgBox "S'han canviat el camps de quantitat, espessor o material i aixó afecta al preu donat al client." + Chr(10) + "SISPLAU REVISEU QUE EL PREU SIGUI CORRECTE", vbCritical, "Atenció"
   End If
   End With
fi:
End Sub
Sub enviaremailsihihauncanvienlescampsqueefectenlestarifes()
  Dim camps As String
  Dim vmsg As String
  Dim vcamp As String
    'he tret #Ltipusadhesiu
  camps = "#Edirenvio#Etubolam#Eampleesq#Eplegatesq#Esolapa#Eespessor#Emicropex#Eoberturaex#Ematerialex#Ecantitatex#Etubbaseext#Inumtreball#IDESSARROLL#INUMEROTINTES#Lampleutil#Lsimulteneitatlam#Rmigelaborat#Ramplereb#Rsimulteneitatreb"
  camps = camps + "#Smigelaboratsol#Samplesol#Sampleplegsol#Slongitudsol#Ssolapasol#Sfuellebasesol#Sfuellebocasol#Sespessorsol#Stroquel#Sansa#Scinta#"
  camps = UCase(camps)
   For i = 0 To Data1.Recordset.Fields.Count - 1
     If InStr(1, camps, UCase(Data1.Recordset.Fields(i).Name + "#")) > 0 Then
      If atrim(rstcontrolcanvis.Fields(i)) <> atrim(Data1.Recordset.Fields(i)) Then
         vcamp = atrim(rstcontrolcanvis.Fields(i).Name)
         'canvio el nomdel camp perque aquest es un camp aprofitat i no s'entendria
         If vcamp = "tubbaseext" Then vcamp = "quantitatdemanada"
         If vcamp <> "direnvio" And vcamp <> "materialex" Then
             vmsg = vmsg + "Camp: " + vcamp + " Valor: " + atrim(rstcontrolcanvis.Fields(i)) + " -> " + atrim(Data1.Recordset.Fields(i)) + Chr(13) + Chr(10)
             vcamp = ""
         End If
         If vcamp = "materialex" And vcamp <> "" Then
             vmsg = vmsg + "Camp: Material  Valor: " + nomdelmaterial(cadbl(rstcontrolcanvis.Fields(i))) + " -> " + nomdelmaterial(cadbl(Data1.Recordset.Fields(i))) + Chr(13) + Chr(10)
             vcamp = ""
         End If
         If vcamp <> "" And vcamp = "direnvio" Then vmsg = vmsg + "Camp: Canvi direcció enviament: " + Label1(147) + Chr(13) + Chr(10)
      End If
     End If
   Next i
   If vmsg <> "" Then
       vcap = "Codi client: " + atrim(Data1.Recordset!client) + "-" + nomclient + vbNewLine
       vcap = vcap + "Ref.Client: " + atrim(Data1.Recordset!refclient) + IIf(atrim(Data1.Recordset!comandaclient) <> "", " Com.Cli: " + atrim(Data1.Recordset!comandaclient), "") + vbNewLine
       vcap = vcap + "Texte Imp: " + atrim(Data1.Recordset!marcailinia) + vbNewLine + vbNewLine + "S'ha canviat:" + vbNewLine
       enviaremailgeneric "comandesrevisarpreus@inplacsa.com", "La comanda " + atrim(Data1.Recordset!comanda) + " revisar preu s'ha modificat paràmetres", vcap + Chr(13) + Chr(10) + Chr(13) + Chr(10) + vmsg
       'comandesrevisarpreus@inplacsa.com
   End If
   
End Sub
Function nomdelmaterial(vcodi As Double)
    Dim rstmat As Recordset
    Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstc!materialex)), , ReadOnly)
    If Not rstmat.EOF Then nomdelmaterial = atrim(rstmat!descripcio)
    Set rstmat = Nothing
End Function
Sub control_de_modificacions()
   Dim i As Long
   Dim rst As Recordset
   Dim vcampafectat As String
   If rstcontrolcanvis.EOF Then Exit Sub
   If cadbl(rstcontrolcanvis!comanda) <> cadbl(Data1.Recordset!comanda) Then Exit Sub
   If Not comandaimpresa(Data1.Recordset!comanda) Then GoTo fi
   'MsgBox rstcontrolcanvis.Fields(i).Name
   For i = 0 To Data1.Recordset.Fields.Count - 1
      If atrim(rstcontrolcanvis.Fields(i)) <> atrim(Data1.Recordset.Fields(i)) Then
         'guardar el control de canvis
         vcampafectat = rstcontrolcanvis.Fields(i).Name
         If vcampafectat = "tubbaseext" Then vcampafectat = "Quan_demanada"
         dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','" + vcampafectat + "','" + atrim(rstcontrolcanvis.Fields(i)) + "','" + atrim(Data1.Recordset.Fields(i)) + "')"
      End If
   Next i
   Set rst = dbtmp.OpenRecordset("select * from comandes_extres where comanda=" + atrim(cadbl(Text1)))
   If rst.EOF Then GoTo fi
   For i = 0 To rst.Fields.Count - 1
      If atrim(rstcontrolcanvis_extres.Fields(i)) <> atrim(rst.Fields(i)) Then
         'guardar el control de canvis
         dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(rst!comanda) + ",'" + nomordinador + "','" + rstcontrolcanvis_extres.Fields(i).Name + "','" + atrim(rstcontrolcanvis_extres.Fields(i)) + "','" + atrim(rst.Fields(i)) + "')"
      End If
   Next i
fi:
   Set rstcontrolcanvis = Nothing
   Set dbcontrolcanvis = Nothing
   Set rst = Nothing
End Sub
Function generarlinialaminacio(numc As Double) As String
   Dim rstpc As Recordset
   Dim rstpc2 As Recordset
   Dim rstcomextra As Recordset
   Dim descpc2 As String
   Dim descpc1 As String
   Dim desc As String
   Dim lot1 As String
   Dim lot2 As String
   Dim carametall As String
   Dim carametallpc1 As String
   Dim primerproces As Boolean
   primerproces = False
   'segon proces
   Set rstpc = dbtmp.OpenRecordset("SELECT comandes.lotmatdesb1,comandes.lotmatdesb2,comandes.linkcomanda1,comandes.linkcomanda2,comandes.producte,comandes.comanda,comandes.refilatd, comandes_extres.carametall FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda where comandes.comanda = " + atrim(numc), , ReadOnly)
   If Not rstpc.EOF Then
      If rstpc!producte = "PC" Or rstpc!producte = "PCP" Then Exit Function
      Set rstpc2 = dbtmp.OpenRecordset("select refilatd,comanda from comandes where (producte='PC2' or producte='PCI3') and (comanda=" + atrim(cadbl(rstpc!comanda)) + " OR comanda=" + atrim(cadbl(rstpc!linkcomanda1)) + " or comanda=" + atrim(cadbl(rstpc!linkcomanda2)) + " )")
      If rstpc2.EOF Then
         primerproces = False
           Else
             If rstpc2!comanda = numc And rstpc2!refilatd <> 0 Then primerproces = True
             If rstpc2!refilatd = 0 And rstpc2!comanda <> numc Then primerproces = True
      End If
      If primerproces Then
         If rstpc!producte = "PC2" Then
             lot1 = " (" + atrim(generardadescomanda(cadbl(rstpc!lotmatdesb1))) + ")"
             lot2 = " (" + atrim(generardadescomanda(cadbl(rstpc!lotmatdesb2))) + ")"
               Else:
                 lot1 = " (" + atrim(generardadescomanda(cadbl(rstpc!lotmatdesb1))) + ")"
                 lot2 = " (" + atrim(generardadescomanda(cadbl(rstpc!lotmatdesb2))) + ")"
         End If
           Else
             If rstpc!producte <> "PC2" And rstpc!producte <> "PCI3" Then
                  If cadbl(rstpc!linkcomanda2) > 0 Then
                        lot1 = " (" + atrim(generardadescomanda(cadbl(rstpc!linkcomanda1))) + "+" + atrim(generardadescomanda(cadbl(rstpc!linkcomanda2))) + ")"
                        lot2 = " (" + atrim(generardadescomanda(cadbl(rstpc!comanda))) + ")"
                      Else
                          primerproces = True
                           lot1 = " (" + atrim(generardadescomanda(cadbl(rstpc!lotmatdesb1))) + ")"
                           lot2 = " (" + atrim(generardadescomanda(cadbl(rstpc!lotmatdesb2))) + ")"
                   End If
                  Else
                    lot1 = " (" + atrim(generardadescomanda(cadbl(rstpc!linkcomanda1))) + "+" + atrim(generardadescomanda(cadbl(rstpc!linkcomanda2))) + ")"
                    'lot1 = "(" + atrim(generardadescomanda(cadbl(rstpc!lotmatdesb1))) + ")"
                    lot2 = " (" + atrim(generardadescomanda(cadbl(rstpc!lotmatdesb2))) + ")"
             End If
      End If
        Else: Exit Function
   End If
   
   
   desc = "LAMINAR "
   carametall = ""
   If rstpc!carametall = "D" Then carametall = " CARA METALL "
   If rstpc!carametall = "F" Then carametall = " CARA NO METALL "
       
   desc = desc + carametall + lot1 ' + carametall
  ' If rstpc!lotmatdesb1 = vlink1 Then desc = desc + lot1 + carametallpc1
  ' If rstpc!lotmatdesb1 = vlink2 Then desc = desc + "(" + atrim(generardadescomanda(vlink2)) + ")"
   desc = desc + " AMB "
   
   desc = desc + lot2
   'If rstpc!lotmatdesb2 = vlink1 Then desc = desc + lot2 + carametallpc1
   'If rstpc!lotmatdesb2 = vlink2 Then desc = desc + "(" + atrim(generardadescomanda(vlink2)) + ")"
   
   desc = IIf(primerproces, "1r. ", "2n. ") + desc
   dbtmp.Execute "update comandes set arxiuext='" + treure_apostruf(desc) + "' where comanda=" + atrim(numc)
   generarlinialaminacio = desc
   Set rstpc = Nothing
   Set rstp2 = Nothing
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
    Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc), , ReadOnly)
    
    If Not rstc.EOF Then
      '  Set rstd1 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(rstc!lotmatdesb1)))
      '  Set rstd2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(rstc!lotmatdesb2)))
      '  If rstd1.EOF Or rstd2.EOF Then Exit Sub
      '  If rstd!refilatd <> 1 Then
           Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstc!materialex)), , ReadOnly)
           If Not rstmat.EOF Then nommaterial = descripciomaterial2(rstmat, True)
           generardadescomanda = nommaterial
           'Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstd2!materialex)))
           'If Not rstmat.EOF Then nommaterial = descripciomaterial(rstmat, True)
           'generardadescomanda = generardadescomanda + " + " + nommaterial
        'End If
        
    End If
    Set rstc = Nothing
    Set rstmat = Nothing
End Function

Function estatclixemod(ByVal ntreball As Double, ByVal ordrem As Double) As String
  estatclixemod = estatdelclixe(ntreball, ordrem)
  imp1.Caption = "Impresora          Estat Clixes: " + estatclixemod
  
End Function
Function estatdelclixe(ByVal ntreball As Double, ByVal ordrem As Double) As String
Dim rst As Recordset
  If ordrem = 0 Then ordrem = 1
  Set rst = dbclixes.OpenRecordset("SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, CLIXES_MODIFI.data_prevista,Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE Clixes_modifi.id_treball=" + atrim(ntreball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordrem) + " AND clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(ntreball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordrem) + ");")
  ' CONSULTA ULTIMA MODIFICACIO AMB DATA FI  VINCULANT "SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE (((Clixes_modifi.id_treball)=" + atrim(id_treball) + ") AND ((Clixes_modifi.ordremodificacio)=" + atrim(ordremodificacio) + ") AND ((Clixes_estats.vinculant)=True and isdate(clixes_modifi.data_fi)) and clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + "));"
  ' CONSULTA ULTIMA MODIFICACIO AMB DATA FI SENSE VINCULANT "SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE (((Clixes_modifi.id_treball)=" + atrim(id_treball) + ") AND ((Clixes_modifi.ordremodificacio)=" + atrim(ordremodificacio) + ") AND (isdate(clixes_modifi.data_fi)) and clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + "));"
  If Not rst.EOF Then
     estatdelclixe = IIf(Not IsDate(rst!data_fi), Format(rst!data_prevista, "dd/mm") + " - ", "") + atrim(rst!descrip)
       Else: estatdelclixe = ""
  End If
End Function
Sub diferenciescomandaitreball()
 'posardiferenciesacomandadeltreball
 If duplicant Then Exit Sub
  If atrim(Data1.Recordset!impressio) = "R" Then
   mirardiferenciescomandaitreball cadbl(Text1)
   If Not imprimirdiferenciescomandaitreball(cadbl(Text1)) Then
      'no ha impres cap canvi per tan passo la modificacio a 1 si estava a 0
    If cadbl(Data1.Recordset!numordremodificacio) = 0 Then
      dbtmp.Execute "update comandes set numordremodificacio=" + atrim(modificaciodeltreballmesgran(cadbl(Data1.Recordset!numtreball))) + " where comanda=" + atrim(cadbl(Text1))
      dbtmp.Execute "update comandes_Extres set aviscanvisambeltreball='Ordre Modificacio a 1, no diferencies' where comanda=" + atrim(cadbl(Text1))
      Text103(3) = atrim(cadbl(Data1.Recordset!numtreball)) + "/" + atrim(modificaciodeltreballmesgran(cadbl(Data1.Recordset!numtreball)))
    End If
   End If
  End If
 comprovarestatclixe True
End Sub
Sub comprovarestatclixe(ensenyarmissatge As Boolean)
  Dim nordre As Integer
  Dim vubicacio As String
  Dim vdatabaixa As String
  nordre = cadbl(Data1.Recordset!numordremodificacio)
  If nordre = 0 Then nordre = 1
  If InStr(1, estatdelclixe(cadbl(Data1.Recordset!numtreball), nordre), "RETORNEM CLIXES") > 0 Then
     If ensenyarmissatge Then MsgBox "NO TENIM ELS CLIXES NOSALTRES" + Chr(10) + " ELS TE EL CLIENT", vbCritical, "ATENCIÓ"
     imp1.BackColor = QBColor(12)
  End If
  vubicacio = ubicaciodelclixe(cadbl(Data1.Recordset!numtreball), vdatabaixa)
  If vdatabaixa <> "" Then imp1.BackColor = &H80FF&
  If Mid(vubicacio + "  ", 1, 2) = "P-" Then
     If ensenyarmissatge Then MsgBox "ELS CLIXES ESTAN EN UN PALET Nº:  " + atrim(vubicacio), vbCritical, "ATENCIÓ"
  End If

End Sub
Function ubicaciodelclixe(ByVal ntreball As Double, vdatabaixa As String) As String
  Dim rst As Recordset
  Set rst = dbclixes.OpenRecordset("select ubicacio,databaixaclixe from clixes where id_treball=" + atrim(ntreball))
  If Not rst.EOF Then
     ubicaciodelclixe = atrim(rst!ubicacio)
     vdatabaixa = atrim(rst!databaixaclixe)
  End If
End Function

Sub passaravisevasinoteidtreball(numc As Double)
   Dim db As Database
   Set db = OpenDatabase(camiclixes)
   db.Execute "insert into Avisoscomandessenseidtreball (comanda) values (" + atrim(numc) + ")"
   Set db = Nothing
   
End Sub
Function valorcontrolpantalla(Camp As String) As String
   Dim objecte As Control
   On Error GoTo proxim
   For Each objecte In formcomandes
     If TypeOf objecte Is MaskEdBox Or TypeOf objecte Is TextBox Or TypeOf objecte Is ComboBox Or TypeOf objecte Is CheckBox Then
        If objecte.DataField = Camp Then valorcontrolpantalla = objecte.Text
     End If
proxim:
   Next
End Function
Function controlarcampsavisaralicia() As Boolean
    Dim i As Long
    i = 0
    While campscontrolalicia(i, 0) <> ""
      If valorcontrolpantalla(campscontrolalicia(i, 0)) <> campscontrolalicia(i, 1) Then
        If campscontrolalicia(i, 0) = "materialex" Then
        'SI ES MATERIALEX MIRAR SI LES FAMILIES ES CORRESPONEN
          If Not mirarsilesfamiliescorresponen(cadbl(campscontrolalicia(i, 1)), valorcontrolpantalla(campscontrolalicia(i, 0))) Then
            GoTo gravar
              Else: GoTo cont
          End If
         End If
gravar:
        
        relacions = hiharelacions(cadbl(Data1.Recordset!comanda), cadbl(Data1.Recordset!linkcomanda1), cadbl(Data1.Recordset!linkcomanda2))
        If relacions Then
           If Not (UCase(campscontrolalicia(i, 0)) = "DATAACTIVACIO" And IsDate(valorcontrolpantalla(campscontrolalicia(i, 0)))) Then
              gravar_avis_alicia cadbl(Data1.Recordset!comanda), " CAMP: " + UCase(campscontrolalicia(i, 0)) + " " + campscontrolalicia(i, 1) + "->" + valorcontrolpantalla(campscontrolalicia(i, 0)), UCase(campscontrolalicia(i, 0)), campscontrolalicia(i, 1), valorcontrolpantalla(campscontrolalicia(i, 0))
           End If
           'MsgBox "Aquesta comanda ja te una " + reservaassignacioocompra + " feta hauries de parlar amb COMPRES per arreglar-ho abans de fer aquests canvis.", vbCritical, "Canvis compres"
           If campscontrolalicia(i, 0) = "dataactivacio" Then
              'If Not IsDate(valorcontrolpantalla(campscontrolalicia(i, 0))) Then
                  'controlarcampsavisaralicia = True
                  MsgBox "Aquesta comanda ja te una " + reservaassignacioocompra + " feta, s'havisarà a compres per arreglar-ho però passarem a desactivada igualment.", vbCritical, "Canvis compres"
              'End If
           End If
           'GoTo sortir
        End If
        If campscontrolalicia(i, 0) = "dataactivacio" Then
           If valorcontrolpantalla(campscontrolalicia(i, 0)) = "" Then
              guardarinformaciodesactivacio cadbl(Data1.Recordset!comanda)
                Else
                  comprovarsihihaamuntadoraeltreballiavisar cadbl(Data1.Recordset!comanda)
                    'treure la informació perque es va desactivar si es que es va fer
                  dbtmp.Execute "update informaciodesactivades set actiu=false where comandaoreferencia='" + atrim(Data1.Recordset!comanda) + "'"
           End If
        End If
cont:
      End If
      i = i + 1
    Wend
sortir:
    reservaassignacioocompra = ""
End Function
Sub comprovarsihihaamuntadoraeltreballiavisar(vcomanda As Double)
    Dim rst As Recordset
    Dim rst2 As Recordset
    Dim vresp As String
    Dim vcos As String
    Dim vnumtreball As Double
    'si no està maracar el check de passar a impresores no cal mirar si hi ha treballs (en principi son tots menys crops)
    Set rst2 = dbtmp.OpenRecordset("select passaraimpresores from comandes_extres where comanda=" + atrim(vcomanda))
    If Not rst2.EOF Then If rst2!passaraimpresores = 0 Then Exit Sub
    
    Set rst = dbtmp.OpenRecordset("select numtreball from comandes where comanda=" + atrim(vcomanda))
    If rst.EOF Then GoTo fi
    vnumtreball = cadbl(rst!numtreball)
    If vnumtreball < 1 Then Exit Sub
    Set rst = dbbaixes.OpenRecordset("SELECT muntadora_ordremuntatge.comanda, comandes.numtreball FROM muntadora_ordremuntatge INNER JOIN comandes ON muntadora_ordremuntatge.comanda = comandes.comanda where muntadora_ordremuntatge.comanda<>" + atrim(vcomanda) + " and numtreball=" + atrim(vnumtreball))
    If Not rst.EOF Then
       While vresp <> "TREBALL A MUNTADORA"
         vresp = UCase(InputBox(vbCr & vbCr & vbCr & vbCr & vbCr + "Hi ha una comanda entrada a muntadora amb aquest mateix treball." & vbCr & "Escriu [TREBALL A MUNTADORA] per continuar." & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr & vbCr, "ATENCIÓ"))
       Wend
       'avisar per email
       vcos = "ATENCIÓ!!! s'ha activat la comanda " + atrim(vcomanda) + " i el mateix treball està entrat a ordre de muntatge de muntadora." + Chr(13) + Chr(10) + "Comanda trobada a muntadora: " + atrim(rst!comanda) + "  NºTreball: " + atrim(rst!numtreball)
       enviaremailgeneric "ComandaAMBtreballaMuntadora", "URGENT!!! Comanda activada amb el treball a muntadora.", treure_apostruf(vcos)
       'tintes@inplacsa.com;impresores@inplacsa.com
    End If
fi:
    Set rst = Nothing
    Set rst2 = Nothing
 End Sub
Sub guardarinformaciodesactivacio(numc As Double)
   Dim vdescripcio As String
   Dim vnomclient As String
   vdescripcio = InputBox("Entra una descripció de perquè desactives aquesta comanda", "Motiu per la desactivació")
   vnomclient = nomclient.Caption
   If vdescripcio = "" Then vdescripcio = "Desactivada sense possar motiu."
   vdescripcio = "[" + nomordinador + "] " + treure_apostruf(UCase(vdescripcio))
   dbtmp.Execute "Insert into informaciodesactivades (comandaoreferencia,data,tipus,nomclient,descripcio) values ('" + atrim(numc) + "',now,'P','" + vnomclient + "','" + vdescripcio + "')"
   
End Sub
Function mirarsilesfamiliescorresponen(matant As Long, matnou As Long) As Boolean
  Dim rstmatant As Recordset
  Dim rstmatnou As Recordset
  On Error GoTo 0
  Set rstmatant = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(matant))
  Set rstmatnou = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(matnou))
  mirarsilesfamiliescorresponen = False
  If rstmatant.EOF Then Exit Function
  If rstmatant!familia <> rstmatnou!familia Then
      Exit Function
  End If
  If rstmatant!subfamilia <> rstmatnou!subfamilia Then
      Exit Function
  End If
  If rstmatant!familiacol <> rstmatnou!familiacol Then
     Exit Function
  End If
  If rstmatant!subfamiliacol <> rstmatnou!subfamiliacol Then
     Exit Function
  End If
  If rstmatant!familiaad <> rstmatnou!familiaad Then
     Exit Function
  End If
  If rstmatant!subfamiliaad <> rstmatnou!subfamiliaad Then
     Exit Function
  End If
     
  mirarsilesfamiliescorresponen = True
End Function
Sub gravar_avis_alicia(numc As Double, descripcio As String, vcamp As String, vinicial As String, vfinal As String)
   Dim rst As Recordset
   Dim descripcioantiga As String
   
   Set rst = dbtmp.OpenRecordset("select * from aviscampsmodificats where comanda=" + atrim(numc))
      'amb aquestes linies seguents recuperava la ultima modificcio de camps i els concatenava ara
        ' cada modificacio ha de quedar un registre
   'If Not rst.EOF Then
   '  descripcioantiga = atrim(rst!descripciomodificacio)
   '  rst.Delete
   '  descripcio = descripcio + "  |  " + descripcioantiga
   'End If
   
   rst.AddNew
   rst!comanda = numc
   rst!descripciomodificacio = Mid(descripcio, 1, 254)
   rst!ordinadorcanvi = atrim(Environ("computername"))
   rst!datacanvi = Now
   rst!Camp = vcamp
   rst!valorinicial = vinicial
   rst!valorfinal = vfinal
   rst.Update
   Set rst = Nothing

End Sub
Function noestaduplicat(avis As String, fitxer As String) As Boolean
   Dim linia As String
   noestaduplicat = True
   If existeix(fitxer) Then
    Open fitxer For Input Access Read As #1
    While Not EOF(1)
     Line Input #1, linia
     If linia = avis Then noestaduplicat = False
    Wend
    Close 1
   End If
End Function
Sub actualitzar_comuns_complexes(numcomanda As Double, numcomanda2 As Double)
    Dim rst As Recordset
    Dim ample As Double
    Dim canviarample As Boolean
    canviarample = True
    wait (1)
    
    Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + Trim(numcomanda2))
    If Not rst.EOF And numcomanda2 > 0 Then
       ample = cadbl(rst!ampleesq)
         Else: ample = cadbl(Data1.Recordset!ampleesq)
    End If
    If InStr(1, Data1.Recordset!producte, "PC") <> 0 Then Exit Sub
    Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + Trim(numcomanda))
    If Not rst.EOF And numcomanda > 0 Then
       If atrim(rst!producte) = "PCP" Or atrim(Data1.Recordset!producte) = "PCP" Then Exit Sub
       If rst!ampleesq <> Data1.Recordset!ampleesq Or ample <> Data1.Recordset!ampleesq Then
          'If MsgBox("Els valors de ample d'extrusora son diferents a les fulles relacionades. Vols canviar-los", vbInformation + vbYesNo, "Atenció") = vbNo Then canviarample = False
          If UCase(InputBox("Els valors de ample d'extrusora son diferents a les fulles relacionades, Vols canviar-les?" + Chr(10) + " Escriu SI per canviar-les.", "Amplades diferents")) <> "SI" Then canviarample = False
       End If
     'canvio els camps comuns
       rst.Edit
       rst!client = Data1.Recordset!client
       rst!direnvio = Data1.Recordset!direnvio
       rst!dataactivacio = Data1.Recordset!dataactivacio
       rst!datamaterial = Data1.Recordset!datamaterial
       rst!refclient = Data1.Recordset!refclient
       rst!refclialt = Data1.Recordset!refclialt
       rst!comandaclient = Data1.Recordset!comandaclient
       rst!datacomanda = Data1.Recordset!datacomanda
       rst!dataentrega = Data1.Recordset!dataentrega
       rst!tipoentrega = Data1.Recordset!tipoentrega
       If canviarample Then rst!ampleesq = Data1.Recordset!ampleesq: rst!mesuracantex = Data1.Recordset!mesuracantex
       rst!cantitatex = Data1.Recordset!cantitatex
        rst!simulteneitatlam = Data1.Recordset!simulteneitatlam
       rst!ampleutil = cadbl(Data1.Recordset!ampleutil)
       rst!amplelaminar = cadbl(rst!ampleutil) * cadbl(Data1.Recordset!simulteneitatlam)
       rst!camisa = cadbl(rst!amplelaminar) + 1
       rst.Update
       dbtmp.Execute "update comandes_Extres set codicomptable=" + atrim(cadbl(Text32(3).Tag)) + " where comanda=" + atrim(cadbl(rst!comanda))
       rst.Edit
       rst!pes1000mtrs = calcular_pes1000kg(atrim(numcomanda))
       rst.Update
    End If
    'ara canvio els de linkcomanda2
    Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + Trim(numcomanda2))
    If Not rst.EOF And numcomanda2 > 0 Then
       'canvio els camps comuns
       rst.Edit
       rst!client = Data1.Recordset!client
       rst!direnvio = Data1.Recordset!direnvio
       rst!dataactivacio = Data1.Recordset!dataactivacio
       rst!refclient = Data1.Recordset!refclient
       rst!refclialt = Data1.Recordset!refclialt
       rst!comandaclient = Data1.Recordset!comandaclient
       rst!datacomanda = Data1.Recordset!datacomanda
       rst!dataentrega = Data1.Recordset!dataentrega
       rst!tipoentrega = Data1.Recordset!tipoentrega
       If canviarample Then rst!ampleesq = Data1.Recordset!ampleesq: rst!mesuracantex = Data1.Recordset!mesuracantex
       rst!cantitatex = Data1.Recordset!cantitatex
       rst!simulteneitatlam = Data1.Recordset!simulteneitatlam
       rst!ampleutil = cadbl(Data1.Recordset!ampleutil)
       rst!amplelaminar = cadbl(rst!ampleutil) * cadbl(Data1.Recordset!simulteneitatlam)
       rst!camisa = cadbl(rst!amplelaminar) + 1
       rst.Update
       dbtmp.Execute "update comandes_Extres set codicomptable=" + atrim(cadbl(Text32(3).Tag)) + " where comanda=" + atrim(cadbl(rst!comanda))
       rst.Edit
       rst!pes1000mtrs = calcular_pes1000kg(atrim(numcomanda2))
       rst.Update
    End If
    ratoli "espera"
    '-------------------------------------
    'ara vaig a les dues linkcomanda i faig un actualitzar
    'If numcomanda > 0 Then
    '  data1.RecordSource = "select * from comandes where comanda=" + atrim(numcomanda)
    '  data1.Refresh
    '  wait 2
    '  data1.Recordset.Edit
     ' wait 1
     '' data1.Recordset.Update
   ' End If
    
   ' If numcomanda2 > 0 Then
    ' data1.RecordSource = "select * from comandes where comanda=" + atrim(numcomanda2)
    '' data1.Refresh
    ' wait 2
    ' data1.Recordset.Edit
    ' wait 1
   '  data1.Recordset.Update
   ' End If
    '----------------------------------------------
    Set rst = Nothing
End Sub
Sub cancelar_registre()
  Dim marcareg As Variant
  'marcareg = Data1.Recordset.Bookmark
  If Data1.Recordset.EditMode > 0 Then
   Data1.Recordset.CancelUpdate
   areadedatos False
    enabled_campscontrolcodiinplacsa True
   cimpressio.Clear
   cimpressio.AddItem "Falta Autoritzar"
   'control_usuaris True
   Text1.Enabled = False
   buscant = False
   DoEvents
   DoEvents
   'On Error GoTo cont
  ' If Not Data1.Recordset.EOF Then
   '    data1.Recordset.MoveNext: data1.Recordset.MovePrevious
    ' Else: If Not data1.Recordset.BOF Then data1.Recordset.MovePrevious: data1.Recordset.MoveNext
   'End If
   'On Error GoTo 0
   
   refrescar
   ratoli "normal"
   Frame1(0).Enabled = True
   On Error GoTo cont
   If marcareg <> Empty Then Data1.Recordset.Bookmark = marcareg
   On Error GoTo 0
   ''''carregar_lookups
   possarvalordcamps
    Else:
       On Error Resume Next
       Unload Me
       On Error GoTo 0
  End If
  activaronocampsimpresio False
  Exit Sub
cont:
  refrescar
  Resume Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo cont
  If duplicant Then KeyAscii = 0: Exit Sub
  If Not buscant And llistadecampsvalids <> "[tots]" And InStr(1, llistadecampsvalids, "[" + atrim(ActiveControl.DataField) + "]") = 0 Then KeyAscii = 0: MsgBox "No pots canviar aquest camp", vbCritical, "Seguretat": Exit Sub
  On Error GoTo 0
cont:
  If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
  If KeyAscii > 50 Then KeyAscii = Asc(UCase(Chr$(KeyAscii)))
  
End Sub

Private Sub Form_Load()

 muntadora = False
 centerscreen Me
 taulapos(0) = -30768
' Data1.Connect = "Driver={SQL Server}; Server=serverprodu; Database=comandes--SQL2; uid=sa; pwd=Ipc123"
' cami = "\\SERVERPRODU\Dades\progcomandes\dades\comandes_copia.mdb"
 Data1.DatabaseName = cami
 Set dbtmp = OpenDatabase(Data1.DatabaseName)
 Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
 Set dbclixesnous = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
 Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
 Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificacio.mdb")
 Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "\palets.mdb")



 Data1.RecordSource = "select * from comandes where  client=0 or cantitatex=null  order by comanda DESC" '"comandes"
 Data1.RecordSource = "select * from comandes where  comanda=0"
 refrescar

 ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
 colorrisc = cap.BackColor
 llistadecampsvalids = llegir_ini("General", "llistadecampsvalids", fitxerini)
 r = llegir_ini("General", "llistadecampsvalids2", fitxerini)
 If r <> "{[}]" Then llistadecampsvalids = llistadecampsvalids + r
 r = llegir_ini("General", "llistadecampsvalids3", fitxerini)
 If r <> "{[}]" Then llistadecampsvalids = llistadecampsvalids + r
If InStr(1, llistadecampsvalids, "[pvp]") > 0 Then
    llistadecampsvalids = llistadecampsvalids + "[pvpdolar]"
      Else: If llistadecampsvalids <> "[tots]" Then Command9(3).Visible = False
End If
If Len(llistadecampsvalids) < 5 Then
  llistadecampsvalids = "": escriure_ini "General", "llistadecampsvalids", "", fitxerini
  modificar.Enabled = False: eliminar.Enabled = False: alta.Enabled = False: Command8.Enabled = False
End If

 Set ultimcontrol = Screen.ActiveControl
llocform = 0
VScroll1.Top = 750
VScroll1.Left = 10215
areadedatos False
If muntadora Then
   Command9(0).Visible = False
   Command9(1).Visible = False
   Command8.Visible = False
   Command23.Visible = False
   Text43.Visible = False
   Text6.Visible = False
   Text8.Visible = False
   eliminar.Visible = False
   alta.Visible = False
End If
'posso el vector de control de camps per avisar a l'alicia
'  campscontrolalicia(0, 0) = "impressio"
campscontrolalicia(4, 0) = "dataactivacio"
campscontrolalicia(0, 0) = "ampleesq"
campscontrolalicia(1, 0) = "espessor"
campscontrolalicia(2, 0) = "cantitatex"
campscontrolalicia(3, 0) = "materialex"
  
 'DESACTIVAR EL FONS DE TOTS ELS LABELS
 For i = 0 To formcomandes.Controls.Count - 1
      If TypeOf formcomandes.Controls(i) Is Label Then Label1(1).BackStyle = 0
 Next i
 Label1(162).BackStyle = 1
'--------------
 vnodemanarcontrasenyapassaraimpresores = False
 dbtmp.Execute "delete * from registreconsultescomandes where datediff('d',horainici,now)>30"
 If llegir_ini("General", "exportant", fitxerini) <> "1" Then comprovarcomandesambEsenseestarimpreses
 'If existeix("c:\temp\docx") Then  Kill "c:\temp\docx\*.*"
End Sub
Function refrescar()
'On Error Resume Next
If InStr(1, Data1.RecordSource, "{[}]") Then Data1.RecordSource = "comandes" 'Exit Function
'MsgBox Data1.RecordSource
If Data1.RecordSource = "#Temporary QueryDef#" Then
    Data1.RecordSource = "select * from comandes where comanda=" + Text1.Text
End If
'data1.RecordSource = "select * from comandes1 where materialex=559"
'If Data1.Recordset.EditMode > 0 Then Data1.Recordset.CancelUpdate
donaerrorelcontroldata_faun_Move0
Data1.Refresh
If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) And Not buscant Then
 Data1.Recordset.MoveLast
 Data1.Recordset.MoveFirst
 possarvalordcamps
End If
End Function
Sub donaerrorelcontroldata_faun_Move0()
    On Error GoTo fi
    If Data1.Recordset.EditMode >= 0 Then Data1.Recordset.Move 0
    Exit Sub
fi:

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ' If Button = 2 And Shift = 2 Then
 '    If MsgBox("vols començar a possar codis?", vbCritical + vbYesNo, "Atenció") = vbYes Then possarreferenciainplacsa
 ' End If
 ' If Button = 1 And Shift = 1 Then
 '    If MsgBox("vols genera el codi refinplacsa d'aquesta comanda?", vbCritical + vbYesNo, "Atenció") = vbYes Then generarrefinplacsadefinitiu cadbl(Text1)
 ' End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 vprimeraentradacomandes = False
  sortir_Click
End Sub

Private Sub formscrooll_Change(ByVal HorzValue As Integer, ByVal VertValue As Integer)
  If formscrooll.Tag = "20" Then Exit Sub
  formscrooll.Tag = "1"
  Text32(12).Visible = False
End Sub

Private Sub Frame1_Click(Index As Integer)
  'Dim vsql As String
  'Dim rst As Recordset
  'vsql = "SELECT comandes.comanda, clients.grupdeclient, Year([datacomanda]) AS Expr1, comandes.datacomanda, Month([datacomanda]) AS Expr2, comandes.pvp "
  'vsql = vsql + " FROM comandes LEFT JOIN clients ON comandes.client = clients.codi "
  ''vsql = vsql + " WHERE (((clients.grupdeclient)='ARDO') AND ((Year([datacomanda]))>2023) AND ((Month([datacomanda]))>9) AND ((comandes.pvp)>0));"
  'vsql = vsql + " WHERE (((clients.grupdeclient)='ARDO') AND ((Year([datacomanda]))>2023) AND ((comandes.pvp)>0));"
  'Set rst = dbtmp.OpenRecordset(vsql)
  'rst.MoveLast
  'rst.MoveFirst
  'While Not rst.EOF
  '   Me.Caption = atrim(rst.RecordCount) + "/" + atrim(rst.AbsolutePosition): DoEvents
  '            dbtmp.Execute "delete * from comandes_firmes where comanda=" + atrim(rst!comanda) + " and tipus='PVP'"
  '            dbtmp.Execute "insert into comandes_firmes (comanda,usuari,tipus,data) values (" + atrim(rst!comanda) + ",'ARDO_PVP1','PVP',now)"
  '            dbtmp.Execute "insert into comandes_firmes (comanda,usuari,tipus,data) values (" + atrim(rst!comanda) + ",'ARDO_PVP2','PVP',now)"
  '  rst.MoveNext
  'Wend

End Sub

Private Sub gravar_Click()


End Sub

Private Sub grmt2_LostFocus()
grmt2 = passaradecimal(grmt2)
possarconsums
End Sub

Private Sub importancia_DblClick()
Dim d As String
  d = InputBox("Entra el valor d'importancia de planificació.", "Entrada de importancia")
  If IsNumeric(d) Then
     importancia = cadbl(d)
     dbplanificacio.Execute "update  planificacioimp set imp_importancia=" + atrim(cadbl(d)) + " where comanda=" + atrim(Text1)
  End If
End Sub

Private Sub Label1_DblClick(Index As Integer)
  Dim v As String
  If Index = 18 Or Index = 181 Then
      If UCase(InputBoxEx("Entra la contrasenya de canvi:", "Canvi Espessor", , , , , , SPassword)) = "INPLACSA" Then
          v = InputBox("Entra el valor de l'espessor que vols possar a aquesta comanda.", "Canvi espessor", Text21)
          If StrPtr(v) <> 0 Then
              Text21 = atrim(cadbl(v))
          End If
      End If
  End If
  If Index = 146 Then
     If MsgBox("Segur que vols cambiar l'estat de Risc de la Comanda?", vbCritical + vbYesNo, "A T E N C I Ó...") = vbYes Then
        puntrisc.Visible = True
        If puntrisc.FillColor = &H80FF80 Then 'verd
         puntrisc.FillColor = colorrisc: Label1(146).Caption = "0": Label1(146).Tag = "1"
        End If
        If puntrisc.FillColor = &HFF& Then 'vermell
             puntrisc.FillColor = &H80FF80
             Label1(146).Caption = "2"
        End If
          'gris
        If cadbl(Label1(146).Caption) = 0 And Label1(146).Tag <> "1" Then puntrisc.FillColor = &HFF&: Label1(146).Caption = "1"
        Label1(146).Tag = ""
        If buscant Then Label1(146).Tag = "9"
     End If
  End If
  If Index = 147 Then
    comboenvios.Clear
    
    Set rsttmp = dbtmp.OpenRecordset("select * from clients_envios where codi=" + atrim(cadbl(Data1.Recordset!client)))
    While Not rsttmp.EOF
      comboenvios.AddItem rsttmp!nome + " | " + rsttmp!poblacioe
      comboenvios.ItemData(comboenvios.NewIndex) = rsttmp!ID
      rsttmp.MoveNext
    Wend
    If comboenvios.ListCount > 0 And comboenvios.Enabled Then
        comboenvios.Visible = True
        comboenvios.Text = "Escull un envio"
        comboenvios.SetFocus
        SendKeys "%{DOWN}"
    End If
  End If
  
End Sub
Sub ensenyar_informaciobasica_Comanda_Postit()
   Static vultimacomanda As Double
   Static vhoraultim As Date
   If vultimacomanda = cadbl(text77(11)) And DateDiff("s", vhoraultim, Now) < 30 Then GoTo ensenyaretiqueta ' si es la mateixa que la ultima vegada no la carrega
   vhoraultim = Now
   Text32(12) = ""
   Text32(12) = carregar_inforbasicacomanda(text77(11))
   
   If cadbl(text77(12)) > 0 Then Text32(12) = Text32(12) + "---------------------------------------------" + carregar_inforbasicacomanda(text77(12))
ensenyaretiqueta:
   If Text32(12) = "" Then Exit Sub
    'col.loco l'etiqueta sota el PVP
   Text32(12).Top = formscrooll.Top + text77(11).Top + text77(11).Height + 60
   Text32(12).Width = 8000
   Text32(12).Height = IIf(cadbl(text77(12)) = 0, 3000, 4000)
   Text32(12).Left = cap.Left + areadatos.Left + (text77(11).Left - Text32(12).Width)
   Text32(12).Visible = True
   Text32(12).Tag = "6"  ' aquest tag es el temps que tardarà a tancar-se 6 es 3 segons perque resta cada 0.5 segons Timer1
   vultimacomanda = cadbl(text77(11))
   'carrego la informació a l'etiqueta
   
End Sub
Function carregar_inforbasicacomanda(vnumc As Double)
   Dim rst As Recordset
   Dim vmsg As String
   Dim vm As String
   Dim vq As String
   Dim rst2 As Recordset
   
   If vnumc = 0 Then Exit Function
   Set rst2 = dbtmp.OpenRecordset("select * from mesureslineals")
   Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
      rst2.FindFirst "codi=" + atrim(cadbl(rst!mesuraesp))
      If Not rst2.NoMatch Then vm = rst2!descripcio
      rst2.FindFirst "codi=" + atrim(cadbl(cadbl(rst!mesuracantex)))
      If Not rst2.NoMatch Then vq = rst2!descripcio
      Set rst2 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(rst!materialex))
      
      vmsg = Chr(13) + Chr(10) + "ToL: " + atrim(rst!tubolam) + "  Ample/Plegat: " + atrim(cadbl(rst!ampleesq)) + "/" + atrim(cadbl(rst!plegatesq)) + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "   Espessor: " + atrim(cadbl(rst!espessor)) + vm + " Quant: " + atrim(cadbl(rst!cantitatex)) + vq
      
      If Not rst2.EOF Then
         vmsg = vmsg + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Material: " + atrim(descripciomaterial(rst2)) + Chr(13) + Chr(10)
         vmsg = vmsg + atrim(possargrmm2(cadbl(rst!materialex), micresmaterial(cadbl(rst!mesuraesp), cadbl(rst!espessor), atrim(rst!tubolam)))) + " G/m2" + vbNewLine
      End If
      
      Set rst2 = dbtmp.OpenRecordset("select materialexacte from comandes_extres where comanda=" + atrim(vnumc))
      If Not rst2.EOF Then
        If rst2.materialexacte Then vmsg = vmsg + Chr(13) + Chr(10) + " ATENCIÓ MATERIAL ESPECIFIC!!!"
      End If
   End If
   carregar_inforbasicacomanda = vmsg
End Function
Sub ensenyar_informacio_PVP()
   If Data1.Recordset.EditMode > 0 Then Exit Sub
  'col.loco l'etiqueta sota el PVP
   Text32(12) = ""
   Text32(12).Top = formscrooll.Top + Text16.Top + Text16.Height + 60
   Text32(12).Left = cap.Left + areadatos.Left + Text16.Left
   Text32(12).Width = 3000
   Text32(12).Height = 1000
   Text32(12).Visible = True
   Text32(12).Tag = "6"  ' aquest tag es el temps que tardarà a tancar-se 6 es 3 segons perque resta cada 0.5 segons Timer1
   
   'carrego la informació a l'etiqueta
   Text32(12) = carregar_PVP_teoric
   
   
End Sub
Function carregar_PVP_teoric() As String
 Dim vunitat As String
 Dim vquantitat As Double
 Dim vtotal As Double
 Dim rst As Recordset
 Dim vmsg As String
 Dim vrebmetres As Double
 Dim vrebpes As Double
 Dim vrebpcs As Double
 Dim vpesxrpeça As Double
 Dim vpesgrmcm2 As Double
 
 
 Set rst = dbtmp.OpenRecordset("SELECT mesures.unitatinterna FROM comandes INNER JOIN mesures ON comandes.mesurapvp = mesures.codi where comanda=" + atrim(cadbl(Text1.Text)), , ReadOnly)
 If rst.EOF Then Exit Function
 vunitat = rst!unitatinterna
 Set rst = dbtmp.OpenRecordset("select solpesgrmcm2 from comandes_extres where comanda=" + atrim(Text1.Text), , ReadOnly)
 vpesgrmcm2 = rst!solpesgrmcm2 'calcularpesunitatsoldadora(data1.Recordset)
 vpesxpeça = calcularpesxrpeça(Data1.Recordset, vpesgrmcm2)
 If InStr(1, ruta, "S") > 0 Then
     vrebpcs = cadbl(Data1.Recordset!cantitatsol)
     vrebpes = Redondejar(vrebpcs * vpesxpeça, 0)
        Else:
         If InStr(1, ruta, "R") > 0 Then
          vrebpcs = cadbl(rebpcs)
          vrebmetres = cadbl(Data1.Recordset!rebmtrs)
          vrebpes = cadbl(Data1.Recordset!rebkilos)
            Else
              If ruta = "E" Then
                vrebpes = cadbl(MaskEdBox6)
                vrebmetres = cadbl(MaskEdBox6)
                vrebpcs = cadbl(MaskEdBox6)
              End If
              If ruta = "EI" Then
                vquantitat = cadbl(MaskEdBox6)
                GoTo calculat
              End If
         End If
 End If
 
 Select Case vunitat
     Case "€/1000U"
       vquantitat = Redondejar(cadbl(vrebpcs) / 1000, 3)
     Case "€/U"
       vquantitat = cadbl(cadbl(vrebpcs))
     Case "€/B"
       If cadbl(Text103(0)) > 0 Then
          vquantitat = Redondejar(cadbl(vrebmetres) / cadbl(Text103(0)), 0)
           Else: vquantitat = cadbl(MaskEdBox6)
       End If
     Case "€/K"
        vquantitat = Redondejar(cadbl(vrebpes), 1)
     Case "€/M"
       vquantitat = cadbl(vrebmetres)
     Case "€/KM"
       vquantitat = Redondejar(cadbl(vrebmetres) / 1000, 2)
     Case "€/FIX"
       vquantitat = 1
     Case "€/M2"
       vquantitat = Redondejar(cadbl(vrebmetres) * (cadbl(Text121) / 1000), 2)
   End Select
calculat:
   vtotal = cadbl(Text6) * vquantitat
   vmsg = "    IMPORT COMANDA" + Chr(10) + "          " + Chr(13) + Chr(10) + "     " + atrim(vtotal) + " €"
   If vrebpes > 0 Then vmsg = vmsg + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "      " + atrim(Redondejar(atrim(vtotal) / vrebpes, 2)) + " €/KG  (" + atrim(Redondejar(vrebpes, 0)) + "Kg)"
   carregar_PVP_teoric = vmsg
   Set rst = Nothing
End Function

Private Sub label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Text32(12).Visible = False
End Sub

Private Sub label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  If Index = 174 Then
       ensenyar_informacio_PVP
  End If
  If Index = 133 And materialexacte(3).Value = 1 Then
       ensenyar_informaciobasica_Comanda_Postit
  End If
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If InStr(1, "12Nn", Chr$(KeyAscii)) = 0 Then
    MaskEdBox1.Text = "": KeyAscii = 0
   Else: MaskEdBox1.Text = ""
End If
End Sub

Private Sub MaskEdBox10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
     triaralgu "Triar Tub Base", "tubbase", MaskEdBox10, MaskEdBox10, "cm_int", 1
  End If

End Sub

Private Sub MaskEdBox11_Change()
'  If MaskEdBox11 <> "" Then
'    MaskEdBox11.Visible = True
'   Else: MaskEdBox11.Visible = False
'  End If
End Sub

Private Sub MaskEdBox11_GotFocus()
   comprovarsitereprintimicroperforats
End Sub

Private Sub MaskEdBox11_LostFocus()
  If cadbl(Text6) = 0 Then MsgBox "El valor en Euros del PVP també ha de tenir un import per poder calcular bé les estadistiques.", vbCritical, "Atenció"
End Sub

Private Sub MaskEdBox12_KeyDown(KeyCode As Integer, Shift As Integer)
 KeyCode = 0
End Sub

Private Sub MaskEdBox12_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub MaskEdBox13_Change()
' If MaskEdBox13.Text <> "" Then
'    MaskEdBox13.Visible = True
'   Else: MaskEdBox13.Visible = False
' End If
   
End Sub

Private Sub MaskEdBox15_KeyPress(KeyAscii As Integer)
If KeyAscii > 60 Then
    If InStr(1, "SN", UCase(Chr$(KeyAscii))) = 0 Then KeyAscii = 0
End If
End Sub

Private Sub MaskEdBox17_KeyPress(KeyAscii As Integer)
If InStr(1, "12Nn", Chr$(KeyAscii)) = 0 Then
    MaskEdBox17.Text = "": KeyAscii = 0
   Else: MaskEdBox17.Text = ""
End If
End Sub





Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
If InStr(1, "SN", Chr$(KeyAscii)) = 0 Then
    MaskEdBox2.Text = "": KeyAscii = 0
   Else: MaskEdBox2.Text = ""
End If
End Sub

Private Sub MaskEdBox20_LostFocus()
  If Trim(MaskEdBox20.Text) <> "" Then
    If Trim(MaskEdBox20.Text) <> "1,14" And Trim(MaskEdBox20.Text) <> "2,80" And Trim(MaskEdBox20.Text) <> "2,54" And Trim(MaskEdBox20.Text) <> "2,84" And Trim(MaskEdBox20.Text) <> "0" Then MsgBox "Valors de 1,14 o 2,54 o 2,84": MaskEdBox20.Text = "": MaskEdBox20.SetFocus
  End If
  
End Sub

Private Sub MaskEdBox22_LostFocus()
If IsDate(MaskEdBox22) And dataentrega2 = "" Then Text5 = MaskEdBox22
End Sub

Private Sub MaskEdBox6_GotFocus()
  Label1(170).Visible = True
  If Text32(7) = "" And Text2 = "6841" Then Text32(7) = "UNITAT": Text32(8) = "8"
End Sub

Private Sub MaskEdBox6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
     triaralgu "Triar Tub Base", "tubbase", MaskEdBox6, MaskEdBox6, "cm_int", 1
  End If
End Sub

Private Sub modificaciotreball_Change()

End Sub

Sub activaronocampsimpresio(activar As Boolean)
'   Dim objecte As Object
   For Each objecte In formcomandes
      If TypeOf objecte Is MaskEdBox Or TypeOf objecte Is TextBox Or TypeOf objecte Is ComboBox Then
          If objecte.HelpContextID = 99 Then
              
              objecte.Enabled = activar
              objecte.BackColor = &HFFC0C0
              If Not activar And objecte.DataField = "arxiu" And Text103(3) = "0/0" Then objecte.Enabled = True
              'MsgBox objecte.Name
          End If
      End If
   Next objecte
   Text2.Enabled = activar
   Text32(5).Enabled = activar
   comboenvios.Enabled = activar
   Text103(4).Enabled = activar
End Sub



Private Sub MaskEdBox6_LostFocus()
   Label1(170).Visible = False
End Sub

Private Sub materialexacte_GotFocus(Index As Integer)
  If Index = 0 Then
   If Screen.ActiveControl.Name = "materialexacte" Then
    If MsgBox("Segur que vols possar o treure aquesta comanda amb material exacte?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
      materialexacte(0).Value = IIf(materialexacte(0).Value = 1, 0, 1)
      dbtmp.Execute "update comandes_extres set materialexacte=" + IIf(materialexacte(0).Value = 1, "True", "False") + "  where comanda=" + atrim(cadbl(Text1.Text))
      dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','Material_exacte','" + IIf(materialexacte(0).Value = 1, "No", "Si") + "','" + IIf(materialexacte(0).Value = 1, "Si", "No") + "')"
    End If
   End If
   If Text18.Enabled Then Text18.SetFocus
  End If
   If Index = 4 Then
   If Screen.ActiveControl.Name = "materialexacte" Then
    If MaskEdBox12 <> "T" Then
      MsgBox "Aquesta casella només es pot desmarcar si la comanda està entregada." + Chr(10) + "Si encara no està impresa es pot tornar a fer el packing-list sense compatible i ja està.", vbCritical, "Error"
      materialexacte(4).Visible = False
      Exit Sub
    End If
    If MsgBox("Segur que vols desmarcar aquesta comanda per NO UTILITZAR MATERIALS COMPATIBLES?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
      materialexacte(4).Value = IIf(materialexacte(4).Value = 1, 0, 1)
      dbtmp.Execute "update comandes_extres set codigrupmaterialcompatible=0  where comanda=" + atrim(cadbl(Text1.Text))
      dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','MateriaCompatible','" + IIf(materialexacte(1).Value = 1, "No", "Si") + "','" + IIf(materialexacte(1).Value = 1, "Si", "No") + "')"
      materialexacte(4).Visible = False
    End If
   End If
   Text18.SetFocus
  End If
  If Index = 1 Then
   If Screen.ActiveControl.Name = "materialexacte" Then
    If MsgBox("Segur que vols possar o treure que el client vindrà a revisar la comanda?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
      materialexacte(1).Value = IIf(materialexacte(1).Value = 1, 0, 1)
      dbtmp.Execute "update comandes_extres set clientvindraarevisarimpresio=" + IIf(materialexacte(1).Value = 1, "True", "False") + "  where comanda=" + atrim(cadbl(Text1.Text))
      dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','clientvindraarevisarimpresio','" + IIf(materialexacte(1).Value = 1, "No", "Si") + "','" + IIf(materialexacte(1).Value = 1, "Si", "No") + "')"
    End If
   End If
   cimpressio.SetFocus
  End If
  If Index = 2 Then
   If Screen.ActiveControl.Name = "materialexacte" Then
    If MsgBox("Segur que vols possar o treure aquesta comanda amb cola exacte?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
      materialexacte(2).Value = IIf(materialexacte(2).Value = 1, 0, 1)
      dbtmp.Execute "update comandes_extres set colaexacte=" + IIf(materialexacte(2).Value = 1, "True", "False") + "  where comanda=" + atrim(cadbl(Text1.Text))
      dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(Data1.Recordset!comanda) + ",'" + nomordinador + "','Cola_exacte','" + IIf(materialexacte(2).Value = 1, "No", "Si") + "','" + IIf(materialexacte(2).Value = 1, "Si", "No") + "')"
    End If
   End If
   adhesiu.SetFocus
  End If
  
End Sub

Private Sub modificar_Click()
   Dim rstcm As Recordset
   If muntadora Then
     If InStr(1, ruta, "I") = 0 Then MsgBox "No pots editar un registre si no te seccio d'impresora": Exit Sub
   End If
   If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    areadedatos True
    refrescartreballimodificacio
    gravar_camps_extres_Data Text1.Text
    gravar_camps_extres_Data text77(11).Text
    gravar_camps_extres_Data text77(12).Text
    
   DoEvents
  ' If InStr(1, data1.Recordset!producte, "PC") = 0 Then
    Set rstcm = dbtmp.OpenRecordset("select data from comandes_extres where comanda=" + atrim(cadbl(Text1.Text)))
    If Not rstcm.EOF Then
     'If DateDiff("d", rstcm!Data, Now) > 0 Then
      alta.Tag = atrim(rstcm!Data)
      carregar_controlscampsalicia
      modificar.Tag = "campsalicia"
     'End If
      Else:
       carregar_controlscampsalicia
       modificar.Tag = "campsalicia"
       alta.Tag = ""
    End If
   'End If
   guardar_controlcanviscomanda
   Data1.Recordset.Edit
   ' Modificacio fet el dia 2/6/22 per en miralles, ara bloquejo per si hi ha refinplacsa
   'If data1.Recordset!proximaseccio = "T" Then
   '   enabled_campscontrolcodiinplacsa False
   '     Else: enabled_campscontrolcodiinplacsa True
   'End If
   If Text32(5) <> "" Then
      enabled_campscontrolcodiinplacsa False
        Else: enabled_campscontrolcodiinplacsa True
   End If
   DoEvents
   If muntadora Then
        desactivar_area_datos
   End If
   possar_direccio_envio True
'    If Text2.Enabled Then
'       dataactivacio.SetFocus
'     Else: If imp1.Visible And imp1.Enabled Then Text71.SetFocus
'    End If
   End If
   
End Sub
Sub enabled_campscontrolcodiinplacsa(venabled As Boolean)
  Dim camps As String
  Dim nomcamp As String
  Dim nomcontrol As String
    ' he tret #Ltipusadhesiu
  camps = "#Etubolam#Eampleesq#Eplegatesq#Esolapa#Eespessor#Emicropex#Eoberturaex#Ematerialex#Inumtreball#Lampleutil#Lsimulteneitatlam#Rmigelaborat#Ramplereb#Rsimulteneitatreb"
  camps = camps + "#Smigelaboratsol#Samplesol#Sampleplegsol#Slongitudsol#Ssolapasol#Sfuellebasesol#Sfuellebocasol#Stroquel#Sansa#Scinta#"
  nomcamp = proximcamp(camps)
  While nomcamp <> ""
    nomcamp = Mid(nomcamp, 2)
    nomcontrol = ""
    nomcontrol = buscarelcontrolambelnom(nomcamp)
    If nomcontrol <> "" Then
       On Error Resume Next
        Controls(nomcontrol).BackColor = IIf(venabled, &HFFFFFF, &HFF00FF)
        Controls(nomcontrol).Enabled = venabled
       On Error GoTo 0
    End If
    nomcamp = proximcamp(camps)
  Wend
  
End Sub
Function buscarelcontrolambelnom(nomcamp As String) As String
   Dim objecte As Object
   On Error Resume Next
   For Each objecte In formcomandes
      If objecte.DataField <> nomcamp Then
         nomcamp = nomcamp
         Else
          buscarelcontrolambelnom = objecte.Name: GoTo fi
      End If
   Next
fi:
End Function
Sub guardar_controlcanviscomanda()
   Set rstcontrolcanvis = Nothing
   Set rstcontrolcanvis_extres = Nothing
   Set dbcontrolcanvis = Nothing
   If existeix("c:\temp\~canviscomanda.mdb") Then Kill "c:\temp\~canviscomanda.mdb"
   DBEngine.CreateDatabase "c:\temp\~canviscomanda.mdb", dbLangGeneral, DatabaseTypeEnum.dbVersion30
   Set dbcontrolcanvis = OpenDatabase("c:\temp\~canviscomanda.mdb")
   Data1.Database.Execute "select * into comandes IN 'c:\temp\~canviscomanda.mdb' from comandes where comanda=" + atrim(Data1.Recordset!comanda)
   Data1.Database.Execute "select * into comandes_extres IN 'c:\temp\~canviscomanda.mdb' from comandes_extres where comanda=" + atrim(Data1.Recordset!comanda)
   Set rstcontrolcanvis = dbcontrolcanvis.OpenRecordset("select * from comandes where comanda=" + atrim(Data1.Recordset!comanda))
   Set rstcontrolcanvis_extres = dbcontrolcanvis.OpenRecordset("select * from comandes_extres where comanda=" + atrim(Data1.Recordset!comanda))
End Sub
Sub carregar_controlscampsalicia(Optional azero As Boolean)
'carrego els valors dels camps actuals per controlar l'avis a l'alicia
    i = 0
    While campscontrolalicia(i, 0) <> ""
      If Not azero Then
          campscontrolalicia(i, 1) = atrim(Data1.Recordset(campscontrolalicia(i, 0)))
        Else:
           campscontrolalicia(i, 1) = ""
           modificar.Tag = ""
      End If
      i = i + 1
    Wend
  'fins aqui
End Sub
Sub recalcular_comandescomplexes()
  Dim comandaact As String
  If checkrecalcular <> 1 Then Exit Sub
  If cadbl(Data1.Recordset!linkcomanda1) = 0 Then Exit Sub
  comandaact = Text1
  frrecalcular.Visible = True
  ratoli "espera"
  frrecalcular.Top = (formcomandes.Height / 2) - (frrecalcular.Height / 2)
  frrecalcular.Left = (formcomandes.Width / 2) - (frrecalcular.Width / 2)
  If cadbl(vlink3) > 0 Then
     buscant = True
     queryorder = ""
     querywhere = "comanda=" + atrim(cadbl(vlink3))
     finalitzarbusqueda 1
     wait 1
     modificar_Click
     wait 1
     gravar_registre
     wait 1
  End If
  If cadbl(vlink2) > 0 Then
     buscant = True
     queryorder = ""
     querywhere = "comanda=" + atrim(cadbl(vlink2))
     finalitzarbusqueda 1
     wait 1
     modificar_Click
     wait 1
     gravar_registre
     wait 1
  End If
  If cadbl(vlink1) > 0 Then
     buscant = True
     queryorder = ""
     querywhere = "comanda=" + atrim(cadbl(vlink1))
     finalitzarbusqueda 1
     wait 1
     modificar_Click
     wait 1
     gravar_registre
     wait 1
  End If
  
  If cadbl(comandaact) > 0 Then
     buscant = True
     queryorder = ""
     querywhere = "comanda=" + atrim(cadbl(comandaact))
     finalitzarbusqueda 1
    
  End If
  
  ratoli "normal"
  frrecalcular.Visible = False
End Sub
Sub control_usuaris(activar As Boolean)
  Dim objecte As Object
  If llistadecampsvalids = "[tots]" Then Exit Sub
  For Each objecte In Me
      'MsgBox objecte.Name
      If TypeOf objecte Is MaskEdBox Or TypeOf objecte Is TextBox Or TypeOf objecte Is ComboBox Or TypeOf objecte Is CheckBox Then
        If activar Then
          If InStr(1, llistadecampsvalids, "[" + atrim(objecte.DataField) + "]") Then
            objecte.Enabled = True
              Else: objecte.Enabled = False
           End If
            Else: objecte.Enabled = True
        End If
      End If
   Next objecte
   If InStr(1, llistadecampsvalids, "PVP") = 0 Then
      Command9(3).Visible = False
       Else: MaskEdBox11.Enabled = True
   End If
   
End Sub
Sub desactivar_area_datos()
  Static objectes(300) As String
'On Error Resume Next
  If Text2.Enabled Then
   i = 0
   For Each objecte In Me
      'MsgBox objecte.Name
      If TypeOf objecte Is MaskEdBox Or TypeOf objecte Is TextBox Or TypeOf objecte Is ComboBox Or TypeOf objecte Is CheckBox Then
        If objecte.Enabled = True And (objecte.Name <> "Text71" And objecte.Name <> "Text70") Then
          objectes(i) = objecte.Name
          objecte.Enabled = False
          i = i + 1
        End If
      End If
      If i = 300 Then Exit For
   Next objecte
    Else
      
      For Each objecte In Me
        If TypeOf objecte Is MaskEdBox Or TypeOf objecte Is TextBox Or TypeOf objecte Is ComboBox Or TypeOf objecte Is CheckBox Then
          For i = 0 To 299
           If objecte.Name = objectes(i) Then Exit For
          Next i
          If objecte.Name = objectes(i) Then objecte.Enabled = True
        End If
      Next objecte
  End If
  'On Error GoTo 0
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub nomclient_Click()
If cadbl(Data1.Recordset!client) > 0 Then
  escriure_ini "General", "clienttmp", atrim(Data1.Recordset!client), fitxerini
  formclients.Show
  formclients.Frame2.Tag = atrim(Data1.Recordset!direnvio)
End If
  
End Sub

Private Sub nomextrussora_Click(Index As Integer)

End Sub

Private Sub noplanificable_Click()
 'LA ALICIA A DIA 19 DE SETEMBRE DE 2022 M'HA FET TREURE AQUEST BOTÓ PERQUÈ DIU QUE NO ES UTIL
   'EN RICARD L'HA APRETAT SENSE VOLER ES VEU... HI HA UN CORREU D'AIXÓ
   
   dbtmp.Execute "update comandes_extres set noplanificable=" + IIf(noplanificable.Value = 1, "True", "False") + "  where comanda=" + atrim(cadbl(Text1.Text))
End Sub

Private Sub pes1_LostFocus()
possarconsums
End Sub

Private Sub pes2_LostFocus()
possarconsums
End Sub

Private Sub primerproces_Click()
  On Error Resume Next
   If Screen.ActiveControl.Name = "primerproces" Then
     If err.Number = 0 Then Data1.Recordset!refilatd = primerproces.Value
   End If
End Sub
Sub possarprimerprocesalpc2()
  Dim numc As Double
  If Data1.Recordset!comanda = vlink2 Then Exit Sub
  
  numc = IIf(Data1.Recordset!comanda = vlink1, vlink3, vlink1)
  dbtmp.Execute "update comandes set refilatd=" + atrim(IIf(Data1.Recordset!refilatd = 1, 0, 1)) + " where comanda=" + atrim(numc)
  dbtmp.Execute "update comandes set refilatd=" + atrim(cadbl(Data1.Recordset!refilatd)) + " where comanda=" + atrim(Data1.Recordset!comanda)
End Sub

Private Sub sortir_Click()
  If Not Frame1(0).Enabled Then Exit Sub
Unload subbusqueda
On Error Resume Next
  If isloaded("formfirmes") Then Unload formfirmes
If cadbl(formcomandes.Left) > 0 Then
  escriure_ini "PosicioFormComandes", "Left", atrim(formcomandes.Left), "comandes.ini"
  escriure_ini "PosicioFormComandes", "Top", atrim(formcomandes.Top), "comandes.ini"
End If
formcomandes.Hide
 Unload formcomandes
 
 AppActivate "Menu de Comandes"
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_LostFocus()
 'If Not buscant And Data1.Recordset.EditMode > 0 Then
 '  Set rsttmp = dbtmp.OpenRecordset("select client from comandes where comanda=" + atrim(cadbl(Text1.Text)))
 '  If rsttmp.RecordCount > 0 Then MsgBox "Aquest codi de client ja existeix haurieu de canviar-lo."
 'End If
End Sub


Private Sub Text103_KeyPress(Index As Integer, KeyAscii As Integer)
 If Not buscant And Index = 3 Then KeyAscii = 0
End Sub

Private Sub Text103_LostFocus(Index As Integer)
  If Index = 3 Then
     possaretiquetasistemaimpibandes
  End If
  If Index = 5 Then
     If atrim(Text103(Index)) <> "C" And atrim(Text103(Index)) <> "P" Then
        MsgBox "Només valen els valors C o P en aquest camp", vbCritical, "Error"
        Text103(Index) = "C"
     End If
     dbtmp.Execute "update comandes_extres set tipusmaterialcanutureb='" + atrim(Text103(Index)) + "' where comanda=" + atrim(cadbl(Text1))
  End If
   
End Sub
Function vernis_reprint(vtreball As Double, vordre As Integer) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("SELECT Tintes.id_treball, Tintes.ordremodificacio, Tintes.color From Tintes WHERE (((Tintes.id_treball)=" + atrim(vtreball) + ") AND ((Tintes.ordremodificacio)=" + atrim(vordre * -1) + ") AND ((Tintes.color)<>''));")
  If Not rst.EOF Then vernis_reprint = atrim(rst!color)
  Set rst = Nothing
End Function
Sub possaretiquetasistemaimpibandes()
   Dim rstsisimp As Recordset
   'Dim dbclixes As Database
   Dim bandes As Byte
   Dim nordre As Integer
   Label1(162).Visible = False
   nordre = cadbl(Data1.Recordset!numordremodificacio)
   If nordre = 0 Then nordre = 1
   
   Set rstsisimp = dbclixes.OpenRecordset("select sistemadimpresio,bandes,reimpres from modificacions where id_treball=" + atrim(cadbl(Data1.Recordset!numtreball)) + " and ordre=" + atrim(nordre))
   If Not rstsisimp.EOF Then
     bandes = atrim(cadbl(rstsisimp!bandes))
     Label1(159) = atrim(rstsisimp!sistemadimpresio) + IIf(bandes > 0, " - Clixes a " + atrim(bandes) + " Bandes", "")
     Label1(159).Tag = atrim(rstsisimp!sistemadimpresio)
     If rstsisimp!reimpres Then
        Label1(162).Visible = True
        Label1(162) = "ATENCIÓ IMPRESIÓ AMB REPRINT " + vernis_reprint(cadbl(Data1.Recordset!numtreball), nordre)
        Label1(162).ZOrder 1
         Else: Label1(162).Visible = False
     End If
       Else: Label1(159) = "": Label1(159).Tag = ""
   End If
   possar_color_tipusimpresio Label1(159).Tag
   Set rstsisimp = Nothing
   Label1(162).BackStyle = 1
   Label1(162).ZOrder 1
End Sub

Private Sub Text109_Click()
  Dim ruta As String
  Dim nomfitxer
  Text109.SetFocus
  nomfitxer = ActiveControl.Text
  If cadbl(Mid(nomfitxer, 1, 6)) = 0 Then nomfitxer = numcarpetaclient + " " + Trim(nomfitxer)
  ruta = ruta_relativa_docs + "\" + nomfitxer ' + Chr$(34)
  If Not existeix(ruta) Then ruta = ruta + "x"
  If existeix(ruta) Then
     obrir_document ruta
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + ruta, vbCritical, "Error"
  End If
  
  
  'obrir_document r + Chr$(34) + ruta_relativa_docs + "\" + ActiveControl.Text + Chr$(34)
End Sub

Private Sub Text11_LostFocus()
 carregar_lookups
End Sub

Private Sub Text110_Change()

End Sub

Private Sub Text111_Click()
  Dim ruta As String
  Dim nomfitxer
 Text111.SetFocus
 nomfitxer = ActiveControl.Text
  If cadbl(Mid(nomfitxer, 1, 6)) = 0 Then nomfitxer = numcarpetaclient + " " + Trim(nomfitxer)
  ruta = ruta_relativa_docs + "\" + nomfitxer ' + Chr$(34)
  If Not existeix(ruta) Then ruta = ruta + "x"
  If existeix(ruta) Then
     obrir_document ruta
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + ruta, vbCritical, "Error"
  End If

'obrir_document r + Chr$(34) + ruta_relativa_docs + "\" + ActiveControl.Text + Chr$(34)
End Sub

Private Sub Text120_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Soldadora", "select * from maquines where maquina='S' order by codi", Text120, nomsoldadora(1)
End Sub

Private Sub Text120_LostFocus()
  possar_lookup_manuals
End Sub

Private Sub Text127_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
     triaralgu "Triar Mesura Espessor", "mesureslineals", Text128, Text127, , 1
  End If
End Sub

Private Sub Text100_LostFocus()
possar_desc_lot Text100.Text, desclot1(1)
End Sub

Private Sub Text129_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
     triaralgu "Triar Soldadura", "tipussoldadura", Text129, Text130, , 1
  End If
End Sub

Private Sub Text133_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then r = "triaraccessoris": triaralgu "Triar Cinta", "select * from accessoris where Tipus_TNC='C'", Text133, Label1(177), , 1
End Sub

Private Sub Text133_LostFocus()
possar_lookup_manuals
End Sub

Private Sub Text134_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then r = "triaraccessoris": triaralgu "Triar Ansa", "select * from accessoris where Tipus_TNC='N'", Text134, ansa(0), , 1
End Sub

Private Sub Text134_LostFocus()
possar_lookup_manuals
End Sub

Private Sub Text135_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then r = "triaraccessoris": triaralgu "Triar Troquel", "select * from accessoris where Tipus_TNC='T'", Text135, truquel(0), , 1
End Sub

Private Sub Text135_LostFocus()
possar_lookup_manuals
End Sub

Private Sub Text138_Change()

End Sub

Private Sub Text15_Change()

End Sub

Private Sub Text16_Click()
comprovarsijahihaalgunalbaraentrataexpedicions
End Sub

Private Sub Text2_GotFocus()
 re = Text2
If llocform <> 0 Then
 llocform = 0
 formscrooll.SetValues formscrooll.Values.HorzValue, taulapos(llocform)
End If

End Sub

Private Sub Text37_Change()

End Sub

Private Sub Text21_LostFocus()
  ' If label1(162).Visible Then
  '     Text21.Text = atrim(cadbl(Text21.Text) + 1000)
  '     MsgBox "Com que es una comanda amb REPRINT s'ha afegit 1000 metres més del que has posat a la comanda.", vbInformation, "Atenció"
  ' End If
  calcular_micres_soldadores
End Sub

Private Sub Text24_LostFocus()
possar_lookup_manuals
End Sub

Private Sub Text25_LostFocus()
  possar_lookup_manuals
  calculcanvimaterial
End Sub

Private Sub Text26_LostFocus()
possar_lookup_manuals
End Sub

Function calcular_pes1000kg(Optional ByRef comanda As String, Optional codimat As Double)
  Dim rstc As Recordset
  Dim espessor As Double
  Dim kgr As Double
  Dim ample As Double
  Dim tuboolamina As String
  Dim vmicres As String
  Dim mmicres As String
  Dim vample As String
  Dim vsolapa As String
  Dim rstpes As Recordset
  Dim rstpes2 As Recordset
  'per saber l'espessor
  If cadbl(comanda) > 0 Then
     Set rstpes = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(comanda)))
     vmicres = atrim(rstpes!espessor)
     vample = atrim(rstpes!ampleesq)
     vsolapa = atrim(rstpes!solapa)
     tuboolamina = atrim(rstpes!tubolam)
     If cadbl(codimat) = 0 And Not rstpes.EOF Then codimat = rstpes!materialex
     Set rstpes2 = dbtmp.OpenRecordset("select grmcm3 from materials where codi=" + atrim(codimat))
     If Not rstpes2.EOF Then vgrmcm3 = rstpes2!grmcm3
     
     Set rstpes2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rstpes!mesuraesp)))
     If Not rstpes2.EOF Then mmicres = atrim(rstpes2!descripcio)
     Set rstpes2 = Nothing
     Set rstpes = Nothing
      Else
       comanda = atrim(Text1.Text)
       mmicres = atrim(Text22.Text)
       vmicres = atrim(Text21.Text)
       vgrmcm3 = atrim(grmcm3.Text)
       vample = atrim(Text18)
       vsolapa = atrim(Text20)
       tuboolamina = Combo8.Text
    End If
  
    If InStr(1, UCase(mmicres), "MICRES") Then
        espessor = cadbl(vmicres)
      Else
         If InStr(1, UCase(mmicres), "GALG") Then
          If (tuboolamina = "L") Then
               espessor = cadbl(vmicres) / 2
             Else: espessor = cadbl(vmicres) / 4
          End If
             Else
               espessor = 0
         End If
    End If
    'persaber els grams mt2
    kgr = (cadbl(vgrmcm3) / 0.000001) * (cadbl(espessor) * 0.000001)
    kgr = kgr / 1000
    If InStr(1, UCase(mmicres), "GR/MT2") Then kgr = cadbl(vmicres) / 1000
    
    'per saber l'ample si es tubo o lamina
    ample = cadbl(vample)
    If (tuboolamina <> "L") Then ample = cadbl(vample) * 2 + cadbl(vsolapa)
    ample = ample / 100
    calcular_pes1000kg = (kgr * ample * 1000)
    
End Function



Private Sub Text29_Change()
  If Data1.Recordset.EditMode > 0 And Not buscant And Not duplicant Then
     Text31.Text = "1"
     Text30.Text = "MTRS."
     Text32(15).Visible = True
     Text32(15).Top = Text29.Top + Text29.Height + 20
     Text32(15).Left = Text29.Left
     Text32(15).Height = Text29.Height
     calcular_pesmtr2imetresrebipesreb
     calcular_pes_metres_rebobinadora
     Text32(15) = rebpes + "KG - " + rebmetres + " MTRS"
   End If
End Sub

Private Sub Text29_GotFocus()
   Text29_Change
   If InStr(1, UCase(nomclient), "VAN DAMME") > 0 Then ensenyar_etiqueta_vandamme True
End Sub
Sub ensenyar_etiqueta_vandamme(vensenyar As Boolean)
  Dim v As String
  If vensenyar Then
    Text32(12).Font = "Courier New"
    Text32(12).FontBold = True
    v = "   ******  % QUANTITATS VAN DAMME   *****" + vbNewLine + vbNewLine
    v = v + "     KG  COMANDA [STD] ['ACTION'] [PECES]" + vbNewLine
    v = v + "   --------------------------------------" + vbNewLine
    v = v + "   0 - 300      - 3 %    + 1 %     + 3 %" + vbNewLine
    v = v + "   300 - 500    - 3 %    + 1 %     + 3 %" + vbNewLine
    v = v + "   500 - 750    - 4 %    + 0 %     + 2 %" + vbNewLine
    v = v + "   751 - a X    - 4 %    + 0 %     + 1 %"
    Text32(12).Text = v
    'Text32(12).Top = 5400
    Text32(12).Top = Text29.Top + Text29.Height + 60
    Text32(12).Width = 6000
    Text32(12).Height = 2430
    Text32(12).Left = 500
    Text32(12).Tag = "20"
    Text32(12).Visible = True
  End If
  If Not vensenyar Then
    Text32(12).Font = "Arial"
    Text32(12).Tag = "0"
    Text32(12).Visible = False
  End If
End Sub
Private Sub Text29_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 114 Then
     Text29 = passarkilosametres(cadbl(Text29))
  End If
End Sub
Sub p100ossarpes1000metresalescomplexes(numc As Double)
    Dim pes1000 As Double
    Dim rstc As Recordset
    Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(numc)))
    If rstc.EOF Then Exit Sub
    If rstc!linkcomanda1 > 0 Then
     pes1000 = calcular_pes1000kg(rstc!linkcomanda1)
     dbtmp.Execute "update comandes set pes1000mtrs=" + passaradecimalpunt(Format(pes1000, "#00.00")) + "  where comanda=" + atrim(cadbl(rstc!linkcomanda1))
    End If
    If rstc!linkcomanda2 > 0 Then
     pes1000 = calcular_pes1000kg(rstc!linkcomanda2)
     dbtmp.Execute "update comandes set pes1000mtrs=" + passaradecimalpunt(Format(pes1000, "#00.00")) + "  where comanda=" + atrim(cadbl(rstc!linkcomanda2))
    End If
    Set rstc = Nothing
End Sub
Function passarkilosametres(kg As Double) As String
   Text33 = calcular_pes1000kg
   p100ossarpes1000metresalescomplexes cadbl(Text1)
   calcular_pesmtr2imetresrebipesreb
   passarkilosametres = 0
   'If (cadbl(Text30.Tag) * (cadbl(Text18) / 100)) > 0 Then passarkilosametres = Format(kg / (cadbl(Text30.Tag) * (cadbl(Text18) / 100)), "#.##0")
   If (cadbl(Text30.Tag) * ((cadbl(Text98) * cadbl(Combo3)) / 100)) > 0 Then passarkilosametres = Format(Redondejar(kg / (cadbl(Text30.Tag) * ((cadbl(Text98) * cadbl(Combo3)) / 100)), 0), "#.##0")
   If cadbl(Text98) = 0 Then passarkilosametres = Redondejar((kg / cadbl(Text33)) * 1000, 0)
End Function

Private Sub Text29_LostFocus()
  Dim vquantdemanada As Double
  ensenyar_etiqueta_vandamme False
  Text32(15).Visible = False
  If cadbl(Text29) <> cadbl(Data1.Recordset!cantitatex) Then
    If cadbl(MaskEdBox6) > 0 Then MsgBox "Has canviat la quantitat de comanda, comprova si has de canviar el valor de Quantitat demanada.", vbInformation, "Atenció"
  End If
  If MaskEdBox6.Tag = "obligatquantitatdemanada" And cadbl(MaskEdBox6) = 0 Then
     While vquantdemanada = 0
       vquantdemanada = cadbl(InputBox("Entra la quantitat demanada pel client.", "Quantitat demanada"))
     Wend
     MaskEdBox6.SetFocus
     MaskEdBox6 = atrim(vquantdemanada)
  End If
  calcular_micres_soldadores
  ' If label1(162).Visible Then
  '     Text29.Text = atrim(cadbl(Text29.Text) + 1000)
  '     MsgBox "Com que es una comanda amb REPRINT s'ha afegit 1000 metres més del que has posat a la comanda.", vbInformation, "Atenció"
  ' End If
End Sub

Private Sub Text32_Click(Index As Integer)
  If Index = 12 Then Text32(12).Visible = False
  If Index = 15 Then Text32(15).Visible = False
End Sub

Private Sub Text32_GotFocus(Index As Integer)
If Index = 1 Then Text32(Index).Locked = True
If Index = 2 Then Text32(2).Height = 1500: Text32(2).Width = 3525: Text32(2).Top = 1900
If Index = 10 Then Command10.Visible = False: Text32(10).Height = 1500: Text32(10).Width = 3000: Text32(10).Top = 1900
If Index = 14 Then Text32(Index).MaxLength = 12
End Sub

Private Sub Text32_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then triarmesuraquantitatdesitjada
End Sub

Private Sub Text32_LostFocus(Index As Integer)
If Index = 2 Then Text32(2).Width = 1380: Text32(2).Height = 285: Text32(2).Top = 3120
If Index = 10 Then Command10.Visible = True: Text32(10).Width = 2190: Text32(10).Height = 285: Text32(10).Top = 3120
If Index = 11 Then If Text32(11) = "" Then Text32(11) = 0
End Sub

Private Sub Text35_Click()
 
End Sub

Private Sub Text33_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub Text33_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text4_LostFocus()
  'calcular_dies_entrega
End Sub

Private Sub Text42_Click()
 ' obrir_word r + Chr$(34) + ruta_relativa_docs + "\" + ActiveControl.Text + Chr$(34)
  obrir_document r + Chr$(34) + ruta_relativa_docs + "\" + ActiveControl.Text + Chr$(34)
  'r = "cmd /c "
  'If existeix("c:\windows\command\start.exe") Then r = "start "
 'r = Shell(r + Chr$(34) + ruta_relativa_docs + "\" + ActiveControl.Text + Chr$(34), vbMinimizedFocus)
End Sub

Private Sub Text51_Change()

End Sub

Private Sub Text44_GotFocus()
   comprovar_sihihabobinesetiquetades
End Sub
Function comprovar_sihihabobinesetiquetades() As Boolean
  If Data1.Recordset!proximaseccio <> "E" Then
    If jahihabobinesetiquetades(cadbl(Text1)) Then
      MsgBox "Aquesta comanda ja té bobines etiquetades a rebobinadora si fas un canvi de referència has de pendre les mesures oportunes.", vbCritical, "Atenció"
      comprovar_sihihabobinesetiquetades = True
    End If
  End If
End Function
Function jahihabobinesetiquetades(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbbaixes.OpenRecordset("SELECT rebobinadores.comanda, bobinesreb.numerodebobina FROM rebobinadores LEFT JOIN bobinesreb ON rebobinadores.Id = bobinesreb.Id WHERE (((rebobinadores.comanda)=" + atrim(vnumc) + "));", , ReadOnly)
   If Not rst.EOF Then jahihabobinesetiquetades = True
   Set rst = Nothing
End Function


Private Sub Text44_LostFocus()
  r = Text44.Text
  If InStr(1, Text32(1), r) = 0 And r <> "" Then
      'If atrim(Text32(1)) <> "" Then Text32(1) = Text32(1) + " | "
      Text32(1) = Text32(1) + r + " | "
      
  End If
End Sub

Private Sub Text5_GotFocus()
If Len(dataentrega2.Text) > 3 Then
    Text5.Enabled = False
      Else: Text5.Enabled = True
End If
End Sub

Private Sub Text57_Change()

End Sub

Private Sub Text59_Change()

End Sub

Private Sub Text6_GotFocus()
     comprovarsijahihaalgunalbaraentrataexpedicions
     comprovarsitereprintimicroperforats
     possarinformacioentradapreuPVP
End Sub
Sub possarinformacioentradapreuPVP()
   Text32(12) = "Si el valor es -1 no es cobrarà al Client."
   Text32(12).Visible = True
   Text32(12).Left = cap.Left + areadatos.Left + Text6.Left
   Text32(12).Top = formscrooll.Top + Text6.Top + Text6.Height + 60
   Text32(12).Tag = "7"
End Sub

Sub comprovarsitereprintimicroperforats()
     Dim vmsg As String
     If Label1(162).Visible Then
       vmsg = "Reprint"
     End If
     If Data1.Recordset.EOF Then Exit Sub
     If atrim(Data1.Recordset!microperforat) = "S" Or atrim(Data1.Recordset!rebmacroperforat) = "S" Or atrim(Data1.Recordset!microperforatsol) = "S" Then
        If vmsg <> "" Then vmsg = vmsg + " i "
        vmsg = vmsg + " Microperforat"
     End If
     If vmsg <> "" Then MsgBox "Atenció aquesta comanda te " + vmsg + " tenir-ho amb compte amb el preu", vbCritical, "Atenció"
End Sub

Private Sub Text6_LostFocus()
    If cadbl(Text6) > 999 Then MsgBox "Aquest valor es molt gran pel preu, assegura que sigui correcte.", vbCritical, "ATENCIÓ"
    Text32(12).Visible = False
End Sub

Private Sub Text63_GotFocus()
   'Text63 = calcular_tinters
End Sub
Function calcular_tinters() As Byte
  Dim tintersa As Byte
  Dim tintersb As Byte
  If atrim(Text40) <> "" Then tintersa = tintersa + 1
  If atrim(Text46) <> "" Then tintersa = tintersa + 1
  If atrim(Text47) <> "" Then tintersa = tintersa + 1
  If atrim(Text48) <> "" Then tintersa = tintersa + 1
  If atrim(Text49) <> "" Then tintersa = tintersa + 1
  If atrim(Text50) <> "" Then tintersa = tintersa + 1
  If atrim(Text141) <> "" Then tintersa = tintersa + 1
  If atrim(Text140) <> "" Then tintersa = tintersa + 1
  
 ' If atrim(Text57) <> "" Then tintersb = tintersb + 1
 ' If atrim(Text58) <> "" Then tintersb = tintersb + 1
 ' If atrim(Text59) <> "" Then tintersb = tintersb + 1
 ' If atrim(Text60) <> "" Then tintersb = tintersb + 1
  'If atrim(Text61) <> "" Then tintersb = tintersb + 1
  'If atrim(Text62) <> "" Then tintersb = tintersb + 1
  'If atrim(Text137) <> "" Then tintersb = tintersb + 1
  'If atrim(Text136) <> "" Then tintersb = tintersb + 1
  
  calcular_tinters = tintersa '+ tintersb
  
End Function

Private Sub Text63_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub Text63_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text64_Change()
If Text64 = "R" Then cimpressio.Text = "Repetida"
If Text64 = "N" Then cimpressio.Text = "Nova"
If Text64 = "M" Then cimpressio.Text = "Modificada"
If Text64 = "F" Then cimpressio.Text = "Falta Autoritzar"
If atrim(Text64) = "" Then cimpressio.Text = ""

End Sub

Private Sub Text65_Change()
If Text65 = "T" Then ctipusimp.Text = "Transparencia"
If Text65 = "N" Then ctipusimp.Text = "Normal"
If atrim(Text65) = "" Then ctipusimp.Text = ""

End Sub

Private Sub Text68_KeyPress(KeyAscii As Integer)
  If Chr$(KeyAscii) <> "N" And Chr$(KeyAscii) <> "C" And Chr$(KeyAscii) <> "1" And Chr$(KeyAscii) <> "2" Then
     KeyAscii = Asc("N")
   Else: Text68.Text = ""
  End If
End Sub

Private Sub Text68_LostFocus()
  If Len(Text68.Text) = 0 Then Text68.Text = "N"
End Sub

Private Sub Text73_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Impressora", "select * from maquines where maquina='I' order by codi", Text73, nomimpressora(1)
End Sub

Private Sub Text75_Change()

End Sub

Private Sub Text77_Change(Index As Integer)
  If cadbl(text77(28)) > 0 Then
       text77(28).BackColor = QBColor(13)
        Else: text77(28).BackColor = QBColor(15)
  End If
  ' If Index = 10 Then
  '    ' si es variable i es null el % l'etiqueta serà NO ASSIGNAT si es 0 serà NO TÉ
  '     If text77(10).Tag = "V" And text77(10) = "" Then
  '        Label1(148).Caption = "NO ASSIGNAT"
  '          Else: Label1(148).Caption = "NO TÉ"
  '     End If
  ' End If
End Sub
Sub comprovar_num_pack()

End Sub
Private Sub Text77_GotFocus(Index As Integer)
  If Index = 28 Then If atrim(Text41) = "" Then MsgBox "Primer has d'entrar el numero de pressupost abans del pack.", vbCritical, "Error": Text41.SetFocus
  If Index = 10 Then
   If text77(10).Tag = "F" Then
      text77(10).Locked = True
      SendKeys "{TAB}"
        Else: text77(10).Locked = False
   End If
  End If
  If Index = 27 Then bxl.Visible = True
  
End Sub

Private Sub text77_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If Index = 16 Then
    If KeyCode = 113 Then triaralgu "Triar tipus etiquetes", "tipusetiquetes", text77(15), text77(16)
  End If

If Index = 13 Then
    If KeyCode = 113 Then triaralgu "Triar embolicades", "bobinesembolicades", text77(14), text77(13)
  End If
If KeyCode = 46 And Index = 16 Then text77(15) = 0: text77(16) = ""
If KeyCode = 46 And Index = 13 Then text77(14) = 0: text77(13) = ""
 
 
 
End Sub

Private Sub text77_LostFocus(Index As Integer)
  If Index = 27 Then If Screen.ActiveControl.Name <> "bxl" Then bxl.Visible = False
  If Index = 28 Then text77(28) = IIf(cadbl(text77(28)) > 0, text77(28), ""): guardar_pack
End Sub
Sub guardar_pack()
     dbtmp.Execute "update comandes_extres set numpack='" + text77(28) + "'  where comanda=" + atrim(Data1.Recordset!comanda)
End Sub
Private Sub text77_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rstb As Recordset
  If Shift = 2 And Index = 17 And atrim(text77(17)) <> "" Then
      If atrim(text77(17)) = "R" Then
        Set rstb = dbbaixes.OpenRecordset("select * from rebobinadorestot where comanda=" + atrim(Text1))
        If rstb.EOF Then
          MsgBox "No hi ha la seccio de Baixa de REBOBINADORA donada d'alta per aquesta comanda. Primer donala d'alta abans de marcar-la coma a acabada."
           Else: dbbaixes.Execute "update  rebobinadorestot set acavada=1 where comanda=" + atrim(Text1)
        End If
      End If
      If atrim(text77(17)) = "L" Then
        Set rstb = dbbaixes.OpenRecordset("select * from laminadorestot where comanda=" + atrim(Text1))
        If rstb.EOF Then
          MsgBox "No hi ha la seccio de Baixa de LAMINADORA donada d'alta per aquesta comanda. Primer donala d'alta abans de marcar-la coma a acabada."
           Else
            dbbaixes.Execute "update  laminadorestot set acavada=1 where comanda=" + atrim(Text1)
        End If
      End If
      If atrim(text77(17)) = "I" Then
         Set rstb = dbbaixes.OpenRecordset("select * from impressorestot where comanda=" + atrim(Text1))
         If rstb.EOF Then
          MsgBox "No hi ha la seccio de Baixa de IMPRESORES donada d'alta per aquesta comanda. Primer donala d'alta abans de marcar-la coma a acabada."
           Else
            dbbaixes.Execute "update  impressorestot set acavada=1 where comanda=" + atrim(Text1)
         End If
      End If
  End If

End Sub

Private Sub Text78_Click()
  Dim ruta As String
  Dim nomfitxer As String
  
  Text78.SetFocus
  nomfitxer = ActiveControl.Text
  If cadbl(Mid(nomfitxer, 1, 6)) = 0 Then nomfitxer = numcarpetaclient + " " + Trim(nomfitxer)
  ruta = ruta_relativa_docs + "\" + nomfitxer ' + Chr$(34)
  If existeix(ruta) Then
     obrir_document ruta
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + ruta, vbCritical, "Error"
  End If
  
End Sub
Function numcarpetaclient() As String
    numcarpetaclient = Mid(ruta_relativa_client, 1, 6)

End Function
Private Sub Text79_Click()
  Dim ruta As String
  Dim nomfitxer As String
 Text79.SetFocus
 '+ numcarpetaclient
   nomfitxer = ActiveControl.Text
  If cadbl(Mid(nomfitxer, 1, 6)) = 0 Then nomfitxer = numcarpetaclient + " " + Trim(nomfitxer)
  ruta = ruta_relativa_docs + "\" + nomfitxer ' + Chr$(34)
  If existeix(ruta) Then
     obrir_document ruta
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + ruta, vbCritical, "Error"
  End If

 'obrir_document r + Chr$(34) + ruta_relativa_docs + "\" + Text79.Text + Chr$(34)

End Sub

Private Sub Text80_LostFocus()
  possar_desc_lot Text80.Text, desclot1(0)
End Sub
Sub possar_desc_lot(numlot As String, desclotx As Control)
  Dim desctmp As String
  Dim rsttmp2 As Recordset
  Dim rst As Recordset
  Dim rstd1 As Recordset
  Dim lot1 As String
  Dim lot2 As String
  If cadbl(numlot) < 1 Then Exit Sub
  Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(Data1.Recordset!comanda), , ReadOnly)
  If rst.EOF Then Exit Sub
  If rst!refilatd <> 1 Then
       If cadbl(numlot) <> cadbl(Data1.Recordset!comanda) Then
        Set rstd1 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(numlot)), , ReadOnly) 'rst!lotmatdesb1
        If rstd1.EOF Then Exit Sub
        lot1 = "(" + generardadescomanda(cadbl(rstd1!lotmatdesb1))
        lot2 = generardadescomanda(cadbl(rstd1!lotmatdesb2)) + ")"
        desctmp = lot1 + " + " + lot2
          Else
            desctmp = "(" + generardadescomanda(cadbl(numlot)) + ")"
       End If
      Else
        desctmp = "(" + generardadescomanda(cadbl(numlot)) + ")"
  End If
  desclotx = desctmp
  Set rsttmp2 = Nothing
  Set rsttmp = Nothing
  Set rstd1 = Nothing
  Set rst = Nothing
End Sub
Sub possar_desc_lot2(numlot As String, desclotx As Control)
  Dim desctmp As String
  Dim rsttmp2 As Recordset
  desctmp = ""
  desclotx = desctmp
  If cadbl(numlot) < 1 Then Exit Sub
  Set rsttmp = dbtmp.OpenRecordset("select materialex,colorex,espessor,mesuraesp from comandes where comanda=" + atrim(cadbl(numlot)), , ReadOnly)
  If Not rsttmp.EOF Then
     Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from materials where codi=" + atrim(cadbl(rsttmp!materialex)), , ReadOnly)
     If Not rsttmp2.EOF Then desctmp = atrim(rsttmp2!descripcio) + " - "
     Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from colorants where codi=" + atrim(cadbl(rsttmp!colorex)), , ReadOnly)
     If Not rsttmp2.EOF Then desctmp = desctmp + rsttmp2!descripcio + "  "
     Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from mesures where codi=" + atrim(cadbl(rsttmp!mesuraesp)), , ReadOnly)
     If Not rsttmp2.EOF Then desctmp = desctmp + rsttmp2!descripcio
  End If
  desclotx = desctmp
  Set rsttmp2 = Nothing
  Set rsttmp = Nothing
End Sub

Private Sub Text81_LostFocus()
  possar_desc_lot Text81.Text, desclot2(1)
End Sub

Private Sub Text82_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then triaralgu "Triar Laminadora", "select * from maquines where maquina='L' order by codi", Text82, nomlaminadora(0)
End Sub

Private Sub Text87_LostFocus()
   calcular_ample_lam
   If Text87 <> Text98 Then
        If MsgBox("El valor d'ample de rebobinat es diferent a l'ample útil." + Chr(10) + "Vols canviar el de rebobinat?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
        Text98.Text = passaradecimal(Text87.Text)
   End If
End Sub

Sub calcular_ample_lam()
   Text89 = cadbl(Text87) * cadbl(Combo2)
End Sub

Private Sub Text89_GotFocus()
calcular_ample_lam
End Sub

Private Sub Text9_Change()

End Sub


Private Sub Text91_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then
     triaralgu "Triar Camisa", "camises", Text91, Text91, "cm", 1
  End If
End Sub

Private Sub Text91_LostFocus()
  comprovarcamisadelcomplexe
End Sub
Sub comprovarcamisadelcomplexe()
  Dim rstco As Recordset
  If estattaula.Caption = "Buscant..." Then Exit Sub
  If atrim(Data1.Recordset!producte) = "PC" Or atrim(Data1.Recordset!producte) = "PC2" Then Exit Sub
  Set rstco = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(Data1.Recordset!linkcomanda2))
  If rstco.EOF Then Exit Sub
  If cadbl(rstco!camisa) <> cadbl(Text91) Then
    If MsgBox("La camisa del segon procés de laminacio es diferent... " + Chr(10) + " Vols copiar els valors de ampleutil,simulteneitat,amplelaminar i camisa al segon procés?", vbInformation + vbYesNo, "Atenció") = vbNo Then Exit Sub
  End If
  rstco.Edit
  rstco!ampleutil = cadbl(Text87)
  rstco!simulteneitatlam = cadbl(Combo2)
  rstco!amplelaminar = cadbl(Text89)
  rstco!camisa = cadbl(Text91)
  rstco.Update
  Set rstco = Nothing
End Sub

Private Sub Text97_Click()
  Dim ruta As String
  Dim nomfitxer As String
   Text97.SetFocus
   nomfitxer = ActiveControl.Text
  If cadbl(Mid(nomfitxer, 1, 6)) = 0 Then nomfitxer = numcarpetaclient + " " + Trim(nomfitxer)
  ruta = ruta_relativa_docs + "\" + nomfitxer ' + Chr$(34)
  If Not existeix(ruta) Then ruta = ruta + "x"
  If existeix(ruta) Then
     obrir_document ruta
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + ruta, vbCritical, "Error"
  End If
   
  'obrir_document r + Chr$(34) + ruta_relativa_docs + "\" + ActiveControl.Text + Chr$(34)
End Sub

Private Sub Text98_LostFocus()
   If lam1.Visible And cadbl(Text87) <> cadbl(Text98) Then
       If MsgBox("El camp de Ample util de laminadora es diferent que el de rebobinadora." + vbNewLine + " La camisa també es canviarà a l'ample + 1cm." + vbNewLine + "VOLS CANVIAR-LO AUTOMÀTICAMENT?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
       Text87 = Text98  'CANVIO AMPLEUTIL
       Text89 = Text98   'CANVIO AMPLELAMINAR
       Text91 = atrim(cadbl(Text87) + 1) 'CANVIO CAMISA
   End If
End Sub

Private Sub Text99_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then triaralgu "Triar Rebobinadora", "select * from maquines where maquina='R' order by codi", Text99, nomrebobinadora(1)
End Sub

Private Sub Timer1_Timer()
Dim color As Double
'Me.Caption = Screen.ActiveControl.Name
controlarseccions

comprovarsihihacomandaxrobrirdeplanificacio
  estattaula.Caption = textestattaula(Data1.EditMode)
  If estattaula.ForeColor <> QBColor(0) Then
     estattaula.ForeColor = QBColor(0)
    Else: estattaula.ForeColor = QBColor(14)
  End If
  possarcoloraltrausuarigravant
 'amagar la etiqueta de informacio PVP
  If cadbl(Text32(12).Tag) = 0 Then
     Text32(12).Visible = False
      Else: Text32(12).Tag = cadbl(Text32(12).Tag) - 1
  End If
  
 'If buscant Then
 '  For Each objecte In Me
 '   If TypeOf objecte Is MaskEdBox Or TypeOf objecte Is TextBox Then
 '     color = objecte.BackColor
 '    If objecte.Tag = "9" And color <> QBColor(11) Then
 '      If Not TypeOf objecte Is CommandButton Then objecte.BackColor = QBColor(14)
 '    End If
 '   End If
 ' Next
 'End If
  'Me.Caption = Screen.ActiveControl.Container.Name
 
End Sub
Sub possarcoloraltrausuarigravant()
   Dim rst As Recordset
   If Data1.Recordset.EditMode > 0 Then
       'Set rst = dbtmp.OpenRecordset("select * from valorsgenerals")
       If hihaalgugravant Then
           'If atrim(rst!gravantnomordinador) <> "" And atrim(rst!gravantnomordinador) <> atrim(nomordinador) Then
              Command11.BackColor = &HC0C0FF
              Label1(33) = Command11.Tag
              Else: Command11.BackColor = &H8000000F
          ' End If
       End If
       Label1(33) = Command11.Tag 'posa el nom de l'usuari que està gravant el registre
       DoEvents
     Else
       If Command11.Tag <> "" Then hihaalgugravant
       Command11.BackColor = &H8000000F
       Label1(33) = Command11.Tag  'borra el nom de l'usuari que està gravant el registre
   End If
   Set rst = Nothing
End Sub

Sub recorregutregistres()
 Dim objecte As Object
 Dim weremodificacio As String
 queryorder = ""
 querywhere = ""
 If Text32(5) <> "" Then GoTo refinplacsa
 'On Error Resume Next
 For Each objecte In Me
    If TypeOf objecte Is MaskEdBox Or TypeOf objecte Is TextBox Or TypeOf objecte Is ComboBox Then
     If objecte.Tag = "9" Or objecte.Text <> "" Then
       If objecte.DataField <> "" Then ' Si Texto es igual "Hola".
         If objecte.Text <> "" Then
           evaluarcontingut objecte.DataField, objecte.Text, Data1.Recordset.Fields(objecte.DataField).Type
           objecte.Text = ""
         End If
      End If
     End If
    End If
 Next
 'excepcio del punt de risc
 If Label1(146).Tag = "9" Then
  If querywhere = "" Then
     querywhere = " puntrisc=" + atrim(cadbl(Label1(146)))
    Else
     querywhere = querywhere + " and " + " puntrisc=" + atrim(cadbl(Label1(146))) + " "
  End If
  Label1(146).Tag = ""
End If
posarnumtreballcorrectement
 If cadbl(Text103(3).Tag) <> 0 Then
   weremodificacio = " and (numordremodificacio=" + atrim(Text103(3).WhatsThisHelpID) + IIf(Text103(3).WhatsThisHelpID = 1, " or numordremodificacio=0", "") + ")"
   wereimpres = " and (producte in (SELECT productes.codi From productes WHERE (((InStr(1,[productes].[ruta],'I'))>0))))"

   If querywhere = "" Then
     querywhere = "( numtreball=" + atrim(cadbl(Text103(3).Tag)) + IIf(Text103(3).WhatsThisHelpID > 0, weremodificacio, "") + wereimpres + ")"
    Else
     querywhere = querywhere + " and " + "( numtreball=" + atrim(cadbl(Text103(3).Tag)) + IIf(Text103(3).WhatsThisHelpID > 0, weremodificacio, "") + wereimpres + ")"
   End If
   Text103(3).Tag = ""
   Text103(3).WhatsThisHelpID = 0
   'MsgBox querywhere
 End If
 Exit Sub
refinplacsa:
  querywhere = " comanda in (select comanda from comandes_extres where refinplacsa like '*" + atrim(Text32(5)) + "*')"

End Sub

Sub posarnumtreballcorrectement()
   Dim numt As Double
   Dim numm As Double
   Dim Camp As String
   Camp = atrim(Text103(3))
   If InStr(1, Camp, "/") > 0 Then
    numt = cadbl(Mid(Camp, 1, InStr(1, Camp, "/") - 1))
    numm = cadbl(Mid(Camp, InStr(1, Camp, "/") + 1))
     Else: numt = cadbl(Camp)
   End If
   If numt > 0 Then
      Text103(3).Tag = atrim(numt)
      Text103(3).WhatsThisHelpID = atrim(numm)
   End If
End Sub
Function evaluarcontingut(Camp As String, valor As String, tipusdato As Byte) As String
  Dim rest As String
  rest = ""
  evaluarcontingut = ""
  If triarordre(Camp, valor) Then Exit Function
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
 
  evaluarcontingut = Camp + rest
  If InStr(1, evaluarcontingut, "texteimpressio") <> 0 Or Camp = "marcailinia" Then
     evaluarcontingut = "(" + evaluarcontingut + " or texteimpressio" + rest + " or obsimp1" + rest + " or marcailinia" + rest + ")"
  End If
  If InStr(1, evaluarcontingut, "refclient") <> 0 Then
     'evaluarcontingut = "(" + evaluarcontingut + " or obsimp1" + rest + " or obspedido1" + rest + " or obsext1" + rest + " or obssol1" + rest + " or refclialt" + rest + ")"
     evaluarcontingut = "(" + evaluarcontingut + " or refclialt" + rest + ")"
  End If

  rest = evaluarcontingut
  
  If querywhere = "" Then
     querywhere = rest
    Else
     querywhere = querywhere + " and " + rest + " "
  End If
  
End Function

Function triarordre(Camp As String, valorord As String) As Boolean
  Dim ord As String
  triarordre = False
  If InStr(1, valorord, "<<") Then ord = Camp + " " + " ASC"
  If InStr(1, valorord, ">>") Then ord = Camp + " " + " DESC"
  If ord <> "" Then
      triarordre = True
    Else: Exit Function
  End If
  If queryorder = "" Then
     queryorder = ord
   Else: queryorder = queryorder + ", " + ord
  End If
  
End Function
Sub finalitzarbusqueda(Optional tipus As Byte)
 ratoli "espera"
 cimpressio.Clear
 cimpressio.AddItem "Falta Autoritzar"
 If cadbl(tipus) = 1 Then GoTo ficonsulta
 recorregutregistres
 possarvalordcamps
 If Data1.Recordset.EditMode > 0 Then
    Text103(3) = ""
    Text103(3).Tag = ""
    Text103(3).WhatsThisHelpID = 0
    Data1.Recordset.CancelUpdate
    enabled_campscontrolcodiinplacsa True
 End If
ficonsulta:
 
 buscant = False
 Text1.Enabled = False
 areadedatos False
 If queryorder <> "" Then
     queryorder = " Order By " + queryorder
    Else: queryorder = " order by comanda DESC  "
 End If
 If querywhere = "()" Then querywhere = ""
 If querywhere <> "" Then querywhere = " Where " + querywhere
 Data1.RecordSource = "select * from comandes " + querywhere + queryorder
 'dbtmp.Execute "insert into registreconsultescomandes (horainici,usuari,consultasql) values (now,'" + nomordinador + "','" + treure_apostruf(Data1.RecordSource) + "')"
 refrescar
 dbtmp.Execute "update registreconsultescomandes set horafi=now where consultasql='" + treure_apostruf(Data1.RecordSource) + "' and usuari='" + nomordinador + "' and horafi=null"
 ratoli "normal"
 'Unload subbusqueda
End Sub
Sub areadedatos(valor As Boolean)
   'areadatos.Enabled = valor
   For Each Control In formcomandes
       If TypeOf Control Is Frame Then
         If Control.Tag <> "100" Then Control.Enabled = valor
        ' MsgBox Control.Name
       End If
   Next
   
End Sub
Sub deixartotblanc()
 For Each objecte In Me
    If TypeOf objecte Is TextBox Or TypeOf objecte Is MaskEdBox Or TypeOf objecte Is ComboBox Then
      If objecte.DataField <> "" Then ' Si Texto es igual "Hola".
        objecte.Tag = ""
        objecte.Text = ""
        If TypeOf objecte Is TextBox Then
           objecte.MaxLength = 255
        End If
        If TypeOf objecte Is MaskEdBox Then
              On Error Resume Next
              'objecte.Format = ""
              objecte.MaxLength = 50
               On Error GoTo 0
        End If
           
     End If
    End If
Next

End Sub

Sub possar_color_tipusimpresio(tipus As String)
  Dim color As Long
  Dim colorclassic As Long
  Dim colorkodak As Long
  Dim coloroffset As Long
  tipusimpresio = tipus
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
   
   cap.BackColor = color
   ext.BackColor = color
   imp1.BackColor = color
   lam1.BackColor = color
   reb.BackColor = color
   sol.BackColor = color
   checkpassaraproduccio.BackColor = color
   noplanificable.BackColor = color
End Sub
Sub possar_boto_refinplacsavalida(vrefvalida As Boolean)
   If vrefvalida Then
         Command26(5).BackColor = &HC0FFC0
          Else: Command26(5).BackColor = &H8080FF
   End If
   Command26(5).Tag = IIf(vrefvalida, "T", "F")
End Sub
Sub carregarvalorscomandesextres()
   Dim rste As Recordset
   Dim rstc As Recordset
   'netejo camps
   Command26(4).Visible = False  'veure boto per veure el codi gtin a pantalla
   Command26(4).Tag = ""
   Text32(3) = ""
   Text32(3).Tag = ""
   Text32(3).BackColor = &H8000000F
   Command9(0).BackColor = Command8.BackColor
   
   If buscant Then GoTo fi
   Set rste = Data1.Database.OpenRecordset("select * from comandes_extres where comanda=" + atrim(cadbl(Text1)))
   If rste.EOF Or cadbl(Text1) = 0 Then Exit Sub
   velclientvolPVPimpostinclos = rste!PVPimpostinclos
   'posso el valor de la comanda d'on es va duplicar
   Label1(178) = ""
   If cadbl(rste!comandaduplicadade) > 0 Then Label1(178) = "Duplicada de: " + atrim(rste!comandaduplicadade)
   
   possar_boto_refinplacsavalida rste!refinplacsa_validada
   
   'lookup de codicomptable
   'label1(32).Caption = "Ref: " + atrim(rste!refinplacsa)
   Text103(5) = atrim(rste!tipusmaterialcanutureb)
   checkpassaraproduccio.Value = cadbl(rste!passaraimpresores)
   'Check1(1).Value = cadbl(rste!pararaexpedicions) 'valor del check parar a expedicions de la taula comandes_extres
   If checkpassaraproduccio.Value = 2 Then
      checkpassaraproduccio.Enabled = False
        Else: checkpassaraproduccio.Enabled = True
   End If
   If rste!clientvindraarevisarimpresio Then
         materialexacte(1).Value = 1
          Else: materialexacte(1).Value = 0
   End If
   If rste!comandaimpresa Then Command9(0).BackColor = &HC0FFC0
   Command9(0).Tag = IIf(rste!comandaimpresa, "Siimpresa", "NoImpresa")
   escullirvalorcombo Combo1(3), cadbl(rste!pararaexpedicions)
   Combo1(0) = atrim(rste!carametall)
   Text32(5) = atrim(rste!refinplacsa)
   Text32(9) = atrim(rste!desarrollclient)
   Text32(10) = atrim(rste!observacionsalbara)
   text77(28) = atrim(rste!numpack)
   Combo1(4) = atrim(rste!est_o_past)
   Label1(173) = atrim(rste!numerobossasoldadores)
   Set rstc = dbtmp.OpenRecordset("select * from clients_codiscomptables where codicomptable=" + atrim(cadbl(rste!codicomptable)) + " and codifabricacio=" + atrim(Data1.Recordset!client))
   Label1(17) = ""
   If Not rstc.EOF Then
      Text32(3) = atrim(rstc!codicomptable) + " - " + atrim(rstc!nomclient): Text32(3).Tag = atrim(rstc!codicomptable)
      Set rstc = dbtmp.OpenRecordset("select * from clients_codissap where codisap=" + atrim(rstc!codicomptable))
      If Not rstc.EOF Then
         Label1(17) = UCase("Credit: " + atrim(rstc!creditsap) + "€   Valor diferencial:" + atrim(Redondejar(cadbl(rstc!valordiferencial), 0)) + "€")
         If cadbl(rstc!valordiferencial) <= 0 Then
           Label1(17).ForeColor = QBColor(12)
            Else:
                Label1(17).ForeColor = &H80000012
         End If
      End If
   End If
   
   If Text32(3) = "" Then Text32(3).BackColor = QBColor(12)
   If Not rste.EOF Then
       If rste!noplanificable Then
           noplanificable.Value = 1
          Else: noplanificable.Value = 0
       End If
       If rste!materialexacte Then
           materialexacte(0).Value = 1
          Else: materialexacte(0).Value = 0
       End If
       If cadbl(rste!codigrupmaterialcompatible) > 0 Then
           materialexacte(4).Value = 1
           materialexacte(4).Caption = "Compatible: " + nomgrupcompatible(rste!codigrupmaterialcompatible)
           materialexacte(4).Visible = True
           materialexacte(0).Visible = False
          Else: materialexacte(4).Value = 0: materialexacte(4).Visible = False: materialexacte(0).Visible = True
       End If
       possarnommaterialexacte materialexacte(0).Value
       If rste!colaexacte Then
           materialexacte(2).Value = 1
          Else: materialexacte(2).Value = 0
       End If
   End If
   If cadbl(rste!gtin14) > 0 Then 'si hi ha codi gtin ensenyo el botó i guardo el valor al tag
      Command26(4).Tag = atrim(rste!gtin14)
      Command26(4).Visible = True
   End If
   If rste!PVPimpostinclos Then
        Text6.BackColor = &HC991FB
              Else: Text6.BackColor = QBColor(15)
   End If
fi:
  Set rste = Nothing
  Set rstc = Nothing
End Sub
Function nomgrupcompatible(vcodi As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select nomdelgrup from grupsmaterialscompatibles where numerodegrup=" + atrim(vcodi))
   If Not rst.EOF Then nomgrupcompatible = atrim(rst!nomdelgrup)
   Set rst = Nothing
End Function
Sub escullirvalorcombo(vcombo As ComboBox, vi As Integer)
   Dim i As Byte
   vcombo.BackColor = QBColor(15)
   For i = 0 To vcombo.ListCount - 1
      If vcombo.ItemData(i) = vi Then
         vcombo.ListIndex = i
         If vcombo.List(i) = "Parar a expedicions." Then vcombo.BackColor = QBColor(12)
         GoTo fi
      End If
   Next i
fi:
End Sub
Sub possarnommaterialexacte(vvalor As Byte)
  Dim rst As Recordset
  Dim vcodi As String
  If Data1.Recordset.EOF Then Exit Sub
  If vvalor = 0 Then nomcolor(23) = "": Exit Sub
  vcodi = atrim(Data1.Recordset!materialex)
  Set rst = dbtmp.OpenRecordset("SELECT refproducte, proveidor FROM materials where codi=" + vcodi + ";")
  If Not rst.EOF Then
    nomcolor(23).Caption = "Prov: " + atrim(rst!proveidor) + "-" + atrim(rst!refproducte)
  End If
  Set rst = Nothing
  End Sub
Sub possarbotocanvis()
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from comandes_controlcanvis where comanda=" + atrim(Data1.Recordset!comanda))
   If rst.EOF Then
       Command9(2).Visible = False
      Else: Command9(2).Visible = True
   End If
   
End Sub



Sub carregar_calloffs()
   Dim rstent As Recordset
   Dim rst As Recordset
   Dim vnumc As Double
   vnumc = cadbl(Text1)
   Combo1(2).Clear
   Set rstent = dbtmp.OpenRecordset("select distinct numcalloff,entregat from bobinesent where (numcalloff<>'' and numcalloff<>null) and comanda=" + atrim(vnumc) + " order by entregat")
   If rstent.EOF Then
      Set rst = dbtmp.OpenRecordset("select * from calloffs_detall where comanda=" + atrim(vnumc))
      If Not rst.EOF Then Combo1(2).AddItem atrim(rst!numcalloff)
       Else
          While Not rstent.EOF
            Combo1(2).AddItem atrim(rstent!numcalloff) + IIf(rstent!entregat = "S", "-E", "")
            rstent.MoveNext
          Wend
   End If
   If Combo1(2).ListCount > 0 Then
          Combo1(2).ListIndex = 0
            Else: Combo1(2) = buscarcalloffgeneric(vnumc)
   End If
End Sub
Function buscarcalloffgeneric(vnumc As Double, Optional vnumcalloffnou As String) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select numcalloff from comandes_extres where comanda=" + atrim(vnumc))
   If vnumcalloffnou = "" Then
         If Not rst.EOF Then buscarcalloffgeneric = atrim(rst!numcalloff)
          Else
             If Not rst.EOF Then
                 rst.Edit
                   Else: rst.AddNew
             End If
             rst!numcalloff = vnumcalloffnou
             rst.Update
   End If
   Set rst = Nothing
End Function

Sub carregar_obsdelatinta()
 Dim rst As Recordset
 cobs1imp = ""
 cobs2imp = ""
 Set rst = dbtmp.OpenRecordset("select * from comandes_observacionstintes where comanda=" + atrim(Data1.Recordset!comanda) + " order by id")
 If rst.EOF Then Exit Sub
 cobs1imp = atrim(rst!observacio)
 rst.MoveNext
 If Not rst.EOF Then cobs2imp = atrim(rst!observacio)
 Set rst = Nothing
End Sub
Sub carregartarifesperreferencia(vcodiclient As String, vref As String)
  Dim rst As Recordset
  Command1(9).BackColor = &H5C31DD
  Command1(9).Tag = ""
  Command1(9).ToolTipText = ""
  Label1(148) = ""
  'If cadbl(vcodiclient) > 0 Then
  '  Set rst = dbtmp.OpenRecordset("select grupdeclient from clients where codi=" + atrim(vcodiclient))
  '  If rst.EOF Then GoTo fi
  '  If rst!grupdeclient <> "" Then vcodiclient = UCase(rst!grupdeclient)
  'End If
  Command1(9).Tag = vcodiclient
  Text32(5).BackColor = &H6BEBB1
  Set rst = dbtmp.OpenRecordset("select * from tarifes_referencies where refinplacsa='" + atrim(vref) + "'")
  If Not rst.EOF Then
      rst.FindFirst "refinplacsa='" + vref + "'"
      If Not rst.NoMatch Then
        If atrim(rst!coditarifa) <> "" Then
          Command1(9).BackColor = &H25EFAD
          Command1(9).Tag = vcodiclient
          Command1(9).ToolTipText = "Codi tarifa: " + atrim(rst!coditarifa)
          Label1(148) = atrim(rst!coditarifa)
        End If
        If rst!inactiva Then
              Text32(5).BackColor = &H8080FF
               Else
                Text32(5).BackColor = &H6BEBB1
        End If
      End If
  End If
fi:
  Set rst = Nothing
End Sub
Sub carregar_firmes()
  Dim rst As Recordset
  Dim vnomordinador As String
  If Data1.Recordset.EOF Then Command9(8).Visible = False
  'If data1.Recordset!producte = "PC" Or data1.Recordset!producte = "PC2" Or data1.Recordset!producte = "PCP" Then Command9(8).Visible = False: Exit Sub
  Command9(8).Visible = True
  'Set rst = dbtmp.OpenRecordset("select * from comandes_firmes where anulada=false and comanda=" + atrim(cadbl(Text1)) + " and TIPUS='COM'")
  If cadbl(Text1) > 0 Then
    vnomordinador = nomordinador
    If vnomordinador = "ORDINADOR_LP" Or vnomordinador = "ORD_JOSEPM" Then
          Set rst = dbtmp.OpenRecordset("select * from comandes_firmes where anulada=false and tipus='COM' and comanda=" + atrim(cadbl(Text1)) + " and (usuari='ORDINADOR_LP' OR usuari='ORD_JOSEPM')")
         Else
           Set rst = dbtmp.OpenRecordset("select * from comandes_firmes where anulada=false and comanda=" + atrim(cadbl(Text1)) + " and usuari='" + vnomordinador + "'")
    End If
    If rst.EOF Then
      Command9(8).BackColor = &H5C31DD 'vermell
        Else: Command9(8).BackColor = &H25EFAD  'verd
    End If
  End If
  Set rst = Nothing
   
End Sub
Sub carregar_observacioPVP()
    Dim rst As Recordset
    Set rst = dbtmp.OpenRecordset("select * from comandes_observacioPVP where comanda=" + atrim(cadbl(Text1)))
    If rst.EOF Then
        Command26(6).ToolTipText = ""
        Command26(6).BackColor = &H8000000F
          Else:
             Command26(6).ToolTipText = IIf(Text6.BackColor = &HC991FB, "PVP impost inclòs. ", "") + atrim(rst!observacio) + vbNewLine + IIf(cadbl(rst!extracost) > 0, vbNewLine + "[Extra Cost = " + atrim(cadbl(rst!extracost)) + "]", "")
             Command26(6).BackColor = &H5C31DD
    End If
    Set rst = dbtmp.OpenRecordset("select * from comandes_observacioPVP where comandesafectades like '*" + atrim(cadbl(Text1)) + "*'")
    If Not rst.EOF Then
      Command26(6).ToolTipText = Command26(6).ToolTipText + " Extracost compartit amb: " + IIf(InStr(1, atrim(rst!comandesafectades), atrim(rst!comanda)) = 0, atrim(rst!comanda) + " i " + atrim(rst!comandesafectades), atrim(rst!comandesafectades))
      Command26(6).BackColor = &H5C31DD
    End If
    Set rst = Nothing
End Sub
Sub carregar_lookups()
 Dim risc As Double
 Dim riscpla As Double
 Dim empresarisc As String
 Dim impagats As Boolean
 Dim direnvio As Byte
 Dim vrisc As TipusVrisc
 
 If Not buscant Then
    lookupde "clients", Text2, nomclient, "nom", "clientsextres"
   Else: lookupde "clients", Text2, nomclient, "nom"
 End If
 If Data1.Recordset.EOF Then Exit Sub
  ruta_relativa_client = carpeta_del_client
 Text6.BackColor = QBColor(15)
 
 velclientvolPVPimpostinclos = False
 'mirar firmes
 carregar_firmes
 
 'carregar calloffs
 carregar_calloffs
 
 'carregar observacions de la tinta
 carregar_obsdelatinta
 
 'possar boto de canvis realitzats
 possarbotocanvis
 
 'possar el color del tipus d'impresio
 possaretiquetasistemaimpibandes
 
 'carrega els valors de comandes_extres
 carregarvalorscomandesextres
 
 'carregar tarifes per referencia
 carregartarifesperreferencia atrim(Data1.Recordset!client), atrim(Text32(5))  'refinplacsa
 
 
 'referencies de client alternatives
 r = Text32(1)
   If InStr(InStr(1, r, "|") + 1, r, "|") > 0 Then
      Command26(1).BackColor = &H8080FF
       Else: Command26(1).BackColor = &HFFFFFF
   End If
  areadatos.BackColor = &H8000000F
  If Data1.Recordset!refilate > 0 Then areadatos.BackColor = &H6BEBB1
 'comprovo si l'etiqueta de rebobinadora està marcada com ok
  
   If Data1.Recordset!etrebvistiplau Then
       Command6.BackColor = &H80FF80
      Else: Command6.BackColor = &H8080FF
   End If
 
 'possar comandaentrega2
   Set rsttmp = dbplanificacio.OpenRecordset("select * from planificaciototes where comanda=" + atrim(cadbl(Text1)))
   If Not rsttmp.EOF Then
        dataentrega2 = atrim(rsttmp!Data2)
        importancia = IIf(IsNull(rsttmp!importancia), "", rsttmp!importancia)
      Else: dataentrega = "": importancia = ""
   End If
   
 'lookup de direnvio
    
   Set rsttmp = dbtmp.OpenRecordset("select * from clients_envios where codi=" + atrim(cadbl(Data1.Recordset!client)))
   direnvio = 0
   If Not rsttmp.EOF Then rsttmp.MoveLast: direnvio = rsttmp.RecordCount
   Set rsttmp = dbtmp.OpenRecordset("select poblacioe,codi,peuimprenta,impostinclosalPVP from clients_envios where id=" + atrim(cadbl(Data1.Recordset!direnvio)))
   Label1(147).Caption = "Envio (" + atrim(cadbl(Data1.Recordset!direnvio)) + "):" ' atrim(direnvio) + "):"
   Label1(153).Caption = ""
   nomclient.Tag = ""
   
   If Not rsttmp.EOF Then
      If rsttmp!codi <> cadbl(Text2) Then
         MsgBox "La direcció d'envio no correspon a aquest client. Es canviarà automaticament la direccio d'enviament o bé REVISEU-LA", vbCritical, "Atenció"
         If Data1.Recordset.EditMode > 0 Then Data1.Recordset!direnvio = 0
        Else:
          velclientvolPVPimpostinclos = rsttmp!impostinclosalPVP
          Label1(147).Caption = "Envio (" + atrim(cadbl(Data1.Recordset!direnvio)) + "):" + atrim(rsttmp!poblacioe)
          nomclient.Tag = atrim(rsttmp!poblacioe)
          'lookupde "peuimprenta", rsttmp!peuimprenta, Label1(153)
          
          Set rsttmp = dbtmp.OpenRecordset("select descripcio from peuimprenta where codi=" + atrim(cadbl(rsttmp!peuimprenta)))
          If Not rsttmp.EOF Then Label1(153).Caption = atrim(rsttmp!descripcio)
          
          'poso la direccio d'envio amb tronja si no es la primera
          Set rsttmp = dbtmp.OpenRecordset("select * from clients_envios where codi=" + atrim(cadbl(Data1.Recordset!client)))
             Label1(147).BackStyle = 1
             If Not rsttmp.EOF Then
                rsttmp.MoveLast
                rsttmp.MoveFirst
                rsttmp.FindFirst "id=" + atrim(Data1.Recordset!direnvio)
                If Not rsttmp.NoMatch Then
                   Label1(147).Tag = atrim(rsttmp.AbsolutePosition + 1)
                  Else: Label1(147).Tag = "1"
                End If
                    Else: Label1(147).Tag = "1"
          End If
      End If
        Else:
          If cadbl(Data1.Recordset!direnvio) > 0 Then
             'If Data1.Recordset.EditMode = 0 Then
             '   MsgBox "Aquest client te una direccio d'envio que ja no existeix. MODIFIQUEU LA DIRECCIO D'ENVIAMENT"
             '  Else: MsgBox "He modificat la direccio d'enviament a cap, si no es correcte MODIFIQUEU A LA OPCIO CORRECTE": Data1.Recordset!direnvio = 0
             'End If
             dbtmp.Execute "update  comandes set direnvio=0 where comanda=" + atrim(cadbl(Data1.Recordset!comanda))
             
             
          End If
   End If
   'miro els nous valors si no hi ha mes d'una direccio envio
   If rsttmp.EOF Then
      'Set rsttmp = dbtmp.OpenRecordset("select peuimprenta from clients where codi=" + atrim(cadbl(data1.Recordset!client)))
      Set rsttmp = dbtmp.OpenRecordset("select poblacioe,codi,peuimprenta from clients_envios where id=" + atrim(cadbl(Data1.Recordset!direnvio)))
      If Not rsttmp.EOF Then
         Set rsttmp = dbtmp.OpenRecordset("select descripcio from peuimprenta where codi=" + atrim(cadbl(rsttmp!peuimprenta)))
         If Not rsttmp.EOF Then Label1(153).Caption = atrim(rsttmp!descripcio)
      End If
   End If
   
   MaskEdBox6.Tag = ""
   'LOOKUP COMISIÓ CLIENT i obligatquantitatdemanada
   Set rsttmp = dbtmp.OpenRecordset("select fix_com_rep,var_com_rep,obligatquantitatdemanada from clients where codi=" + atrim(cadbl(Data1.Recordset!client)))
   If Not rsttmp.EOF Then
     text77(10).Tag = IIf(rsttmp!fix_com_rep, "F", IIf(rsttmp!var_com_rep, "V", ""))
     If rsttmp!obligatquantitatdemanada Then MaskEdBox6.Tag = "obligatquantitatdemanada"
   End If
   Text77_Change 10
 'carregar observacio de PVP
 carregar_observacioPVP
   
  'LOOKUP DE producte
   Set rsttmp = dbtmp.OpenRecordset("select descripcio,ruta from productes where codi='" + atrim((Text3.Text)) + "'")
   If Not rsttmp.EOF Then
     nomproducte.Caption = atrim(rsttmp!descripcio)
     ruta = atrim(rsttmp!ruta)
    Else: nomproducte.Caption = "": ruta = ""
   End If
   'etiqueta de micromacroperforat a la capçalera de la seccio
   Label1(163).Visible = False
   Label1(164).Visible = False
   If InStr(1, ruta, "R") Then
      If (atrim(Data1.Recordset!microperforat) <> "N" And atrim(Data1.Recordset!microperforat) <> "") Or atrim(Data1.Recordset!rebmacroperforat) = "S" Then
         Label1(163).Visible = True
         Label1(163).BackStyle = 1
      End If
   End If
   If InStr(1, ruta, "R") Then
      If atrim(Data1.Recordset!microperforatsol) = "S" Then
         Label1(164).Visible = True
      End If
   End If
  'lookup de tipussoldadura
  Set rsttmp = dbtmp.OpenRecordset("select descripcio from tipussoldadura where codi='" + atrim((Text129.Text)) + "'")
  If Not rsttmp.EOF Then
     Text130 = atrim(rsttmp!descripcio)
    Else: Text130 = ""
  End If
  If cadbl(Data1.Recordset!tincclixes) = 1 Then
     tincclixes.Value = 1
    Else: tincclixes.Value = 0
  End If
  
  '--------- color vermell de l'estat de comanda
  If MaskEdBox12.Text = MaskEdBox14.Text And atrim(MaskEdBox12.Text) <> "" Then
    MaskEdBox12.BackColor = QBColor(12)
    MaskEdBox14.BackColor = QBColor(12)
   Else: MaskEdBox12.BackColor = QBColor(15): MaskEdBox14.BackColor = QBColor(15)
 End If

'------------  RISC CLIENTS   ------------------
puntrisc.Visible = True
lookupde "clients", Text2, Text22, "importrisc"
risc = cadbl(Text22)
lookupde "clients", Text2, Text22, "importriscpla"
riscpla = cadbl(Text22)
lookupde "clients", Text2, Text22, "companyiacredit"
empresarisc = atrim(Text22)
lookupde "clients", Text2, Text22, "impagats"
If Text22 = "" Then
   impagats = False
   Else: impagats = CBool(Text22)
End If
'label1(17) = ""
''calcular_credit_delclient cadbl(Text32(3).Tag), vrisc
''label1(17) = UCase("Credit: " + atrim(vrisc.creditsap) + "€   Valor diferencial:" + atrim(Redondejar(vrisc.valordiferencial, 0)) + "€")
'Label1(17) = "EMPRESA RISC-->" + empresarisc + "  ---- R I S C INP--> " + atrim(risc) + " --- PLA-->" + atrim(riscpla) + IIf(impagats, "    IMPAGATS", "")
'If (risc = 0 And riscpla = 0 And empresarisc <> "") Or impagats Then
'If empresarisc = "" Then Label1(17).ForeColor = QBColor(9)
Text22 = ""
DoEvents
 Label1(146).Refresh
' If cadbl(Label1(146).Caption) > 0 Then
'    puntrisc.Visible = True
'     Else: puntrisc.Visible = False
' End If
 colorrisc = cap.BackColor
Select Case cadbl(Label1(146).Caption)
  Case 2 'verd
   puntrisc.FillColor = &H80FF80
  Case 1 ' vermell
   puntrisc.FillColor = &HFF&
  Case Else 'sense color
   puntrisc.FillColor = colorrisc
 End Select
'-------------------------------------

lookupde "tipusentregues", Text11, Label3
lookupde "bobinesembolicades", text77(14), text77(13)
lookupde "tipusetiquetes", text77(15), text77(16)
lookupde "mesures", Text7, Text16
lookupde "mesureslineals", Text23, Text22
lookupde "mesureslineals", Text31, Text30
lookupde "mesureslineals", Text32(8), Text32(7)
'lookupde "colorants", Text24, nomcolor(23)
lookupde "materials", Text25, nommaterial(23)
possarcolordelmaterial Text25.Text, cmarcmaterial
lookupde "select grmcm3 from materials where codi=" + atrim(cadbl(Text25)), , grmcm3
lookupde "aditius", Text26, nomadditiu(23)
lookupde "accessoris", Text133, Label1(177)
lookupde "accessoris", Text134, ansa(0)
lookupde "accessoris", Text135, truquel(0)
'lookupde "select descripcio from maquines where maquina='E' and codi=" + atrim(cadbl((Text27.Text))), , nomextrussora(0)
lookupde "select descripcio from maquines where maquina='I' and codi=" + atrim(cadbl((Text73.Text))), , nomimpressora(1)
lookupde "select descripcio from maquines where maquina='L' and codi=" + atrim(cadbl((Text82.Text))), , nomlaminadora(0)
lookupde "select descripcio from maquines where maquina='S' and codi=" + atrim(cadbl((Text120.Text))), , nomsoldadora(0)
lookupde "select descripcio from maquines where maquina='R' and codi=" + atrim(cadbl((Text99.Text))), , nomrebobinadora(1)
primerproces.Value = cadbl(Data1.Recordset!refilatd)

possar_desc_lot Text80.Text, desclot1(0)
possar_desc_lot Text81.Text, desclot2(1)
possar_desc_lot Text100.Text, desclot1(1)
lookupde "mesureslineals", Text128, Text127
possar_noms_adhesius True
possarconsums
calcular_pesmtr2imetresrebipesreb
calcular_pes_metres_rebobinadora
possar_nous_materials
possarreduccionscilindre
Label1(137) = atrim(possargrmm2(cadbl(Data1.Recordset!materialex), micresmaterial(cadbl(Data1.Recordset!mesuraesp), cadbl(Data1.Recordset!espessor), atrim(Data1.Recordset!tubolam)))) + " G/m2"
 'mirar la posicio en la ruta
 If Not buscant Then
   text77(17) = posicioenlaruta(Text1, atrim(Data1.Recordset!proximaseccio), ruta)
     Else: text77(17) = ""
 End If
'carregar modificacions del treball actiu
  refrescartreballimodificacio
 'carregarmodificacionsdeltreball cadbl(data1.Recordset!numtreball)
Set rsttmp = Nothing

End Sub
Function possargrmm2(vcodimat As Double, vmicres As Double) As Double
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(vcodimat))
   If Not rst.EOF Then
       possargrmm2 = vmicres * rst!grmcm3
   End If
   Set rst = Nothing
End Function
Sub possarcolordelmaterial(vcodimaterial As String, vcontrolnom As Shape)
  Dim codicolor As Double
  Dim rst As Recordset
  If cadbl(vcodimaterial) = 0 Then GoTo fi
  Set rst = dbtmp.OpenRecordset("SELECT materials.codi, materials.familia, materials.subfamilia, subfamiliesmaterials.color AS elcolor FROM subfamiliesmaterials RIGHT JOIN materials ON subfamiliesmaterials.codi = materials.subfamilia WHERE (((materials.codi)=" + atrim(cadbl(vcodimaterial)) + "))", , ReadOnly)
  If rst.EOF Then GoTo fi
  codicolor = &H8000000F
  Select Case rst!elcolor
    Case "VERD"
       codicolor = QBColor(10)
    Case "TARONJA"
       codicolor = &H62B1F2
    Case "BLAU"
       codicolor = QBColor(9)
    Case "ROSA"
       codicolor = &HC78DFA
    Case "GROC"
       codicolor = QBColor(6)
    Case "VERMELL"
       codicolor = QBColor(12)
    Case "BLANC"
       codicolor = QBColor(15)
    Case Else
       codicolor = &H8000000F
  End Select
  vcontrolnom.BorderColor = codicolor
  Set rst = Nothing
  Exit Sub
fi:
  vcontrolnom.BorderColor = &H8000000F
End Sub
Function hihaarxiu(treball As Double, modificacio As Double, client As Double, direnvio As Double, impopdf As String) As Boolean
  Dim rst As Recordset
  If modificacio = 0 Then modificacio = 1
  If impopdf = "imp" Then
        Set rst = dbclixes.OpenRecordset("select * from clientsvinculats where id_treball=" + atrim(treball) + " and ordremodificacio=" + atrim(modificacio) + " and codiclient=" + atrim(client) + " and direnvio=" + atrim(direnvio))
        If Not rst.EOF Then
           If impopdf = "imp" Then hihaarxiu = rst!arxiuimp
        End If
  End If
  If impopdf = "pdf" Then
        Set rst = dbclixes.OpenRecordset("select pdfvalid from modificacions where id_treball=" + atrim(treball) + " and ordre=" + atrim(modificacio))
        If Not rst.EOF Then hihaarxiu = rst!pdfvalid
  End If
  Set rst = Nothing
End Function
Sub colocardensitatiliniaturatintes(numt As Double, numordre As Double, vnumtinters As Byte)
     Dim rstt As Recordset
     Dim vnumbandes As Byte
     Dim vcilindre As Double
     Dim vdesarroll As Double
     Dim i As Integer
     For i = 0 To 7
       cdetalltinter(i) = ""
     Next i
     Set rstt = dbclixes.OpenRecordset("select bandes from modificacions where id_treball=" + atrim(numt) + " and ordre=" + atrim(numordre))
     If Not rstt.EOF Then vnumbandes = cadbl(rstt!bandes)
     
     Set rstt = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(numt) + " and ordremodificacio=" + atrim(numordre) + " order by ordretinter")
     i = 19
     On Error Resume Next
     While Not rstt.EOF
        If vdesarroll = 0 Then vdesarroll = cadbl(rstt!desarroll): vcilindre = cadbl(rstt!cilindre)
        If atrim(rstt!color) <> "" Or cadbl(rstt!tinterlinkambid_treball) > 0 Then vnumtinters = vnumtinters + 1
        If atrim(rstt!densitatutilitzada) <> "" Then
            text77(i) = atrim(rstt!aniloxclixe) + "/" + atrim(rstt!densitatutilitzada) + IIf(cadbl(rstt!volum) > 0, "-v" + atrim(cadbl(rstt!volum)), "")
              Else: text77(i) = ""
        End If
        cdetalltinter(rstt!ordretinter - 1) = atrim(Data1.Recordset.Fields("detalltinter" + atrim(rstt!ordretinter)))
        If cdetalltinter(rstt!ordretinter - 1) = "" Then cdetalltinter(rstt!ordretinter - 1) = atrim(rstt!detalltinter)
        If atrim(rstt!clixeosleeve) <> "" And atrim(rstt!clixeosleeve) <> "Clixé" Then cdetalltinter(rstt!ordretinter - 1) = cdetalltinter(rstt!ordretinter - 1) + "(" + atrim(rstt!clixeosleeve) + ")"
        If cdetalltinter(rstt!ordretinter - 1) <> "" Then
            cdetalltinter(rstt!ordretinter - 1).Width = Len(cdetalltinter(rstt!ordretinter - 1)) * 80
            cdetalltinter(rstt!ordretinter - 1).Left = 3825 - cdetalltinter(rstt!ordretinter - 1).Width
            cdetalltinter(rstt!ordretinter - 1).Top = text77(rstt!ordretinter + 1).Top
            cdetalltinter(rstt!ordretinter - 1).Visible = True
            cdetalltinter(rstt!ordretinter - 1).ZOrder 0
              Else: cdetalltinter(rstt!ordretinter - 1).Visible = False
        End If
        i = i + 1
        rstt.MoveNext
     Wend
     If vdesarroll = 0 Then vdesarroll = 1: vcilindre = 1
     Label1(179) = atrim(vnumbandes * (vcilindre / vdesarroll)) + " Motius"
End Sub
Sub refrescartreballimodificacio()
  Dim activarcamps As Boolean
  Dim vnumtinters As Byte
   Command1(0).Enabled = False
    Command1(2).Enabled = False
   Command1(4).Enabled = False
  Text103(3) = atrim(cadbl(Data1.Recordset!numtreball)) + "/" + atrim(cadbl(Data1.Recordset!numordremodificacio))
  If cadbl(Data1.Recordset!numtreball) > 0 Then
    'If existeix_imp_treball(cadbl(data1.Recordset!numtreball), cadbl(data1.Recordset!numordremodificacio), cadbl(data1.Recordset!client), cadbl(data1.Recordset!direnvio)) Then
    If hihaarxiu(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), cadbl(Data1.Recordset!client), cadbl(Data1.Recordset!direnvio), "imp") Then
       Command1(2).Enabled = True
      Else
         Command1(2).Enabled = False
    End If
    'If existeix_pdf_treball(cadbl(data1.Recordset!numtreball), cadbl(data1.Recordset!numordremodificacio)) Then
    If hihaarxiu(cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), cadbl(Data1.Recordset!client), cadbl(Data1.Recordset!direnvio), "pdf") Then
       Command1(0).Enabled = True
      Else
         Command1(0).Enabled = False
    End If
  End If
  'miro si la comanda es nova i deixo editar els camps, si no els bloquejo
  If atrim(Data1.Recordset!impressio) = "N" And cadbl(Data1.Recordset!numtreball) = 0 Then
     activarcamps = True
       Else: activarcamps = False
  End If
  colocardensitatiliniaturatintes cadbl(Data1.Recordset!numtreball), cadbl(Data1.Recordset!numordremodificacio), vnumtinters
  Text63.BackColor = &HFFC0C0: Text63.Enabled = False
  If atrim(Data1.Recordset!arxiuimpressora) <> "" Then Command1(4).Enabled = True
  If Not buscant Then activaronocampsimpresio activarcamps
  If cadbl(Text63) <> vnumtinters Then Text63.Enabled = True:  Text63.BackColor = QBColor(12)
  comprovarestatclixe False

End Sub
Sub posarmarcailinia(treball As Double)
   Dim rst As Recordset
   If treball < 1 Then Text103(4) = "": Exit Sub
   Set rst = dbclixesnous.OpenRecordset("select marca,linia from clixes where id_Treball=" + atrim(treball))
   If Not rst.EOF Then
     Text103(4) = Mid(treure_apostruf(atrim(rst!marca)) & " - " & treure_apostruf(atrim(rst!linia)), 1, 60)
      Else: Text103(4) = "NO HI HA LINIA DE PRODUCTE DEFINIDA"
   End If
End Sub
Sub possar_nous_materials()
  Dim rstmat As Recordset
  Text26.Visible = True
  Text24.Visible = True
  If cadbl(Text25) < 500 Then Exit Sub
  Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(Text25)), , ReadOnly)
  If Not rstmat.EOF Then
    nommaterial(23) = descripciomaterial(rstmat)
    Text26.Visible = False
    Text24.Visible = False
    'nomcolor(23) = ""
    nomadditiu(23) = ""
  End If
End Sub

Function descripciomaterial(rstmat As Recordset, Optional nomesfamilia As Boolean) As String
  Dim desc As String
  Dim rstfam As Recordset
  Set rstfam = dbtmp.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)), , ReadOnly)
  If Not rstfam.EOF Then desc = desc + atrim(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from subfamiliesmaterials where codi=" + atrim(cadbl(rstmat!subfamilia)), , ReadOnly)
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(rstmat!familiacol)), , ReadOnly)
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from subfamiliescolorants where codi=" + atrim(cadbl(rstmat!subfamiliacol)), , ReadOnly)
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
 If nomesfamilia Then descripciomaterial = desc: Exit Function
  Set rstfam = dbtmp.OpenRecordset("select descripcio from familiesaditius where codi=" + atrim(cadbl(rstmat!familiaad)), , ReadOnly)
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmp.OpenRecordset("select descripcio from subfamiliesaditius where codi=" + atrim(cadbl(rstmat!subfamiliaad)), , ReadOnly)
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  descripciomaterial = desc
  Set rstfam = Nothing
End Function
Function descripciomaterial2(rstmat As Recordset, Optional nomesfamilia As Boolean) As String
  Dim desc As String
  Dim rstfam As Recordset
  Set rstfam = dbtmp.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)), , ReadOnly)
  If Not rstfam.EOF Then desc = desc + atrim(rstfam!descripcio)
 
  Set rstfam = dbtmp.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(rstmat!familiacol)), , ReadOnly)
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  
  descripciomaterial2 = desc
  Set rstfam = Nothing
End Function

Function af(v As Variant) As String
  v = atrim(v)
  If Len(v) > 1 Then
     v = " - " + v
    Else: v = ""
  End If
  af = v
End Function

Sub calcular_pesmtr2imetresrebipesreb()
   Dim pesmtr2 As Double
   Dim metresreb As Double
   Dim pesreb As Double
   Dim pes1 As Double
   Dim pes2 As Double
   Dim rebkilosa As Double
   Dim rebmetresa As Double
   Dim pesgrmcm2 As Double
   Dim pesxpeça As Double
   
   If cadbl(Data1.Recordset!linkcomanda1) > 0 Then
    Set rsttmp = dbtmp.OpenRecordset("select pes1000mtrs from comandes where comanda=" + atrim(cadbl(Data1.Recordset!linkcomanda1)), , dbReadOnly )
    If Not rsttmp.EOF Then pes1 = cadbl(rsttmp!pes1000mtrs)
  End If
  If cadbl(Data1.Recordset!linkcomanda2) Then
   Set rsttmp = dbtmp.OpenRecordset("select pes1000mtrs from comandes where comanda=" + atrim(cadbl(Data1.Recordset!linkcomanda2)), , dbReadOnly)
   If Not rsttmp.EOF Then pes2 = cadbl(rsttmp!pes1000mtrs)
  End If
  rebpes = ""
  rebmetres = ""
   If ((cadbl(Text18) * 1000) / 100) > 0 Then
     pesmtr2 = (cadbl(Text33) + pes1 + pes2) / ((cadbl(Text18) * 1000) / 100)
       Else: Exit Sub
   End If
   rebmetresa = cadbl(rebmetres)
   rebkilosa = cadbl(rebkilos)
   Text30.Tag = atrim(pesmtr2)
   If InStr(1, Text30, "MTRS") Then
      ampler = cadbl(Text98) * IIf(Combo9 = "L", 1, 2)
      If (Combo8 <> "L" And Combo9 <> "L") Then ampler = cadbl(Text98) 'excepcio si extrussora es tubo i rebobinadora també
      If InStr(1, ruta, "R") = 0 Then ampler = cadbl(Text18)
      pesreb = (ampler / 100) * (cadbl(Text29) * cadbl(Combo3))
      pesreb = pesreb * pesmtr2
      metresreb = cadbl(Text29) * cadbl(Combo3)
      rebmetres = atrim(Format(metresreb, "#,##0"))
      rebpes = atrim(Format(pesreb, "#,##0"))
      
      dbtmp.Execute "update comandes set rebmtrs=" + passaradecimalpunt(Format(rebmetres, "#00.00")) + ",rebkilos=" + passaradecimalpunt(Format(rebpes, "#00.00")) + "  where comanda=" + atrim(cadbl(Data1.Recordset!comanda))
   End If
   If InStr(1, ruta, "I") > 0 And cadbl(Data1.Recordset!dessarroll) > 0 Then
      peces = cadbl(metresreb) / (cadbl(Data1.Recordset!dessarroll) / 1000)
        Else: peces = 0
   End If
   rebpcs = Format(peces, "#,##0")
   
   pesgrmcm2 = calcularpesunitatsoldadora(Data1.Recordset)
   pesxpeça = calcularpesxrpeça(Data1.Recordset, pesgrmcm2)
   solpes = "Pes de " + Format(cadbl(Data1.Recordset!cantitatsol), "#,##0") + " Unitats -> " + Format(cadbl(Data1.Recordset!cantitatsol) * pesxpeça, "#,##0") + " Kg"
   solpes.Tag = atrim(cadbl(Data1.Recordset!cantitatsol) * pesxpeça)
   
   
   If InStr(1, Text30, "KG") Then
      If pesmtr2 > 0 Then
        metresreb = cadbl(Text29) / pesmtr2
      End If
      If cadbl(Text98) > 0 Then
       metresreb = Format(cadbl(metresreb) / (cadbl(Text98) / 100), "#,##0")
      End If
      pesreb = cadbl(Text29)
      rebmetres = atrim(Format(metresreb, "#,##0"))
      rebpes = atrim(Format(pesreb, "#,##0"))
      'dbtmp.Execute "update comandes set rebmetres=" + Format(cadbl(rebmetres), "###0,00") + " and rebkilos=" + Format(cadbl(rebpes), "###0,00") + " where comanda=" + atrim(cadbl(data1.Recordset!comanda))
      dbtmp.Execute "update comandes set rebmtrs=" + passaradecimalpunt(Format(rebmetres, "#00.00")) + ",rebkilos=" + passaradecimalpunt(Format(rebpes, "#00.00")) + "  where comanda=" + atrim(cadbl(Data1.Recordset!comanda))
   End If
   
   Set rsttmp = dbtmp.OpenRecordset("select rebmtrs,rebkilos from comandes where comanda=" + atrim(cadbl(Data1.Recordset!comanda)))
   
   
   
End Sub
Function calcularpesxrpeça(rst As Recordset, pesgrmcm2 As Double) As Double
    calcularpesxrpeça = pesgrmcm2 * ((cadbl(rst!amplesol)) * (cadbl(rst!longitudsol) + (cadbl(rst!solapasol) / 2)))
    calcularpesxrpeça = calcularpesxrpeça * IIf(rst!migelaboratsol = "L", 1, 2)
End Function
Function calcularpesunitatsoldadora(rst As Recordset) As Double
   Dim pesgrmcm2 As Double
   Dim amplecomanda As Double
   If rst.EOF Then Exit Function
   amplecomanda = cadbl(rst!ampleesq) + cadbl(rst!solapa)
   
   'pesgrmcm2 = ((calcularpescomanda(rst!comanda) / IIf(rst!tubolam = "L", 1, 2)) / 100000) / amplecomanda
  If amplecomanda > 0 Then pesgrmcm2 = Redondejar((calcularpescomanda(rst!comanda) / 100000) / amplecomanda, 9)
   dbtmp.Execute "update comandes_extres set solpesgrmcm2=" + passaradecimalpunt(atrim(pesgrmcm2)) + "  where comanda=" + atrim(rst!comanda)
   calcularpesunitatsoldadora = pesgrmcm2
End Function

Function calcularpescomanda(numc As Double) As Double
   Dim rstc As Recordset
   Set rstc = dbtmp.OpenRecordset("select linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(numc))
   If Not rstc.EOF Then
     Set rstc = dbtmp.OpenRecordset("select tubolam,pes1000mtrs,comanda from comandes where comanda=" + atrim(numc) + " or comanda=" + atrim(cadbl(rstc!linkcomanda1)) + " or comanda=" + atrim(cadbl(rstc!linkcomanda2)))
     calcularpescomanda = 0
     While Not rstc.EOF
        If rstc!comanda > 0 Then
           calcularpescomanda = calcularpescomanda + (cadbl(rstc!pes1000mtrs) / IIf(rstc!tubolam = "L", 1, 2))
        End If
        rstc.MoveNext
     Wend
   End If
   Set rstc = Nothing
End Function
Function calcularfactor(vpes1000 As Double, vvalor As Double) As Double
    Dim vample As Double
    vample = cadbl(Data1.Recordset!ampleesq)
    If (atrim(Data1.Recordset!migelaborat) <> "L") Then vample = cadbl(vample) * 2
    'vample = vample
    v = vpes1000 * 100 / vample
    v = v + vvalor
    vample = cadbl(Data1.Recordset!amplereb)
    If (atrim(Data1.Recordset!migelaborat) <> "L") Then vample = cadbl(vample) * 2
    v = (v / 100) * vample
    calcularfactor = v
End Function
Sub calcular_pes_metres_rebobinadora()
  Dim pes10001 As Double
  Dim pes10002 As Double
  Dim totalspes As Double
  Dim totalspesMesTA As Double
  Dim totalmtrs As Double
  Dim vfactormes As Double
  vfactormes = 2
  If cadbl(Data1.Recordset!linkcomanda1) > 0 Then
    vfactormes = vfactormes + 2
    Set rsttmp = dbtmp.OpenRecordset("select pes1000mtrs from comandes where comanda=" + atrim(cadbl(Data1.Recordset!linkcomanda1)), , dbReadOnly )
    If Not rsttmp.EOF Then pes10001 = cadbl(rsttmp!pes1000mtrs)
  End If
  If cadbl(Data1.Recordset!linkcomanda2) Then
   vfactormes = vfactormes + 2
   Set rsttmp = dbtmp.OpenRecordset("select pes1000mtrs from comandes where comanda=" + atrim(cadbl(Data1.Recordset!linkcomanda2)), , dbReadOnly)
   If Not rsttmp.EOF Then pes10002 = cadbl(rsttmp!pes1000mtrs)
  End If
  
  If InStr(1, Text30, "MTRS") Then
     
  ' pes material+tinta+adhesiu
    ' If cadbl(pes10001) > 0 Then vfactormes = calcularfactor(cadbl(pes10001), 2) '2 es el valor a afegir s'ha de convertir al factor
    ' If pes10001 > 0 Then totalspesMesTA = ((cadbl(Text29)) / 1000) * vfactormes
    ' If cadbl(pes10002) > 0 Then vfactormes = calcularfactor(cadbl(pes10002), 2) '2 es el valor a afegir s'ha de convertir al factor
    ' If pes10002 > 0 Then totalspesMesTA = totalspesMesTA + ((cadbl(Text29)) / 1000) * vfactormes
    ' If cadbl(Text33) > 0 Then vfactormes = calcularfactor(cadbl(Text33), 2) '2 es el valor a afegir s'ha de convertir al factor
    ' If cadbl(Text33) > 0 Then totalspesMesTA = totalspesMesTA + ((cadbl(Text29)) / 1000) * (vfactormes)
     
  ' pes material
     totalspes = ((cadbl(Text29)) / 1000) * pes10001
     totalspes = totalspes + ((cadbl(Text29)) / 1000) * pes10002
     totalspes = totalspes + ((cadbl(Text29)) / 1000) * cadbl(Text33)
     totalsmtrs = cadbl(Text29) * cadbl(Combo3)
     totalspesMesTA = Redondejar(((vfactormes / 1000) * (cadbl(Data1.Recordset!amplereb) / 100) * totalsmtrs) + cadbl(rebpes), 0)
     'rebpes = rebpes + "  (" + Format(totalspes, "#,##0") + ")  +TiA[" + Format(totalspesMesTA, "#,##0") + "]"
     rebpes = rebpes + "  +TiA " + Format(totalspesMesTA, "#,##0")
     rebmetres = rebmetres '+ "(" + Format(totalsmtrs, "#,##0") + ")"
     dbtmp.Execute "update comandes_extres set totalspesMesTA=" + atrim(cadbl(totalspesMesTA)) + " where comanda=" + atrim(Data1.Recordset!comanda)
  End If
  If InStr(1, Text30, "KG") Then
     totalspes = cadbl(Text29)
     If cadbl(Text33) <> 0 Then
        totalsmtrs = (cadbl(Text29) * 1000) / cadbl(Text33)
       Else: totalsmtrs = 0
     End If
     rebpes = rebpes + " (" + Trim(Format(totalspes, "#,##0")) + ")"
     rebmetres = rebmetres + " (" + Trim(Format(totalsmtrs, "#,##0") * cadbl(Combo3)) + ")"
     
  End If
End Sub
Sub llistatlookupde(taula As String, Optional control1 As String, Optional control2 As Control, Optional Camp As String, Optional altres As String)
Dim rsttmp2 As Recordset
If Camp = "" Then Camp = "descripcio"
If altres = "clientsextres" Then Camp = Camp + ",observacions1,observacions2,obsext1,obsext2,obsimp1,obsimp2,obslam1,obslam2,obsreb1,obsreb2,obssol1,obssol2"
If Len(taula) < 20 Then
    Set rsttmp2 = dbtmp.OpenRecordset("select " + Camp + " from " + taula + " where codi=" + atrim(cadbl(control1)), , ReadOnly)
   Else: Set rsttmp2 = dbtmp.OpenRecordset(taula, , ReadOnly)
End If
If Not rsttmp2.EOF Then
     control2 = atrim(rsttmp2.Fields(0))
    Else: control2 = ""
End If

End Sub
Sub actualitzacamp(nomcamp As String, valorcamp As String, numc As Double, Optional valor As Byte, Optional valoractual As String)
   If atrim(valorcamp) = atrim(valoractual) Then Exit Sub
   If Len(nomcamp) > 2 Then
    'If InStr(1, valorcamp, "'") > 0 Then
    'MsgBox valorcamp
    dbtmp.Execute "update comandes set " + atrim((nomcamp)) + "='" + treure_apostruf(atrim(valorcamp)) + "' where comanda=" + Trim((numc))
   End If
End Sub
Sub lookupde(taula As String, Optional control1 As Control, Optional control2 As Control, Optional Camp As String, Optional altres As String)
If Camp = "" Then Camp = "descripcio"
If altres = "clientsextres" Then Camp = Camp + ",observacions1,observacions2,obsext1,obsext2,obsimp1,obsimp2,obslam1,obslam2,obsreb1,obsreb2,obssol1,obssol2"
If Len(taula) < 20 Then
    Set rsttmp = dbtmp.OpenRecordset("select " + Camp + " from " + taula + " where codi=" + atrim(cadbl(control1.Text)), , ReadOnly)
   Else: Set rsttmp = dbtmp.OpenRecordset(taula, , ReadOnly)
End If
If Not rsttmp.EOF Then
     If rsttmp.Fields.Count = 0 Then GoTo fi
     control2 = atrim(rsttmp.Fields(0))
     If altres = "clientsextres" Then
      Text32(0) = atrim(rsttmp.Fields(1))
      'GoTo fi
      'Text12 = atrim(rsttmp.Fields(2))
      actualitzacamp Text32(0).DataField, atrim(rsttmp!observacions1), cadbl(Text1), , Text32(0).Text
      'If atrim(Text34) = "" Then
      actualitzacamp Text34.DataField, atrim(rsttmp!obsext1), cadbl(Text1), , Text34.Text
      'Text34 = atrim(rsttmp.Fields(3).Name)
  ' actualitzacamp Text35.DataField, atrim(rsttmp!obsext2), cadbl(Text1)
      'Text35 = atrim(rsttmp.Fields(4))
      actualitzacamp text77(0).DataField, atrim(rsttmp!obsimp1), cadbl(Text1), , text77(0).Text
      'text77(0) = atrim(rsttmp.Fields(5))
      'actualitzacamp text76.DataField, atrim(rsttmp!obsimp2), cadbl(Text1)
      'text76 = atrim(rsttmp.Fields(6))
      actualitzacamp Text93.DataField, atrim(rsttmp!obslam1), cadbl(Text1), , Text93.Text
      'Text93 = atrim(rsttmp.Fields(7))
      'actualitzacamp Text94, atrim(rsttmp!obslam2), cadbl(Text1)
      'Text94 = atrim(rsttmp.Fields(8))
      actualitzacamp Text108.DataField, atrim(rsttmp!obsreb1), cadbl(Text1), , Text108.Text
      'Text108 = atrim(rsttmp.Fields(9))
'      actualitzacamp Text110.DataField, atrim(rsttmp!obsreb2), cadbl(Text1)
      'Text110 = atrim(rsttmp.Fields(10))
      actualitzacamp Text17.DataField, atrim(rsttmp!obssol1), cadbl(Text1), , Text17.Text
      'Text17 = atrim(rsttmp.Fields(11))
'      actualitzacamp Text88.DataField, atrim(rsttmp!obssol2), cadbl(Text1)
      'Text88 = atrim(rsttmp.Fields(12))

     End If
    Else: control2 = ""
End If
fi:
End Sub

Sub possarvalordcamps(Optional tamany As Byte)
Dim t As Double
 If Data1.Recordset.EOF Then Exit Sub
 If cadbl(tamany) = 0 Then t = tamany
'On Error Resume Next
 For Each objecte In formcomandes
    If TypeOf objecte Is Label Then If objecte.WhatsThisHelpID = 0 Then objecte.BackStyle = 0
    If TypeOf objecte Is TextBox Or TypeOf objecte Is MaskEdBox Then
      If objecte.DataField <> "" And objecte.WhatsThisHelpID = 0 Then
         'If objecte.DataField = "desarrollclient" Then Stop
         If cadbl(tamany) = 0 And objecte.DataField <> "refinplacsa" And objecte.DataField <> "unitatsquantitatdemanada" Then
            t = tamany_camp(Data1.Recordset.Fields(objecte.DataField))
         End If
         
        ' objecte.Name
       
          'assigno el format standard a tots els controls
          
         If TypeOf objecte Is MaskEdBox Then
          If objecte.Format = "" Then
               If Not duplicant Then objecte.Format = format_camp(Data1.Recordset.Fields(objecte.DataField))
          End If
           Else: objecte.MaxLength = t
         End If
         
      End If
    End If
Next
'Label1(162).BackStyle = 0
Label1(163).BackStyle = 0
Label1(164).BackStyle = 0
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triartipusentrega
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarmesura
End Sub

Private Sub Text16_LostFocus()
 carregar_lookups
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarclient
End Sub

Private Sub Text2_LostFocus()
 possar_direccio_envio
 carregar_lookups
 If re <> Text2 And Not buscant Then Command23_Click
End Sub

Private Sub Text22_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarmesuraespesor
End Sub

Private Sub Text22_LostFocus()
   
   carregar_lookups
   calcular_micres_soldadores
End Sub
Function micresdelmaterialgrms(vmaterial As Double) As Double
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(vmaterial))
  If Not rst.EOF Then
      micresdelmaterialgrms = cadbl(rst!micresdelsgrm2)
  End If
  Set rst = Nothing
End Function
Sub calcular_micres_soldadores()
   Dim rst2 As Recordset
   Dim v As Double
   Dim vespessor As Double
   If InStr(1, ruta, "S") = 0 Then Exit Sub
   If cadbl(Text23) = 10 Then vespessor = cadbl(Text21)
   If cadbl(Text23) = 11 Then vespessor = micresdelmaterialgrms(cadbl(Data1.Recordset!materialex))
   If Data1.Recordset!linkcomanda1 > 0 Then
        Set rst2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(Data1.Recordset!linkcomanda1))
        If Not rst2.EOF Then
             If cadbl(rst2!mesuraesp) = 10 Then v = cadbl(rst2!espessor)
             If cadbl(rst2!mesuraesp) = 11 Then v = micresdelmaterialgrms(cadbl(rst2!materialex))
             vespessor = vespessor + v
             v = 0
        End If
   End If
   If Data1.Recordset!linkcomanda2 > 0 Then
        Set rst2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(Data1.Recordset!linkcomanda2))
        If Not rst2.EOF Then
             If cadbl(rst2!mesuraesp) = 10 Then v = cadbl(rst2!espessor)
             If cadbl(rst2!mesuraesp) = 11 Then v = micresdelmaterialgrms(cadbl(rst2!materialex))
             vespessor = vespessor + v
             v = 0
        End If
   End If
   
   Text126 = atrim(vespessor)
   Text128 = 10
   Text127 = "MICRES"
   
End Sub
Private Sub Text24_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Colorant", "colorants", Text24, nomcolor(23)
End Sub
Private Sub Text25_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim db As Database
 Dim rstm As Recordset
 Dim rstmatactual As Recordset
 Dim rstmatnou As Recordset
 Dim vcodimat As TextBox
 Dim vnommat As TextBox
If KeyCode = 113 Then
   Set rstmatactual = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(Text25)))
   triaralgu "Triar Material", "materials", Text25, nommaterial(23)
   Set rstm = dbtmp.OpenRecordset("select material2cares from materials where codi=" + atrim(cadbl(Text25)))
   If Not rstm.EOF Then
      If rstm!material2cares And MaskEdBox17 <> "2" Then MsgBox "OJU... AQUEST MATERIAL ESTÀ TRACTAT A DOS CARES I LA COMANDA POSSA MATERIAL QUE NO... FES ELS CANVIS OPORTUNS...", vbCritical, "ATENCIOOOOO"
      If Not rstm!material2cares And atrim(Data1.Recordset!producte) = "PC" And (cadbl(Data1.Recordset!linkcomanda1) > 0 And cadbl(Data1.Recordset!linkcomanda2) > 0) Then
          MsgBox "Atenció el material escullit no es tractat a dos cares i el material del mig ha de ser-ho", vbCritical, "Error"
          If Not rstmatactual.EOF Then
            Text25 = atrim(rstmatactual!codi)
            nommaterial(23) = atrim(rstmatactual!descripcio)
          End If
          Exit Sub
      End If
   End If
   
   Set rstmatnou = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(Text25)))
   If Not rstmatactual.EOF And Not rstmatnou.EOF Then
        If cadbl(rstmatactual!familia) <> cadbl(rstmatnou!familia) Then MsgBox "ATENCIÓ EL MATERIAL QUE HAS CANVIAT ES D'UNA FAMILIA DIFERENT QUE EL QUE JA HI HAVIA", vbCritical, "ATENCIÓ"
   End If
   
   calculcanvimaterial
 End If
Set rstm = Nothing
End Sub
Sub possar_lookup_manuals()
lookupde "colorants", Text24, nomcolor(23)
lookupde "materials", Text25, nommaterial(23)
lookupde "aditius", Text26, nomadditiu(23)
lookupde "accessoris", Text133, Label1(177)
lookupde "accessoris", Text134, ansa(0)
lookupde "accessoris", Text135, truquel(0)
lookupde "select descripcio from maquines where maquina='S' and codi=" + atrim(cadbl((Text120.Text))), , nomsoldadora(0)

End Sub
Sub calculcanvimaterial()
 Dim altres As String
 altres = "colorex=0,aditiuex=0,"
 If Data1.Recordset.EditMode > 0 And Not buscant Then
   Data1.Database.Execute "update comandes set " + IIf(cadbl(Text25) > 499, altres, "") + "materialex=" + atrim(cadbl(Text25)) + " where comanda=" + atrim(Text1)
   lookupde "select grmcm3 from materials where codi=" + atrim(cadbl(Text25)), , grmcm3
   Data1.Recordset!materialex = cadbl(Text25)
   Text33 = calcular_pes1000kg(atrim(Text1), cadbl(Text25))
'   calcular_pesmtr2imetresrebipesreb
'   calcular_pes_metres_rebobinadora
 End If
End Sub

Private Sub Text26_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Aditiu", "aditius", Text26, nomadditiu(23)
End Sub





Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarproducte
End Sub

Private Sub Text3_LostFocus()
'carregar_lookups
 Set rsttmp = dbtmp.OpenRecordset("select descripcio,ruta from productes where codi='" + atrim((Text3.Text)) + "'")
  If Not rsttmp.EOF Then
     nomproducte.Caption = atrim(rsttmp!descripcio)
     ruta = atrim(rsttmp!ruta)
    Else: nomproducte.Caption = "": ruta = ""
  End If

End Sub
Sub calcular_dies_entrega()
 Dim dies As Double
 Set rsttmp = dbtmp.OpenRecordset("select dies from productes where codi='" + atrim(Text3) + "'")
     If Not rsttmp.EOF Then
        dies = rsttmp!dies
      Else: dies = 0
     End If
 If atrim(Text4.Text) <> "" Then Text5 = DateAdd("d", dies, CVDate(atrim(Text4.Text)))
 Set rsttmp = Nothing
End Sub
Private Sub Text30_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = 113 Then triarmesuraquantitat
End Sub

Private Sub Text30_LostFocus()
 carregar_lookups
End Sub

Private Sub Timer2_Timer()
 Dim color As Double
 
'canviarelscolorsdelscontrolsalentrar
 On Error Resume Next
 'LES 4 LINIES SEGUENTS SON LES DE AMUNT I AVALL AMB EL CURSOR
 'If Not (TypeOf Screen.ActiveControl Is DBGrid) Then
 ' If tecla = 38 And (Screen.ActiveForm.Name = "formcomandes" Or Screen.ActiveForm.Name = "subbusqueda") Then tecla = 0: SendKeys ("+{TAB}"): tecla = 0
 ' If tecla = 40 And (Screen.ActiveForm.Name = "formcomandes" Or Screen.ActiveForm.Name = "subbusqueda") Then tecla = 0: SendKeys ("{TAB}"): tecla = 0
 'End If
 If buscant And (tecla > 42 And tecla < 112) Then
    ActiveControl.Tag = "9"
 End If
 Exit Sub
err:

End Sub
Sub canviarcolor()
'aquesta funcio es temporal no està funcionant ara mateix
 'retorna el color del control anterior
 If ultimcontrol.Name <> Screen.ActiveControl.Name And ultimcontrol.WhatsThisHelpID >= 90 Then
    If Not TypeOf Screen.ActiveControl Is CommandButton Then Screen.ActiveControl.BackColor = QBColor(colororiginal)
    If Not TypeOf Screen.ActiveControl Is CommandButton Then ultimcontrol.BackColor = QBColor(colororiginal)
    proxima_seccio ultimcontrol.WhatsThisHelpID
    DoEvents
    If Not TypeOf Screen.ActiveControl Is CommandButton Then Screen.ActiveControl.BackColor = QBColor(11)
    Set ultimcontrol = Screen.ActiveControl
 End If
 On Error GoTo 0
 
 
'fa canviar el color del control que te el focus
On Error Resume Next

   If TypeOf Screen.ActiveControl Is TextBox Or TypeOf Screen.ActiveControl Is MaskEdBox Or TypeOf Screen.ActiveControl Is ComboBox Then
     'color = Screen.ActiveControl.BackColor
     If color <> QBColor(11) Then
         seleccionartotelcontrol
         If Not TypeOf Screen.ActiveControl Is CommandButton Then
            Screen.ActiveControl.BackColor = QBColor(11) 'possar aqui el color
         End If
         If TypeOf ultimcontrol Is TextBox Or TypeOf ultimcontrol Is MaskEdBox Or TypeOf ultimcontrol Is ComboBox Then
          If Not TypeOf Screen.ActiveControl Is CommandButton Then ultimcontrol.BackColor = QBColor(colororiginal)
         End If
          Set ultimcontrol = Screen.ActiveControl
     End If
     Else:
        
        If TypeOf ultimcontrol Is TextBox Or TypeOf ultimcontrol Is MaskEdBox Or TypeOf ultimcontrol Is ComboBox Then
         ultimcontrol.BackColor = QBColor(colororiginal)
        End If
        Set ultimcontrol = Screen.ActiveControl
   End If
   
End Sub
Sub seleccionartotelcontrol()
 Screen.ActiveControl.SelStart = 0
     Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 1
End Sub
Sub proxima_seccio(ultimaseccio As Byte)
 Dim trobat As Boolean
 trobat = False
 ultimaseccio = ultimaseccio - 89
 For Each objecte In Me
   On Error Resume Next
   If objecte.WhatsThisHelpID = ultimaseccio Then objecte.SetFocus: trobat = True
  Next
 If trobat Then
  
   llocform = ultimaseccio - 1
  Else
   dataactivacio.SetFocus
    llocform = 0
 End If
   formscrooll.SetValues formscrooll.Values.HorzValue, taulapos(llocform)
End Sub
Sub possarconsums()
Dim valorscol
Dim val1, val2, val3, dens1, dens2 As Double
On Error Resume Next
val1 = 0: val2 = 0: val3 = 0: dens1 = 0
'val1 = (cadbl(Text89.Text) / 100) * 1000
val1 = (cadbl(Text91.Text) / 100) * 1000
val2 = cadbl(grmt2) / 1000
val3 = cadbl(pes1.Text) / (cadbl(pes1.Text) + cadbl(pes2.Text))
dens1 = (val1 * val2 * val3) / cadbl(grcm1(0))

val1 = 0: val2 = 0: val3 = 0: dens2 = 0
'val1 = (cadbl(Text89.Text) / 100) * 1000
val1 = (cadbl(Text91.Text) / 100) * 1000
val2 = cadbl(grmt2) / 1000
val3 = cadbl(pes2.Text) / (cadbl(pes1.Text) + cadbl(pes2.Text))
dens2 = (val1 * val2 * val3) / cadbl(grcm2(0))

On Error GoTo 0
valorscol = Array(1000, 2000, 3000, 4000, 5000, 7500, 10000, 15000, 20000, 30000, 40000, 50000, 75000, 100000, 150000, 200000)
'prepara l'amplada de les reixes i possa els titols
For i = 0 To 15
 If reixaconsums.Tag = "1" Then
  reixaconsums.ColWidth(i) = 620
  If i < 5 Then reixaconsums.ColWidth(i) = 520
 End If
 reixaconsums.col = i
 reixaconsums.row = 0
 If reixaconsums.Text <> valorscol(i) Then reixaconsums.Text = valorscol(i)
 reixaconsums.row = 1
 If reixaconsums.Text <> Format(dens1 * (valorscol(i) / 1000), "##,##0.00") Then
   reixaconsums.Text = Format(dens1 * (valorscol(i) / 1000), "##,##0.00")
 End If
 reixaconsums.row = 2
 If reixaconsums.Text <> Format(dens2 * (valorscol(i) / 1000), "##,##0.00") Then
   reixaconsums.Text = Format(dens2 * (valorscol(i) / 1000), "##,##0.00")
 End If
Next i
If reixaconsums.Tag = "1" Then
 reixaconsums.ColWidth(15) = 650
 reixaconsums.Width = (590 * 16) + 100
End If
reixaconsums.Tag = ""
reixaconsums.row = 1
reixaconsums.col = 15
litres1(1) = "-"
litres2(2) = "-"
If cadbl(reixaconsums.Text) <> 0 Then litres1(1) = (cadbl(reixaconsums.Text) / cadbl(reixaconsums.Text)) * 100
reixaconsums.row = 2
reixaconsums.col = 15
If cadbl(litres1(1)) <> 0 Then litres2(2) = reixaconsums.Text

End Sub


Private Sub tincclixes_Click()
On Error Resume Next
  If Data1.Recordset.EditMode > 0 And Screen.ActiveControl.Name = "tincclixes" Then Data1.Recordset!tincclixes = tincclixes.Value
End Sub

Private Sub tipusimpresio_Change()

End Sub



Private Sub VScroll1_Change()
  Dim vfactor As Double
  vfactor = ((formscrooll.Values.VertMax - formscrooll.Values.VertMin) / VScroll1.Max)
  vfactor = vfactor * IIf(VScroll1.Value > 1, VScroll1.Value, 0)
     
  formscrooll.SetValues formscrooll.Values.HorzValue, formscrooll.Values.VertMin + vfactor
End Sub
Sub possar_tarifes_client(rst As Recordset, vclient As String, vdata As Date, vvigents As Boolean, vproducte As String, vunitatpvp As Double)
 Dim vsql As String
 Dim vunpvp As String
 Dim vversio As String
 Dim vsqlvigents As String
 Dim vsqlaltres As String
 Dim rst2 As Recordset
 
 Set rst = dbtarifes.OpenRecordset("select * from tarifes_capcalera where client=-9999")
 Set rst2 = dbtmp.OpenRecordset("select unitatinterna from mesures where codi=" + atrim(vunitatpvp))
 If Not rst2.EOF Then vunpvp = rst2!unitatinterna
 If vunpvp = "" Then Exit Sub
 If vvigents Then vsqlvigents = "valid_inici<=#" + Format(vdata, "mm/dd/yy") + "# and valid_fi>=#" + Format(vdata, "mm/dd/yy") + "# and "
 If vproducte <> "" Then vsqlaltres = " and codiproducte='" + atrim(vproducte) + "' and unitat_facturacio='" + vunpvp + "'"
 If cadbl(vclient) > 0 Then
   vsql = crearllistaidtarifes("SELECT * From Tarifes_capcalera where " + vsqlvigents + " client=" + vclient + vsqlaltres + " order by Tarifes_capcalera.numerotarifa,versio desc;")
   If (vsql <> "") Then vsql = " and idtarifa in (" + vsql + ")"
   If vsql <> "" Then
     If vversio <> "" Then vsql = vversio
     Set rst = dbtarifes.OpenRecordset("select * from tarifes_capcalera where " + vsqlvigents + " client=" + vclient + vsql)
   End If
    Else:
       vsql = crearllistaidtarifes("SELECT * From Tarifes_capcalera where " + vsqlvigents + " grupclients='" + vclient + "'" + vsqlaltres + " order by Tarifes_capcalera.numerotarifa,versio desc;")
       If (vsql <> "") Then vsql = " and idtarifa in (" + vsql + ")"
       If vsql <> "" Then
         If vversio <> "" Then vsql = vversio
         Set rst = dbtarifes.OpenRecordset("select * from tarifes_capcalera where grupclients='" + vclient + "' " + vsql)
       End If
 End If
 
End Sub
Function crearllistaidtarifes(vsql As String) As String
  Dim rst As Recordset
  Dim vtarifa As Double
  Dim v As String
  Set rst = dbtarifes.OpenRecordset(vsql)
  While Not rst.EOF
    vtarifa = rst!numerotarifa
    crearllistaidtarifes = crearllistaidtarifes + IIf(crearllistaidtarifes <> "", "," + atrim(rst!idtarifa), atrim(rst!idtarifa))
    While vtarifa = rst!numerotarifa
      rst.MoveNext
      If rst.EOF Then GoTo fi
    Wend
    crearllistaidtarifes = crearllistaidtarifes + IIf(crearllistaidtarifes <> "", "," + atrim(rst!idtarifa), atrim(rst!idtarifa))
    rst.MoveNext
  Wend
fi:
  Set rst = Nothing
End Function
Sub buscar_tarifa_corresponent()
  Dim rsttarifa As Recordset
  Dim vclient As String
  Dim rstcli As Recordset
  Dim vidsvalids As String

  vclient = atrim(Data1.Recordset!client)
  Set rstcli = dbtmp.OpenRecordset("select grupdeclient from clients where codi=" + atrim(vclient))
  If Not rstcli.EOF Then If rstcli!grupdeclient <> "" Then vclient = atrim(rstcli!grupdeclient)
  
  Set dbtarifes = OpenDatabase(rutadelfitxer(cami) + "Tarifes.mdb", , False)
  possar_tarifes_client rsttarifa, vclient, Data1.Recordset!datacomanda, True, Data1.Recordset!producte, Data1.Recordset!mesurapvp
  If rsttarifa.EOF Then
     MsgBox "No hi ha cap tarifa que concordi"
      Else
         rsttarifa.MoveLast: rsttarifa.MoveFirst
         While Not rsttarifa.EOF
            vidsvalids = vidsvalids + atrim(rsttarifa!idtarifa) + ", ": rsttarifa.MoveNext
         Wend
         'vidsvalids = Mid(vidsvalids, 1, Len(vidsvalids) - 2)
         MsgBox "Hi ha " + atrim(rsttarifa.RecordCount) + " tarifes que concorden. " + atrim(vidsvalids)
  End If
  
  Set dbtarifes = Nothing
  Set rstcli = Nothing
  Set rsttarifa = Nothing
End Sub

