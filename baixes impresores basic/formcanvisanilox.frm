VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form formcanvisanilox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Canvis d'anilox"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14895
   Icon            =   "formcanvisanilox.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramedeltaE 
      BackColor       =   &H006BEBB1&
      Caption         =   "DeltaE proximes comandes"
      Height          =   4260
      Left            =   45
      TabIndex        =   102
      Top             =   4260
      Width           =   14760
      Begin MSFlexGridLib.MSFlexGrid reixadeltae 
         Height          =   3840
         Left            =   90
         TabIndex        =   103
         Top             =   315
         Width           =   14190
         _ExtentX        =   25030
         _ExtentY        =   6773
         _Version        =   393216
         Rows            =   6
         Cols            =   9
         AllowBigSelection=   0   'False
         SelectionMode   =   2
         Appearance      =   0
      End
   End
   Begin VB.Frame framebeliminar 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   8250
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   345
      Begin VB.Image Image1 
         Height          =   285
         Left            =   30
         Picture         =   "formcanvisanilox.frx":058A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.CommandButton Command8 
      Height          =   405
      Left            =   12780
      Picture         =   "formcanvisanilox.frx":0676
      Style           =   1  'Graphical
      TabIndex        =   100
      ToolTipText     =   "Imprimeix aniloxos per netejar."
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton Command6 
      Height          =   420
      Left            =   13845
      Picture         =   "formcanvisanilox.frx":1140
      Style           =   1  'Graphical
      TabIndex        =   98
      ToolTipText     =   "Llista d'aniloxos."
      Top             =   -15
      Width           =   945
   End
   Begin VB.Frame framebuscant 
      Height          =   780
      Left            =   3450
      TabIndex        =   68
      Top             =   1350
      Visible         =   0   'False
      Width           =   7965
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscant..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F3B378&
         Height          =   420
         Left            =   3105
         TabIndex        =   69
         Top             =   225
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   420
      Left            =   10125
      Picture         =   "formcanvisanilox.frx":1E0A
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Imprimir ultim canvi"
      Top             =   0
      Width           =   960
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   60
      Top             =   75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command3 
      Height          =   420
      Left            =   9135
      Picture         =   "formcanvisanilox.frx":2494
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Imprimir ultim canvi"
      Top             =   15
      Width           =   945
   End
   Begin VB.Frame framedades 
      Height          =   3735
      Left            =   9135
      TabIndex        =   1
      Top             =   360
      Width           =   5685
      Begin VB.CommandButton Command7 
         Height          =   330
         Left            =   4320
         Picture         =   "formcanvisanilox.frx":2A1E
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Borra els colors entrats per poder sel.leccionar-los de nou."
         Top             =   120
         Width           =   1290
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   480
         Picture         =   "formcanvisanilox.frx":2CEA
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Netejar les dades entrades."
         Top             =   150
         Width           =   480
      End
      Begin VB.CommandButton botobuscarcolor 
         Height          =   285
         Index           =   7
         Left            =   5325
         Picture         =   "formcanvisanilox.frx":3174
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   2895
         Width           =   300
      End
      Begin VB.CommandButton botobuscarcolor 
         Height          =   285
         Index           =   6
         Left            =   5325
         Picture         =   "formcanvisanilox.frx":36FE
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   2547
         Width           =   300
      End
      Begin VB.CommandButton botobuscarcolor 
         Height          =   285
         Index           =   5
         Left            =   5325
         Picture         =   "formcanvisanilox.frx":3C88
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   2200
         Width           =   300
      End
      Begin VB.CommandButton botobuscarcolor 
         Height          =   285
         Index           =   4
         Left            =   5325
         Picture         =   "formcanvisanilox.frx":4212
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   1853
         Width           =   300
      End
      Begin VB.CommandButton botobuscarcolor 
         Height          =   285
         Index           =   3
         Left            =   5325
         Picture         =   "formcanvisanilox.frx":479C
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   1506
         Width           =   300
      End
      Begin VB.CommandButton botobuscarcolor 
         Height          =   285
         Index           =   2
         Left            =   5325
         Picture         =   "formcanvisanilox.frx":4D26
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   1159
         Width           =   300
      End
      Begin VB.CommandButton botobuscarcolor 
         Height          =   285
         Index           =   1
         Left            =   5325
         Picture         =   "formcanvisanilox.frx":52B0
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   812
         Width           =   300
      End
      Begin VB.CommandButton botobuscarcolor 
         Height          =   285
         Index           =   0
         Left            =   5325
         Picture         =   "formcanvisanilox.frx":583A
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   465
         Width           =   300
      End
      Begin VB.CommandButton boto_neteja 
         Height          =   285
         Index           =   7
         Left            =   690
         Picture         =   "formcanvisanilox.frx":5DC4
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Escullir tipus de neteja."
         Top             =   2880
         Width           =   300
      End
      Begin VB.CommandButton boto_neteja 
         Height          =   285
         Index           =   6
         Left            =   690
         Picture         =   "formcanvisanilox.frx":634E
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Escullir tipus de neteja."
         Top             =   2532
         Width           =   300
      End
      Begin VB.CommandButton boto_neteja 
         Height          =   285
         Index           =   5
         Left            =   690
         Picture         =   "formcanvisanilox.frx":68D8
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Escullir tipus de neteja."
         Top             =   2190
         Width           =   300
      End
      Begin VB.CommandButton boto_neteja 
         Height          =   285
         Index           =   4
         Left            =   690
         Picture         =   "formcanvisanilox.frx":6E62
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Escullir tipus de neteja."
         Top             =   1848
         Width           =   300
      End
      Begin VB.CommandButton boto_neteja 
         Height          =   285
         Index           =   3
         Left            =   690
         Picture         =   "formcanvisanilox.frx":73EC
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Escullir tipus de neteja."
         Top             =   1506
         Width           =   300
      End
      Begin VB.CommandButton boto_neteja 
         Height          =   285
         Index           =   2
         Left            =   690
         Picture         =   "formcanvisanilox.frx":7976
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Escullir tipus de neteja."
         Top             =   1164
         Width           =   300
      End
      Begin VB.CommandButton boto_neteja 
         Height          =   285
         Index           =   1
         Left            =   690
         Picture         =   "formcanvisanilox.frx":7F00
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Escullir tipus de neteja."
         Top             =   822
         Width           =   300
      End
      Begin VB.CommandButton boto_neteja 
         Height          =   285
         Index           =   0
         Left            =   690
         Picture         =   "formcanvisanilox.frx":848A
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Escullir tipus de neteja."
         Top             =   480
         Width           =   300
      End
      Begin VB.TextBox cneteja 
         Height          =   330
         Index           =   7
         Left            =   2340
         TabIndex        =   78
         Top             =   2865
         Width           =   420
      End
      Begin VB.TextBox cneteja 
         Height          =   330
         Index           =   6
         Left            =   2340
         TabIndex        =   77
         Top             =   2517
         Width           =   420
      End
      Begin VB.TextBox cneteja 
         Height          =   330
         Index           =   5
         Left            =   2340
         TabIndex        =   76
         Top             =   2175
         Width           =   420
      End
      Begin VB.TextBox cneteja 
         Height          =   330
         Index           =   4
         Left            =   2340
         TabIndex        =   75
         Top             =   1833
         Width           =   420
      End
      Begin VB.TextBox cneteja 
         Height          =   330
         Index           =   3
         Left            =   2340
         TabIndex        =   74
         Top             =   1491
         Width           =   420
      End
      Begin VB.TextBox cneteja 
         Height          =   330
         Index           =   2
         Left            =   2340
         TabIndex        =   73
         Top             =   1149
         Width           =   420
      End
      Begin VB.TextBox cneteja 
         Height          =   330
         Index           =   1
         Left            =   2340
         TabIndex        =   72
         Top             =   807
         Width           =   420
      End
      Begin VB.TextBox cneteja 
         Height          =   330
         Index           =   0
         Left            =   2340
         TabIndex        =   71
         Top             =   465
         Width           =   420
      End
      Begin VB.TextBox ccolor 
         Height          =   330
         Index           =   7
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   65
         Top             =   2865
         Width           =   2550
      End
      Begin VB.TextBox ccolor 
         Height          =   330
         Index           =   6
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   64
         Top             =   2520
         Width           =   2550
      End
      Begin VB.TextBox ccolor 
         Height          =   330
         Index           =   5
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   63
         Top             =   2175
         Width           =   2550
      End
      Begin VB.TextBox ccolor 
         Height          =   330
         Index           =   4
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   62
         Top             =   1830
         Width           =   2550
      End
      Begin VB.TextBox ccolor 
         Height          =   330
         Index           =   3
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   61
         Top             =   1485
         Width           =   2550
      End
      Begin VB.TextBox ccolor 
         Height          =   330
         Index           =   2
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   60
         Top             =   1155
         Width           =   2550
      End
      Begin VB.TextBox ccolor 
         Height          =   330
         Index           =   1
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   59
         Top             =   810
         Width           =   2550
      End
      Begin VB.TextBox ccolor 
         Height          =   330
         Index           =   0
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   58
         Top             =   465
         Width           =   2550
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00008000&
         Height          =   315
         Left            =   75
         Picture         =   "formcanvisanilox.frx":8A14
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Copiar tots els anilox que hi havia."
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox cnumanilox 
         Height          =   345
         Index           =   7
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   2865
         Width           =   315
      End
      Begin VB.TextBox cnumanilox 
         Height          =   345
         Index           =   6
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2520
         Width           =   315
      End
      Begin VB.TextBox cnumanilox 
         Height          =   345
         Index           =   5
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   2175
         Width           =   315
      End
      Begin VB.TextBox cnumanilox 
         Height          =   345
         Index           =   4
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1830
         Width           =   315
      End
      Begin VB.TextBox cnumanilox 
         Height          =   345
         Index           =   3
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1500
         Width           =   315
      End
      Begin VB.TextBox cnumanilox 
         Height          =   345
         Index           =   2
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1155
         Width           =   315
      End
      Begin VB.TextBox cnumanilox 
         Height          =   345
         Index           =   1
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   810
         Width           =   315
      End
      Begin VB.TextBox cnumanilox 
         Height          =   345
         Index           =   0
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   465
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Height          =   420
         Left            =   60
         Picture         =   "formcanvisanilox.frx":8F9E
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Borra ultim canvi"
         Top             =   3240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton bacceptarcanvis 
         Height          =   420
         Left            =   900
         Picture         =   "formcanvisanilox.frx":9528
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3240
         Width           =   1170
      End
      Begin VB.TextBox cvolum 
         Height          =   330
         Index           =   7
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2865
         Width           =   390
      End
      Begin VB.TextBox cvolum 
         Height          =   330
         Index           =   6
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2520
         Width           =   390
      End
      Begin VB.TextBox cvolum 
         Height          =   330
         Index           =   5
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2175
         Width           =   390
      End
      Begin VB.TextBox cvolum 
         Height          =   330
         Index           =   4
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1830
         Width           =   390
      End
      Begin VB.TextBox cvolum 
         Height          =   330
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1485
         Width           =   390
      End
      Begin VB.TextBox cvolum 
         Height          =   330
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1155
         Width           =   390
      End
      Begin VB.TextBox cvolum 
         Height          =   330
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   810
         Width           =   390
      End
      Begin VB.TextBox cvolum 
         Height          =   330
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   465
         Width           =   390
      End
      Begin VB.TextBox cliniatura 
         Height          =   330
         Index           =   7
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   25
         Top             =   2865
         Width           =   465
      End
      Begin VB.TextBox cliniatura 
         Height          =   330
         Index           =   6
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   24
         Top             =   2520
         Width           =   465
      End
      Begin VB.TextBox cliniatura 
         Height          =   330
         Index           =   5
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   23
         Top             =   2175
         Width           =   465
      End
      Begin VB.TextBox cliniatura 
         Height          =   330
         Index           =   4
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   22
         Top             =   1830
         Width           =   465
      End
      Begin VB.TextBox cliniatura 
         Height          =   330
         Index           =   3
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   21
         Top             =   1500
         Width           =   465
      End
      Begin VB.TextBox cliniatura 
         Height          =   330
         Index           =   2
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   20
         Top             =   1155
         Width           =   465
      End
      Begin VB.TextBox cliniatura 
         Height          =   330
         Index           =   1
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   19
         Top             =   810
         Width           =   465
      End
      Begin VB.TextBox cliniatura 
         Height          =   330
         Index           =   0
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   18
         Top             =   465
         Width           =   465
      End
      Begin VB.CommandButton boto_igual 
         Height          =   285
         Index           =   7
         Left            =   90
         Picture         =   "formcanvisanilox.frx":9AB2
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2880
         Width           =   300
      End
      Begin VB.CommandButton boto_nou 
         Height          =   285
         Index           =   7
         Left            =   375
         Picture         =   "formcanvisanilox.frx":A03C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2880
         Width           =   300
      End
      Begin VB.CommandButton boto_igual 
         Height          =   285
         Index           =   6
         Left            =   90
         Picture         =   "formcanvisanilox.frx":A5C6
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2535
         Width           =   300
      End
      Begin VB.CommandButton boto_nou 
         Height          =   285
         Index           =   6
         Left            =   375
         Picture         =   "formcanvisanilox.frx":AB50
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2535
         Width           =   300
      End
      Begin VB.CommandButton boto_igual 
         Height          =   285
         Index           =   5
         Left            =   90
         Picture         =   "formcanvisanilox.frx":B0DA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2190
         Width           =   300
      End
      Begin VB.CommandButton boto_nou 
         Height          =   285
         Index           =   5
         Left            =   375
         Picture         =   "formcanvisanilox.frx":B664
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2190
         Width           =   300
      End
      Begin VB.CommandButton boto_igual 
         Height          =   285
         Index           =   4
         Left            =   90
         Picture         =   "formcanvisanilox.frx":BBEE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1845
         Width           =   300
      End
      Begin VB.CommandButton boto_nou 
         Height          =   285
         Index           =   4
         Left            =   375
         Picture         =   "formcanvisanilox.frx":C178
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1845
         Width           =   300
      End
      Begin VB.CommandButton boto_igual 
         Height          =   285
         Index           =   3
         Left            =   90
         Picture         =   "formcanvisanilox.frx":C702
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1515
         Width           =   300
      End
      Begin VB.CommandButton boto_nou 
         Height          =   285
         Index           =   3
         Left            =   375
         Picture         =   "formcanvisanilox.frx":CC8C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1515
         Width           =   300
      End
      Begin VB.CommandButton boto_igual 
         Height          =   285
         Index           =   2
         Left            =   90
         Picture         =   "formcanvisanilox.frx":D216
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1170
         Width           =   300
      End
      Begin VB.CommandButton boto_nou 
         Height          =   285
         Index           =   2
         Left            =   375
         Picture         =   "formcanvisanilox.frx":D7A0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1170
         Width           =   300
      End
      Begin VB.CommandButton boto_igual 
         Height          =   285
         Index           =   1
         Left            =   90
         Picture         =   "formcanvisanilox.frx":DD2A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   825
         Width           =   300
      End
      Begin VB.CommandButton boto_nou 
         Height          =   285
         Index           =   1
         Left            =   375
         Picture         =   "formcanvisanilox.frx":E2B4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   825
         Width           =   300
      End
      Begin VB.CommandButton boto_igual 
         Height          =   285
         Index           =   0
         Left            =   90
         Picture         =   "formcanvisanilox.frx":E83E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   300
      End
      Begin VB.CommandButton boto_nou 
         Height          =   285
         Index           =   0
         Left            =   375
         Picture         =   "formcanvisanilox.frx":EDC8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   300
      End
      Begin VB.Label etcomandarelacionada 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2325
         TabIndex        =   95
         Top             =   3300
         Width           =   3000
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Neteja"
         Height          =   270
         Left            =   2430
         TabIndex        =   70
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Color i intensitat"
         Height          =   270
         Left            =   3120
         TabIndex        =   66
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   1020
         TabIndex        =   56
         Top             =   2895
         Width           =   90
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   1020
         TabIndex        =   55
         Top             =   2535
         Width           =   90
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   1020
         TabIndex        =   54
         Top             =   2220
         Width           =   90
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1020
         TabIndex        =   53
         Top             =   1845
         Width           =   90
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   1020
         TabIndex        =   52
         Top             =   1545
         Width           =   90
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   1020
         TabIndex        =   51
         Top             =   1170
         Width           =   90
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1020
         TabIndex        =   50
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   1020
         TabIndex        =   49
         Top             =   495
         Width           =   90
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "(Nº)"
         Height          =   270
         Left            =   1635
         TabIndex        =   47
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volum"
         Height          =   270
         Left            =   1950
         TabIndex        =   35
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Liniatura"
         Height          =   270
         Left            =   990
         TabIndex        =   34
         Top             =   225
         Width           =   750
      End
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   3945
      Left            =   105
      TabIndex        =   0
      Top             =   30
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6959
      _Version        =   393216
      Rows            =   10
      Cols            =   19
      RowHeightMin    =   350
      BackColorSel    =   16744576
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FillStyle       =   1
      SelectionMode   =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label etstatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Dos clics per carregar els valors d'aquesta comanda."
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
      Height          =   240
      Left            =   2145
      TabIndex        =   97
      Top             =   3930
      Visible         =   0   'False
      Width           =   5025
   End
   Begin VB.Label Label3 
      Caption         =   "Canvis d'anilox"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   11220
      TabIndex        =   37
      Top             =   135
      Width           =   1860
   End
   Begin VB.Menu mbuscarcomanda 
      Caption         =   "Buscar comanda"
   End
End
Attribute VB_Name = "formcanvisanilox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vllistafaltenposar As Double
Dim vcopiarcolors As Boolean
Dim rsttimeline As Recordset
Dim vcomandaactiva As String

Private Sub bacceptarcanvis_Click()
   guardarcanvianiloxos
End Sub
Sub demanarcomandarelacionada()
   Dim vnumcomanda As String
   Dim rst As Recordset
   Dim vsqltinters As String
   
   vnumcomanda = InputBox("Entra el numero de comanda relacionat amb aquest canvi.", "Comanda relacionada")
   If cadbl(vnumcomanda) = 0 Then Exit Sub
   vnumcomanda = Trim(cadbl(vnumcomanda))
   Set rst = dbtmpb.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(cadbl(vnumcomanda)), , ReadOnly)
   If rst.EOF Then MsgBox "Aquesta comanda no existeix", vbCritical, "Error": GoTo fi
   If rst!proximaseccio <> "E" And rst!proximaseccio <> "I" Then MsgBox "Aquesta comanda no està apunt per imprimir", vbCritical, "Error": GoTo fi
   etcomandarelacionada.tag = vnumcomanda
   etcomandarelacionada = "Comanda: " + vnumcomanda
   
   vsqltinters = "SELECT IIf([tinterlinkambid_treball]<>0,[tinterlinkambid_treball],[id_tinter]) AS vidtinter FROM comandes RIGHT JOIN Tintes ON (comandes.numtreball = Tintes.id_treball) AND (comandes.numordremodificacio = Tintes.ordremodificacio) "
   vsqltinters = vsqltinters + " WHERE (((comandes.comanda)=" + etcomandarelacionada.tag + "));"
   
   
  ' MsgBox "SELECT Tintes.ordretinter, Tintes.color,tintes.id_tinter From Tintes WHERE Tintes.color<>'' AND Tintes.id_tinter In (" + atrim(cadbl(numerosdetinters(vsqltinters))) + ")  ORDER BY Tintes.ordretinter;"
   treuretintesquenocoincideixenamblacomanda "SELECT Tintes.ordretinter, Tintes.color,tintes.id_tinter From Tintes WHERE Tintes.color<>'' AND Tintes.id_tinter In (" + numerosdetinters(vsqltinters) + ")  ORDER BY Tintes.ordretinter;"
fi:
End Sub
Sub guardarcanvianiloxos()
   Dim i As Byte
   Dim rst As Recordset
   Dim vnumcomanda As String
   Dim vunaniloxpassatdemetres As Boolean
  ' If vllistafaltenposar > 0 Then MsgBox "Falten tintes per posar.", vbExclamation, "Tintes": Exit Sub
  
   rsttimeline.MoveFirst
  
   If cadbl(etcomandarelacionada.tag) = 0 Then demanarcomandarelacionada
   If cadbl(etcomandarelacionada.tag) = 0 Or vllistafaltenposar <> 0 Then
      If vllistafaltenposar > 0 Then MsgBox "Encara falten " + atrim(vllistafaltenposar) + " colors per posar."
      Exit Sub
   End If
   vnumcomanda = etcomandarelacionada.tag
   
   'AQUEST BLOC ES PER SI VOLS CONTROLAR QUE NOMÉS ENTRIN UN COP LA COMANDA
   'Set rst = dbtmpb.OpenRecordset("select * from aniloxtimeline where nummaquina=" + atrim(nummaq) + " AND comanda=" + atrim(vnumcomanda) + " order by data desc", , ReadOnly)
   'If Not rst.EOF Then
   '   If MsgBox("Aquesta comanda ja vas afegir un canvi. " + atrim(Format(rst!Data, "dd/mm/yy hh:nn")) + Chr(10) + "VOLS BORRAR L'ANTERIOR I DEIXAR EL QUE ESTAS ENTRANT?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
   '        rst.Delete
   '   End If
   'End If
   
   Set rst = dbtmpb.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(vnumcomanda))
   If rst.EOF Then MsgBox "Aquesta comanda no existeix", vbCritical, "Error": GoTo fi
   If rst!proximaseccio <> "E" And rst!proximaseccio <> "I" Then MsgBox "Aquesta comanda no està apunt per imprimir", vbCritical, "Error": GoTo fi
   Set rst = dbtmpb.OpenRecordset("select acabada from muntadoratot where comanda=" + atrim(cadbl(vnumcomanda)))
   If rst.EOF Then
     MsgBox "Aquesta comanda no està a la llista de comandes muntades.", vbCritical, "Error"
      Else: If Not rst!acabada Then MsgBox "Aquesta comanda no està acabada de muntar.", vbCritical, "Error"
   End If
   Set rst = Nothing
   rsttimeline.AddNew
   rsttimeline!comanda = vnumcomanda
   rsttimeline!Data = Now
   rsttimeline!numoperari = numop
   rsttimeline!nummaquina = nummaq
   For i = 0 To 7
      rsttimeline.Fields("anilox" + atrim(i + 1)) = IIf(cliniatura(i) = "", Null, cliniatura(i))
      rsttimeline.Fields("volum" + atrim(i + 1)) = IIf(cvolum(i) = "", Null, cvolum(i))
      rsttimeline.Fields("numanilox" + atrim(i + 1)) = IIf(cnumanilox(i) = "", Null, cnumanilox(i))
      rsttimeline.Fields("color" + atrim(i + 1)) = IIf(ccolor(i) = "", Null, ccolor(i))
      rsttimeline.Fields("neteja" + atrim(i + 1)) = IIf(cneteja(i) = "", Null, cneteja(i))
      rsttimeline.Fields("idtinter" + atrim(i + 1)) = IIf(ccolor(i).tag = "", Null, ccolor(i).tag)
      rsttimeline.Fields("matricula" + atrim(i + 1)) = buscarnumanilox(cadbl(rsttimeline!nummaquina), cadbl(rsttimeline.Fields("anilox" + atrim(i + 1))), cadbl(rsttimeline.Fields("volum" + atrim(i + 1))), cadbl(rsttimeline.Fields("numanilox" + atrim(i + 1))))
      cliniatura(i) = ""
      cvolum(i) = ""
      cnumanilox(i) = ""
      ccolor(i) = ""
      ccolor(i).tag = ""
      cneteja(i) = ""
      If cliniatura(i).BackColor = QBColor(12) Then vunaniloxpassatdemetres = True
   Next i
   etcomandarelacionada.tag = ""
   etcomandarelacionada = ""
   vllistafaltenposar = 999
   
   
   rsttimeline.Update
   If vunaniloxpassatdemetres Then MsgBox "ATENCIÓ HI HA UN ANILOX PASSAT DE METRES S'HAURIA D'AVISAR PER NETEJAR-LO.", vbCritical, "ATENCIÓ"
   carregar_timeline
   
fi:
End Sub
Function buscarnumanilox(vnummaq As Double, vanilox As Double, vvolum As Double, vnumanilox As Double) As String
   Dim rst As Recordset
   Dim vsql As String
   vsql = "SELECT aniloxos.nummaquina, aniloxos.lineatura, aniloxos.volum, aniloxos_informacio.matricula_inplacsa, aniloxos_informacio.matricula, aniloxos_informacio.actiu FROM aniloxos RIGHT JOIN aniloxos_informacio ON aniloxos.id = aniloxos_informacio.idanilox "
   vsql = vsql + " WHERE (((aniloxos.nummaquina)=" + atrim(vnummaq) + ") and matricula_inplacsa='" + atrim(vnumanilox) + "' AND ((aniloxos.lineatura)=" + atrim(vanilox) + ") AND ((aniloxos.volum)=" + passaradecimalpunt(atrim(vvolum)) + ") AND ((aniloxos_informacio.actiu)=True));"
   Set rst = dbtmp.OpenRecordset(vsql)
   If Not rst.EOF Then buscarnumanilox = atrim(rst!matricula)
   Set rst = Nothing
End Function
Sub ensenyarinformacio(vcomanda As Double, vhora As String)
   Dim i As Byte
   Dim rsttimeline As Recordset
   Set rsttimeline = dbtmpb.OpenRecordset("select * from aniloxtimeline order by data desc")
   rsttimeline.FindFirst "comanda=" + atrim(vcomanda) + " and format(data,'hh:nn')='" + atrim(vhora) + "'"
   If rsttimeline.NoMatch Then
        Exit Sub
   End If
   For i = 0 To 7
     If rsttimeline!nummaquina = nummaq Then
      cliniatura(i) = atrim(rsttimeline.Fields("anilox" + atrim(i + 1)))
      cvolum(i) = atrim(rsttimeline.Fields("volum" + atrim(i + 1)))
      cnumanilox(i) = atrim(rsttimeline.Fields("numanilox" + atrim(i + 1)))
        Else
            cliniatura(i) = ""
            cvolum(i) = ""
            cnumanilox(i) = ""
     End If
      ccolor(i) = atrim(rsttimeline.Fields("color" + atrim(i + 1)))
      ccolor(i).tag = atrim(rsttimeline.Fields("idtinter" + atrim(i + 1)))
      cneteja(i) = atrim(rsttimeline.Fields("neteja" + atrim(i + 1)))
   Next i
   etcomandarelacionada.tag = atrim(vcomanda)
   etcomandarelacionada = "Comanda: " + Trim(vcomanda)
   Set rsttimeline = Nothing
End Sub
Sub netejarinformacio()
   Dim i As Byte
   For i = 0 To 7
      cliniatura(i) = ""
      cvolum(i) = ""
      cnumanilox(i) = ""
      ccolor(i) = ""
      cneteja(i) = ""
      ccolor(i).tag = ""
      
      'canvicolorcontrols
       cliniatura(Index).BackColor = QBColor(15)
       cnumanilox(Index).BackColor = QBColor(15)
       cvolum(Index).BackColor = QBColor(15)
   Next i
   etcomandarelacionada.tag = ""
   etcomandarelacionada = ""

End Sub



Private Sub beliminar_Click()
  Dim v As String
  Dim vhora As String
  If MsgBox("Segur que vols eliminar aquests canvis d'anilox?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  v = reixa.TextMatrix(1, reixa.col) + "    "
  If Mid(v, 1, 3) = "Op:" Then v = reixa.TextMatrix(1, reixa.col - 1) + "    "
  v = Mid(v, 1, InStr(1, v, "-") - 1)
  vhora = reixa.TextMatrix(0, reixa.col + 1)
  If cadbl(v) > 0 Then
   rsttimeline.FindFirst "comanda=" + atrim(cadbl(v)) + " and format(data,'hh:nn')='" + atrim(vhora) + "'"
   If rsttimeline.NoMatch Then Exit Sub
   rsttimeline.Delete
   carregar_timeline
  End If
End Sub

Private Sub boto_igual_Click(Index As Integer)
   Dim rst As Recordset
   Dim vultimid As Double
   
   Set rst = dbtmpb.OpenRecordset("select * from aniloxtimeline where nummaquina=" + atrim(nummaq) + " order by data desc", , ReadOnly)
   'If Not rst.EOF Then vultimid = rst!id
   'Set rst = dbtmpb.OpenRecordset("select * from aniloxtimeline where nummaquina=" + atrim(nummaq) + " and anilox" + atrim(Index + 1) + ">0 order by data desc", , ReadOnly)
   'If rst.EOF Then Exit Sub
'   If rst!id <> vultimid Then If MsgBox("En el canvi anterior el tinter " + atrim(Index + 1) + " estava buit" + Chr(10) + "Vols copiar l'ultim valor que tenia?", vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then GoTo fi
   cliniatura(Index) = atrim(rst.Fields("anilox" + atrim(Index + 1)))
   cnumanilox(Index) = atrim(rst.Fields("numanilox" + atrim(Index + 1)))
   cvolum(Index) = atrim(rst.Fields("volum" + atrim(Index + 1)))
   If vcopiarcolors Then
        ccolor(Index) = atrim(rst.Fields("color" + atrim(Index + 1)))
        ccolor(Index).tag = atrim(rst.Fields("idtinter" + atrim(Index + 1)))
   End If
   possarvermellsSipassendemetres Index, rst.Fields("matricula" + atrim(Index + 1))
   cneteja(Index) = ""
fi:
   Set rst = Nothing
End Sub
Sub possarvermellsSipassendemetres(Index As Integer, vmatricula As String)
   Dim rstestadistica As Recordset
   cliniatura(Index).BackColor = QBColor(15)
   cnumanilox(Index).BackColor = QBColor(15)
   cvolum(Index).BackColor = QBColor(15)
   Set rstestadistica = dbtmp.OpenRecordset("select * from aniloxos_estadistica WHERE matricula='" + atrim(vmatricula) + "'")
   If Not rstestadistica.EOF Then
      If InStr(1, UCase(rstestadistica!observacioneteja), "METRES") > 0 Then
       cliniatura(Index).BackColor = QBColor(12)
       cnumanilox(Index).BackColor = QBColor(12)
       cvolum(Index).BackColor = QBColor(12)
      End If
      If InStr(1, UCase(rstestadistica!observacioneteja), "DIES") > 0 Then
       cliniatura(Index).BackColor = &HC78DFA
       cnumanilox(Index).BackColor = &HC78DFA
       cvolum(Index).BackColor = &HC78DFA
      End If
   End If
   Set rstestadistica = Nothing
End Sub

Private Sub boto_neteja_Click(Index As Integer)
   formcanvisanilox.tag = "escullint"
   Unload formseleccioneteja
   Load formseleccioneteja
   formseleccioneteja.Show 1
   cneteja(Index) = formseleccioneteja.etNeteges.tag
   Unload formseleccioneteja
   formcanvisanilox.tag = ""
End Sub

Private Sub boto_nou_Click(Index As Integer)
   Dim vliniatura As String
   Dim vvolum As String
   Dim vnumanilox As String
   Dim vmatricula As String
   vliniatura = InputBox("Entra la liniatura que vols utilitzar." + Chr(10) + "Escriu CAP per borrar la que hi ha.", "Liniatura")
   If UCase(vliniatura) = "CAP" Then cliniatura(Index) = "": cvolum(Index) = "": cnumanilox(Index) = "": Exit Sub
   If StrPtr(vliniatura) = 0 Then Exit Sub
   'If cadbl(vliniatura) = 0 Then Exit Sub
   
  ' vnumanilox = InputBox("Entra el Nº d'anilox que vols utilitzar.", "Nº")
  ' If cadbl(vnumanilox) = 0 Then Exit Sub
   
  ' vvolum = InputBox("Entra el volum que vols utilitzar.", "Volum")
  ' If cadbl(vvolum) = 0 Then Exit Sub
   formcanvisanilox.tag = "escullint"
   reixa.Enabled = False
   esculliranilox vliniatura, vvolum, vnumanilox, vmatricula
   formcanvisanilox.tag = ""
   If cadbl(vvolum) = 0 Then GoTo fi
   
   cliniatura(Index) = atrim(vliniatura)
   cvolum(Index) = atrim(vvolum)
   cnumanilox(Index) = atrim(vnumanilox)
   possarvermellsSipassendemetres Index, atrim(vmatricula)
fi:
   reixa.Enabled = True
End Sub
Sub esculliranilox(vnumliniatura As String, vvolum As String, vnumanilox As String, vmatricula As String)
   Dim rst As Recordset
   Dim vsql As String
   vvolum = ""
   vnumanilox = ""
   vsql = "SELECT aniloxos_informacio.matricula_inplacsa, aniloxos.volum, aniloxos_informacio.situacio, aniloxos_informacio.matricula, aniloxos_informacio.observacio,aniloxos.lineatura, aniloxos.nummaquina "
   vsql = vsql + " FROM aniloxos_informacio LEFT JOIN aniloxos ON aniloxos_informacio.idanilox = aniloxos.id "
   vsql = vsql + " WHERE (((aniloxos.lineatura)" + IIf(cadbl(vnumliniatura) = 0, ">", "=") + atrim(cadbl(vnumliniatura)) + ") AND ((aniloxos.nummaquina)=" + atrim(nummaq) + "))"
   vsql = vsql + " and (aniloxos_informacio.informacio = ""DATA ENTRADA DE L'ANILOX"" AND aniloxos_informacio.actiu=True)"
   vsql = vsql + " order by lineatura desc,situacio desc,volum asc;"
   Set rst = dbtmp.OpenRecordset(vsql, , dbReadOnly)
   If Not rst.EOF Then
         Load formseleccionou
         formseleccionou.caption = "Selecciona ANILOX"
         formseleccionou.Data1.DatabaseName = rutadelfitxer(cami) + "comandes.mdb"
         formseleccionou.Data1.RecordSource = vsql
         formseleccionou.refrescar
         formseleccionou.DBGrid2.Columns(0).width = 750
         formseleccionou.DBGrid2.Columns(1).width = 810
         formseleccionou.DBGrid2.Columns(2).width = 470
         formseleccionou.DBGrid2.Columns(3).width = 1100
         formseleccionou.DBGrid2.Columns(4).width = 10000
         formseleccionou.DBGrid2.Columns(5).width = 500
         formseleccionou.DBGrid2.Columns(6).visible = False
         formseleccionou.width = 15000
         formseleccionou.sortirs.tag = "filtre"
         formseleccionou.botofiltre.tag = 3
         formseleccionou.Show 1
         If seleccioret = 1 Then
          vvolum = atrim(formseleccionou.Data1.Recordset!volum)
          vnumanilox = atrim(formseleccionou.Data1.Recordset!matricula_inplacsa)
          vnumliniatura = atrim(formseleccionou.Data1.Recordset!lineatura)
          vmatricula = atrim(formseleccionou.Data1.Recordset!matricula)
         End If
         Unload formseleccionou
   End If
   
End Sub

Private Sub botobuscarcolor_Click(Index As Integer)
   formcanvisanilox.tag = "escullint"
   If cadbl(etcomandarelacionada.tag) = 0 Then demanarcomandarelacionada
   If cadbl(etcomandarelacionada.tag) = 0 Then Exit Sub
   demanarcolordeltreball Index, True
   formcanvisanilox.tag = ""
End Sub
Function numerosdetinters(vsql As String) As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset(vsql, , dbReadOnly)
   While Not rst.EOF
      numerosdetinters = numerosdetinters + IIf(numerosdetinters = "", "", ",") + atrim(rst!vidtinter)
      rst.MoveNext
   Wend
   If numerosdetinters = "" Then numerosdetinters = "0"
   Set rst = Nothing
End Function
Sub treuretintesquenocoincideixenamblacomanda(vsql As String)
  Dim rst As Recordset
  Dim i As Byte
  Dim vcont As Byte
  Set rst = dbclixes.OpenRecordset(vsql, , dbReadOnly)
  If rst.EOF Then GoTo fi
  For i = 0 To 7
    vtrobat = -1
    rst.FindFirst "id_tinter = " + atrim(cadbl(ccolor(i).tag))
    If rst.NoMatch Then
      ccolor(i) = "": ccolor(i).tag = ""
       Else: vcont = vcont + 1
    End If
  Next i
  If vcont = rst.RecordCount Then vllistafaltenposar = rst.RecordCount - vcont
fi:
  Set rst = Nothing
End Sub
Sub demanarcolordeltreball(vtinter As Integer, Optional vnoesreprint)
  Dim vsql As String
  Dim vllistaposats As String
  Dim vsqltinters As String
  ratoli "espera"
  If vllistaposats = "" Then vllistaposats = "0"
  
  vordremodificacio = "Tintes.ordremodificacio"
  '22/02/22 'If estemfentreprint And Not vnoesreprint Then vordremodificacio = "(tintes.ordremodificacio*-1)  "
  If estemfentreprint Then
     If MsgBox("Com que estàs fent reprint només ensenyaré tintes de REPRINT." + vbNewLine + "Fes SI per REPRINT o NO per NO-REPRINT.", vbInformation + vbDefaultButton2 + vbYesNo, "REPRINT") = vbYes Then
        vordremodificacio = "(tintes.ordremodificacio*-1)  "
         Else: vordremodificacio = "(tintes.ordremodificacio)  "
     End If
  End If
  
  vsqltinters = "SELECT IIf([tinterlinkambid_treball]>0,[tinterlinkambid_treball],[id_tinter]) AS vidtinter FROM comandes RIGHT JOIN Tintes ON (comandes.numtreball = Tintes.id_treball) AND (comandes.numordremodificacio = " + vordremodificacio + ") "
  vsqltinters = vsqltinters + " WHERE (((comandes.comanda)=" + etcomandarelacionada.tag + "));"
 
  
  treuretintesquenocoincideixenamblacomanda "SELECT Tintes.ordretinter, Tintes.color,tintes.id_tinter From Tintes WHERE Tintes.color<>'' AND Tintes.id_tinter In (" + numerosdetinters(vsqltinters) + ")  ORDER BY Tintes.ordretinter;"
  
  vllistaposats = llistatintersposats
  vsql = "SELECT Tintes.ordretinter, Tintes.color,tintes.id_tinter From Tintes WHERE (((Tintes.ordretinter) Not In (0)) AND ((Tintes.color)<>'') AND ((Tintes.id_tinter) In (" + numerosdetinters(vsqltinters) + "))) and id_tinter not in (" + vllistaposats + ") ORDER BY Tintes.ordretinter;"

  
'  Clipboard.Clear
'  Clipboard.SetText vsql
  Load formseleccionou
  formseleccionou.caption = "Selecciona color"
  formseleccionou.Data1.DatabaseName = rutadelfitxer(cami) + "clixesnous.mdb"
  formseleccionou.Data1.RecordSource = vsql
  formseleccionou.refrescar
 ' Clipboard.Clear
 ' Clipboard.SetText vsql
  'If formseleccionou.Data1.Recordset.EOF Then GoTo fi
  If Not formseleccionou.Data1.Recordset.EOF Then
    formseleccionou.Data1.Recordset.MoveLast: formseleccionou.Data1.Recordset.MoveFirst
  End If
  'formseleccionou.DBGrid2.Columns(0).visible = False
  formseleccionou.DBGrid2.Columns(0).width = 500
  formseleccionou.DBGrid2.Columns(1).width = 4000
  formseleccionou.DBGrid2.Columns(2).visible = False
  formseleccionou.DBGrid2.width = 5500
  formseleccionou.width = 6000
  ratoli "normal"
  formseleccionou.Show 1
  If seleccioret = 1 Then
   vllistafaltenposar = formseleccionou.Data1.Recordset.RecordCount - 1
   ccolor(vtinter) = atrim(formseleccionou.Data1.Recordset!color)
   ccolor(vtinter).tag = atrim(formseleccionou.Data1.Recordset!id_tinter)
  End If
  If seleccioret = 9 Then
   If ccolor(vtinter) <> "" Then vllistafaltenposar = formseleccionou.Data1.Recordset.RecordCount + 1
   ccolor(vtinter) = ""
   ccolor(vtinter).tag = ""
  End If
fi:
  Unload formseleccionou
End Sub
Function llistatintersposats() As String
   Dim i As Byte
   llistatintersposats = "0"
   For i = 0 To 7
      If ccolor(i).tag <> "" Then llistatintersposats = llistatintersposats + "," + atrim(ccolor(i).tag)
   Next i
End Function

Private Sub Command1_Click()
   If MsgBox("Segur que vols eliminar l'ultim canvi? " + reixa.TextMatrix(1, reixa.Cols - 2), vbCritical + vbDefaultButton2 + vbYesNo, "Error") = vbYes Then
       If Not rsttimeline.EOF Then
            rsttimeline.MoveFirst
            rsttimeline.Delete
            carregar_timeline
       End If
   End If
End Sub

Private Sub Command2_Click()
  Dim j As Integer
  mbuscarcomanda.tag = ""
  vcopiarcolors = False
  If MsgBox("Vols copiar també els colors?", vbInformation + vbYesNo + vbDefaultButton1, "Copiar colors") = vbYes Then vcopiarcolors = True
  ratoli "espera"
  carregar_timeline
  netejarinformacio
  For j = 0 To 7
    boto_igual_Click j
  Next j
  ratoli "normal"
End Sub

Private Sub Command3_Click()
  ' Dim llistat As Report
   Dim vnumc1 As String
   Dim cnumc2 As String
   Dim vidmax As Double
   Dim vidmax2 As Double
   Dim i As Byte
   carregar_timeline
   If rsttimeline.EOF Then Exit Sub
   rsttimeline.MoveFirst
   vidmax = rsttimeline!id
   vnumc2 = rsttimeline!comanda
   rsttimeline.MoveNext
   If rsttimeline.EOF Then vnumc1 = 0
   vidmax2 = cadbl(rsttimeline!id)
   If vidmax = 0 Then MsgBox "No trobo els dos ultims registres.", vbCritical, "Error": GoTo fi
   vnumc1 = rsttimeline!comanda
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "canvianiloximpresores.rpt"
   llistat.DataFiles(0) = rutadelfitxer(cami) + "Baixes.mdb"
   For i = 0 To 20
    llistat.Formulas(i) = ""
   Next i
   
'   llistat.SelectionFormula = "{aniloxtimeline.comanda}=" + vnumc1 + " or {aniloxtimeline.comanda}=" + atrim(vnumc2)
   llistat.SelectionFormula = "{aniloxtimeline.id}=" + atrim(vidmax) + " or {aniloxtimeline.id}=" + atrim(vidmax2)
   
   llistat.Destination = crptToPrinter
   'llistat.Destination = crptToWindow
   llistat.CopiesToPrinter = 1
   
   If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
   llistat.Action = 1
   
fi:
   
End Sub

Private Sub Command4_Click()
' Dim llistat As Report
   Dim vnumc1 As String
   Dim cnumc2 As String
   Dim vidmax As Double
   Dim vidmax2 As Double
   Dim rst As Recordset
   Dim v As String
   'mbuscarcomanda.tag = "comanda IN(223985)"
   carregar_timeline
   If rsttimeline.EOF Then Exit Sub
   rsttimeline.MoveFirst
   v = InputBox("Vols imprimir aquesta linia de temps?" + vbNewLine + "POSA LA DATA I HORA DE LA QUE VOLS IMPRIMIR", "Impresió", Format(rsttimeline!Data, "dd/mm/yy hh:nn"))
   If v = "" Then Exit Sub
   While Format(rsttimeline!Data, "dd/mm/yy hh:nn") <> v
     rsttimeline.MoveNext
     If rsttimeline.EOF Then MsgBox "No he trobat aquesta linia de temps": Exit Sub
   Wend
   vidmax = rsttimeline!id
   vnumc1 = rsttimeline!comanda
   Set rst = dbclixes.OpenRecordset("SELECT comandes.comanda, comandes.numordremodificacio, Modificacions.id_treball, Modificacions.ordre, Fotogravadors.nomfotogravador FROM comandes LEFT JOIN (Modificacions LEFT JOIN Fotogravadors ON Modificacions.fotograbador = Fotogravadors.codi) ON (comandes.numordremodificacio = Modificacions.ordre) AND (comandes.numtreball = Modificacions.id_treball)WHERE (((comandes.comanda)=" + atrim(vnumc1) + "));", , dbReadOnly)
 '  Clipboard.Clear
 '  Clipboard.SetText "SELECT comandes.comanda,comandes.numordremodificacio, Modificacions.id_treball,modificacions.ordre, Fotogravadors.nomfotogravador FROM comandes LEFT JOIN (Modificacions LEFT JOIN Fotogravadors ON Modificacions.fotograbador = Fotogravadors.codi) ON comandes.numtreball = Modificacions.id_treball WHERE (((comandes.comanda)=" + atrim(vnumc1) + "));"
   If rst.EOF Then Exit Sub
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "canvianiloximpresores_colors.rpt"
   llistat.DataFiles(0) = rutadelfitxer(cami) + "Baixes.mdb"
   
'   llistat.SelectionFormula = "{aniloxtimeline.comanda}=" + vnumc1 + " or {aniloxtimeline.comanda}=" + atrim(vnumc2)
   llistat.SelectionFormula = "{aniloxtimeline.id}=" + atrim(vidmax)
   llistat.Formulas(0) = "nt='NT: " + atrim(rst!id_treball) + "/" + atrim(rst!numordremodificacio) + "'"
   llistat.Formulas(1) = "nomfotogravador='FotoG: " + atrim(rst!nomfotogravador) + "'"
   possar_kgtinta cadbl(vidmax)
   llistat.Destination = crptToPrinter
   llistat.CopiesToPrinter = 1
   
   If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
   llistat.Action = 1
   Set rst = Nothing
fi:
   
End Sub
Sub possar_kgtinta(vidmax As Double)
   Dim rstc As Recordset
   Dim rst As Recordset
   Dim numtinter As Double
   Dim rstlink As Recordset
   Dim vnumc As Double
   Dim i As Byte
   Dim vkgtinta As Double
   
   Set rst = dbtmp.OpenRecordset("select * from aniloxtimeline where id=" + atrim(vidmax), , ReadOnly)
   
   If rst.EOF Then Exit Sub
   vnumc = rst!comanda
   Set rstc = dbtmp.OpenRecordset("select numtreball,numordremodificacio,cantitatex from comandes where comanda=" + atrim(vnumc), , ReadOnly)
   For i = 1 To 8
      vkgtinta = 0
      Set rstlink = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(cadbl(rst.Fields("idtinter" + atrim(i)))), , ReadOnly)
      If Not rstlink.EOF Then
             vkgtinta = Redondejar(calcular_kgmetreteoric(rstlink) * cadbl(rstc!cantitatex), 1)
             llistat.Formulas(i + 2) = "KgTinta" + atrim(i) + "=" + passaradecimalpunt(atrim(vkgtinta))
              Else: llistat.Formulas(i + 2) = "KgTinta" + atrim(i) + "=" + passaradecimalpunt(atrim(vkgtinta))
      End If
      DoEvents
   Next i
   Set rstc = Nothing
   Set rst = Nothing
   Set rstlink = Nothing
   
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
Private Sub Command5_Click()
   netejarinformacio
   mbuscarcomanda.tag = ""
   carregar_timeline
End Sub

Private Sub Command6_Click()
   Dim rst As Recordset
   Dim vsql As String
   Dim v As String
   formcanvisanilox.tag = "escullint"
   vvolum = ""
   vnumanilox = ""
   'vsql = "SELECT aniloxos_informacio.matricula_inplacsa, aniloxos.lineatura, aniloxos.volum, aniloxos_informacio.situacio, aniloxos_informacio.matricula, aniloxos_estadistica.metres as [Mtrs_Neteja], aniloxos_estadistica.metrestotal as [Mtrs_Total], aniloxos_informacio.observacio, aniloxos.nummaquina "
   'vsql = vsql + " FROM (aniloxos_informacio LEFT JOIN aniloxos ON aniloxos_informacio.idanilox = aniloxos.id) INNER JOIN aniloxos_estadistica ON aniloxos_informacio.idanilox = aniloxos_estadistica.id"
   vsql = "SELECT aniloxos_informacio.matricula_inplacsa, iif(aniloxos.lineatura=1,'GT',trim(aniloxos.lineatura)) as Liniatura, iif(aniloxos.volum=1,'L',trim(aniloxos.volum)) as Volum_, aniloxos_informacio.situacio, aniloxos_informacio.matricula, aniloxos_estadistica.observacioneteja AS Toca_Neteja,aniloxos_estadistica.metres AS Mtrs_Neteja, aniloxos_estadistica.metrestotal AS Mtrs_Total, aniloxos_informacio.observacio, aniloxos.nummaquina"
   vsql = vsql + " FROM (aniloxos_informacio LEFT JOIN aniloxos ON aniloxos_informacio.idanilox = aniloxos.id) LEFT JOIN aniloxos_estadistica ON aniloxos_informacio.matricula = aniloxos_estadistica.matricula "
   vsql = vsql + " WHERE aniloxos.nummaquina=" + atrim(nummaq)
   vsql = vsql + " and (aniloxos_informacio.informacio = ""DATA ENTRADA DE L'ANILOX"" AND aniloxos_informacio.actiu=True)"
   vsql = vsql + " order by lineatura DESC,situacio desc,volum asc;"
   Clipboard.Clear
   Clipboard.SetText vsql
   Set rst = dbtmp.OpenRecordset(vsql, , dbReadOnly)
  ' Clipboard.Clear
  ' Clipboard.SetText vsql
   If Not rst.EOF Then
         Unload formseleccionou
         Load formseleccionou
         formseleccionou.caption = "Selecciona ANILOX"
         formseleccionou.Data1.DatabaseName = rutadelfitxer(cami) + "comandes.mdb"
         formseleccionou.Data1.RecordSource = vsql
         
         formseleccionou.refrescar
         formseleccionou.DBGrid2.Columns(0).width = 750
         formseleccionou.DBGrid2.Columns(1).width = 500
         formseleccionou.DBGrid2.Columns(2).width = 810
         formseleccionou.DBGrid2.Columns(3).width = 470
         formseleccionou.DBGrid2.Columns(4).width = 1100
         formseleccionou.DBGrid2.Columns(5).width = 1500
         formseleccionou.DBGrid2.Columns(6).width = 1200
         formseleccionou.DBGrid2.Columns(7).width = 1100
         formseleccionou.DBGrid2.Columns(8).width = 7500
         formseleccionou.DBGrid2.Columns(9).visible = False
         formseleccionou.Top = 50
         formseleccionou.width = 16000
         formseleccionou.Height = 11500
         
         formseleccionou.sortirs.tag = "filtre"
         formseleccionou.botofiltre.tag = 3
         formseleccionou.caption = "Escull anilox"
         formseleccionou.Show 1
         If seleccioret = 1 Then
            v = InputBox("Escriu la observació d'aquest anilox.", "Observació", atrim(formseleccionou.Data1.Recordset!observacio))
            If StrPtr(v) = 0 Then GoTo fi
            v = UCase(Mid(treure_apostruf(v) + " ", 1, 255))
            formseleccionou.Data1.Recordset.Edit
            formseleccionou.Data1.Recordset!observacio = v
            formseleccionou.Data1.Recordset.Update
         End If
        Unload formseleccionou
   End If
fi:
    Set rst = Nothing
    formcanvisanilox.tag = ""
End Sub

Private Sub Command7_Click()
  Dim i As Byte
  If MsgBox("Segur que vols borrar tots els colors?", vbExclamation + vbDefaultButton2 + vbYesNo, "Borrar colors") = vbYes Then
      For i = 0 To 7
        ccolor(i).tag = ""
        ccolor(i) = ""
      Next i
      vllistafaltenposar = 999
  End If
End Sub

Private Sub Command8_Click()
 
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vnummaq As Byte

  vnummaq = cadbl(InputBox("Escriu el numero de màquina que vols consultar. [7,9]" + Chr(10) + "Ex: 7 ", "Escull màquina"))
  If vnummaq = 0 Then Exit Sub
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatdaniloxospernetejar.rpt", 1) '"etiqueta_llaunes.rpt"
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "comandes.mdb"
  oreport.RecordSelectionFormula = "{resumestadisticaaniloxos.nummaquina}=" + atrim(vnummaq) + " and {resumestadisticaaniloxos.Toca_Neteja}<>''"
  'oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
 ' oreport.Sections("D").ReportObjects.Item("serie2").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  oreport.DiscardSavedData
  'oreport.Sections("D").ReportObjects.Item("recuperador").Suppress = True
  oreport.FormulaFields.GetItemByName("titol").text = "'Aniloxos per netejar. Maq:" + atrim(vnummaq) + "'"
  oreport.DisplayProgressDialog = False
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
    Else: oreport.PrintOut False, 1
  End If
End Sub

Private Sub etcomandarelacionada_DblClick()
   etcomandarelacionada = ""
   etcomandarelacionada.tag = ""
   
   demanarcomandarelacionada
End Sub

Private Sub Form_Activate()
   If nummaq = 0 Then nummaq = 7
   If formcanvisanilox.tag = "escullint" Then Exit Sub
   If existeix("c:\ordprog.ini") Then Command1.visible = True
   vcomandaactiva = form1.comanda
   If InStr(1, formcanvisanilox.tag, "nomesdelta") = 0 And Not isloaded("formtintes") Then
       carregar_timeline
   End If
   ensenyar_deltaE
End Sub

Private Sub Form_Load()
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "\clixesnous.mdb")
   vllistafaltenposar = 999
   configurar_reixa
   
End Sub
Function mtrsminut(vnumc As Double) As Double
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select mtrsminut from impressores where comanda=" + atrim(vnumc) + " and mtrsminut>0 order by id desc")
  If Not rst.EOF Then mtrsminut = cadbl(rst!mtrsminut)
  Set rst = Nothing
End Function
Sub carregar_timeline()
  Dim rst As Recordset
  Dim cont As Byte
  Dim vnumaniloxsihiha As String
  Dim vnummaq As String
  Dim rstaniloxos As Recordset
  
  framebuscant.visible = True
  DoEvents
  Set rstaniloxos = dbtmp.OpenRecordset("select distinct matricula,matricula_inplacsa,situacio,diesneteja,metresneteja, actiu from aniloxos_informacio where informacio=""DATA ENTRADA DE L'ANILOX""")
  Set rst = dbtmpb.OpenRecordset("select * from aniloxtimeline where " + IIf(InStr(1, mbuscarcomanda.tag, "IN(") > 0, "", "nummaquina=" + atrim(nummaq)) + mbuscarcomanda.tag + " order by data desc", , dbReadOnly)
  framebuscant.visible = False
  DoEvents
  cont = 1
  reixa.Clear
  configurar_reixa
  If reixa.visible Then
       reixa.SetFocus
        Else: Exit Sub
  End If
  While Not rst.EOF And cont < (reixa.Cols - 2)
     vnummaq = ""
     If InStr(1, mbuscarcomanda.tag, "IN(") > 0 Then vnummaq = "M:" + atrim(rst!nummaquina) + " "
     reixa.TextMatrix(1, reixa.Cols - (cont + 1)) = atrim(rst!comanda) + "-" + atrim(mtrsminut(rst!comanda))
     reixa.TextMatrix(1, reixa.Cols - (cont)) = "Op:" + atrim(rst!numoperari)
     reixa.TextMatrix(0, reixa.Cols - (cont + 1)) = vnummaq + Format(rst!Data, "dd/mm/yy")
     reixa.TextMatrix(0, reixa.Cols - (cont)) = Format(rst!Data, "hh:nn")
     For i = 1 To 8
      rstaniloxos.FindFirst "matricula='" + atrim(rst.Fields("matricula" + atrim(i)) + "'")
      vnumaniloxsihiha = atrim(rst.Fields("numanilox" + atrim(i)))
      If cadbl(rst.Fields("volum" + atrim(i))) <> 0 Then vnumaniloxsihiha = "  (" + vnumaniloxsihiha + ")"
      reixa.TextMatrix(i + 1, reixa.Cols - (cont + 1)) = atrim(rst.Fields("anilox" + atrim(i))) + vnumaniloxsihiha
      reixa.TextMatrix(i + 1, reixa.Cols - cont) = atrim(rst.Fields("volum" + atrim(i)))
      If Not rstaniloxos.NoMatch Then
          If rstaniloxos!situacio = "C" Then
           reixa.row = i + 1: reixa.col = reixa.Cols - cont
           reixa.CellBackColor = QBColor(12)
          End If
      End If
     Next i
     rst.MoveNext
     cont = cont + 2
  Wend
 seleccionarlaultimacolumna
 Set rsttimeline = dbtmpb.OpenRecordset("select * from aniloxtimeline where nummaquina=" + atrim(nummaq) + " order by data desc")
End Sub
Sub seleccionarlaultimacolumna()
  
 reixa.SetFocus
 Sendkeys "{END}"
 reixa.SelectionMode = flexSelectionFree
 wait 1
 reixa.row = 0
 reixa.col = reixa.Cols - 2
 reixa.SelectionMode = flexSelectionByRow
 'reixa.RowSel = reixa.Rows - 1
 'reixa.ColSel = reixa.Cols - 1
 reixa.SelectionMode = flexSelectionByColumn
End Sub
Sub configurar_reixa()
  Dim vcolor As Double
  Dim i As Double
  reixa.Cols = 79
  reixa.col = 0
  reixa.row = 1
  reixa.text = "Lot-Mtrsmin"
  reixa.row = 2
  reixa.text = "Tinter 1"
  reixa.row = 3
  reixa.text = "Tinter 2"
  reixa.row = 4
  reixa.text = "Tinter 3"
  reixa.row = 5
  reixa.text = "Tinter 4"
  reixa.row = 6
  reixa.text = "Tinter 5"
  reixa.row = 7
  reixa.text = "Tinter 6"
  reixa.row = 8
  reixa.text = "Tinter 7"
  reixa.row = 9
  reixa.text = "Tinter 8"
  reixa.row = 0
  vcolor = &H80000002
  For i = 1 To reixa.Cols - 2 Step 2
    reixa.ColWidth(i) = 1200
    reixa.ColWidth(i + 1) = 600
    reixa.row = 0
    reixa.col = i
    reixa.RowSel = reixa.Rows - 1
    reixa.ColSel = i + 1
    reixa.CellBackColor = vcolor
    If vcolor = &H80000003 Then
       vcolor = &H80000002
         Else: vcolor = &H80000003
    End If
  Next i
'  reixa
End Sub
Function buscarcomandesefectades(vtreball As String) As String
   Dim rst As Recordset
   Dim vr As String
   Dim vt As String
   Dim vm As String
   framebuscant.visible = True
   DoEvents
   If InStr(1, vtreball, "/") = 0 Then
       Set rst = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(cadbl(vtreball)))
       If rst.EOF Then GoTo fi
       vtreball = atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio)
       vt = atrim(rst!numtreball)
       vm = atrim(rst!numordremodificacio)
         Else
           vt = atrim(cadbl(Mid(vtreball, 1, InStr(1, vtreball, "/") - 1)))
           vm = atrim(cadbl(Mid(vtreball, InStr(1, vtreball, "/") + 1)))
   End If
   
   'Set rst = dbtmp.OpenRecordset("select comanda from comandes where trim(numtreball)+'/'+trim(numordremodificacio)='" + vtreball + "'")
   Set rst = dbtmp.OpenRecordset("select comanda from comandes where numtreball=" + vt + " and numordremodificacio=" + vm)
   While Not rst.EOF
      vr = vr + IIf(vr <> "", "," + atrim(rst!comanda), atrim(rst!comanda))
      rst.MoveNext
   Wend
   If vr <> "" Then buscarcomandesefectades = "comanda IN(" + vr + ")"
fi:
   Set rst = Nothing
   framebuscant.visible = False
   DoEvents
End Function

Private Sub Form_Unload(Cancel As Integer)
  If Not isloaded("formtintes") Then  'com que comparteix aquest form amb Tintes trec lo que no correspon
   If formannex.etcomanda <> atrim(vcomandaactiva) Then formannex.carregarcomanda cadbl(vcomandaactiva)
   formannex.BackColor = &H8000000F
  End If
End Sub

Private Sub Image1_Click()
   beliminar_Click
End Sub

Private Sub mbuscarcomanda_Click()
   Dim vnumc As String
   framedades.Enabled = True
   vnumc = InputBox("Entra la comanda que vols buscar a la linia de temps." + Chr(10) + "O el Treball/versió ex: 12345/1 ", "Buscar comanda o versió")
  ' If InStr(1, vnumc + " ", "/") = 0 Then
  '    mbuscarcomanda.tag = IIf(cadbl(vnumc) > 0, " and comanda=" + vnumc, "")
  '     Else
   mbuscarcomanda.tag = buscarcomandesefectades(vnumc)
  ' End If
   carregar_formannex vnumc
   If vnumc = "" Then
        mbuscarcomanda.tag = ""
         Else: formcanvisanilox.Height = 5000: carregar_timeline
   End If
End Sub
Sub carregar_formannex(vnumc As String)
   Dim rst As Recordset
   Dim vnumtreball As String
   Dim vnumt As String
   If InStr(1, vnumc, "/") = 0 Then
       Set rst = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
       If Not rst.EOF Then
            vnumtreball = atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio)
             Else: GoTo fi
       End If
        Else: vnumtreball = vnumc
   End If
   'Set rst = dbtmp.OpenRecordset("SELECT impresores_ordreimpresio.comanda, [numtreball] & '/' & [numordremodificacio] AS treball FROM impresores_ordreimpresio INNER JOIN comandes ON impresores_ordreimpresio.comanda = comandes.comanda;")
   'Set rst = dbtmp.OpenRecordset("SELECT impresores_ordreimpresio.comanda, numtreball FROM impresores_ordreimpresio INNER JOIN comandes ON impresores_ordreimpresio.comanda = comandes.comanda order by comandes.comanda desc;")
   vnumt = Mid(vnumtreball, 1, InStr(1, vnumtreball, "/") - 1)
   Set rst = dbtmp.OpenRecordset("select comanda,numtreball from comandes where numtreball=" + atrim(vnumt) + " order by comanda desc")
   rst.FindFirst "numtreball=" + atrim(cadbl(vnumt))
   If Not rst.NoMatch Then
        
        formannex.carregarcomanda cadbl(rst!comanda)
        formannex.BackColor = &HF1B75F
        formcanvisanilox.Top = 500
          Else: MsgBox "Aquest treball o comanda no està a la llista d'ordre d'impresio.", vbCritical, "Error": vnumc = ""
   End If
fi:
   Set rst = Nothing
End Sub

Private Sub reixa_Click()
  Dim v As String
  Dim vhora As String
  v = reixa.TextMatrix(1, reixa.col) + "    "
  If Mid(v, 1, 3) = "Op:" Then v = reixa.TextMatrix(1, reixa.col - 1) + "    ": reixa.col = reixa.col - 1
  v = Mid(v, 1, InStr(1, v, "-") - 1)
  vhora = reixa.TextMatrix(0, reixa.col + 1)
  
  If estatcomanda(CDbl(v)) = "I" Then
        framebeliminar.Left = reixa.ColPos(reixa.col) + 150
        framebeliminar.Top = reixa.Top + 70
        framebeliminar.visible = True
        framebeliminar.BackColor = reixa.CellBackColor
         Else: framebeliminar.visible = False
  End If
End Sub
Function estatcomanda(vnumc As Double) As String
  Dim rst As Recordset
  If vnumc = 0 Then Exit Function
  Set rst = dbtmp.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then estatcomanda = rst!proximaseccio
  Set rst = Nothing
End Function
Private Sub reixa_DblClick()
  Dim v As String
  Dim vhora As String
  
  If MsgBox("Vols carregar aquestes dades?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  v = reixa.TextMatrix(1, reixa.col) + "    "
  If Mid(v, 1, 3) = "Op:" Then v = reixa.TextMatrix(1, reixa.col - 1) + "    ": reixa.col = reixa.col - 1
  v = Mid(v, 1, InStr(1, v, "-") - 1)
  vhora = reixa.TextMatrix(0, reixa.col + 1)
  If cadbl(v) > 0 Then ensenyarinformacio cadbl(v), vhora
End Sub

Private Sub reixa_GotFocus()
  etstatus.visible = True
  
End Sub

Private Sub reixa_LostFocus()
  framebeliminar.visible = False
   etstatus.visible = False
End Sub


Sub possarcolorreixadeltae(rste As Recordset, vrow As Double, vcol As Double)
  Dim vtolVerd As Double
  Dim vtolTaronja As Double
  Dim vtolMagenta As Double
  Dim vcolor As String
  Dim vini As String
  vini = rutadelfitxer(cami) + "valorsprograma.ini"
  
  If rste!color = "N" Then vcolor = "Negre"
  If rste!color = "G" Then vcolor = "Groc"
  If rste!color = "M" Then vcolor = "Magenta"
  If rste!color = "C" Then vcolor = "Cyan"
  vtolVerd = llegir_ini("ToleranciesCuatricomia", vcolor + "_toleranciaV", vini)
  vtolTaronja = llegir_ini("ToleranciesCuatricomia", vcolor + "_toleranciaT", vini)
  vtolMagenta = llegir_ini("ToleranciesCuatricomia", vcolor + "_toleranciaM", vini)
  reixadeltae.row = vrow: reixadeltae.col = vcol
  If rste!deltaE > vtolVerd Then vcolor = "V"
  If rste!deltaE > vtolTaronja Then vcolor = "T"
  If rste!deltaE > vtolMagenta Then vcolor = "M"
  If vcolor = "V" Then reixadeltae.CellBackColor = &H6BEBB1
  If vcolor = "T" Then reixadeltae.CellBackColor = &H80FF&
  If vcolor = "M" Then reixadeltae.CellBackColor = &H5C31DD
End Sub
Function nommaquina(vnummaq As Integer) As String
  nommaquina = "??"
  If vnummaq = 7 Then nommaquina = "FW"
  If vnummaq = 9 Then nommaquina = "F2"
  
End Function
Sub possar_valors_deltaE(vcomandes As String)
  Dim rst As Recordset
  Dim rste As Recordset
  Dim vrow As Double
  Dim vcolor As Double
  Dim vcol As Double
  Dim vlletracolor As String
  Dim vposlletra As Byte
  Dim vlletres As String
  vlletres = "NGMC"
  vrow = 1
  vcol = 1
  If vcomandes = "" Then Exit Sub
  Set rst = dbtintes.OpenRecordset("select comanda,numtreball,numordremodificacio from comandes where comanda in (" + vcomandes + ")")
  If Not rst.EOF Then rst.MoveLast
  While Not rst.BOF
    For vposlletra = 1 To 4
       vlletracolor = Mid(vlletres, vposlletra, 1)
       possar_reixadeltaEsobreelcolorcorrecte vlletracolor
       possar_reixadeltaEsobrelacomandacorrecte rst!comanda
       vrow = reixadeltae.row
       vcol = reixadeltae.col
        Set rste = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where nummaq=" + atrim(nummaq) + " and color='" + vlletracolor + "' and numtreballiversio='" + atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio) + "'")
        If nummaq = 9 Then If rste.EOF Then Set rste = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where nummaq=7 and color='" + vlletracolor + "' and numtreballiversio='" + atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio) + "'")
        If nummaq = 7 Then If rste.EOF Then Set rste = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where nummaq=9 and color='" + vlletracolor + "' and numtreballiversio='" + atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio) + "'")
        If rste.EOF Then Set rste = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where nummaq=1 and color='" + vlletracolor + "' and numtreballiversio='" + atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio) + "'")
        While Not rste.EOF
          reixadeltae.TextMatrix(vrow, vcol) = atrim(rste!deltaE)
          'reixadeltae.TextMatrix(0, vcol + 1) = atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio) + IIf(rste!nummaq <> nummaq, "(" + nommaquina(rste!nummaq) + ")", "")
          reixadeltae.TextMatrix(0, vcol + 1) = atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio) + " " + nommaquina(rste!nummaq)
          possarcolorreixadeltae rste, vrow, vcol
          reixadeltae.TextMatrix(vrow, vcol + 1) = atrim(rste!aniloxivolum)
          rste.MoveNext
          If Not rste.EOF Then
               reixadeltae.col = 0: reixadeltae.row = vrow
               vcolor = reixadeltae.CellBackColor
               reixadeltae.AddItem vlletra, vrow + 1
               vrow = vrow + 1
               reixadeltae.row = vrow
               reixadeltae.RowHeight(vrow) = reixadeltae.RowHeight(vrow - 1)
               reixadeltae.CellBackColor = vcolor
                Else: vrow = vrow + 1
          End If
        Wend
    Next vposlletra
    If InStr(1, UCase(App.EXEName), "MANTENIMENT TINTES") Then
         reixadeltae.TextMatrix(5, vcol) = buscar_material_comanda(rst!comanda)
         reixadeltae.TextMatrix(5, vcol + 1) = buscar_camisa_comanda(rst!comanda)
    End If
    rst.MovePrevious
    'vcol = vcol + 2
  Wend
  Set rst = Nothing
  Set rste = Nothing
End Sub
Function buscar_material_comanda(vnumc As Double) As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("SELECT familiesmaterials.descripcio FROM familiesmaterials RIGHT JOIN (comandes LEFT JOIN materials ON comandes.materialex = materials.codi) ON familiesmaterials.codi = materials.familia WHERE (((comandes.comanda)=" + atrim(vnumc) + "));")
   If Not rst.EOF Then buscar_material_comanda = atrim(rst!descripcio)
   Set rst = Nothing
End Function
Function buscar_camisa_comanda(vnumc As Double) As String
   Dim rst As Recordset
   Dim vdes As Double
   'Set rst = dbclixes.OpenRecordset("SELECT Tintes.cilindre FROM Tintes RIGHT JOIN (comandes LEFT JOIN Modificacions ON (comandes.numordremodificacio = Modificacions.ordre) AND (comandes.numtreball = Modificacions.id_treball)) ON (Tintes.ordremodificacio = Modificacions.ordre) AND (Tintes.id_treball = Modificacions.id_treball) WHERE (((comandes.comanda)=" + atrim(vnumc) + "));")
   Set rst = dbclixes.OpenRecordset("SELECT cilindres FROM comandes where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
       vdes = cadbl(rst!cilindres) / 10
       If vdes <= 56 Then buscar_camisa_comanda = "Camisa: 1"
       If vdes > 56 Then buscar_camisa_comanda = "Camisa: 2"
   End If
   Set rst = Nothing
End Function
Sub possar_reixadeltaEsobrelacomandacorrecte(vnumc As Double)
  Dim i As Double
  For i = 1 To reixadeltae.Cols
    If cadbl(reixadeltae.TextMatrix(0, i)) = vnumc Then Exit For
  Next i
  reixadeltae.col = i
End Sub
Sub possar_reixadeltaEsobreelcolorcorrecte(vlletra As String)
  Dim i As Double
  For i = 1 To reixadeltae.Rows
    If Mid(reixadeltae.TextMatrix(i, 0) + " ", 1, 1) = vlletra Then Exit For
  Next i
  reixadeltae.row = i
End Sub
Sub ensenyar_deltaE()
  Dim vcomandes As String
  config_reixa_deltaE vcomandes
  If InStr(1, formcanvisanilox.tag, "nomesdelta") <> 0 Then
       vcomandes = substituir(formcanvisanilox.tag, "nomesdelta ", "")
       possar_vcomandes_alareixa vcomandes
  End If
  FramedeltaE.tag = vcomandes
  possar_valors_deltaE vcomandes
  FramedeltaE.visible = True
End Sub
Sub possar_vcomandes_alareixa(vcomandes As String)
  Dim vcomandes2 As String
  Dim i As Byte
  Dim vnumc As String
  vcomandes2 = vcomandes
  For i = 1 To reixadeltae.Cols - 1
      vnumc = ""
      If i Mod 2 <> 0 Then
         If InStr(1, vcomandes2, ",") > 0 Then
             vnumc = Mid(vcomandes2, 1, InStr(1, vcomandes2, ",") - 1)
             vcomandes2 = Mid(vcomandes2, InStr(1, vcomandes2, ",") + 1)
              Else: vnumc = vcomandes2: vcomandes2 = ""
         End If
         reixadeltae.TextMatrix(0, i) = vnumc
      End If
  Next i
End Sub
Sub config_reixa_deltaE(vcomandes As String)
   Dim vcol As Byte
   Dim rst As Recordset
   Dim vultima As Double
   Dim vnumc As Double
   
   reixadeltae.Clear
   reixadeltae.col = 0
   reixadeltae.row = 1
   reixadeltae.ColWidth(0) = 1200
   reixadeltae.text = "Negre"
   reixadeltae.CellBackColor = &H80000012
   reixadeltae.CellForeColor = QBColor(15)
   reixadeltae.RowHeight(0) = 350
   reixadeltae.RowHeight(1) = 350
   reixadeltae.row = 2
   reixadeltae.text = "Groc"
   reixadeltae.CellBackColor = QBColor(14)
   reixadeltae.RowHeight(2) = 350
   reixadeltae.row = 3
   reixadeltae.text = "Magenta"
   reixadeltae.CellBackColor = QBColor(13)
   reixadeltae.RowHeight(3) = 350
   reixadeltae.row = 4
   reixadeltae.text = "Cyan"
   reixadeltae.CellBackColor = QBColor(11)
   reixadeltae.RowHeight(4) = 350
   reixadeltae.row = 0
   If Not isloaded("formtintes") Then vultima = ultimacomandaimpresa(nummaq)
   Set rst = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio where maquina=" + atrim(nummaq) + " order by ordre,dataprogramada")
   rst.MoveLast: rst.MoveFirst
   vcol = 1
   For i = 1 To reixadeltae.Cols - 1
     reixadeltae.col = i
     reixadeltae.CellFontBold = True
     reixadeltae.CellAlignment = 3
     If i Mod 2 = 0 Then
         reixadeltae.ColWidth(i) = 1000
          Else:
            reixadeltae.ColWidth(i) = 1500
            'If reixa.Rows > vcol Then
            If Not rst.EOF Then
              vnumc = rst!comanda
              If i = 1 Then vnumc = vultima
              If Not rst.EOF Then
                 reixadeltae.text = atrim(vnumc)
                 vcomandes = vcomandes + IIf(vcomandes <> "", ",", "") + atrim(vnumc)
                     If i > 1 Then rst.MoveNext
              End If
            'End If
            End If
            vcol = vcol + 1
     End If
   Next i
   reixadeltae.row = 0
   reixadeltae.col = 1: reixadeltae.CellBackColor = &HC0C0FF
   reixadeltae.col = 2: reixadeltae.CellBackColor = &HC0C0FF
   
   reixadeltae.col = 3: reixadeltae.CellBackColor = &H80C0FF
   reixadeltae.col = 4: reixadeltae.CellBackColor = &H80C0FF
   
   reixadeltae.col = 5: reixadeltae.CellBackColor = &HC0C0FF
   reixadeltae.col = 6: reixadeltae.CellBackColor = &HC0C0FF
   
   reixadeltae.col = 7: reixadeltae.CellBackColor = &H80C0FF
   reixadeltae.col = 8: reixadeltae.CellBackColor = &H80C0FF
   Set rst = Nothing
End Sub

Function ultimacomandaimpresa(vnummaq As Byte) As Double
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select dataimpressio,comanda as numc from impressorestot where impressora=" + atrim(vnummaq) + " order by dataimpressio desc")
  ultimacomandaimpresa = rst!numc
  Set rst = Nothing
End Function

Private Sub reixadeltae_DblClick()
  Dim vcomandes As String
  vnumc = cadbl(reixadeltae.TextMatrix(0, reixadeltae.col))
  vcomandes = FramedeltaE.tag
  If vnumc = 0 Then Exit Sub
  If isloaded("formtintes") Then
      vcomandes = substituir(vcomandes, atrim(vnumc), "0")
      formcanvisanilox.tag = "nomesdelta " + vcomandes
      ensenyar_deltaE
  End If
  'FramedeltaE.tag = vcomandes
End Sub
