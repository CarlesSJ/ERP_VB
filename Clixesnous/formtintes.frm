VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form formtintes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment de les Tintes "
   ClientHeight    =   6030
   ClientLeft      =   8070
   ClientTop       =   4650
   ClientWidth     =   14745
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   14745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framebotons2 
      Height          =   585
      Left            =   90
      TabIndex        =   28
      Top             =   -45
      Width           =   14610
      Begin VB.CommandButton bobstintes 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Obs.Tintes"
         Height          =   345
         Left            =   4260
         Style           =   1  'Graphical
         TabIndex        =   227
         Top             =   165
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Timer Timer1 
         Interval        =   900
         Left            =   330
         Top             =   255
      End
      Begin VB.CommandButton bcomandespendents 
         Height          =   360
         Left            =   13200
         Picture         =   "formtintes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   226
         ToolTipText     =   "Comandes pendents relacionades"
         Top             =   165
         Width           =   390
      End
      Begin VB.CommandButton bveurepdf 
         Height          =   360
         Left            =   12750
         Picture         =   "formtintes.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   225
         ToolTipText     =   "Veure els PDF d'aquest treball"
         Top             =   165
         Width           =   450
      End
      Begin VB.ComboBox etestatrevisiotintes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005C31DD&
         Height          =   345
         ItemData        =   "formtintes.frx":0B14
         Left            =   9495
         List            =   "formtintes.frx":0B24
         TabIndex        =   224
         Top             =   180
         Width           =   3255
      End
      Begin VB.CommandButton bokdisseny 
         Height          =   405
         Left            =   9015
         Picture         =   "formtintes.frx":0B6B
         Style           =   1  'Graphical
         TabIndex        =   223
         ToolTipText     =   "Fer OK DISSENY de la revisió de tintes."
         Top             =   135
         Width           =   420
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   2
         Left            =   1935
         Picture         =   "formtintes.frx":10F5
         Style           =   1  'Graphical
         TabIndex        =   210
         TabStop         =   0   'False
         ToolTipText     =   "Llista de canvis realitzats a les tintes."
         Top             =   150
         Width           =   315
      End
      Begin VB.CommandButton breprint 
         Caption         =   "Reprint"
         Height          =   345
         Left            =   2385
         Style           =   1  'Graphical
         TabIndex        =   198
         Top             =   165
         Width           =   1860
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   13605
         Picture         =   "formtintes.frx":167F
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Imprimir historia d'impresions amb aquesta versió de treball."
         Top             =   165
         Width           =   450
      End
      Begin VB.CommandButton copiartintes 
         Height          =   360
         Left            =   1485
         Picture         =   "formtintes.frx":1C09
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Copiar tintes d'una altra versió o treball."
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   495
         Picture         =   "formtintes.frx":2193
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Modificar Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton guardar 
         Height          =   360
         Left            =   930
         Picture         =   "formtintes.frx":271D
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   105
         Picture         =   "formtintes.frx":2CA7
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   14085
         Picture         =   "formtintes.frx":3231
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Sortir"
         Top             =   135
         Width           =   390
      End
      Begin Crystal.CrystalReport llistat 
         Left            =   1815
         Top             =   105
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label etrevtintes 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Estat revisió tintes:"
         Height          =   255
         Left            =   6240
         TabIndex        =   222
         Top             =   240
         Width           =   2760
      End
      Begin VB.Label estatedicio 
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
         Left            =   5355
         TabIndex        =   41
         Top             =   165
         Width           =   2085
      End
   End
   Begin VB.Frame fdades 
      Enabled         =   0   'False
      Height          =   5190
      Left            =   90
      TabIndex        =   0
      Top             =   465
      Width           =   14610
      Begin VB.CommandButton btintesalternatives 
         Caption         =   "+"
         Height          =   285
         Index           =   7
         Left            =   405
         Style           =   1  'Graphical
         TabIndex        =   235
         TabStop         =   0   'False
         ToolTipText     =   "Tintes alternatives també vàlides"
         Top             =   2955
         Width           =   285
      End
      Begin VB.CommandButton btintesalternatives 
         Caption         =   "+"
         Height          =   285
         Index           =   6
         Left            =   405
         Style           =   1  'Graphical
         TabIndex        =   234
         TabStop         =   0   'False
         ToolTipText     =   "Tintes alternatives també vàlides"
         Top             =   2595
         Width           =   285
      End
      Begin VB.CommandButton btintesalternatives 
         Caption         =   "+"
         Height          =   285
         Index           =   5
         Left            =   405
         Style           =   1  'Graphical
         TabIndex        =   233
         TabStop         =   0   'False
         ToolTipText     =   "Tintes alternatives també vàlides"
         Top             =   2235
         Width           =   285
      End
      Begin VB.CommandButton btintesalternatives 
         Caption         =   "+"
         Height          =   285
         Index           =   4
         Left            =   405
         Style           =   1  'Graphical
         TabIndex        =   232
         TabStop         =   0   'False
         ToolTipText     =   "Tintes alternatives també vàlides"
         Top             =   1875
         Width           =   285
      End
      Begin VB.CommandButton btintesalternatives 
         Caption         =   "+"
         Height          =   285
         Index           =   3
         Left            =   405
         Style           =   1  'Graphical
         TabIndex        =   231
         TabStop         =   0   'False
         ToolTipText     =   "Tintes alternatives també vàlides"
         Top             =   1515
         Width           =   285
      End
      Begin VB.CommandButton btintesalternatives 
         Caption         =   "+"
         Height          =   285
         Index           =   2
         Left            =   405
         Style           =   1  'Graphical
         TabIndex        =   230
         TabStop         =   0   'False
         ToolTipText     =   "Tintes alternatives també vàlides"
         Top             =   1155
         Width           =   285
      End
      Begin VB.CommandButton btintesalternatives 
         Caption         =   "+"
         Height          =   285
         Index           =   1
         Left            =   405
         Style           =   1  'Graphical
         TabIndex        =   229
         TabStop         =   0   'False
         ToolTipText     =   "Tintes alternatives també vàlides"
         Top             =   795
         Width           =   285
      End
      Begin VB.CommandButton btintesalternatives 
         Caption         =   "+"
         Height          =   285
         Index           =   0
         Left            =   405
         Style           =   1  'Graphical
         TabIndex        =   228
         TabStop         =   0   'False
         ToolTipText     =   "Tintes alternatives també vàlides"
         Top             =   435
         Width           =   285
      End
      Begin VB.ComboBox comboformaimpresio 
         Height          =   315
         ItemData        =   "formtintes.frx":37BB
         Left            =   4950
         List            =   "formtintes.frx":37C5
         TabIndex        =   199
         Top             =   105
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   3240
         Left            =   6510
         TabIndex        =   132
         Top             =   90
         Width           =   3975
         Begin VB.CommandButton buscarvolum 
            Height          =   315
            Left            =   2955
            Picture         =   "formtintes.frx":37E0
            Style           =   1  'Graphical
            TabIndex        =   212
            Top             =   345
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.TextBox tanx100 
            Height          =   315
            Index           =   7
            Left            =   3420
            TabIndex        =   209
            ToolTipText     =   "% de cobertura del clixé."
            Top             =   2880
            Width           =   540
         End
         Begin VB.TextBox tanx100 
            Height          =   315
            Index           =   6
            Left            =   3420
            TabIndex        =   208
            ToolTipText     =   "% de cobertura del clixé."
            Top             =   2517
            Width           =   540
         End
         Begin VB.TextBox tanx100 
            Height          =   315
            Index           =   5
            Left            =   3420
            TabIndex        =   207
            ToolTipText     =   "% de cobertura del clixé."
            Top             =   2155
            Width           =   540
         End
         Begin VB.TextBox tanx100 
            Height          =   315
            Index           =   4
            Left            =   3420
            TabIndex        =   206
            ToolTipText     =   "% de cobertura del clixé."
            Top             =   1793
            Width           =   540
         End
         Begin VB.TextBox tanx100 
            Height          =   315
            Index           =   3
            Left            =   3420
            TabIndex        =   205
            ToolTipText     =   "% de cobertura del clixé."
            Top             =   1431
            Width           =   540
         End
         Begin VB.TextBox tanx100 
            Height          =   315
            Index           =   2
            Left            =   3420
            TabIndex        =   204
            ToolTipText     =   "% de cobertura del clixé."
            Top             =   1069
            Width           =   540
         End
         Begin VB.TextBox tanx100 
            Height          =   315
            Index           =   1
            Left            =   3420
            TabIndex        =   203
            ToolTipText     =   "% de cobertura del clixé."
            Top             =   707
            Width           =   540
         End
         Begin VB.TextBox tanx100 
            Height          =   315
            Index           =   0
            Left            =   3420
            TabIndex        =   202
            ToolTipText     =   "% de cobertura del clixé."
            Top             =   345
            Width           =   540
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00EEE4D7&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   3015
            TabIndex        =   195
            Top             =   2880
            Width           =   375
         End
         Begin VB.TextBox volum 
            Height          =   315
            Index           =   7
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   194
            Top             =   2880
            Width           =   375
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00EEE4D7&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   3015
            TabIndex        =   193
            Top             =   2517
            Width           =   375
         End
         Begin VB.TextBox volum 
            Height          =   315
            Index           =   6
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   192
            Top             =   2517
            Width           =   375
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00EEE4D7&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   3015
            TabIndex        =   191
            Top             =   2155
            Width           =   375
         End
         Begin VB.TextBox volum 
            Height          =   315
            Index           =   5
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   190
            Top             =   2155
            Width           =   375
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00EEE4D7&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   3015
            TabIndex        =   189
            Top             =   1793
            Width           =   375
         End
         Begin VB.TextBox volum 
            Height          =   315
            Index           =   4
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   188
            Top             =   1793
            Width           =   375
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00EEE4D7&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3015
            TabIndex        =   187
            Top             =   1431
            Width           =   375
         End
         Begin VB.TextBox volum 
            Height          =   315
            Index           =   3
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   186
            Top             =   1431
            Width           =   375
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00EEE4D7&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3015
            TabIndex        =   185
            Top             =   1069
            Width           =   375
         End
         Begin VB.TextBox volum 
            Height          =   315
            Index           =   2
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   184
            Top             =   1069
            Width           =   375
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00EEE4D7&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   3015
            TabIndex        =   183
            Top             =   707
            Width           =   375
         End
         Begin VB.TextBox volum 
            Height          =   315
            Index           =   1
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   182
            Top             =   707
            Width           =   375
         End
         Begin VB.TextBox viscositat 
            BackColor       =   &H00EEE4D7&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   3015
            TabIndex        =   181
            Top             =   345
            Width           =   375
         End
         Begin VB.TextBox volum 
            Height          =   315
            Index           =   0
            Left            =   2595
            Locked          =   -1  'True
            TabIndex        =   180
            Top             =   345
            Width           =   375
         End
         Begin VB.CommandButton buscar 
            Height          =   315
            Left            =   450
            Picture         =   "formtintes.frx":3D6A
            Style           =   1  'Graphical
            TabIndex        =   147
            Top             =   330
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.TextBox anilox 
            Height          =   315
            Index           =   0
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   174
            Top             =   345
            Width           =   465
         End
         Begin VB.TextBox anilox 
            Height          =   315
            Index           =   1
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   173
            Top             =   705
            Width           =   465
         End
         Begin VB.TextBox anilox 
            Height          =   315
            Index           =   2
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   172
            Top             =   1065
            Width           =   465
         End
         Begin VB.TextBox anilox 
            Height          =   315
            Index           =   3
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   171
            Top             =   1425
            Width           =   465
         End
         Begin VB.TextBox anilox 
            Height          =   315
            Index           =   4
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   170
            Top             =   1800
            Width           =   465
         End
         Begin VB.TextBox anilox 
            Height          =   315
            Index           =   5
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   169
            Top             =   2160
            Width           =   465
         End
         Begin VB.TextBox anilox 
            Height          =   315
            Index           =   6
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   168
            Top             =   2520
            Width           =   465
         End
         Begin VB.TextBox anilox 
            Height          =   315
            Index           =   7
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   167
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox cilindre 
            Height          =   315
            Index           =   0
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   166
            Top             =   345
            Width           =   465
         End
         Begin VB.TextBox cilindre 
            Height          =   315
            Index           =   1
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   165
            Top             =   705
            Width           =   465
         End
         Begin VB.TextBox cilindre 
            Height          =   315
            Index           =   2
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   164
            Top             =   1065
            Width           =   465
         End
         Begin VB.TextBox cilindre 
            Height          =   315
            Index           =   3
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   163
            Top             =   1425
            Width           =   465
         End
         Begin VB.TextBox cilindre 
            Height          =   315
            Index           =   4
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   162
            Top             =   1800
            Width           =   465
         End
         Begin VB.TextBox cilindre 
            Height          =   315
            Index           =   5
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   161
            Top             =   2160
            Width           =   465
         End
         Begin VB.TextBox cilindre 
            Height          =   315
            Index           =   6
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   160
            Top             =   2520
            Width           =   465
         End
         Begin VB.TextBox cilindre 
            Height          =   315
            Index           =   7
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   159
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox desarroll 
            Height          =   315
            Index           =   0
            Left            =   2085
            Locked          =   -1  'True
            TabIndex        =   158
            Top             =   345
            Width           =   465
         End
         Begin VB.TextBox desarroll 
            Height          =   315
            Index           =   1
            Left            =   2085
            Locked          =   -1  'True
            TabIndex        =   157
            Top             =   705
            Width           =   465
         End
         Begin VB.TextBox desarroll 
            Height          =   315
            Index           =   2
            Left            =   2085
            Locked          =   -1  'True
            TabIndex        =   156
            Top             =   1065
            Width           =   465
         End
         Begin VB.TextBox desarroll 
            Height          =   315
            Index           =   3
            Left            =   2085
            Locked          =   -1  'True
            TabIndex        =   155
            Top             =   1425
            Width           =   465
         End
         Begin VB.TextBox desarroll 
            Height          =   315
            Index           =   4
            Left            =   2085
            Locked          =   -1  'True
            TabIndex        =   154
            Top             =   1800
            Width           =   465
         End
         Begin VB.TextBox desarroll 
            Height          =   315
            Index           =   5
            Left            =   2085
            Locked          =   -1  'True
            TabIndex        =   153
            Top             =   2160
            Width           =   465
         End
         Begin VB.TextBox desarroll 
            Height          =   315
            Index           =   6
            Left            =   2085
            Locked          =   -1  'True
            TabIndex        =   152
            Top             =   2520
            Width           =   465
         End
         Begin VB.TextBox desarroll 
            Height          =   315
            Index           =   7
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   151
            Top             =   2880
            Width           =   465
         End
         Begin VB.CommandButton buscarcilindre 
            Height          =   315
            Left            =   2055
            Picture         =   "formtintes.frx":42F4
            Style           =   1  'Graphical
            TabIndex        =   150
            Top             =   345
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.TextBox densitat 
            Height          =   315
            Index           =   0
            Left            =   525
            TabIndex        =   149
            Top             =   345
            Width           =   465
         End
         Begin VB.TextBox aniloxfotogravador 
            Height          =   315
            Index           =   0
            Left            =   1035
            TabIndex        =   148
            Top             =   345
            Width           =   465
         End
         Begin VB.TextBox aniloxfotogravador 
            Height          =   315
            Index           =   1
            Left            =   1035
            TabIndex        =   146
            Top             =   705
            Width           =   465
         End
         Begin VB.TextBox aniloxfotogravador 
            Height          =   315
            Index           =   2
            Left            =   1035
            TabIndex        =   145
            Top             =   1065
            Width           =   465
         End
         Begin VB.TextBox aniloxfotogravador 
            Height          =   315
            Index           =   3
            Left            =   1035
            TabIndex        =   144
            Top             =   1425
            Width           =   465
         End
         Begin VB.TextBox aniloxfotogravador 
            Height          =   315
            Index           =   4
            Left            =   1035
            TabIndex        =   143
            Top             =   1800
            Width           =   465
         End
         Begin VB.TextBox aniloxfotogravador 
            Height          =   315
            Index           =   5
            Left            =   1035
            TabIndex        =   142
            Top             =   2160
            Width           =   465
         End
         Begin VB.TextBox aniloxfotogravador 
            Height          =   315
            Index           =   6
            Left            =   1035
            TabIndex        =   141
            Top             =   2520
            Width           =   465
         End
         Begin VB.TextBox aniloxfotogravador 
            Height          =   315
            Index           =   7
            Left            =   1035
            TabIndex        =   140
            Top             =   2880
            Width           =   465
         End
         Begin VB.TextBox densitat 
            Height          =   315
            Index           =   1
            Left            =   525
            TabIndex        =   139
            Top             =   705
            Width           =   465
         End
         Begin VB.TextBox densitat 
            Height          =   315
            Index           =   2
            Left            =   525
            TabIndex        =   138
            Top             =   1065
            Width           =   465
         End
         Begin VB.TextBox densitat 
            Height          =   315
            Index           =   3
            Left            =   525
            TabIndex        =   137
            Top             =   1425
            Width           =   465
         End
         Begin VB.TextBox densitat 
            Height          =   315
            Index           =   4
            Left            =   525
            TabIndex        =   136
            Top             =   1800
            Width           =   465
         End
         Begin VB.TextBox densitat 
            Height          =   315
            Index           =   5
            Left            =   525
            TabIndex        =   135
            Top             =   2160
            Width           =   465
         End
         Begin VB.TextBox densitat 
            Height          =   315
            Index           =   6
            Left            =   525
            TabIndex        =   134
            Top             =   2520
            Width           =   465
         End
         Begin VB.TextBox densitat 
            Height          =   315
            Index           =   7
            Left            =   525
            TabIndex        =   133
            Top             =   2880
            Width           =   465
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "%Cob"
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
            Left            =   3465
            TabIndex        =   201
            ToolTipText     =   "% de cobertura del clixé."
            Top             =   150
            Width           =   630
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Adhesiu"
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
            Left            =   2910
            TabIndex        =   197
            Top             =   0
            Width           =   630
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Volum"
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
            Left            =   2550
            TabIndex        =   196
            Top             =   105
            Width           =   465
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Anilox"
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
            Left            =   0
            TabIndex        =   179
            Top             =   75
            Width           =   630
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Cilin/Desar."
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
            Left            =   1725
            TabIndex        =   178
            Top             =   105
            Width           =   945
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Densitat"
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
            Left            =   480
            TabIndex        =   177
            Top             =   105
            Width           =   630
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Liniatura"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1050
            TabIndex        =   176
            Top             =   -15
            Width           =   1020
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Treball"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1095
            TabIndex        =   175
            Top             =   180
            Width           =   405
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   3165
         Left            =   10275
         TabIndex        =   60
         Top             =   105
         Width           =   3825
         Begin VB.CheckBox modificatperinplacsa 
            Height          =   210
            Index           =   7
            Left            =   1890
            TabIndex        =   220
            ToolTipText     =   "Clixé modificat per inplacsa."
            Top             =   2925
            Width           =   225
         End
         Begin VB.CheckBox modificatperinplacsa 
            Height          =   210
            Index           =   6
            Left            =   1890
            TabIndex        =   219
            ToolTipText     =   "Clixé modificat per inplacsa."
            Top             =   2562
            Width           =   225
         End
         Begin VB.CheckBox modificatperinplacsa 
            Height          =   210
            Index           =   5
            Left            =   1890
            TabIndex        =   218
            ToolTipText     =   "Clixé modificat per inplacsa."
            Top             =   2200
            Width           =   225
         End
         Begin VB.CheckBox modificatperinplacsa 
            Height          =   210
            Index           =   4
            Left            =   1890
            TabIndex        =   217
            ToolTipText     =   "Clixé modificat per inplacsa."
            Top             =   1838
            Width           =   225
         End
         Begin VB.CheckBox modificatperinplacsa 
            Height          =   210
            Index           =   3
            Left            =   1890
            TabIndex        =   216
            ToolTipText     =   "Clixé modificat per inplacsa."
            Top             =   1476
            Width           =   225
         End
         Begin VB.CheckBox modificatperinplacsa 
            Height          =   210
            Index           =   2
            Left            =   1890
            TabIndex        =   215
            ToolTipText     =   "Clixé modificat per inplacsa."
            Top             =   1114
            Width           =   225
         End
         Begin VB.CheckBox modificatperinplacsa 
            Height          =   210
            Index           =   1
            Left            =   1890
            TabIndex        =   214
            ToolTipText     =   "Clixé modificat per inplacsa."
            Top             =   752
            Width           =   225
         End
         Begin VB.CheckBox modificatperinplacsa 
            Height          =   210
            Index           =   0
            Left            =   1890
            TabIndex        =   213
            ToolTipText     =   "Clixé modificat per inplacsa."
            Top             =   390
            Width           =   225
         End
         Begin VB.CheckBox continuu 
            Caption         =   "Check1"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   124
            Top             =   330
            Width           =   225
         End
         Begin VB.CheckBox continuu 
            Caption         =   "Check1"
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   123
            Top             =   690
            Width           =   225
         End
         Begin VB.CheckBox continuu 
            Caption         =   "Check1"
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   122
            Top             =   1050
            Width           =   225
         End
         Begin VB.CheckBox continuu 
            Caption         =   "Check1"
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   121
            Top             =   1410
            Width           =   225
         End
         Begin VB.CheckBox continuu 
            Caption         =   "Check1"
            Height          =   315
            Index           =   4
            Left            =   240
            TabIndex        =   120
            Top             =   1785
            Width           =   225
         End
         Begin VB.CheckBox continuu 
            Caption         =   "Check1"
            Height          =   315
            Index           =   5
            Left            =   240
            TabIndex        =   119
            Top             =   2145
            Width           =   225
         End
         Begin VB.CheckBox continuu 
            Caption         =   "Check1"
            Height          =   315
            Index           =   6
            Left            =   240
            TabIndex        =   118
            Top             =   2505
            Width           =   225
         End
         Begin VB.CheckBox continuu 
            Caption         =   "Check1"
            Height          =   315
            Index           =   7
            Left            =   240
            TabIndex        =   117
            Top             =   2865
            Width           =   225
         End
         Begin VB.ComboBox clixeosleeve 
            Height          =   315
            Index           =   0
            ItemData        =   "formtintes.frx":487E
            Left            =   465
            List            =   "formtintes.frx":4888
            TabIndex        =   116
            Top             =   330
            Width           =   900
         End
         Begin VB.ComboBox clixeosleeve 
            Height          =   315
            Index           =   1
            ItemData        =   "formtintes.frx":489B
            Left            =   465
            List            =   "formtintes.frx":48A5
            TabIndex        =   115
            Top             =   690
            Width           =   900
         End
         Begin VB.ComboBox clixeosleeve 
            Height          =   315
            Index           =   2
            ItemData        =   "formtintes.frx":48B8
            Left            =   465
            List            =   "formtintes.frx":48C2
            TabIndex        =   114
            Top             =   1050
            Width           =   900
         End
         Begin VB.ComboBox clixeosleeve 
            Height          =   315
            Index           =   3
            ItemData        =   "formtintes.frx":48D5
            Left            =   465
            List            =   "formtintes.frx":48DF
            TabIndex        =   113
            Top             =   1410
            Width           =   900
         End
         Begin VB.ComboBox clixeosleeve 
            Height          =   315
            Index           =   4
            ItemData        =   "formtintes.frx":48F2
            Left            =   465
            List            =   "formtintes.frx":48FC
            TabIndex        =   112
            Top             =   1785
            Width           =   900
         End
         Begin VB.ComboBox clixeosleeve 
            Height          =   315
            Index           =   5
            ItemData        =   "formtintes.frx":490F
            Left            =   465
            List            =   "formtintes.frx":4919
            TabIndex        =   111
            Top             =   2145
            Width           =   900
         End
         Begin VB.ComboBox clixeosleeve 
            Height          =   315
            Index           =   6
            ItemData        =   "formtintes.frx":492C
            Left            =   465
            List            =   "formtintes.frx":4936
            TabIndex        =   110
            Top             =   2505
            Width           =   900
         End
         Begin VB.ComboBox clixeosleeve 
            Height          =   315
            Index           =   7
            ItemData        =   "formtintes.frx":4949
            Left            =   465
            List            =   "formtintes.frx":4953
            TabIndex        =   109
            Top             =   2865
            Width           =   900
         End
         Begin VB.CheckBox afectatspelcanvi 
            Height          =   210
            Index           =   0
            Left            =   2505
            TabIndex        =   108
            Top             =   390
            Width           =   225
         End
         Begin VB.CheckBox afectatspelcanvi 
            Height          =   210
            Index           =   1
            Left            =   2505
            TabIndex        =   107
            Top             =   752
            Width           =   225
         End
         Begin VB.CheckBox afectatspelcanvi 
            Height          =   210
            Index           =   2
            Left            =   2505
            TabIndex        =   106
            Top             =   1114
            Width           =   225
         End
         Begin VB.CheckBox afectatspelcanvi 
            Height          =   210
            Index           =   3
            Left            =   2505
            TabIndex        =   105
            Top             =   1476
            Width           =   225
         End
         Begin VB.CheckBox afectatspelcanvi 
            Height          =   210
            Index           =   4
            Left            =   2505
            TabIndex        =   104
            Top             =   1838
            Width           =   225
         End
         Begin VB.CheckBox afectatspelcanvi 
            Height          =   210
            Index           =   5
            Left            =   2505
            TabIndex        =   103
            Top             =   2200
            Width           =   225
         End
         Begin VB.CheckBox afectatspelcanvi 
            Height          =   210
            Index           =   6
            Left            =   2505
            TabIndex        =   102
            Top             =   2562
            Width           =   225
         End
         Begin VB.CheckBox afectatspelcanvi 
            Height          =   210
            Index           =   7
            Left            =   2505
            TabIndex        =   101
            Top             =   2925
            Width           =   225
         End
         Begin VB.TextBox polimers 
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   100
            Top             =   330
            Width           =   330
         End
         Begin VB.TextBox polimers 
            Height          =   315
            Index           =   1
            Left            =   1500
            TabIndex        =   99
            Top             =   690
            Width           =   330
         End
         Begin VB.TextBox polimers 
            Height          =   315
            Index           =   2
            Left            =   1500
            TabIndex        =   98
            Top             =   1050
            Width           =   330
         End
         Begin VB.TextBox polimers 
            Height          =   315
            Index           =   3
            Left            =   1500
            TabIndex        =   97
            Top             =   1410
            Width           =   330
         End
         Begin VB.TextBox polimers 
            Height          =   315
            Index           =   4
            Left            =   1500
            TabIndex        =   96
            Top             =   1785
            Width           =   330
         End
         Begin VB.TextBox polimers 
            Height          =   315
            Index           =   5
            Left            =   1500
            TabIndex        =   95
            Top             =   2145
            Width           =   330
         End
         Begin VB.TextBox polimers 
            Height          =   315
            Index           =   6
            Left            =   1500
            TabIndex        =   94
            Top             =   2505
            Width           =   330
         End
         Begin VB.TextBox polimers 
            Height          =   315
            Index           =   7
            Left            =   1500
            TabIndex        =   93
            Top             =   2865
            Width           =   330
         End
         Begin VB.CheckBox compartit 
            Height          =   210
            Index           =   0
            Left            =   2820
            TabIndex        =   92
            Top             =   390
            Width           =   225
         End
         Begin VB.CheckBox compartit 
            Height          =   210
            Index           =   1
            Left            =   2820
            TabIndex        =   91
            Top             =   752
            Width           =   225
         End
         Begin VB.CheckBox compartit 
            Height          =   210
            Index           =   2
            Left            =   2820
            TabIndex        =   90
            Top             =   1114
            Width           =   225
         End
         Begin VB.CheckBox compartit 
            Height          =   210
            Index           =   3
            Left            =   2820
            TabIndex        =   89
            Top             =   1476
            Width           =   225
         End
         Begin VB.CheckBox compartit 
            Height          =   210
            Index           =   4
            Left            =   2820
            TabIndex        =   88
            Top             =   1838
            Width           =   225
         End
         Begin VB.CheckBox compartit 
            Height          =   210
            Index           =   5
            Left            =   2820
            TabIndex        =   87
            Top             =   2200
            Width           =   225
         End
         Begin VB.CheckBox compartit 
            Height          =   210
            Index           =   6
            Left            =   2820
            TabIndex        =   86
            Top             =   2562
            Width           =   225
         End
         Begin VB.CheckBox compartit 
            Height          =   210
            Index           =   7
            Left            =   2820
            TabIndex        =   85
            Top             =   2925
            Width           =   225
         End
         Begin VB.CheckBox portasang 
            Height          =   210
            Index           =   0
            Left            =   2190
            TabIndex        =   84
            Top             =   390
            Width           =   225
         End
         Begin VB.CheckBox portasang 
            Height          =   210
            Index           =   1
            Left            =   2190
            TabIndex        =   83
            Top             =   752
            Width           =   225
         End
         Begin VB.CheckBox portasang 
            Height          =   210
            Index           =   2
            Left            =   2190
            TabIndex        =   82
            Top             =   1114
            Width           =   225
         End
         Begin VB.CheckBox portasang 
            Height          =   210
            Index           =   3
            Left            =   2190
            TabIndex        =   81
            Top             =   1476
            Width           =   225
         End
         Begin VB.CheckBox portasang 
            Height          =   210
            Index           =   4
            Left            =   2190
            TabIndex        =   80
            Top             =   1838
            Width           =   225
         End
         Begin VB.CheckBox portasang 
            Height          =   210
            Index           =   5
            Left            =   2205
            TabIndex        =   79
            Top             =   2200
            Width           =   225
         End
         Begin VB.CheckBox portasang 
            Height          =   210
            Index           =   6
            Left            =   2190
            TabIndex        =   78
            Top             =   2562
            Width           =   225
         End
         Begin VB.CheckBox portasang 
            Height          =   210
            Index           =   7
            Left            =   2190
            TabIndex        =   77
            Top             =   2925
            Width           =   225
         End
         Begin VB.CheckBox macula 
            Caption         =   "Check1"
            Height          =   210
            Index           =   0
            Left            =   3120
            TabIndex        =   76
            Top             =   390
            Width           =   225
         End
         Begin VB.CheckBox arrastre 
            Caption         =   "Check1"
            Height          =   210
            Index           =   0
            Left            =   3435
            TabIndex        =   75
            Top             =   390
            Width           =   225
         End
         Begin VB.CheckBox macula 
            Caption         =   "Check1"
            Height          =   210
            Index           =   1
            Left            =   3120
            TabIndex        =   74
            Top             =   752
            Width           =   225
         End
         Begin VB.CheckBox macula 
            Caption         =   "Check1"
            Height          =   210
            Index           =   2
            Left            =   3120
            TabIndex        =   73
            Top             =   1114
            Width           =   225
         End
         Begin VB.CheckBox macula 
            Caption         =   "Check1"
            Height          =   210
            Index           =   3
            Left            =   3120
            TabIndex        =   72
            Top             =   1476
            Width           =   225
         End
         Begin VB.CheckBox macula 
            Caption         =   "Check1"
            Height          =   210
            Index           =   4
            Left            =   3120
            TabIndex        =   71
            Top             =   1838
            Width           =   225
         End
         Begin VB.CheckBox macula 
            Caption         =   "Check1"
            Height          =   210
            Index           =   5
            Left            =   3120
            TabIndex        =   70
            Top             =   2200
            Width           =   225
         End
         Begin VB.CheckBox macula 
            Caption         =   "Check1"
            Height          =   210
            Index           =   6
            Left            =   3120
            TabIndex        =   69
            Top             =   2562
            Width           =   225
         End
         Begin VB.CheckBox macula 
            Caption         =   "Check1"
            Height          =   210
            Index           =   7
            Left            =   3120
            TabIndex        =   68
            Top             =   2925
            Width           =   225
         End
         Begin VB.CheckBox arrastre 
            Caption         =   "Check1"
            Height          =   210
            Index           =   1
            Left            =   3435
            TabIndex        =   67
            Top             =   752
            Width           =   225
         End
         Begin VB.CheckBox arrastre 
            Caption         =   "Check1"
            Height          =   210
            Index           =   2
            Left            =   3435
            TabIndex        =   66
            Top             =   1114
            Width           =   225
         End
         Begin VB.CheckBox arrastre 
            Caption         =   "Check1"
            Height          =   210
            Index           =   3
            Left            =   3435
            TabIndex        =   65
            Top             =   1476
            Width           =   225
         End
         Begin VB.CheckBox arrastre 
            Caption         =   "Check1"
            Height          =   210
            Index           =   4
            Left            =   3435
            TabIndex        =   64
            Top             =   1838
            Width           =   225
         End
         Begin VB.CheckBox arrastre 
            Caption         =   "Check1"
            Height          =   210
            Index           =   5
            Left            =   3435
            TabIndex        =   63
            Top             =   2200
            Width           =   225
         End
         Begin VB.CheckBox arrastre 
            Caption         =   "Check1"
            Height          =   210
            Index           =   6
            Left            =   3435
            TabIndex        =   62
            Top             =   2562
            Width           =   225
         End
         Begin VB.CheckBox arrastre 
            Caption         =   "Check1"
            Height          =   210
            Index           =   7
            Left            =   3435
            TabIndex        =   61
            Top             =   2925
            Width           =   225
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            X1              =   3525
            X2              =   3525
            Y1              =   180
            Y2              =   3075
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            X1              =   3210
            X2              =   3210
            Y1              =   300
            Y2              =   3075
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            X1              =   2910
            X2              =   2910
            Y1              =   195
            Y2              =   3090
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            X1              =   2595
            X2              =   2595
            Y1              =   300
            Y2              =   3075
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            X1              =   2280
            X2              =   2280
            Y1              =   150
            Y2              =   3045
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderWidth     =   2
            X1              =   1980
            X2              =   1980
            Y1              =   330
            Y2              =   3030
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Modificat"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1770
            TabIndex        =   221
            Top             =   165
            Width           =   570
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Continuu"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   225
            TabIndex        =   211
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Clixé/Sleeve"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   495
            TabIndex        =   131
            Top             =   15
            Width           =   1185
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Polimers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1470
            TabIndex        =   130
            Top             =   45
            Width           =   600
         End
         Begin VB.Label Label8 
            Caption         =   "Afectats"
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
            Left            =   2400
            TabIndex        =   129
            Top             =   150
            Width           =   525
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Compartit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2685
            TabIndex        =   128
            Top             =   15
            Width           =   975
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Sang"
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
            Left            =   2160
            TabIndex        =   127
            Top             =   -15
            Width           =   660
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Màcula"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3030
            TabIndex        =   126
            Top             =   165
            Width           =   510
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Arrastre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3345
            TabIndex        =   125
            Top             =   15
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Observacions de les tintes"
         Height          =   1815
         Left            =   75
         TabIndex        =   57
         Top             =   3255
         Width           =   14310
         Begin VB.CommandButton alta 
            Height          =   360
            Left            =   60
            Picture         =   "formtintes.frx":4966
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Alta treball"
            Top             =   240
            Width           =   375
         End
         Begin VB.Data dataliniesobs 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   11385
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "select * from tintes_observacions "
            Top             =   840
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSDBGrid.DBGrid reixa 
            Bindings        =   "formtintes.frx":4EF0
            Height          =   1545
            Left            =   450
            OleObjectBlob   =   "formtintes.frx":4F08
            TabIndex        =   58
            Top             =   225
            Width           =   13755
         End
      End
      Begin VB.CheckBox creimpres 
         Caption         =   "Reimprès"
         Height          =   210
         Left            =   4335
         TabIndex        =   56
         Top             =   150
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton buscartinta 
         Height          =   315
         Left            =   4515
         Picture         =   "formtintes.frx":5745
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   405
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox color 
         Height          =   315
         Index           =   7
         Left            =   690
         TabIndex        =   16
         Top             =   2940
         Width           =   4140
      End
      Begin VB.TextBox color 
         Height          =   315
         Index           =   6
         Left            =   690
         TabIndex        =   15
         Top             =   2580
         Width           =   4140
      End
      Begin VB.TextBox color 
         Height          =   315
         Index           =   5
         Left            =   705
         TabIndex        =   14
         Top             =   2220
         Width           =   4140
      End
      Begin VB.TextBox color 
         Height          =   315
         Index           =   4
         Left            =   690
         TabIndex        =   13
         Top             =   1860
         Width           =   4140
      End
      Begin VB.TextBox color 
         Height          =   315
         Index           =   3
         Left            =   690
         TabIndex        =   12
         Top             =   1485
         Width           =   4140
      End
      Begin VB.TextBox color 
         Height          =   315
         Index           =   2
         Left            =   690
         TabIndex        =   11
         Top             =   1125
         Width           =   4140
      End
      Begin VB.TextBox color 
         Height          =   315
         Index           =   1
         Left            =   690
         TabIndex        =   10
         Top             =   765
         Width           =   4140
      End
      Begin VB.TextBox color 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   690
         TabIndex        =   9
         Top             =   405
         Width           =   4140
      End
      Begin VB.CommandButton buscardetalltinter 
         Height          =   315
         Left            =   6135
         Picture         =   "formtintes.frx":5CCF
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   405
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox detalltinter 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   7
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   2955
         Width           =   1530
      End
      Begin VB.TextBox detalltinter 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   6
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2595
         Width           =   1530
      End
      Begin VB.TextBox detalltinter 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   5
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2235
         Width           =   1530
      End
      Begin VB.TextBox detalltinter 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   4
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1875
         Width           =   1530
      End
      Begin VB.TextBox detalltinter 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   3
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1500
         Width           =   1530
      End
      Begin VB.TextBox detalltinter 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   2
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1140
         Width           =   1530
      End
      Begin VB.TextBox detalltinter 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   1
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   780
         Width           =   1530
      End
      Begin VB.TextBox detalltinter 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   0
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   420
         Width           =   1530
      End
      Begin VB.CommandButton linkat 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   7
         Left            =   14115
         MaskColor       =   &H00FFFFFF&
         Picture         =   "formtintes.frx":6259
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2940
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton linkat 
         BackColor       =   &H00C78DFA&
         Height          =   315
         Index           =   6
         Left            =   14115
         MaskColor       =   &H00FFFFFF&
         Picture         =   "formtintes.frx":67E3
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2577
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton linkat 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   5
         Left            =   14115
         MaskColor       =   &H00FFFFFF&
         Picture         =   "formtintes.frx":6D6D
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2215
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton linkat 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   4
         Left            =   14115
         MaskColor       =   &H00FFFFFF&
         Picture         =   "formtintes.frx":72F7
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1853
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton linkat 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   3
         Left            =   14115
         MaskColor       =   &H00FFFFFF&
         Picture         =   "formtintes.frx":7881
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1491
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton linkat 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   2
         Left            =   14115
         MaskColor       =   &H00FFFFFF&
         Picture         =   "formtintes.frx":7E0B
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1129
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton linkat 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   1
         Left            =   14115
         MaskColor       =   &H00FFFFFF&
         Picture         =   "formtintes.frx":8395
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   767
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton linkat 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Index           =   0
         Left            =   14115
         MaskColor       =   &H00FFFFFF&
         Picture         =   "formtintes.frx":891F
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   405
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox observacions 
         Height          =   315
         Index           =   7
         Left            =   13950
         MaxLength       =   30
         TabIndex        =   26
         Top             =   2940
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox observacions 
         Height          =   315
         Index           =   6
         Left            =   13950
         MaxLength       =   30
         TabIndex        =   25
         Top             =   2577
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox observacions 
         Height          =   315
         Index           =   5
         Left            =   13950
         MaxLength       =   30
         TabIndex        =   24
         Top             =   2215
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox observacions 
         Height          =   315
         Index           =   4
         Left            =   13950
         MaxLength       =   30
         TabIndex        =   23
         Top             =   1853
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox observacions 
         Height          =   315
         Index           =   3
         Left            =   13950
         MaxLength       =   30
         TabIndex        =   22
         Top             =   1491
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox observacions 
         Height          =   315
         Index           =   2
         Left            =   13950
         MaxLength       =   30
         TabIndex        =   21
         Top             =   1129
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox observacions 
         Height          =   315
         Index           =   1
         Left            =   13950
         MaxLength       =   30
         TabIndex        =   20
         Top             =   767
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox observacions 
         Height          =   315
         Index           =   0
         Left            =   13950
         MaxLength       =   30
         TabIndex        =   19
         Top             =   405
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox ordre 
         Height          =   315
         Index           =   7
         Left            =   75
         TabIndex        =   8
         Top             =   2940
         Width           =   330
      End
      Begin VB.TextBox ordre 
         Height          =   315
         Index           =   6
         Left            =   75
         TabIndex        =   7
         Top             =   2580
         Width           =   330
      End
      Begin VB.TextBox ordre 
         Height          =   315
         Index           =   5
         Left            =   75
         TabIndex        =   6
         Top             =   2205
         Width           =   330
      End
      Begin VB.TextBox ordre 
         Height          =   315
         Index           =   4
         Left            =   75
         TabIndex        =   5
         Top             =   1845
         Width           =   330
      End
      Begin VB.TextBox ordre 
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   4
         Top             =   1485
         Width           =   330
      End
      Begin VB.TextBox ordre 
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   3
         Top             =   1125
         Width           =   330
      End
      Begin VB.TextBox ordre 
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   2
         Top             =   750
         Width           =   330
      End
      Begin VB.TextBox ordre 
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   405
         Width           =   330
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Impresió"
         Height          =   225
         Left            =   3840
         TabIndex        =   200
         Top             =   150
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label18 
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
         Height          =   210
         Left            =   2805
         TabIndex        =   54
         Top             =   135
         Width           =   960
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
         Height          =   225
         Left            =   14205
         TabIndex        =   43
         Top             =   165
         Width           =   360
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacions"
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
         Left            =   13875
         TabIndex        =   27
         Top             =   135
         Visible         =   0   'False
         Width           =   210
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
         Left            =   960
         TabIndex        =   18
         Top             =   135
         Width           =   1590
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tinter"
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
         Left            =   135
         TabIndex        =   17
         Top             =   135
         Width           =   600
      End
   End
End
Attribute VB_Name = "formtintes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const COLORFONDOLINKAT = &H80FFFF
Private Const COLORFONDOLINKATnomesXL = &HC78DFA
Dim vcanvisrevisioitintes As String


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

Private Sub bcomandespendents_Click()
'""
 Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select comanda,datacomanda from comandes where (producte<>'PC' and producte<>'PCP' and producte<>'PC2') and numtreball=" + atrim(id_treball) + " and (proximaseccio <>'T') order by comanda Desc"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.width = 6000
   'formseleccio.DBGrid2.Width = formseleccio.DBGrid2.Width + 1000
   formseleccio.DBGrid2.Columns(0).width = 2000
   formseleccio.DBGrid2.Columns(1).width = 2000
   'formseleccio.DBGrid2.Columns(2).Width = 1000
   'formseleccio.DBGrid2.Columns(3).Width = 1000
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
      formclixes.cridarcomandes formseleccio.Data1.Recordset!comanda
   End If
   Unload formseleccio
End Sub

Private Sub bobstintes_Click()
   Dim v As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from  tintes_observacions where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(900)) '+ ordremodificacio))
   If Not rst.EOF Then v = rst!observacio
   v = InputBox("Escriu la observació de tintes, aquesta informació surtirà només a la revisió de comandes del Manteniment de tintes al peu.", "Observació", v)
   If StrPtr(v) = 0 Then Exit Sub
   If Len(v) > 255 Then MsgBox "Aquest missatge es massa llarg, es retallarà a 255 caràcters", vbCritical, "Error"
   v = Mid(v + " ", 1, 255)
   If atrim(v) = "" Then dbclixes.Execute "delete * from tintes_observacions where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(900): GoTo fi ' + ordremodificacio): GoTo fi
   If rst.EOF Then
       rst.AddNew
       rst!id_treball = id_treball
       rst!ordre = 900 ' + ordremodificacio
         Else: rst.Edit
   End If
   rst!observacio = atrim(v)
   rst.Update
fi:
   Set rst = Nothing
End Sub

Private Sub bokdisseny_Click()
   If estatedicio <> "" Then MsgBox "NO POTS ACCEPTAR LA REVISIÓ SI ESTÀS EDITANT EL REGISTRE.", vbCritical, "ERROR": Exit Sub
   If MsgBox("Segur que vols donar l'OK " + atrim(bokdisseny.tag) + " a aquesta revisió de tintes?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   dbclixes.Execute "insert into tintes_controlcanvis (numtreball,usuari,campafectat,valoranterior,valoractual) values ('" + atrim(id_treball) + "/" + atrim(ordremodificacio) + "','" + nomordinador + "','EstatREVtintes','" + atrim(formclixes.modificacions.Recordset!estatrevisiotintes) + "','" + atrim(formclixes.modificacions.Recordset!estatrevisiotintes) + bokdisseny.tag + "')"
   formclixes.modificacions.Recordset.Edit
   formclixes.modificacions.Recordset!estatrevisiotintes = formclixes.modificacions.Recordset!estatrevisiotintes + bokdisseny.tag
   formclixes.modificacions.Recordset.Update
   
   Unload Me
End Sub

Private Sub breprint_Click()
   formtintes.tag = "reprint"
   formtintes.Hide
   
End Sub

Private Sub btintesalternatives_Click(Index As Integer)
  Unload formaltarep
  If color(Index) = "" Then Exit Sub
  Load formaltarep
  
  formaltarep.caption = "Manteniment Tintes alternatives"
  formaltarep.Data1.DatabaseName = rutadelfitxer(cami) + "clixesnous.mdb"
  
  'formaltarep.DBGrid1.Columns(1).width = formaltarep.DBGrid1.Columns(1).width * 2
  'formaltarep.width = formaltarep.width - 1800
  formaltarep.Data1.RecordSource = "select id_tinter,coditinta,color from tintes_alternatives where id_tinter=" + atrim(ordre(Index).tag)
  formaltarep.Data1.Refresh
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).visible = False
  formaltarep.DBGrid1.Columns(1).width = 1000
  formaltarep.DBGrid1.Columns(2).width = 4000
  formaltarep.width = 7500
  formaltarep.etcoditintaalternativa = color(Index).ToolTipText
  formaltarep.tag = ordre(Index).tag
  formaltarep.Show 1
  posarcolors_tintesalternatives
End Sub
Sub eliminar_tintesalternatives(vindex As Long)
  Dim rst As Recordset
  Set rst = dbclixes.OpenRecordset("select id_tinter,coditinta,color from tintes_alternatives where id_tinter=" + atrim(ordre(vindex).tag))
  If Not rst.EOF Then
     dbclixes.Execute "delete * from tintes_alternatives where id_tinter=" + atrim(ordre(vindex).tag)
     posarcolors_tintesalternatives
     MsgBox "Al fer canvi de tinta també s'ha eliminat totes les tintes alternatives.", vbInformation, "ATENCIÓ"
  End If
  Set rst = Nothing
End Sub
Private Sub buscardetalltinter_Click()
  triardetalltinter
End Sub

Private Sub buscarvolum_Click()
 Dim i As Byte
  triarvolum
  buscar.visible = False
  'If cadbl(buscar.tag) = 0 Then
  '    If MsgBox("Vols copiar aquest anilox a totes les altres tintes?", vbInformation + vbYesNo + vbDefaultButton2, "Copiar anilox") = vbYes Then
  '       For i = 1 To 7
  '         If atrim(color(i)) <> "" Then
  '            anilox(i) = anilox(0)
  '           Else: anilox(i) = "0"
  '         End If
  '       Next i
  '    End If
  'End If
End Sub
Sub triarvolum()
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select distinct nommaquina,volum  from aniloxos where lineatura=" + atrim(cadbl(anilox(buscarvolum.tag))) + " order by volum"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   'formseleccio.Width = formseleccio.Width + 1000
   'formseleccio.DBGrid2.Width = formseleccio.DBGrid2.Width + 1000
   formseleccio.DBGrid2.Columns(0).width = 2000
   formseleccio.DBGrid2.Columns(1).width = 500
   'formseleccio.DBGrid2.Columns(2).Width = 1000
   'formseleccio.DBGrid2.Columns(3).Width = 1000
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           volum(cadbl(buscarvolum.tag)) = formseleccio.DBGrid2.Columns("volum")
        End If
   End If
    If seleccioret = 9 Then
        volum(cadbl(buscarvolum.tag)) = 0
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub bveurepdf_Click()
  formclixes.carregar_veure_pdfs
End Sub

Private Sub comboformaimpresio_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyCode = 0
End Sub

Private Sub Command1_Click()
   imprimir_comandesfetesambaquestsclixes
End Sub

Sub imprimir_comandesfetesambaquestsclixes()
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
 
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatcomandesfetesambclixe.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "clixesnous.mdb"
  oreport.RecordSelectionFormula = "{Tintes.id_treball}=" + atrim(id_treball) + " and {Tintes.ordremodificacio}=" + atrim(ordremodificacio)
  
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
Private Sub imprimir_Click()

End Sub

Private Sub anilox_DblClick(Index As Integer)
  buscar_Click
End Sub

Private Sub anilox_GotFocus(Index As Integer)
  buscar.Left = anilox(Index).Left + anilox(Index).width + 20
  buscar.Top = anilox(Index).Top
  buscar.visible = True
  buscar.tag = atrim(Index)
  buscar.ZOrder 0
End Sub

Private Sub anilox_LostFocus(Index As Integer)
   If Screen.ActiveControl.Name <> "buscar" Then buscar.visible = False
End Sub

Private Sub aniloxfotogravador_LostFocus(Index As Integer)
 If Index = 0 Then
  If MsgBox("Vols copiar aquesta Liniatura a totes les altres tintes?", vbInformation + vbYesNo + vbDefaultButton2, "Copiar anilox") = vbYes Then
         For i = 1 To 7
          If atrim(color(i)) <> "" Then
            aniloxfotogravador(i) = aniloxfotogravador(0)
              Else
                aniloxfotogravador(i) = "0"
          End If
         Next i
      End If
 End If
End Sub

Private Sub arrastre_Click(Index As Integer)
  Dim i As Byte
  Static esticadins As Boolean
  If esticadins Then Exit Sub
  esticadins = True
  For i = 0 To 7
     If i <> Index Then arrastre(i).Value = 0
  Next i
  esticadins = False
End Sub

Private Sub buscar_Click()
  Dim i As Byte
  triaranilox
  buscar.visible = False
  If cadbl(buscar.tag) = 0 Then
      If MsgBox("Vols copiar aquest anilox a totes les altres tintes?", vbInformation + vbYesNo + vbDefaultButton2, "Copiar anilox") = vbYes Then
         For i = 1 To 7
           If atrim(color(i)) <> "" Then
              anilox(i) = anilox(0)
             Else: anilox(i) = "0"
           End If
         Next i
      End If
  End If
End Sub
Sub triaranilox()
Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select lineatura as liniatura,sum(quantitat) as Quants_tenim from aniloxos group by lineatura order by lineatura"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   'formseleccio.Width = formseleccio.Width + 1000
   'formseleccio.DBGrid2.Width = formseleccio.DBGrid2.Width + 1000
   formseleccio.DBGrid2.Columns(0).width = 1000
   formseleccio.DBGrid2.Columns(1).width = 1500
   'formseleccio.DBGrid2.Columns(2).Width = 1000
   'formseleccio.DBGrid2.Columns(3).Width = 1000
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           anilox(cadbl(buscar.tag)) = formseleccio.DBGrid2.Columns("liniatura")
        End If
   End If
    If seleccioret = 9 Then
        anilox(cadbl(buscar.tag)) = 0
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub buscarcilindre_Click()
   triarcilindre
  buscarcilindre.visible = False
  If MsgBox("Vols copiar aquest Cilindre/Desaroll a totes les altres tintes?", vbInformation + vbYesNo + vbDefaultButton2, "Copiar anilox") = vbYes Then
         For i = 1 To 7
          If atrim(color(i)) <> "" Then
            cilindre(i) = cilindre(0)
            desarroll(i) = desarroll(0)
              Else
                cilindre(i) = "0"
                desarroll(i) = "0"
          End If
         Next i
      End If
End Sub

Function buscarmaquinahafetbaixacomanda(numc As Double) As Byte
  Dim rstb As Recordset
  buscarmaquinahafetbaixacomanda = 0
  Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
  Set rstb = dbbaixes.OpenRecordset("select * from impressores where comanda=" + atrim(numc))
  If Not rstb.EOF Then buscarmaquinahafetbaixacomanda = cadbl(rstb!numeromaquina)
End Function
Sub triarcilindre()
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  des = cadbl(InputBox("Quin desarroll vols utilitzar? en m/m.", "Desarroll"))
  If des = 0 Then Exit Sub
  'des = des * 10
'  caigudes = cadbl(InputBox("Quantes caigudes vols?", "Caigudes"))
'  If caigudes = 0 Then Exit Sub

  If comandabuscada > 0 Then nummaq = buscarmaquinahafetbaixacomanda(comandabuscada)

  sql = "SELECT  desarrolls.cilindre, desarrolls.desarroll, desarrolls.divisor as caigudes, MAQUINES.descripcio FROM maquines INNER JOIN (Cilindres INNER JOIN desarrolls ON Cilindres.id_cilindre = desarrolls.id_cilindre) ON maquines.codi = Cilindres.nummaquinaprincipal WHERE maquines.maquina='I' "
  were = "and desarroll=" + atrim(des) + IIf(comandabuscada > 0 And nummaq > 0, " and cilindres.nummaquinaprincipal=" + atrim(nummaq), "") + " order by cilindre "
  Set rst = dbcomandes.OpenRecordset(sql + were)
  If rst.EOF Then were = "and (desarroll>=" + atrim(des) + "-2 and desarroll<=" + atrim(des + 2) + ")  order by cilindre"  'and divisor=" + atrim(caigudes)
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = sql + were
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
    cilindre(cadbl(buscarcilindre.tag)) = atrim(cadbl(formseleccio.Data1.Recordset!cilindre))
    desarroll(cadbl(buscarcilindre.tag)) = atrim(cadbl(formseleccio.Data1.Recordset!desarroll))
  End If
  If seleccioret = 9 Then
    cilindre(cadbl(buscarcilindre.tag)) = 0
    desarroll(cadbl(buscarcilindre.tag)) = 0
  End If
 '  Data1.Recordset!client = Text2.Text
 '  nomclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
  
 ' End If
  Unload formseleccio
End Sub

Private Sub carregartintes_Click()
   
End Sub
Sub carregar_tintes_modificacio_anterior()
     Dim rstactual As Recordset
     Dim rstanterior As Snapshot
     Dim ultimid As Long
     Dim vordremodificacioxrcopiar As Double
     Dim vid_treballxrcopiar As Double
     'If vordremodificacio = 1 Then MsgBox "Aquesta es la primera modificació no es poden copiar tintes. ", vbCritical, "Atenció": Exit Sub
     vid_treballxrcopiar = cadbl(InputBox("Entra el treball que vols copiar els tinters.", "Copiar tinters", id_treball))
     If vid_treballxrcopiar < 1 Then Exit Sub
     vordremodificacioxrcopiar = cadbl(InputBox("Entra la versió que vols copiar els tinters del treball " + atrim(vid_treballxrcopiar), "Copiar tinters", IIf(vid_treballxrcopiar = id_treball, IIf(ordremodificacio > 1, ordremodificacio - 1, 0), 0)))
     Set rstactual = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
     Set rstanterior = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(vid_treballxrcopiar) + " and ordremodificacio=" + atrim(vordremodificacioxrcopiar) + " order by ordretinter asc", dbOpenSnapshot)
     If rstactual.EOF Or rstanterior.EOF Then MsgBox "No s'ha trobat els tinters per copiar", vbCritical, "Error": Exit Sub
     If Not te_tinters_repetits(rstanterior) Then
        copiar_tintes_anterior_a_actual rstanterior, rstactual
        carregartintes
         Else: MsgBox "Aquesta modificacio te tinters repetits, arregla-la primer.", vbCritical, "Error": Exit Sub
     End If
End Sub
Function te_tinters_repetits(rsta As Recordset) As Boolean
   Dim i As Byte
   Dim j As Byte
   Dim rst2 As Recordset
   If rsta.EOF Then Exit Function
   Set rst2 = rsta.Clone
   te_tinters_repetits = False
   While Not rsta.EOF
     rst2.FindFirst "ordretinter=" + atrim(rsta!ordretinter)
     If Not rst2.NoMatch Then
        rst2.FindNext "ordretinter=" + atrim(rsta!ordretinter)
        If Not rst2.NoMatch Then te_tinters_repetits = True: GoTo fi
         Else: GoTo fi
     End If
     rsta.MoveNext
   Wend
fi:
   Set rst2 = Nothing
   rsta.MoveFirst
End Function
Sub copiar_tintes_anterior_a_actual(rstanterior As Recordset, rstactual As Recordset)
    Dim vmodificatperinplacsa As Boolean
    
    Dim resp As String
    
    If Not rstanterior.EOF Then
        rstanterior.MoveLast
        ultimid = rstanterior!id_tinter
        rstanterior.MoveFirst
    End If
    While Not rstanterior.EOF
        If rstanterior!modificatperinplacsa Then vmodificatperinplacsa = True
        rstactual.FindFirst "ordretinter=" + atrim(cadbl(rstanterior!ordretinter))
        rstactual.Edit
        For i = 0 To rstanterior.Fields.Count - 1
          If rstanterior.Fields(i).Name <> "id_tinter" Then
              rstactual.Fields(i) = rstanterior.Fields(i)
          End If
        Next i
        rstactual!id_treball = id_treball
        rstactual!ordremodificacio = ordremodificacio
        rstactual!id_tinter_anterior = rstanterior!id_tinter
        rstactual.Update
        dbclixes.Execute "update tintes set tinterlinkambid_treball=" + atrim(rstactual!id_tinter) + " where tinterlinkambid_treball=" + atrim(rstanterior!id_tinter)
        'dbclixes.Execute "update tintes set ordremodificacio=" + atrim(ordremodificacio) + ",id_treball=" + atrim(id_treball) + " where id_tinter=" + atrim(rstanterior!id_tinter)
        If rstanterior!id_tinter <> ultimid Then
           rstanterior.MoveNext
          Else: rstanterior.MoveLast: rstanterior.MoveNext
        End If
        
    Wend
    If vmodificatperinplacsa Then
      While UCase(resp) <> "OK"
           resp = InputBox("Hi ha clixes MODIFICATS PER INPLACSA, tingues-ho en compte al fer la comanda." + Chr(10) + "Escriu OK per continuar.", "Atenció")
      Wend
    End If
End Sub
Sub triardetalltinter()
 Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  
  sql = "SELECT  detall from detallsdelstinters "
  were = " order by detall"
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(camiclixes) + "tintes.mdb"
  formseleccio.Data1.RecordSource = sql + were
  formseleccio.width = 6000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.DBGrid2.col = 0
  formseleccio.refrescar
  formseleccio.colocar_botofiltre 0
  formseleccio.Show 1
  If seleccioret = 1 Then
    detalltinter(cadbl(buscardetalltinter.tag)) = atrim(formseleccio.Data1.Recordset!detall)
  End If
  If seleccioret = 9 Then
    detalltinter(cadbl(buscardetalltinter.tag)) = ""
  End If
 '  Data1.Recordset!client = Text2.Text
 '  nomclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
  
 ' End If
  If detalltinter(cadbl(buscardetalltinter.tag)) = "" Then
       color(cadbl(buscardetalltinter.tag)).width = 2845
         Else: color(cadbl(buscardetalltinter.tag)).width = 2265
    End If
  formtintes.Refresh
  Unload formseleccio
End Sub
Sub triartinta()
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  
  sql = "SELECT  codi,descripcio,referenciacolor,nominterndelbido as Bidó from tintes_tot "
  were = " order by descripcio"
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(camiclixes) + "tintes.mdb"
  formseleccio.Data1.RecordSource = sql + were
  formseleccio.width = 14000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 500
  formseleccio.Show 1
  If seleccioret = 1 Then
    color(cadbl(buscartinta.tag)) = atrim(formseleccio.Data1.Recordset!descripcio)
    color(cadbl(buscartinta.tag)).ToolTipText = atrim(formseleccio.Data1.Recordset!codi)
  End If
  If seleccioret = 9 Then
    color(cadbl(buscartinta.tag)) = ""
    color(cadbl(buscartinta.tag)).ToolTipText = ""
  End If
 '  Data1.Recordset!client = Text2.Text
 '  nomclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
  
 ' End If
  Unload formseleccio
End Sub

Private Sub buscartinta_Click()
   Dim vtintaabans As String
   vtintaabans = color(cadbl(buscartinta.tag))
   triartinta
   If vtintaabans <> color(cadbl(buscartinta.tag)) Then eliminar_tintesalternatives cadbl(buscartinta.tag)
   color_GotFocus cadbl(buscartinta.tag)
End Sub

Private Sub cilindre_DblClick(Index As Integer)
buscarcilindre_Click
End Sub

Private Sub cilindre_GotFocus(Index As Integer)
 buscarcilindre.Left = Me.Controls(Screen.ActiveControl.Name)(Index).Left + Me.Controls(Screen.ActiveControl.Name)(Index).width + 20
  buscarcilindre.Top = Me.Controls(Screen.ActiveControl.Name)(Index).Top
  buscarcilindre.visible = True
  buscarcilindre.tag = atrim(Index)
  buscarcilindre.ZOrder 0
End Sub

Private Sub cilindre_LostFocus(Index As Integer)
    If Screen.ActiveControl.Name <> "buscarcilindre" Then buscarcilindre.visible = False
End Sub

Private Sub clixeosleeve_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub
Sub crear_tintes()
  Dim rst As Recordset
  Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
  If rst.EOF Then
     For i = 1 To 8
      rst.AddNew
      rst!id_treball = id_treball
      rst!ordremodificacio = ordremodificacio
      rst!ordretinter = i
      rst!numpolimers = 1
      If ordremodificacio < 2 Then rst!afectatspelcanvi = True
      rst.Update
     Next i
  End If
  Set rst = Nothing
End Sub

Private Sub color_GotFocus(Index As Integer)
  buscartinta.Left = color(Index).Left + color(Index).width - buscartinta.width 'color(Index).Left + color(Index).Width + 20
  buscartinta.Top = color(Index).Top
  buscartinta.visible = True
  buscartinta.tag = atrim(Index)
  buscartinta.ZOrder 0
  If color(Index).ToolTipText <> "" Then
     color(Index).BackColor = &HC0E0FF
     color(Index).Locked = True
      Else:
         color(Index).BackColor = QBColor(15)
         color(Index).Locked = False
  End If
End Sub

Private Sub color_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  copiarvalorsdelaprimeratinta Index
End Sub
Sub copiarvalorsdelaprimeratinta(i As Integer)
   If atrim(color(i)) = "" Then
      If cadbl(cilindre(i)) = 0 Then cilindre(i) = cilindre(0)
      If cadbl(anilox(i)) = 0 Then anilox(i) = anilox(0)
      If cadbl(desarroll(i)) = 0 Then desarroll(i) = desarroll(0)
      If cadbl(polimers(i)) = 0 Then polimers(i) = polimers(0)
      continuu(i) = continuu(0)
   End If
End Sub

Private Sub color_LostFocus(Index As Integer)
  If Screen.ActiveControl.Name <> "buscartinta" Then buscartinta.visible = False
End Sub

Private Sub Command2_Click()
   
   
End Sub

Private Sub Command9_Click(Index As Integer)
  If Index = 2 Then
    llistademodificacions
  End If
End Sub
Sub llistademodificacions()
  Unload formseleccio
  Load formseleccio
  'formseleccio.Command3.tag = "filtre"
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = formclixes.clixes.DatabaseName
  formseleccio.Data1.RecordSource = "select * from tintes_controlcanvis where numtreball='" + atrim(id_treball) + "/" + atrim(ordremodificacio) + "' order by data desc"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(1).width = 1000
  formseleccio.DBGrid2.Columns(2).width = 900
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(2).NumberFormat = "dd/mm/yy"
  formseleccio.DBGrid2.Columns(3).width = 1400
  formseleccio.DBGrid2.Columns(4).width = 2500
  formseleccio.DBGrid2.Columns(5).width = 3500
  formseleccio.DBGrid2.Columns(6).width = 3500
  formseleccio.DBGrid2.col = 4
  formseleccio.width = 15000
  formseleccio.colocar_botofiltre 4
  DoEvents
  If formseleccio.Data1.Recordset.EOF Then Exit Sub
  formseleccio.Show 1
  Unload formseleccio
End Sub


Private Sub copiartintes_Click()
 ' If UCase(InputBox("Aixó copiará totes les tintes de la modificació anterior i sobreescriurà aquestes tintes." + Chr(10) + "ES CORRECTE?   escriu [SI] per sobreescriure.")) = "SI" Then
    carregar_tintes_modificacio_anterior
 ' End If
End Sub

Private Sub densitat_LostFocus(Index As Integer)
 If Index = 0 Then
    If MsgBox("Vols copiar aquesta densitat a totes les altres tintes?", vbInformation + vbYesNo + vbDefaultButton2, "Copiar anilox") = vbYes Then
         For i = 1 To 7
          If atrim(color(i)) <> "" Then
            densitat(i) = densitat(0)
              Else
                densitat(i) = "0"
          End If
         Next i
    End If
 End If
End Sub

Private Sub desarroll_DblClick(Index As Integer)
buscarcilindre_Click
End Sub

Private Sub desarroll_GotFocus(Index As Integer)
  buscarcilindre.Left = Me.Controls(Screen.ActiveControl.Name)(Index).Left + Me.Controls(Screen.ActiveControl.Name)(Index).width + 20
  buscarcilindre.Top = Me.Controls(Screen.ActiveControl.Name)(Index).Top
  buscarcilindre.visible = True
  buscarcilindre.tag = atrim(Index)
  buscarcilindre.ZOrder 0
End Sub

Private Sub desarroll_LostFocus(Index As Integer)
  If Screen.ActiveControl.Name <> "buscarcilindre" Then buscarcilindre.visible = False
End Sub

Private Sub detalltinter_GotFocus(Index As Integer)
  buscardetalltinter.Left = detalltinter(Index).Left + detalltinter(Index).width - buscardetalltinter.width
  buscardetalltinter.Top = detalltinter(Index).Top
  buscardetalltinter.visible = True
  buscardetalltinter.tag = atrim(Index)
  buscardetalltinter.ZOrder 0
End Sub

Private Sub detalltinter_LostFocus(Index As Integer)
  If Screen.ActiveControl.Name <> "buscardetalltinter" Then buscardetalltinter.visible = False
End Sub

Private Sub eliminar_Click()
  If ultimtinter = 99 Then MsgBox "Primer escull el numero de tinter fent clic dins de la casella del tinter.", vbCritical, "Borrar tinter": Exit Sub
  If MsgBox("Segur que vols eliminar el tinter  " + atrim(ultimtinter + 1), vbInformation + vbYesNo, "Atenció") = vbNo Then Exit Sub
    color(ultimtinter) = ""
    anilox(ultimtinter) = 0
    cilindre(ultimtinter) = 0
    desarroll(ultimtinter) = 0
    continuu(ultimtinter) = False
    aniloxfotogravador(ultimtinter) = 0
    densitat(ultimtinter) = 0
    clixeosleeve(ultimtinter) = ""
    polimers(ultimtinter) = 0
    modificatperinplacsa(ultimtinter) = False
    afectatspelcanvi(ultimtinter) = False
    compartit(ultimtinter).Value = False
    observacions(ultimtinter) = ""
    dbclixes.Execute "update tintes set tinterlinkambid_treball=0 where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " and id_tinter=" + ordre(ultimtinter).tag
  guardar_Click
End Sub

Private Sub etestatrevisiotintes_Click()
   If etestatrevisiotintes <> atrim(formclixes.modificacions.Recordset!estatrevisiotintes) Then
      If MsgBox("Vols canviar la situació de les revisions de tintes?", vbDefaultButton2 + vbYesNo + vbExclamation, "Atenció") = vbNo Then
         ' etestatrevisiotintes = atrim(formclixes.modificacions.Recordset!estatrevisiotintes)
          SendKeys "{tab}"
         Else
          dbclixes.Execute "insert into tintes_controlcanvis (numtreball,usuari,campafectat,valoranterior,valoractual) values ('" + atrim(id_treball) + "/" + atrim(ordremodificacio) + "','" + nomordinador + "','EstatREVtintes','" + atrim(formclixes.modificacions.Recordset!estatrevisiotintes) + "','" + atrim(etestatrevisiotintes) + "')"
          formclixes.modificacions.Recordset.Edit
          formclixes.modificacions.Recordset!estatrevisiotintes = etestatrevisiotintes
          formclixes.modificacions.Recordset.Update
          etestatrevisiotintes.tag = "tancar"
          SendKeys "{tab}"
      End If
   End If
End Sub

Private Sub etestatrevisiotintes_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub etestatrevisiotintes_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub etestatrevisiotintes_LostFocus()
   If etestatrevisiotintes.tag = "tancar" Then Unload Me
   etestatrevisiotintes = atrim(formclixes.modificacions.Recordset!estatrevisiotintes)
End Sub

Private Sub Form_Click()
  'If MsgBox("ok?", vbCritical + vbYesNo, "Atenció") = vbYes Then passarobstintesaliniesobs
End Sub
Sub passarobstintesaliniesobs()
  Dim rst As Recordset
  Dim rsttintes As Recordset
  Dim vtext As String
  Dim v1 As String
  Dim v2 As String
  Static jahisoc As Boolean
  If jahisoc Then Exit Sub
  jahisoc = True
  Set rst = dbclixes.OpenRecordset("SELECT * FROM modificacions order by id_Treball,ordre;")
  While Not rst.EOF
    Set rsttintes = dbclixes.OpenRecordset("Select * from tintes where id_treball=" + atrim(rst!id_treball) + " and ordremodificacio=" + atrim(rst!ordre) + " order by ordretinter")
    vtext = "": v1 = "": v2 = ""
    While Not rsttintes.EOF
       vtext = atrim(vtext) + IIf(atrim(rsttintes!observacions) <> "", " " + atrim(cadbl(rsttintes!ordretinter)) + "# " + atrim(rsttintes!observacions), "")
       rsttintes.MoveNext
    Wend
   ' While Len(vtext) > 160
   '    vtext = InputBox("Modifica per fer 160 caracters", "atenció", vtext)
   ' Wend
    'If vtext <> "" Then Stop
    v1 = Mid(vtext, 1, 80)
    v2 = Mid(vtext, 81, 160)
    If v1 <> "" Then dbclixes.Execute "insert into tintes_observacions (id_treball,ordre,observacio) values (" + atrim(cadbl(rst!id_treball)) + "," + atrim(cadbl(rst!ordre)) + ",'" + treure_apostruf(v1) + "')"
    If v2 <> "" Then
      dbclixes.Execute "insert into tintes_observacions (id_treball,ordre,observacio) values (" + atrim(cadbl(rst!id_treball)) + "," + atrim(cadbl(rst!ordre)) + ",'" + treure_apostruf(v2) + "')"
    End If
    rst.MoveNext
    Me.caption = atrim(rst!id_treball) + "/" + atrim(rst!ordre)
    DoEvents
  Wend
  Set rst = Nothing
  Set rsttintes = Nothing
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then guardar_Click
End Sub

Private Sub Form_Load()
  copiartintes.Enabled = True
  'If ordremodificacio = 1 Then copiartintes.Enabled = False
  dataliniesobs.DatabaseName = camiclixes
  If ordremodificacio < 0 Then
     breprint.visible = False
     creimpres.visible = False
     comboformaimpresio.visible = True
     Label21.visible = True
     fdades.BackColor = &H80FF80
  End If
  crear_tintes
  carregartintes
  Me.caption = "Manteniment de les Tintes    " + atrim(id_treball) + "/" + atrim(ordremodificacio)
End Sub
Sub comprovar_simaterialblanc()
  Dim i As Byte
  Dim hihablanc As Boolean
  
  For i = 0 To 7
     If InStr(1, UCase(color(i)), "BLANC") > 0 Then hihablanc = True
  Next i
  'material transparent
  If InStr(1, UCase(formclixes.materialultimacomanda), "TRANSPA") > 0 Or InStr(1, UCase(formclixes.materialultimacomanda), "METALI") > 0 Then
     If InStr(1, UCase(formclixes.materialultimacomanda), "BLANC") = 0 Then
      If Not hihablanc Then MsgBox "Aquesta comanda està preparada amb material TRANSPARENT/METALITZAT i no hi ha el color BLANC revisa que sigui correcte.", vbCritical, "Atenció"
     End If
  End If
  
  'material blanc
  If InStr(1, UCase(formclixes.materialultimacomanda), "BLANC") > 0 Then
      If hihablanc Then MsgBox "Aquesta comanda té el material BLANC i has afegir color BLANC revisa que sigui correcte.", vbCritical, "Atenció"
  End If
End Sub
Sub guardartintes()
   Dim rst As Recordset
   Dim rstlink As Recordset
   Dim i As Byte
   Dim vhihaalgu As Boolean
   Dim vordremodificacioant As Long
   comprovar_simaterialblanc
   vcanvisrevisioitintes = ""
   Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " order by ordretinter ASC")
   i = 0
   While Not rst.EOF
      If cadbl(rst!tinterlinkambid_treball) > 0 Then
        Set rstlink = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(rst!tinterlinkambid_treball))
        If Not rstlink.EOF Then
            guardardatosaliniatinta rst, i, True
           Else: MsgBox "Hi ha hagut un error al carregar el tinter linkat al tinter " + atrim(i + 1), vbCritical, "Atenció"
        End If
         Else: guardardatosaliniatinta rst, i, False
      End If
      If atrim(color(i)) <> "" Then vhihaalgu = True
      i = i + 1
      rst.MoveNext
   Wend
   'Set rst = dbclixes.OpenRecordset("select reimpres from modificacions where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio))
   If ordremodificacio < 0 Then
      'rst.Edit
      formclixes.modificacions.Recordset.Edit
      If vhihaalgu Then
        formclixes.modificacions.Recordset!reimpres = True 'IIf(creimpres.Value = 1, True, False)
        formclixes.modificacions.Recordset!reprintformaimpres = atrim(Mid(comboformaimpresio + " ", 1, 1))
          Else: formclixes.modificacions.Recordset!reimpres = False
                formclixes.modificacions.Recordset!reprintformaimpres = ""
      End If
      vordremodificacioant = ordremodificacio
       formclixes.modificacions.Recordset.Update
       ordremodificacio = vordremodificacioant
      'rst.Update
   End If
   guardar_observacions
   If vcanvisrevisioitintes <> "" Then enviar_a_tintes vcanvisrevisioitintes
   Set rst = Nothing
   Set rstlink = Nothing
End Sub
Sub enviar_a_tintes(vmsg As String)
    enviaremailgeneric "tintes@inplacsa.com", "Canvis a un treball REVISAT.", treure_apostruf(vmsg)
End Sub
Sub guardar_observacions()
   dataliniesobs.Refresh
   While Not dataliniesobs.Recordset.EOF
      If atrim(dataliniesobs.Recordset!observacio) = "" Then
        dataliniesobs.Recordset.Delete
       ' dataliniesobs.Refresh
      End If
      dataliniesobs.Recordset.MoveNext
   Wend
   dataliniesobs.Refresh
End Sub
Sub guardardatosaliniatinta(rstdatos As Recordset, i As Byte, estalinkat As Boolean)
    Dim valorsantics(100) As String
    Dim k As Byte
    For k = 0 To (rstdatos.Fields.Count - 1)
      valorsantics(k) = atrim(rstdatos.Fields(k))
    Next k
    'i = 0
    If i > 7 Then Exit Sub
    rstdatos.Edit
    If rstdatos!id_tinter <> cadbl(ordre(i).tag) Then MsgBox "Id no correspon error gravant.": Exit Sub
    rstdatos!ordretinter = cadbl(ordre(i))
    rstdatos!observacions = atrim(observacions(i))
    rstdatos!aniloxclixe = cadbl(aniloxfotogravador(i))
    rstdatos!densitatutilitzada = cadbl(densitat(i))
    If Not estalinkat Then
      If atrim(color(i).ToolTipText) <> "" And atrim(rstdatos!coloranterior) = "" Then rstdatos!coloranterior = atrim(rstdatos!color)
      rstdatos!color = atrim(color(i))
      rstdatos!detalltinter = atrim(detalltinter(i))
      rstdatos!coditinta = color(i).ToolTipText
      rstdatos!anilox = cadbl(anilox(i))
      rstdatos!cilindre = cadbl(cilindre(i))
      rstdatos!desarroll = cadbl(desarroll(i))
      rstdatos!volum = cadbl(volum(i))
      'rstdatos!viscositat = cadbl(viscositat(i))
      rstdatos!tanx100cobertura = cadbl(tanx100(i))
      rstdatos!continuu = IIf(continuu(i) = 0, False, True)
      rstdatos!clixeosleeve = atrim(clixeosleeve(i))
      rstdatos!numpolimers = cadbl(polimers(i))
      rstdatos!afectatspelcanvi = IIf(afectatspelcanvi(i) = 0, False, True)
      rstdatos!modificatperinplacsa = IIf(modificatperinplacsa(i) = 0, False, True)
      rstdatos!comparteix = IIf(compartit(i).Value = 0, False, True)
      rstdatos!macula = IIf(macula(i).Value = 0, False, True)
      rstdatos!arrastre = IIf(arrastre(i).Value = 0, False, True)
      rstdatos!portasang = IIf(portasang(i).Value = 0, False, True)
      rstdatos!aniloxclixe = cadbl(aniloxfotogravador(i))
      rstdatos!densitatutilitzada = cadbl(densitat(i))
        Else
          rstdatos!color = "": rstdatos!anilox = 0: rstdatos!cilindre = 0: rstdatos!desarroll = 0: rstdatos!continuu = False: rstdatos!clixeosleeve = "": rstdatos!numpolimers = 0: rstdatos!modificatperinplacsa = False: rstdatos!afectatspelcanvi = False: rstdatos!comparteix = False
          rstdatos!detalltinter = ""
          rstdatos!coditinta = ""
        '  rstdatos!aniloxclixe = 0: rstdatos!densitatutilitzada = 0
    End If
    
    rstdatos.Update
    For k = 0 To (rstdatos.Fields.Count - 1)
      If ((rstdatos.Fields(k).Type <> 10 And rstdatos.Fields(k).Type <> 1) And cadbl(valorsantics(k)) <> cadbl(rstdatos.Fields(k).Value)) Or ((rstdatos.Fields(k).Type = 10 Or rstdatos.Fields(k).Type = 1) And atrim(valorsantics(k)) <> atrim(rstdatos.Fields(k).Value)) Then
          'gravar la modificació i la persona i ordinador
          dbclixes.Execute "insert into tintes_controlcanvis (numtreball,usuari,campafectat,valoranterior,valoractual) values ('" + atrim(rstdatos!id_treball) + "/" + atrim(rstdatos!ordremodificacio) + "','" + nomordinador + "','Tinter" + atrim(i + 1) + ": " + rstdatos.Fields(k).Name + "','" + treure_apostruf(valorsantics(k)) + "','" + treure_apostruf(rstdatos.Fields(k)) + "')"
          'MsgBox rstdatos.Fields(k).Name
          If InStr(1, atrim(rstdatos!color), "P-") > 0 Then
            If InStr(1, etestatrevisiotintes, "+TIN") > 0 Then
                 If rstdatos.Fields(k).Name = "desarroll" Or rstdatos.Fields(k).Name = "volum" Then
                    vcanvisrevisiotintes = vcanvisrevisiotintes + "Canvis de desarroll o volum al treball " + atrim(rstdatos!id_treball) + "/" + atrim(rstdatos!ordremodificacio) + vbNewLine
                    vcanvisrevisioitintes = vcanvisrevisiotintes + "     " + nomordinador + "  Tinter" + atrim(i + 1) + ": " + rstdatos.Fields(k).Name + " " + treure_apostruf(valorsantics(k)) + "  -> " + treure_apostruf(rstdatos.Fields(k)) + vbNewLine + vbNewLine
                    MsgBox vcanvisrevisioitintes
                 End If
            End If
        End If
     End If
    Next k
    
    
End Sub

Sub carregartintes()
   Dim rst As Recordset
   Dim rstlink As Recordset
   Dim i As Byte
   dataliniesobs.RecordSource = "select * from tintes_observacions where id_Treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio) + " order by id"
   dataliniesobs.Refresh
   Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " order by ordretinter ASC")
   i = 0
   While Not rst.EOF And i < 8
      
      If cadbl(rst!tinterlinkambid_treball) > 0 Then
        Set rstlink = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(rst!tinterlinkambid_treball))
        If Not rstlink.EOF Then
            passardatosaliniatinta rst, i, True, rstlink
           Else:
             MsgBox "Hi ha hagut un error al carregar el tinter linkat al tinter " + atrim(i + 1), vbCritical, "Atenció"
             passardatosaliniatinta rst, i, False
        End If
         Else: passardatosaliniatinta rst, i, False
      End If
      i = i + 1
      rst.MoveNext
   Wend
   While Not rst.EOF
     rst.Delete
     rst.MoveNext
   Wend
   'Set rst = dbclixes.OpenRecordset("select reimpres from modificacions where id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio))
   If formclixes.modificacions.Recordset!reimpres Then
     creimpres.Value = IIf(formclixes.modificacions.Recordset!reimpres, 1, 0)
     If atrim(formclixes.modificacions.Recordset!reprintformaimpres) <> "" Then
         comboformaimpresio = IIf(atrim(formclixes.modificacions.Recordset!reprintformaimpres) = "N", "Normal", "Transparencia")
           Else: comboformaimpresio = ""
     End If
     If ordremodificacio > 0 Then breprint.visible = True
     If creimpres.Value = 1 Then breprint.BackColor = &H80FF80
   End If
   bokdisseny.visible = False
   
   If modificartintes Then etestatrevisiotintes.Enabled = False: bobstintes.visible = True
   etestatrevisiotintes = atrim(formclixes.modificacions.Recordset!estatrevisiotintes)
   If InStr(1, etestatrevisiotintes, "OK DISSENY") = 0 And InStr(1, etestatrevisiotintes, "+TIN") > 0 And InStr(1, etestatrevisiotintes, "+IMP") > 0 Then
      bokdisseny.visible = True
      bokdisseny.tag = "+OK DISSENY"
   End If
   If modificartintes Then
    If atrim(arguments(6)) = "+TIN" And InStr(1, etestatrevisiotintes, "+TIN") = 0 Then
        bokdisseny.visible = True
        bokdisseny.tag = "+TIN"
        
    End If
    If atrim(arguments(6)) = "+IMP" And InStr(1, etestatrevisiotintes, "+IMP") = 0 Then
        bokdisseny.visible = True
        bokdisseny.tag = "+IMP"
    End If
   End If
   etrevtintes = bokdisseny.tag + " " + "Estat revisió tintes:"
   bokdisseny.ToolTipText = "Fer OK [" + bokdisseny.tag + "] de la revisió de tintes."
   possar_ultims_adhesius_utilitzats cadbl(id_treball), cadbl(ordremodificacio)
   posarcolors_tintesalternatives
   Set rst = Nothing
   Set rstlink = Nothing
End Sub
Sub posarcolors_tintesalternatives()
   Dim rst As Recordset
   Dim i As Byte
   For i = 0 To 7
        btintesalternatives(i).BackColor = &H8000000F
        Set rst = dbclixes.OpenRecordset("select * from tintes_alternatives where id_tinter=" + atrim(cadbl(ordre(i).tag)))
        If Not rst.EOF Then btintesalternatives(i).BackColor = QBColor(11)
   Next i
   Set rst = Nothing
End Sub
Sub possar_ultims_adhesius_utilitzats(vtreball As Double, vversio As Double)
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim rstad As Recordset
   Dim i As Byte
   For i = 0 To 7: viscositat(i) = "": Next i
   Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vversio))
   Set rstad = dbclixes.OpenRecordset("select * from adhesiusmuntadora")
   While Not rst.EOF
      Set rst2 = dbclixes.OpenRecordset("select * from muntadorescilindres where id_tinter=" + atrim(rst!id_tinter) + " order by datamuntatge DESC")
      If Not rst2.EOF Then
          rstad.FindFirst "codiintern='" + atrim(rst2!idadhesiu) + "'"
          If Not rstad.NoMatch Then viscositat(rst!ordretinter - 1) = atrim(rstad!inicialsfoam)
      End If
      rst.MoveNext
   Wend
   Set rst = Nothing
   Set rstad = Nothing
   Set rst2 = Nothing
End Sub
Sub possarcolorfondocontrols(vcolor As Double, i As Byte)
   ' ordre(i).BackColor = vcolor
    color(i).BackColor = vcolor
    detalltinter(i).BackColor = vcolor
    anilox(i).BackColor = vcolor
    cilindre(i).BackColor = vcolor
    desarroll(i).BackColor = vcolor
    continuu(i).BackColor = vcolor
    clixeosleeve(i).BackColor = vcolor
    polimers(i).BackColor = vcolor
    afectatspelcanvi(i).BackColor = vcolor
    modificatperinplacsa(i).BackColor = vcolor
    compartit(i).BackColor = vcolor
   ' observacions(i).BackColor = vcolor
End Sub
Sub passardatosaliniatinta(ByVal rstdatos As Recordset, i As Byte, estalinkat As Boolean, Optional rstlink As Recordset)
      If estalinkat Then
            linkat(i).BackColor = COLORFONDOLINKAT
            linkat(i).tag = atrim(rstdatos!id_tinter)
            color(i).tag = atrim(rstdatos!id_treball)
            possarcolorfondocontrols COLORFONDOLINKAT, i
          Else
                linkat(i).BackColor = QBColor(15)
                linkat(i).tag = ""
                color(i).tag = ""
                possarcolorfondocontrols QBColor(15), i
                If atrim(rstdatos!coditinta) <> "" And atrim(rstdatos!coditinta) <> "0" Then color(i).BackColor = &HC0E0FF
                detalltinter(i).BackColor = &H80C0FF
      End If
      If cadbl(rstdatos!tinterlinkambid_treball) < 0 Then
         linkat(i).BackColor = COLORFONDOLINKATnomesXL
         linkat(i).tag = atrim(rstdatos!tinterlinkambid_treball * -1)
         color(i).tag = atrim(rstdatos!id_treball)
         compartit(i).Enabled = False
           Else: compartit(i).Enabled = Not estalinkat
      End If
      linkat(i).visible = True
    
    ordre(i).Enabled = True
    color(i).Enabled = Not estalinkat
    detalltinter(i).Enabled = Not estalinkat
    anilox(i).Enabled = Not estalinkat
    cilindre(i).Enabled = Not estalinkat
    desarroll(i).Enabled = Not estalinkat
    continuu(i).Enabled = Not estalinkat
    volum(i).Enabled = Not estalinkat
    viscositat(i).Enabled = False
    tanx100(i).Enabled = Not estalinkat
    clixeosleeve(i).Enabled = Not estalinkat
    polimers(i).Enabled = Not estalinkat
    portasang(i).Enabled = Not estalinkat
    afectatspelcanvi(i).Enabled = Not estalinkat
    modificatperinplacsa(i).Enabled = Not estalinkat

    observacions(i).Enabled = True
    
    ordre(i).tag = atrim(cadbl(rstdatos!id_tinter))
    ordre(i) = atrim(cadbl(rstdatos!ordretinter))
    observacions(i) = atrim(rstdatos!observacions)
    aniloxfotogravador(i) = atrim(cadbl(rstdatos!aniloxclixe))
    densitat(i) = atrim(cadbl(rstdatos!densitatutilitzada))
    
    'ara els linkats
    If estalinkat Then Set rstdatos = rstlink
    color(i) = atrim(rstdatos!color)
    color(i).ToolTipText = IIf(atrim(rstdatos!coditinta) <> "0", atrim(rstdatos!coditinta), "")
    detalltinter(i) = atrim(rstdatos!detalltinter)
    anilox(i) = atrim(cadbl(rstdatos!anilox))
    cilindre(i) = atrim(cadbl(rstdatos!cilindre))
    desarroll(i) = atrim(cadbl(rstdatos!desarroll))
    volum(i) = atrim(cadbl(rstdatos!volum))
    'viscositat(i) = atrim(cadbl(rstdatos!viscositat))
    tanx100(i) = atrim(cadbl(rstdatos!tanx100cobertura))
    continuu(i) = IIf(rstdatos!continuu, 1, 0)
    macula(i) = IIf(rstdatos!macula, 1, 0)
    arrastre(i) = IIf(rstdatos!arrastre, 1, 0)
    clixeosleeve(i) = atrim(rstdatos!clixeosleeve)
    polimers(i) = atrim(cadbl(rstdatos!numpolimers))
    portasang(i) = IIf(rstdatos!portasang, 1, 0)
    afectatspelcanvi(i) = IIf(rstdatos!afectatspelcanvi, 1, 0)
    modificatperinplacsa(i) = IIf(rstdatos!modificatperinplacsa, 1, 0)
    compartit(i).Value = IIf(rstdatos!comparteix, 1, 0)
    If cadbl(aniloxfotogravador(i)) = "0" Then aniloxfotogravador(i) = atrim(cadbl(rstdatos!aniloxclixe))
    If cadbl(densitat(i)) = "0" Then densitat(i) = atrim(cadbl(rstdatos!densitatutilitzada))
    color(i).width = 4140
    'If detalltinter(i) = "" Then
    '   color(i).width = 2845
    '     Else: color(i).width = 2265
    'End If
   
End Sub
Sub comprovarimpresiocentrada()
    If formclixes.comboimpresiocentrada = "" Then
        formclixes.modificacions.Recordset.Edit
        If MsgBox("Diga'm si aquesta impresió es centrada o No.", vbInformation + vbYesNo + vbDefaultButton4, "Impresió centrada") = vbYes Then
             formclixes.modificacions.Recordset!impresiocentrada = "Si"
               Else
                formclixes.modificacions.Recordset!impresiocentrada = "No"
        End If
        formclixes.modificacions.Recordset.Update
    End If
End Sub
Private Sub guardar_Click()
  If Not baixaclixes Then comprovarimpresiocentrada
  If ordremodificacio < 0 And comboformaimpresio = "" Then MsgBox "Escull una forma d'impresió pel REPRINT", vbExclamation, "Atenció": comboformaimpresio.SetFocus: Exit Sub
  If comprovarrepetits Then MsgBox "Hi ha numero de tinter repetit primer arregla-ho", vbCritical, "Atenció": Exit Sub
  guardartintes
  carregartintes
  fdades.Enabled = False
  ultimtinter = 99
  estatedicio = ""
End Sub

Sub ensenyarambquicomparteix(Index As Integer)
   Dim msg As String
   Dim rst As Recordset
   Dim tinterprincial As Double
   
   Set rst = dbclixes.OpenRecordset("SELECT Tintes.id_tinter, Tintes.id_treball,tintes.ordremodificacio,tintes.comparteix from Tintes WHERE (Tintes.tinterlinkambid_treball=" + atrim(cadbl(ordre(Index).tag)) + ") ;")
   If Not rst.EOF Then
     msg = msg + " Aquest tinter l'utilitzen els seguents treballs:"
     While Not rst.EOF
        msg = msg + Chr(10) + " Treball nº: " + atrim(rst!id_treball) + "/" + atrim(rst!ordremodificacio)
        rst.MoveNext
     Wend
       Else
          msg = msg + Chr(10) + " Aquest clixe no l'utilitza ningu."
   End If
   If msg <> "" Then
     MsgBox msg, vbInformation, "Informació"
   End If
   
   
End Sub

Private Sub Label24_Click()

End Sub

Private Sub linkat_Click(Index As Integer)
   guardartintes
   modificar_Click
   If cadbl(linkat(Index).tag) = 0 Then
      If compartit(Index).Value = 0 Then
         escullirtintercompartit CByte(Index)
           Else
             ensenyarambquicomparteix Index
      End If
      Else
        If escullirnoulinkat(Index) Then escullirtintercompartit CByte(Index)
   End If
End Sub
Function escullirnoulinkat(Index As Integer) As Boolean
   Dim msg As String
   Dim rst As Recordset
   Dim tinterprincial As Double
   Dim vnomesXL As Boolean
   escullirnoulinkat = False
   Set rst = dbclixes.OpenRecordset("select tinterlinkambid_treball from tintes where id_tinter=" + atrim(cadbl(ordre(Index).tag)))
   If rst.EOF Then Exit Function
   tinterprincipal = cadbl(rst!tinterlinkambid_treball)
   If tinterprincipal < 0 Then tinterprincipal = tinterprincipal * -1: vnomesXL = True
   Set rst = dbclixes.OpenRecordset("SELECT Tintes.id_tinter, tintes.id_treball,tintes.ordremodificacio FROM tintes WHERE (((Tintes.id_tinter)=" + atrim(tinterprincipal) + "));")
   If Not rst.EOF Then
     msg = "Aquest tinter estira el clixe del treball: " + atrim(rst!id_treball) + "/" + atrim(rst!ordremodificacio) + Chr(10) + IIf(vnomesXL, " (NOMÉS UBICACIÓ DE LA BOSSA.)" + Chr(10), "") + Chr(10)
   End If
   Set rst = dbclixes.OpenRecordset("SELECT Tintes.id_tinter, Tintes.id_treball,tintes.ordremodificacio,tintes.comparteix from Tintes WHERE ((not tintes.comparteix and (Tintes.tinterlinkambid_treball)=" + atrim(tinterprincipal) + ") and id_tinter<>" + atrim(cadbl(ordre(Index).tag)) + ");")
   If Not rst.EOF Then
     msg = msg + " Aquest tinter també l'utilitzen els seguents treballs:"
     While Not rst.EOF
        msg = msg + Chr(10) + " Treball nº: " + atrim(rst!id_treball) + "/" + atrim(rst!ordremodificacio)
        rst.MoveNext
     Wend
       Else
          msg = msg + Chr(10) + " Aquest clixe no es comparteix amb ningu mes."
   End If
   If msg <> "" Then
     msg = msg + Chr(10) + "SI VOLS MODIFICAR AQUEST LINK AMB EL TINTER D'UN ALTRE TREBALL ESCRIU: [modificar]"
     If UCase(InputBox(msg, "Informació link")) = "MODIFICAR" Then
        escullirnoulinkat = True
     End If
   End If
   
   
End Function
Sub duplicarregistre(vidoriginal As Long, viddesti As Long)
   Dim rst1 As Recordset
   Dim rst2 As Recordset
   Set rst1 = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(cadbl(vidoriginal)))
   Set rst2 = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(cadbl(viddesti)))
   If rst1.EOF Or rst2.EOF Then Exit Sub
   rst2.Edit
   For i = 4 To rst1.Fields.Count - 1
        If rst1.Fields(i).Name <> "tinterlinkambid_treball" Then
             rst2.Fields(i) = rst1.Fields(i)
        End If
   Next i
   rst2!comparteix = False
   rst2.Update
End Sub
Sub escullirtintercompartit(numtinter As Byte)
    Dim rst As Recordset
    Dim v As String
    Dim id_versiocompartit As Integer
    Dim id_treballcompartit As Long
    Dim id_tintercompartit As Long
    
    v = InputBox("Entra el numero de treball d'on vols agafar aquest tinter compartit." + Chr(10) + "RECORDA QUE AQUEST TREBALL HA DE TENIR ALGUN CLIXÉ MARCAT COM A COMPARTEIX." + Chr(10) + "Si vols deskinkar-lo escriu 0 o res.", "Escullir treball", color(numtinter).tag)
    If StrPtr(v) = 0 Then Exit Sub
    id_treballcompartit = cadbl(v)
    If id_treballcompartit > 0 Then
        v = InputBox("Entra la versió del treball d'on vols agafar aquest tinter compartit." + Chr(10) + "RECORDA QUE AQUEST TREBALL HA DE TENIR ALGUN CLIXÉ MARCAT COM A COMPARTEIX." + Chr(10) + "Si vols deskinkar-lo escriu 0 o res.", "Escullir la versió")
        If StrPtr(v) = 0 Then Exit Sub
        id_versiocompartit = cadbl(v)
        If ordremodificacio < 0 Then id_versiocompartit = id_versiocompartit * -1
    End If
    
    
    If id_treballcompartit > 0 Then id_tintercompartit = escullirtintercompartitdeltreball(id_treballcompartit, id_versiocompartit)
    If id_tintercompartit = 0 Then
           If MsgBox("No s'ha escullit cap clixé vols deslinkar aquest clixé compartit?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo cont
             Else
               If id_tintercompartit = cadbl(ordre(numtinter).tag) Then MsgBox "No pots linkar el mateix tinter que estàs compartint.", vbCritical, "Error": Exit Sub
               If MsgBox("Vols compartir tota la informació del clixé compartit [SI]" + Chr(10) + "o només l'XL on està guardat? [NO]", vbInformation + vbDefaultButton1 + vbYesNo, "Atenció") = vbNo Then
                  id_tintercompartit = id_tintercompartit * -1
               End If
    End If
    dbclixes.Execute "update tintes set tinterlinkambid_treball=" + atrim(id_tintercompartit) + " where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " and id_tinter=" + ordre(numtinter).tag
    'si es negatiu vol dir que només vol l'XL per tan copio tota la informació del registre original i que canvii
    'la informació que vulgui l'usuari.
    'AIXO ES VA FER PODER LINKAR UN TREBALL AMBUN ALTRA PERÒ PODER CANVIAR EL PANTONE QUE S'UTILITZARÀ
    If id_tintercompartit < 0 Then duplicarregistre id_tintercompartit * -1, cadbl(ordre(numtinter).tag)
    
cont:
 '   End If
    carregartintes
End Sub
Function escullirtintercompartitdeltreball(id_treballcompartit As Long, id_versiocompartit As Integer) As Long
  Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select id_tinter,ordretinter as [Ordre], color as [Color del tinter] from tintes where (id_treball=" + atrim(id_treballcompartit) + " and ordremodificacio=" + atrim(id_versiocompartit) + ") and comparteix"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns(0).visible = False
   formseleccio.DBGrid2.Columns(1).width = 600
   formseleccio.DBGrid2.Columns(2).width = 3400
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           escullirtintercompartitdeltreball = formseleccio.DBGrid2.Columns("id_tinter")
        End If
   End If
    If seleccioret = 9 Then
        escullirtintercompartitdeltreball = 0
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Function

Private Sub modificar_Click()
   fdades.Enabled = True
   ultimtinter = 99
   estatedicio = "Editant..."
End Sub

Private Sub ordre_Change(Index As Integer)
   comprovarrepetits
End Sub
Function comprovarrepetits() As Boolean
  Dim trobat As Byte
  For i = 0 To 7
    trobat = 0
    For j = 0 To 7
       If cadbl(ordre(i)) = cadbl(ordre(j)) And cadbl(ordre(i)) > 0 Then trobat = trobat + 1
    Next j
    If trobat > 1 Then
       ordre(i).BackColor = QBColor(12): comprovarrepetits = True
      Else: ordre(i).BackColor = QBColor(15)
    End If
  Next i
End Function

Private Sub ordre_GotFocus(Index As Integer)
  ultimtinter = Index
End Sub

Private Sub reixa_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim vnomcamp As String
  vnomcamp = reixa.Columns(reixa.col).DataField
  If Len(reixa.Text) >= dataliniesobs.Recordset.Fields(vnomcamp).Size Then
     If KeyCode > 31 And KeyCode < 127 Then KeyCode = 0
  End If
End Sub

Private Sub reixa_KeyPress(KeyAscii As Integer)
Dim vnomcamp As String
  vnomcamp = reixa.Columns(reixa.col).DataField
  If Len(reixa.Text) >= dataliniesobs.Recordset.Fields(vnomcamp).Size Then
     If KeyAscii > 31 And KeyAscii < 127 Then KeyAscii = 0
  End If
End Sub

Private Sub sortir_Click()
  If comprovarrepetits Then MsgBox "Hi ha numero de tinter repetit primer arregla-ho", vbCritical, "Atenció": Exit Sub
  If estatedicio <> "" Then If MsgBox("Estas editant les tintes i si surts perdràs els canvis," + Chr(10) + "VOLS SORTIR SENSE GUARDAR?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
  estatedicio = ""
  Unload Me
End Sub



Private Sub Timer1_Timer()
  If formclixes.llistadecomandespendents.ListCount > 0 Then
     If bcomandespendents.BackColor = &H5C31DD Then
           bcomandespendents.BackColor = &H8000000F
          Else: bcomandespendents.BackColor = &H5C31DD
     End If
  End If
End Sub

Private Sub volum_DblClick(Index As Integer)
   triarvolum
End Sub

Private Sub volum_GotFocus(Index As Integer)
 buscarvolum.Left = volum(Index).Left + volum(Index).width + 20
  buscarvolum.Top = volum(Index).Top
  buscarvolum.visible = True
  buscarvolum.tag = atrim(Index)
  buscarvolum.ZOrder 0
End Sub

Private Sub volum_LostFocus(Index As Integer)
If Screen.ActiveControl.Name <> "buscarvolum" Then buscarvolum.visible = False
End Sub
