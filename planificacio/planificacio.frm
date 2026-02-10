VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form planificacio 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Planificació"
   ClientHeight    =   9900
   ClientLeft      =   2565
   ClientTop       =   1965
   ClientWidth     =   20250
   Icon            =   "planificacio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fcontrols 
      Height          =   9795
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   21300
      Begin VB.CommandButton btarifaipressupost 
         BackColor       =   &H00FF80FF&
         Caption         =   "TiP"
         Height          =   255
         Left            =   2835
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Veure la Tarifa o el Pressupost"
         Top             =   1680
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Frame Frameentregues 
         Caption         =   "Entregues"
         Height          =   780
         Left            =   1380
         TabIndex        =   48
         Top             =   105
         Visible         =   0   'False
         Width           =   5985
         Begin VB.CommandButton Command9 
            Caption         =   "Una setmana"
            Height          =   450
            Left            =   1470
            TabIndex        =   52
            Top             =   225
            Width           =   1215
         End
         Begin VB.CheckBox Checknomesclixes 
            Caption         =   "Clixes"
            Height          =   195
            Left            =   4710
            TabIndex        =   50
            Top             =   120
            Width           =   1185
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Entre dates"
            Height          =   450
            Left            =   195
            TabIndex        =   49
            Top             =   225
            Width           =   1215
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   800
         Left            =   12675
         Top             =   -60
      End
      Begin VB.CommandButton Command67 
         Height          =   465
         Index           =   11
         Left            =   15630
         Picture         =   "planificacio.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Assignar linies d'impresió 001#1"
         Top             =   345
         Width           =   600
      End
      Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
         Height          =   4665
         Left            =   1245
         TabIndex        =   40
         Top             =   3180
         Visible         =   0   'False
         Width           =   7050
         _cx             =   5080
         _cy             =   5080
      End
      Begin VB.CommandButton botoreclamar 
         Height          =   480
         Left            =   15015
         Picture         =   "planificacio.frx":0E8F
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Reclamar comandes en paper."
         Top             =   330
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame framereclamar 
         BackColor       =   &H00F3B378&
         Caption         =   "Reclamar comandes amb paper"
         Height          =   4995
         Left            =   8325
         TabIndex        =   29
         Top             =   2805
         Visible         =   0   'False
         Width           =   5010
         Begin VB.Frame Framereclamades 
            Caption         =   "Reclamades"
            Height          =   4260
            Left            =   -3120
            TabIndex        =   43
            Top             =   210
            Visible         =   0   'False
            Width           =   4635
            Begin VB.CommandButton Command7 
               Height          =   480
               Left            =   3255
               Picture         =   "planificacio.frx":1839
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   "No reclamar la comanda escullida."
               Top             =   3675
               Width           =   1155
            End
            Begin VB.ListBox llistareclamades 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3300
               Left            =   180
               TabIndex        =   44
               Top             =   255
               Width           =   4230
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Si la comanda porta * està reactivada."
               Height          =   225
               Left            =   120
               TabIndex        =   46
               Top             =   3690
               Width           =   2910
            End
         End
         Begin VB.CommandButton Command6 
            Height          =   480
            Left            =   210
            Picture         =   "planificacio.frx":1DC3
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Historial de comandes reclamades."
            Top             =   4470
            Width           =   1035
         End
         Begin VB.ListBox llistacomandes 
            Height          =   255
            Left            =   1575
            TabIndex        =   41
            Top             =   3975
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox observacionsdemanades 
            BackColor       =   &H00EAD9CE&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   270
            MaxLength       =   50
            TabIndex        =   35
            Top             =   1830
            Width           =   4500
         End
         Begin VB.CommandButton Command5 
            Height          =   480
            Left            =   2370
            Picture         =   "planificacio.frx":24AD
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Afegir la comanda sel.leccionada"
            Top             =   3945
            Width           =   1035
         End
         Begin VB.CommandButton Command4 
            Height          =   480
            Left            =   3585
            Picture         =   "planificacio.frx":2A37
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Enviar per email a qui correspongui."
            Top             =   3945
            Width           =   1110
         End
         Begin VB.CommandButton Command2 
            Height          =   480
            Left            =   210
            Picture         =   "planificacio.frx":2FC1
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Borrar comandes apuntades."
            Top             =   3945
            Width           =   1035
         End
         Begin VB.TextBox llistadecomandes 
            BackColor       =   &H00EAD9CE&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   225
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   2775
            Width           =   4500
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Per afegir fer Click a Ok o dos clics a la comanda."
            Height          =   225
            Left            =   1350
            TabIndex        =   39
            Top             =   4560
            Width           =   4170
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Comandes demanades"
            Height          =   180
            Left            =   315
            TabIndex        =   37
            Top             =   2550
            Width           =   3885
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Observació"
            Height          =   180
            Left            =   285
            TabIndex        =   36
            Top             =   1590
            Width           =   1080
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"planificacio.frx":354B
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1365
            Left            =   180
            TabIndex        =   30
            Top             =   360
            Width           =   4815
         End
      End
      Begin VB.Timer timercontrol 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   12885
         Top             =   150
      End
      Begin VB.CommandButton bordre 
         Height          =   315
         Left            =   0
         Picture         =   "planificacio.frx":364D
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Ordenar per..."
         Top             =   1200
         Width           =   285
      End
      Begin VB.CommandButton exportaraxls 
         Height          =   480
         Left            =   16215
         Picture         =   "planificacio.frx":3BD7
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Exportar a Excel"
         Top             =   330
         Width           =   615
      End
      Begin VB.CommandButton exportarapdf 
         Height          =   480
         Left            =   16785
         Picture         =   "planificacio.frx":4E61
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exportar a PDF"
         Top             =   330
         Width           =   615
      End
      Begin VB.CommandButton canvidemaquina 
         BackColor       =   &H00FF8080&
         Caption         =   "Nº Maq."
         Height          =   525
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   345
         Width           =   810
      End
      Begin VB.CommandButton canviordre 
         BackColor       =   &H008080FF&
         Caption         =   "Ordre"
         Height          =   255
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   615
         Width           =   810
      End
      Begin VB.CheckBox multiseleccio 
         Caption         =   "Multiselecció"
         Height          =   195
         Left            =   75
         TabIndex        =   21
         Top             =   150
         Width           =   1275
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   100
         Top             =   7920
      End
      Begin VB.TextBox postit 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   4890
         TabIndex        =   18
         Top             =   1680
         Visible         =   0   'False
         Width           =   7575
      End
      Begin VB.Frame factualitzant 
         ForeColor       =   &H00FF8080&
         Height          =   1020
         Left            =   4920
         TabIndex        =   16
         Top             =   3120
         Visible         =   0   'False
         Width           =   5010
         Begin VB.Shape liniaprogres 
            BackColor       =   &H00FF8080&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   255
            Left            =   165
            Top             =   690
            Width           =   105
         End
         Begin VB.Label etactualitzant 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Actualitzant . . ."
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   18
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   435
            Left            =   675
            TabIndex        =   17
            Top             =   195
            Width           =   3885
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   315
            Left            =   105
            Top             =   660
            Width           =   4815
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   480
         Left            =   18015
         Picture         =   "planificacio.frx":516B
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Refrescar"
         Top             =   330
         Width           =   615
      End
      Begin VB.CommandButton sortir 
         Height          =   480
         Left            =   18630
         Picture         =   "planificacio.frx":56F5
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Sortir"
         Top             =   330
         Width           =   615
      End
      Begin VB.CommandButton botomaquina 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   885
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   330
         Width           =   1890
      End
      Begin VB.CommandButton botomaquina 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   2745
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   330
         Width           =   1890
      End
      Begin VB.CommandButton botomaquina 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   2
         Left            =   4650
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   330
         Width           =   1890
      End
      Begin VB.CommandButton botomaquina 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   3
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   330
         Width           =   1890
      End
      Begin VB.CommandButton botomaquina 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   4
         Left            =   8430
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   330
         Width           =   1890
      End
      Begin VB.CommandButton botomaquina 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   5
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   330
         Width           =   1890
      End
      Begin VB.TextBox filtre 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   0
         Left            =   255
         TabIndex        =   3
         ToolTipText     =   "Pots buscar valors separats per comes i a client pots posar el codi de client."
         Top             =   915
         Width           =   630
      End
      Begin VB.CommandButton treurefiltre 
         Height          =   270
         Left            =   15
         Picture         =   "planificacio.frx":5C7F
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminar totes les linies"
         Top             =   915
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Height          =   480
         Left            =   17400
         Picture         =   "planificacio.frx":6209
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir Packing-List"
         Top             =   330
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid reixa 
         Height          =   8340
         Left            =   225
         TabIndex        =   13
         Top             =   1170
         Width           =   20835
         _ExtentX        =   36751
         _ExtentY        =   14711
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorSel    =   16756318
         ForeColorSel    =   16711680
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   3
      End
      Begin VB.Image cNomesLectura 
         Height          =   360
         Left            =   12585
         Picture         =   "planificacio.frx":6793
         ToolTipText     =   "Només lectura... No pots fer canvis"
         Top             =   390
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label etultimaactualitzacio 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   18645
         TabIndex        =   28
         Top             =   135
         Width           =   2535
      End
      Begin VB.Label etmsgajuda 
         BackColor       =   &H0000FFFF&
         Height          =   270
         Left            =   1380
         TabIndex        =   27
         Top             =   60
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label etiquetaoperaris 
         BackStyle       =   0  'Transparent
         Caption         =   "Operaris"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   570
         Left            =   12975
         TabIndex        =   24
         Top             =   300
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label escullirmaqdesti 
         BackStyle       =   0  'Transparent
         Caption         =   "Escull la màquina on vols moure la comanda sel.leccionada."
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
         Height          =   180
         Left            =   1470
         TabIndex        =   15
         Top             =   150
         Visible         =   0   'False
         Width           =   5520
      End
      Begin VB.Label tempsproximrefresc 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   17895
         TabIndex        =   20
         Top             =   120
         Width           =   705
      End
      Begin VB.Label eseccio 
         BackStyle       =   0  'Transparent
         Caption         =   "Impresores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7755
         TabIndex        =   19
         Top             =   60
         Width           =   7275
      End
      Begin VB.Label registres 
         BackStyle       =   0  'Transparent
         Caption         =   "Comandes"
         Height          =   270
         Left            =   210
         TabIndex        =   14
         Top             =   9510
         Width           =   6120
      End
   End
   Begin VB.Menu mhoraris 
      Caption         =   "Horaris Màquines"
   End
   Begin VB.Menu meines 
      Caption         =   "Eines"
      Begin VB.Menu mcomananoreal 
         Caption         =   "Afegir comanda NO REAL"
      End
   End
   Begin VB.Menu mllistats 
      Caption         =   "Llistats/Imprimir"
      Begin VB.Menu llestatfabricacio 
         Caption         =   "Llistat d'estat de fabricació"
         Begin VB.Menu mllimpresores 
            Caption         =   "Impresores"
         End
         Begin VB.Menu mimpperlam 
            Caption         =   "Impresores per Laminadores"
         End
         Begin VB.Menu mllmuntadores 
            Caption         =   "Muntadores"
         End
         Begin VB.Menu mlllaminadores 
            Caption         =   "Laminadores"
         End
         Begin VB.Menu mllistattotes 
            Caption         =   "Totes les Seccions (M,I,L)"
         End
      End
      Begin VB.Menu mimprimirveurecomanda 
         Caption         =   "Imprimir/Veure Comanda                         "
      End
      Begin VB.Menu lldescansop 
         Caption         =   "Llistat temps descans dels operaris"
      End
      Begin VB.Menu mllistathoradeprogramacio 
         Caption         =   "Llistat hora de programació"
      End
      Begin VB.Menu mllistatentregades 
         Caption         =   "Llistat comandes entregades"
      End
   End
   Begin VB.Menu m1 
      Caption         =   ""
   End
   Begin VB.Menu m 
      Caption         =   ""
   End
   Begin VB.Menu m2 
      Caption         =   ""
   End
   Begin VB.Menu mgeneral 
      Caption         =   "General"
   End
   Begin VB.Menu mimpresores 
      Caption         =   "Impresores"
   End
   Begin VB.Menu mlaminadores 
      Caption         =   "Laminadores"
   End
   Begin VB.Menu mrebobinadores 
      Caption         =   "Rebobinadores"
   End
   Begin VB.Menu msoldadores 
      Caption         =   "Soldadores"
   End
   Begin VB.Menu mentregues 
      Caption         =   "Entregues"
   End
   Begin VB.Menu m4 
      Caption         =   ""
   End
   Begin VB.Menu M8 
      Caption         =   ""
   End
   Begin VB.Menu mexpedicions1 
      Caption         =   "Expedicions"
      Begin VB.Menu mnoenviat 
         Caption         =   "No Enviat"
      End
      Begin VB.Menu menviat 
         Caption         =   "Enviat"
      End
      Begin VB.Menu mpujaraexpedicions 
         Caption         =   "Pujar a Expedicions"
      End
   End
   Begin VB.Menu mreixa1 
      Caption         =   "menureixa1"
      Visible         =   0   'False
      Begin VB.Menu mcomandastandby 
         Caption         =   "Comanda a StandBy"
      End
   End
   Begin VB.Menu menureclam 
      Caption         =   "menureclam"
      Visible         =   0   'False
      Begin VB.Menu smno 
         Caption         =   "No"
      End
      Begin VB.Menu sms1ja 
         Caption         =   "S1-Ja"
      End
      Begin VB.Menu sms2 
         Caption         =   "S2-12h"
      End
      Begin VB.Menu sms3 
         Caption         =   "S3-24h"
      End
   End
   Begin VB.Menu mexpedicio 
      Caption         =   "menuexpedicio"
      Visible         =   0   'False
      Begin VB.Menu menviarja 
         Caption         =   "ENVIAR(Possar Data Expedició)"
      End
      Begin VB.Menu mnoenviarencara 
         Caption         =   "No ENVIAR (Treure Data Expedició)"
      End
      Begin VB.Menu mentregaparcialtotal 
         Caption         =   "Entrega PARCIAL(GROC)/TOTAL(BLANC)"
      End
   End
   Begin VB.Menu m3 
      Caption         =   "                                            "
      Enabled         =   0   'False
   End
   Begin VB.Menu mreclamacions 
      Caption         =   "Reclamacions"
   End
   Begin VB.Menu macceptarCdL 
      Caption         =   "menuacceptarCdL"
      Visible         =   0   'False
      Begin VB.Menu CdLacceptar 
         Caption         =   "Acceptar-la (VERD)"
      End
      Begin VB.Menu mrebutjarCdL 
         Caption         =   "Rebutjar-la (VERMELL)"
      End
   End
   Begin VB.Menu m9 
      Caption         =   ""
   End
   Begin VB.Menu mtaripres 
      Caption         =   "menutarifaipressupost"
      Visible         =   0   'False
      Begin VB.Menu mveuretarifa 
         Caption         =   "Veure la Tarifa"
      End
      Begin VB.Menu mveurepressupost 
         Caption         =   "Veure Pressupost"
      End
   End
End
Attribute VB_Name = "planificacio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim whereultimfiltre As String
Dim generarelfitxertemporal As Boolean
Dim rstCdL As Recordset
Dim rstCdLestats As Recordset
Dim vhihanCdLtaronja As Boolean
Dim dbsap As Database
Function ObtenerLíneaComando(Optional MaxArgs)
    'Declara las variables.
    Dim C, LíneaComando, LonLínComando, ArgIn, i, NúmArgs
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
        C = Mid(LíneaComando, i, 1)
        'Comprueba espacio o tabulación.
        If (C <> " " And C <> vbTab) Then
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

ArgArray(NúmArgs) = ArgArray(NúmArgs) + C
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

Private Sub Command9_Click()
    'empleno la taula d'entregues
    ratoli "espera"
    copiaregistreatemporalentregues
    mentregues_Click
    ratoli "normal"
End Sub

Private Sub mllistatentregades_Click()
  Dim vcodi As Double
  Dim vnomclient As String
  Dim vfi As String
  Dim vinici As String
  Dim vfitxertemp As String
  If existeix(vfitxertemp) Then Kill vfitxertemp
  escullir_client vcodi, vnomclient
  If vcodi = 0 Then Exit Sub
  vinici = InputBox("Entra la data d'inici del llistat", "Data inici")
  If Not IsDate(vinici) Then GoTo fi
  vfi = InputBox("Entra la data de fi del llistat", "Data fi")
  If Not IsDate(vfi) Then GoTo fi
  vfitxertemp = "c:\temp\llistatentregadesexportat.csv"
  Open vfitxertemp For Output As #3
  Print #3, ""
  Print #3, "                                            COMANDES ENTREGADES"
  Print #3, vbNewLine
  Print #3, "CLIENT: " + atrim(vcodi) + "-" + atrim(vnomclient)
  Print #3, "Comandes entregades Data fi: " + atrim(vfi)
  Print #3, "Comandes entregades Data inici: " + atrim(vinici)
  Print #3, ""
  Print #3, "Data Comanda;Data Entrega;NºLot;Fulla;NºContracte/NºComanda client;NºCall-Off;Ref.Client;Quantitat teorica demanada;Quantitat entregada;Kg entregats;Pvp venta;Unitats;Total venta;Nom client entrega"
  posarliniesLLISTATENTREGA vinici, vfi, vcodi
  
  
  Close #3
  If existeix(vfitxertemp) Then obrir_document vfitxertemp
fi:
  ratoli "normal"

End Sub
Sub posarliniesLLISTATENTREGA(vinici As String, vfi As String, vcodi As Double)
  Dim rstt As Recordset
  Dim rsttemp As Recordset
  Dim amplemax As Double
  Dim taulaplanificacio As String
  Dim tintes As Byte
  Dim rstalb As Recordset
  Dim dbvendes As Database
  Dim vsql As String
  Dim rstobs As Recordset
  Dim vunitat As String
  Dim vcalloff As String
  ratoli "espera"
  Set dbvendes = OpenDatabase(rutadelfitxer(cami) + "vendes.mdb")
  taulaplanificacio = "planificacioent"
  'Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio, dbOpenSnapshot, dbReadOnly)
  vsql = "SELECT comandes.comandaclient,Clients_envios.nome,capcaleraalbara.*, liniesalbara.*, clients.nom, comandes.refilate,comandes.comandaclient,comandes.datacomanda, comandes.impressio,comandes.pvp, comandes.tubbaseext, comandes.numtreball,comandes.refclientdeclient, pressupostos.preu FROM (((capcaleraalbara RIGHT JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara) LEFT JOIN (Clients_envios LEFT JOIN clients ON Clients_envios.codi = clients.codi) ON capcaleraalbara.id_direnvio = Clients_envios.id) LEFT JOIN comandes ON liniesalbara.lotinplacsa = comandes.comanda) LEFT JOIN pressupostos ON liniesalbara.lotinplacsa = pressupostos.lotambelqueshafacturat "
  vsql = vsql + " WHERE comandes.client=" + atrim(vcodi) + " and [dataalbara]<=#" + format(vfi, "mm/dd/yy") + "# and [dataalbara]>=#" + format(vinici, "mm/dd/yy") + "#"
  Set rstalb = dbvendes.OpenRecordset(vsql)
  While Not rstalb.EOF
    vunitat = ""
    vcalloff = atrim(rstalb!numcalloff)
    vlinia = atrim(rstalb!datacomanda) + ";"
    vlinia = vlinia + atrim(rstalb!dataalbara) + ";"
    vlinia = vlinia + atrim(cadbl(rstalb!lotinplacsa)) + ";"
    vlinia = vlinia + atrim(cadbl(rstalb!refilate)) + ";"  'fulla de la comanda
    vlinia = vlinia + atrim(cadbl(rstalb!comandaclient)) + ";"
    vlinia = vlinia + atrim(vcalloff) + ";"
    vlinia = vlinia + atrim(rstalb!refclient) + ";"
    vlinia = vlinia + atrim(cadbl(rstalb!tubbaseext)) + ";"
    vlinia = vlinia + atrim(cadbl(rstalb!quantitat)) + ";"
    vlinia = vlinia + atrim(cadbl(rstalb!kgtotalsbruts)) + ";"
    vlinia = vlinia + atrim(Redondejar(cadbl(rstalb!preuvenda), 4)) + ";"
    '!preu = cadbl(rstalb!pvp)
    vlinia = vlinia + atrim(mesurainterna(atrim(rstalb!unitatmesura))) + ";"
    vlinia = vlinia + atrim(Redondejar(rstalb!quantitat * rstalb!preuvenda, 3)) + ";"
    vlinia = vlinia + atrim(rstalb!nome) + ";"
    Print #3, vlinia
    rstalb.MoveNext
  Wend
  Set rstt = Nothing
  Set dbvendes = Nothing
  Set rstalb = Nothing
  Set rstt = Nothing
  Set rsttemp = Nothing
  Set rstobs = Nothing
  ratoli "normal"
End Sub
Function buscarsihihacalloff(vnumc As Double) As String
  Dim rst As Recordset
     Set rst = dbbaixes.OpenRecordset("select numcalloff from bobinesent where comanda=" + atrim(vnumc))
     If Not rst.EOF Then buscarsihihacalloff = atrim(rst!numcalloff)
  Set rst = Nothing
End Function

Sub escullir_client(vcodi As Double, vnomclient As String)
    Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,nom from clients"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(0).width = 1500
  formseleccio.DBGrid2.Columns(1).width = 4000
  'formseleccio.Width = 9000
  'formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           vcodi = formseleccio.DBGrid2.Columns("codi")
           vnomclient = formseleccio.DBGrid2.Columns("nom")
        End If
   End If
End Sub

Sub mveurepressupost_click()
  Dim rst As Recordset
  Dim ruta_documentacio_pressupostos As String
  Dim vnumc As Double
  Dim vnomfitxer As String
   
  vnumc = cadbl(reixa.TextMatrix(reixa.Row, numcol("NºLot")))
  Set rst = dbcomandes.OpenRecordset("SELECT numpressupost, client from comandes where comanda=" + atrim(vnumc) + ";")
  ruta_documentacio_pressupostos = llegir_ini("ruta", "ruta_documentacio_pressupostos", rutadelfitxer(cami) + "valorsprograma.ini")
  vnomfitxer = ruta_documentacio_pressupostos + "\" + atrim(rst!client) + "\" + atrim(rst!numpressupost) + IIf(InStr(1, rst!numpressupost, "_") = 0, "_" + atrim(vnumc), "") + ".pdf"
  If existeix(vnomfitxer) Then Shell "cmd /c start chrome.exe " + vnomfitxer
  
  Set rst = Nothing
End Sub
Sub mveuretarifa_click()
   Dim rst As Recordset
   Dim vnumc As Double
   vnumc = cadbl(reixa.TextMatrix(reixa.Row, numcol("NºLot")))
   Set rst = dbcomandes.OpenRecordset("SELECT tarifes_referencies.coditarifa, tarifes_referencies.codiclient, comandes_extres.comanda FROM comandes_extres RIGHT JOIN tarifes_referencies ON comandes_extres.refinplacsa = tarifes_referencies.refinplacsa WHERE (((comandes_extres.comanda)=" + atrim(vnumc) + "));")
   If Not rst.EOF Then buscar_tarifa_referencies rst!codiclient, cadbl(rst!coditarifa)
   Set rst = Nothing
End Sub
Private Sub btarifaipressupost_Click()
    Me.PopupMenu mtaripres, , btarifaipressupost.Left, btarifaipressupost.Top
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
   Set rst = dbcomandes.OpenRecordset("select grupdeclient from clients where codi=" + atrim(vcodiclient))
   If Not rst.EOF Then
      If atrim(rst!grupdeclient) <> "" Then vcodiclient = atrim(rst!grupdeclient)
       Set rst = Nothing
   End If
   If vruta = "" Then Exit Sub
   If cadbl(vcodiclient) > 0 Then vcodiclient = format(vcodiclient, "000000")
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
      vdir = Dir(vrutacomplerta + "\" + format(cadbl(vcoditarifa), "000") + " -*.*", vbArchive)
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
Private Sub Checknomesclixes_Click()
   filtre_LostFocus 1
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 And Shift = 2 Then
      copiar_dades_del_servidor
      mgeneral_Click
   End If
End Sub

Private Sub Command67_Click(Index As Integer)
   
   
   Formagrupartreballs.Show
End Sub

Private Sub Command8_Click()
  Dim vinici As String
  Dim vfi As String
  Dim vultimdia As Date
  
  If WeekDay(Now, vbMonday) > 1 Then
       vultimdia = DateAdd("d", -1, Now)
        Else: vultimdia = DateAdd("d", -3, Now)
  End If
  vinici = InputBox("Entra la data d'inici de la consulta", "Data inici", format(vultimdia, "dd/mm/yy"))
  If StrPtr(vinici) = 0 Then Exit Sub
  vfi = InputBox("Entra la data de fi de la consulta", "Data fi", format(vultimdia, "dd/mm/yy"))
  If StrPtr(vfi) = 0 Then Exit Sub
  copiaregistreatemporalentregues "[dataalbara]>=#" + format(vinici, "mm/dd/yy") + "# and [dataalbara]<=#" + format(vfi, "mm/dd/yy") + "#"
  mentregues_Click
End Sub

Private Sub fcontrols_Click()
'copiaregistreatemporalentregues
  
   'comprovarsihihaamuntadoraeltreballiavisar 1
   'CreateObject("WScript.Shell").Popup "No es pot borrar el fitxer temporal." + Chr(13) + "Mira que no hi hagi una altra planificació oberta", 1800, "Error", 0
 

 
   
   
End Sub

Private Sub lldescansop_Click()
Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vlimitdetemps As String
  Dim vdatainici As String
  Dim vdatafi As String
  
  Set oapp = New CRAXDDRT.Application
  vdatainici = InputBox("Escriu la data d'inici de la consulta. Ex: 01/06/" + atrim(format(Now, "yy")), "Data inici")
  vdatafi = InputBox("Escriu la data de fi de la consulta. Ex: 07/06/" + atrim(format(Now, "yy")), "Data fi")
  vlimitdetemps = InputBox("Escriu quin avís de limit de temps vols utilitzar (en minuts). EX:30 ", "Limit de temps", 30)
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistat de descansos operaris.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "baixes.mdb"
  oreport.Database.Tables.Item(2).Location = rutadelfitxer(cami) + "comandes.mdb"
  oreport.DiscardSavedData
  oreport.FormulaFields.GetItemByName("DataInici").Text = "#" + format(vdatainici, "mm/dd/yy") + "#"
  oreport.FormulaFields.GetItemByName("DataFi").Text = "#" + format(vdatafi, "mm/dd/yy") + "#"
  oreport.FormulaFields.GetItemByName("limitdetemps").Text = cadbl(vlimitdetemps)
  Load veurereport
  veurereport.CRViewer.ReportSource = oreport
  veurereport.CRViewer.DisplayGroupTree = False
  veurereport.CRViewer.ViewReport
  veurereport.Show 1, Me
  
End Sub

 Private Sub mcomandastandby_click()
    Dim numc As Double
    Dim v As String
    If Screen.ActiveControl.Name <> "reixa" Then MsgBox "Escull una fila primer": Exit Sub
    v = reixa.TextMatrix(reixa.Row, numcol("NºLot"))
    numc = cadbl(v)
    If InStr(1, v, "R") > 2 Then numc = cadbl(Mid(atrim(v), 1, InStr(1, v, "R") - 1))
    If MsgBox("Vols passar la comanda a StandBy?", vbInformation + vbYesNo + vbDefaultButton2, "StandBy") = vbYes Then
        reixa.TextMatrix(reixa.Row, numcol("StandBy")) = "S"
        dbcomandes.Execute "update comandes_extres set passaraimpresores=0 where comanda=" + atrim(numc)
        dbconsulta.Execute "update planificaciototes set standbyimpresio='S' where comanda=" + atrim(numc)
          Else:
            treurestandby numc
    End If
 End Sub
 Sub comprovarsihihaamuntadoraeltreballiavisar(vcomanda As Double)
    Dim rst As Recordset
    Dim vresp As String
    Dim vcos As String
    Dim vnumtreball As Double
    
    Set rst = dbcomandes.OpenRecordset("select numtreball from comandes where comanda=" + atrim(vcomanda))
    If rst.EOF Then GoTo fi
    vnumtreball = cadbl(rst!numtreball)
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
 End Sub

Private Sub bordre_Click()
 etmsgajuda = "Prem sobre la columna que vols ordenar."
 etmsgajuda.width = 3000
 etmsgajuda.Left = multiseleccio.Left + multiseleccio.width + 100
 etmsgajuda.visible = True
 bordre.BackColor = &HFFFF&
 reixa.BackColorFixed = &HFFFF&
End Sub
Private Sub CdLacceptar_Click()
 canviestatCdl "A"
    reixa.CellBackColor = QBColor(10)
End Sub

Private Sub mentregues_Click()
 Dim comanda As Double
 Set dbsap = OpenDatabase(rutadelfitxer(cami) + "connexiosap.mdb")
 'nummaquina = cadbl(msoldadores.tag)
 Frameentregues.visible = True
 comanda = cadbl(mgeneral.tag)
 eseccio = "Entregues"
 carregarmaquines "E"
 taulaplanificacio = "planificacioent"
 carregarllistadecampstemporals "E"
 ordrereixa = " order by dataalbara Desc"
 configreixa
  ' reordenarregistres
 poblarlareixa 0
End Sub

Private Sub mrebutjarCdL_click()
    canviestatCdl "R"
    reixa.CellBackColor = QBColor(12)
End Sub
Private Sub botoreclamar_Click()
   If framereclamar.visible Then
       amagarframereclamar
        Else: framereclamar.visible = True
    End If
End Sub

Private Sub canviordre_Click()
   Dim canvi As Boolean
   Dim i As Double
   Dim ordre As String
   Dim comanda As Double
   For i = seleccionats("inici") To seleccionats("fi")
        comanda = cadbl(reixa.TextMatrix(i, numcol("NºLot")))
        canviarordrecomanda canvi, comanda, ordre
        If canvi Then actualitzarlareixa comanda
   Next i
End Sub

Private Sub Command2_Click()
  amagarframereclamar
End Sub
Sub amagarframereclamar()
  llistadecomandes = ""
  llistacomandes.Clear
  observacionsdemanades = ""
  framereclamar.visible = False
End Sub

Private Sub Command4_Click()
   enviar_mail_reclamarcomandes
   guardar_reclamades
   amagarframereclamar
End Sub
Sub guardar_reclamades()
   Dim i As Byte
   If llistacomandes.ListCount = 0 Then Exit Sub
   While i < llistacomandes.ListCount
     dbbaixes.Execute "delete * from planificacio_reclamades where numcomanda=" + atrim(cadbl(llistacomandes.List(i)))
     dbbaixes.Execute "insert into planificacio_reclamades (numcomanda,datareclamacio) values (" + atrim(cadbl(llistacomandes.List(i))) + ",now)"
     i = i + 1
   Wend
   llistacomandes.Clear
End Sub
Sub enviar_mail_reclamarcomandes()
   Dim vcos As String
   
   If MsgBox("Vols enviar aquest e-mail reclamant comandes?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
       vcos = treure_apostruf(observacionsdemanades + Chr(10) + llistadecomandes)
       enviaremailgeneric "liniesimpresio@inplacsa.com", "Reclam comandes en paper per Impresores", vcos
       MsgBox "Missatge enviat", vbInformation, "Atenció"
   End If
End Sub
Function treuresimbols(desc As String) As String
   desc = substituir(desc, ":", "_")
   desc = substituir(desc, "'", "´")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ";", "_")
   treuresimbols = desc
End Function
Function substituir(cadena As String, buscar As String, canviar As String) As String
   comença = InStr(1, "  " + cadena, buscar) - 1
   If comença < 1 Then substituir = cadena: Exit Function
   acaba = comença + Len(buscar) + 1
   cadena = Mid("  " + cadena, 1, comença) + canviar + Mid("  " + cadena, acaba)
   substituir = cadena
   'MsgBox linia
End Function
Sub enviaremailgeneric(destinatari As String, assumpte As String, cos As String)
   Dim dbenvio As Database
   If atrim(cos) = "" Then Exit Sub
   Set dbenvio = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
   dbenvio.Execute "insert into envios_mails (data,destinatari,assumpte,cos) values (now,'" + destinatari + "','" + treuresimbols(assumpte) + "','" + treuresimbols(cos) + "')"
   Set dbenvio = Nothing
End Sub

Private Sub Command5_Click()
 afegir_comandaperreclamar
End Sub
Sub afegir_comandaperreclamar()
 Dim vnumc As Double
 Dim rst As Recordset
 Dim v As String
  v = reixa.TextMatrix(reixa.Row, numcol("NºLot"))
  If InStr(1, v, "R") > 0 Then MsgBox "Aquesta comanda ja està reclamada.", vbCritical, "Reclamar": Exit Sub
  vnumc = cadbl(v)
  'vnumc = cadbl(Mid(v, 1, IIf(InStr(1, v, "R"), InStr(1, v, "R") - 1, 10)))
  Set rst = dbcomandes.OpenRecordset("select passaraimpresores from comandes_Extres where comanda=" + atrim(vnumc))
  If Not rst.EOF Then If cadbl(rst!passaraimpresores) = 0 Then MsgBox "Aquesta comanda està en Standby de Planificació.", vbExclamation, "Atenció": GoTo fi
  If vnumc > 0 Then llistadecomandes = llistadecomandes + " " + atrim(vnumc): llistacomandes.AddItem atrim(vnumc)
fi:
  Set rst = Nothing
End Sub
Private Sub Command6_Click()
 'planificacio_reclamades
   carregarllistareclamades
   
  Framereclamades.visible = Not Framereclamades.visible
  Framereclamades.Left = 200
  Framereclamades.Top = 200
End Sub
Sub carregarllistareclamades()
   Dim rst As Recordset
   Set rst = dbbaixes.OpenRecordset("select * from planificacio_reclamades order by numcomanda", , ReadOnly)
   llistareclamades.Clear
   While Not rst.EOF
     llistareclamades.AddItem atrim(rst!numcomanda) + IIf(rst!reactivada, "*", "") + Chr(9) + format(rst!datareclamacio, "dd/mmmm")
     llistareclamades.ItemData(llistareclamades.NewIndex) = cadbl(rst!numcomanda)
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub Command7_Click()
   Dim vnumc As String
   If llistareclamades.ListIndex <> -1 Then
        vnumc = atrim(cadbl(llistareclamades.ItemData(llistareclamades.ListIndex)))
        If MsgBox("Vols treure la comanda " + vnumc + " de la llista?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
            dbbaixes.Execute "delete * from planificacio_reclamades where numcomanda=" + vnumc
        End If
        carregarllistareclamades
      Else: MsgBox "Has d'escullir una comanda per treura-la de la llista", vbInformation, "Atenció"
   End If
   
End Sub

Private Sub exportarapdf_Click()
  If taulaplanificacio = "planificacioent" Then Exit Sub
  exportarapdf.tag = "exportar"
  Command3_Click
  obrir_document "c:\temp\llistatexportat.pdf"
  exportarapdf.tag = ""
End Sub

Private Sub exportaraxls_Click()
  If taulaplanificacio <> "planificacioent" Then
        exportaraxls.tag = "exportar"
        Command3_Click
        obrir_document "c:\temp\llistatexportat.xls"
         Else
           exportar_entregues
           obrir_document "c:\temp\llistatexportat.csv"
  End If
  exportaraxls.tag = ""
End Sub
Sub exportar_entregues()
  Dim vfitxertemp As String
  Dim vcol As Double
  Dim vrow As Double
  Dim vlinia As String
  vfitxertemp = "c:\temp\llistatexportat.csv"
  Open vfitxertemp For Output As #3
  For vrow = 0 To reixa.Rows - 1
    vlinia = ""
    For vcol = 0 To reixa.Cols - 1
        vlinia = vlinia + IIf(vlinia = "", "", ";") + reixa.TextMatrix(vrow, vcol)
    Next vcol
    Print #3, vlinia
  Next vrow
  Close #3
End Sub
Sub generareltemporalsical()
  Static heentrat As Boolean
  Dim vTempsGenerantTEMP As String
   If generarelfitxertemporal And Not heentrat Then
      heentrat = True
      If llegir_ini("Planificacio", "EsticGenerantTEMP", "comandes.ini") = "S" Then
        vTempsGenerantTEMP = llegir_ini("Planificacio", "EsticGenerantTEMP_Hora", "comandes.ini")
        If Not IsDate(vTempsGenerantTEMP) Then vTempsGenerantTEMP = "01/01/2000"
        If DateDiff("n", vTempsGenerantTEMP, Now) > 4 Then
               escriure_ini "Planificacio", "EsticGenerantTEMP_Hora", Now, "comandes.ini"
               escriure_ini "Planificacio", "EsticGenerantTEMP", "N", "comandes.ini"
                Else: GoTo fi
        End If
      End If
      escriure_ini "Planificacio", "EsticGenerantTEMP", "S", "comandes.ini"
      timercontrol.Enabled = True
      Command1_Click
      If existeix("c:\temp\planificaciotmp.mdb") Then
         If existeix(rutadelfitxer(cami) + "planificaciotemporal.mdb") Then
            Kill rutadelfitxer(cami) + "planificaciotemporal.mdb"
         End If
         dbplanificacio.Close
         dbconsulta.Close
         Set dbplanificacio = Nothing
         Set dbconsulta = Nothing
         Set dbplanificacioalicia = Nothing
         Set dbplanificaciooperaris = Nothing
         wait 1
         
         Copiar_Fitxer "c:\temp\planificaciotmp.mdb", rutadelfitxer(cami) + "planificaciotemporal.mdb"
         escriure_ini "Planificacio", "EsticGenerantTEMP", "N", "comandes.ini"
         escriure_ini "Planificacio", "ultimaactualitzacio", Now, rutadelfitxer(cami) + "\actualitzacioplanificacio.ini"
      End If
fi:
      heentrat = False
      End
   End If
End Sub


Function buscarcomandavinculada(vCdL As String) As Double
    Dim rst As Recordset
    Set rst = dbconsulta.OpenRecordset("select comanda from planificaciototes where numeroliniaimpresio like '" + Mid(vCdL, 1, 3) + "*' and estat='I'")
    If Not rst.EOF Then buscarcomandavinculada = rst!comanda
    Set rst = Nothing
End Function
Sub canviestatCdl(vestat As String)
    Dim vnumc As Double
    Dim vnumcvinc As String
    vnumc = cadbl(reixa.TextMatrix(reixa.Row, numcol("NºLot")))
    vnumcvinc = buscarcomandavinculada(reixa.TextMatrix(reixa.Row, numcol("NºLinia Imp")))
    If vnumc = vnumcvinc Then Exit Sub
    dbplanificacio.Execute "delete * from estatsCdL where comanda=" + atrim(vnumc)
    dbplanificacio.Execute "insert into estatsCdL (comanda,comandavinculada,estat) values (" + atrim(vnumc) + "," + atrim(vnumcvinc) + ",'" + vestat + "')"


End Sub
Private Sub Form_Click()

'  AcroPDF1.LoadFile "c:\etiqueta_llaunes.pdf"
'  AcroPDF1.src = "c:\etiqueta_llaunes.pdf"
'  ultimaactualitzacio = DateAdd("n", -300, ultimaactualitzacio)
'MsgBox reixa.SelectionMode
  
End Sub

Private Sub Form_Resize()
If planificacio.Height - reixa.Top - 800 < 1 Then Exit Sub
   fcontrols.width = planificacio.width - 500
   fcontrols.Height = planificacio.Height - 1000
   reixa.width = fcontrols.width - 300
   reixa.Height = fcontrols.Height - reixa.Top - 300
   fcontrols.Left = planificacio.width - fcontrols.width - 300
   registres.Top = fcontrols.Height - 300
   If planificacio.tag <> "canvianttamany" Then
    escriure_ini "TamanyFormPlanificacio", "ample", atrim(planificacio.width), iniconfigreixa
    escriure_ini "TamanyFormPlanificacio", "alt", atrim(planificacio.Height), iniconfigreixa
   End If
End Sub
Sub carregartamanyform()
  If cadbl(llegir_ini("TamanyFormPlanificacio", "ample", iniconfigreixa)) > 0 Then
   planificacio.tag = "canvianttamany"
   planificacio.width = llegir_ini("TamanyFormPlanificacio", "ample", iniconfigreixa)
   planificacio.Height = llegir_ini("TamanyFormPlanificacio", "alt", iniconfigreixa)
   planificacio.tag = ""
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   tancartaules
  ' mSubclassForm.Hook
End Sub

Private Sub mSubclassForm_MouseWheel(ByVal wKeys As Long, ByVal zDelta As Long, ByVal XPos As Long, ByVal YPos As Long)
    ' Cuando se mueve la rueda del ratón
    'Me.caption = "Se ha movido la rueda del ratón " + atrim(wKeys) + "  -  " + atrim(zDelta)
    If zDelta = 120 Then
       reixa.TopRow = IIf(reixa.TopRow > 1, reixa.TopRow = reixa.TopRow - 1, reixa.TopRow)
         Else
           reixa.TopRow = IIf(reixa.TopRow < reixa.Rows, reixa.TopRow = reixa.TopRow + 1, reixa.TopRow)
   End If
End Sub


Sub guardar_amples_reixa()
Dim j As Integer
If iniconfigreixa <> "" Then
  For j = 0 To reixa.Cols - 1
   escriure_ini "AmplesReixa", UCase(reixa.TextMatrix(0, j)), atrim(reixa.ColWidth(j)), iniconfigreixa
 Next j
End If
End Sub
Sub carregar_amples_reixa()
 Dim ample As String
 Dim x As Long
 Dim j As Integer
 If iniconfigreixa <> "" Then ' existeix("c:\windows\" + iniconfigreixa) Then
 
  x = reixa.Left + 35
  For j = 0 To reixa.Cols - 1
   ample = llegir_ini("AmplesReixa", UCase(reixa.TextMatrix(0, j)), iniconfigreixa)
   If ample <> "{[}]" Then
    reixa.ColWidth(j) = cadbl(ample)
    If x < reixa.width Then
     filtre(j).Left = x
     filtre(j).width = cadbl(ample)
     If cadbl(ample) > 100 Then filtre(j).visible = True
     filtre(j).ForeColor = &H808080
      Else: If filtre.Count < j - 1 Then filtre(j).visible = False
    End If
    x = x + cadbl(ample)
   End If
 Next j
End If
filtre(0).width = filtre(0).width - 50
filtre(0).Left = filtre(0).Left + 50
End Sub

Function numcol(nom As String) As Byte
   numcol = 0
   For i = 0 To reixa.Cols - 1
     If reixa.TextMatrix(0, i) = nom Then numcol = i: Exit For
   Next i
   
End Function

Function seleccionats(iniciofi As String) As Double
    If reixa.Row > reixa.RowSel Then
      If iniciofi = "inici" Then seleccionats = reixa.RowSel
      If iniciofi = "fi" Then seleccionats = reixa.Row
        Else
          If iniciofi = "inici" Then seleccionats = reixa.Row
          If iniciofi = "fi" Then seleccionats = reixa.RowSel
    End If
      
End Function

Private Sub botomaquina_Click(Index As Integer)
   Dim i As Double
   Dim ordre As String
    If escullirmaqdesti.visible Then
        ordre = InputBox("Entra el numero d'ordre que vols per aquesta comanda.", "Atenció")
        For i = seleccionats("inici") To seleccionats("fi")
         canviarmaquinaalacomanda cadbl(Mid(reixa.TextMatrix(i, numcol("NºLot")), 1, 6)), cadbl(botomaquina(Index).tag), cadbl(ordre)
       Next i
       escullirmaqdesti.visible = False
       reordenarregistres
       'Exit Sub
    End If
   ratoli "espera"
   reixa.visible = False
   ordrereixa = triar_ordre_reixa
    Command3.tag = ""
   
   nummaquina = cadbl(botomaquina(Index).tag)
   For i = 0 To 5
     botomaquina(i).BackColor = &HC0FFC0
   Next i
   botomaquina(Index).BackColor = QBColor(10)
   nummaquina = cadbl(botomaquina(Index).tag)
   carregar_ordre_correcte nummaquina
   possarhoraprevista
   configreixa
   reordenarregistres
   poblarlareixa nummaquina
   If eseccio = "Rebobinadores" And programaoperaris Then comprovarcomandesquenopodenfersealarebobinadoraescullida
   
   ratoli "normal"
   reixa.visible = True
   
 
End Sub
Function maquinapotferzipper(vmaq As Double) As Boolean
    If vmaq = 3 Then maquinapotferzipper = True
End Function
Function comprovar_siporta_Zipper_i_potferla(vnumc As Double, vnovamaquina As Double) As Boolean
   Dim rst As Recordset
   Dim vsql As String
   comprovar_siporta_Zipper_i_potferla = True
   vsql = "SELECT comandes.comanda, Accessoris.Descripcio as A1, Accessoris_1.Descripcio as A2 FROM ((comandes LEFT JOIN Accessoris ON comandes.ansa = Accessoris.codi) LEFT JOIN Accessoris AS Accessoris_1 ON comandes.cinta = Accessoris_1.codi) LEFT JOIN productes ON comandes.producte = productes.codi "
   vsql = vsql + " WHERE comandes.comanda=" + atrim(vnumc) + " AND InStr(1,[ruta],'S')>0"
   Set rst = dbcomandes.OpenRecordset(vsql)
   'Clipboard.Clear
   'Clipboard.SetText vsql
   If Not rst.EOF Then
       If InStr(1, UCase(rst!A1), "ZIP") Or InStr(1, UCase(rst!A2), "ZIP") Then
           If Not maquinapotferzipper(vnovamaquina) Then comprovar_siporta_Zipper_i_potferla = False
       End If
   End If
   Set rst = Nothing
End Function
Sub canviarmaquinaalacomanda(comanda As Double, novamaquina As Double, Optional ordre As Double)
   If cadbl(ordre) = 0 Then
      ordre = 999
     Else: ordre = ordre - 0.1
   End If
   If eseccio = "Laminadores" And programaoperaris Then If Not comprovar_siporta_Zipper_i_potferla(comanda, novamaquina) Then MsgBox "Aquesta màquina escullida no pot fer ZIPPER.", vbCritical, "Error": GoTo fi
   dbplanificacio.Execute "insert into " + taulaplanificacio + " (comanda) values (" + atrim(comanda) + ")"
   dbplanificacio.Execute "update " + taulaplanificacio + " set maquina=" + atrim(novamaquina) + " where comanda=" + atrim(comanda)
   dbplanificacio.Execute "update " + taulaplanificacio + " set ordre=" + passaradecimalpunt(atrim(ordre)) + " where comanda=" + atrim(comanda)
   dbconsulta.Execute "update " + taulaplanificacio + " set maquina=" + atrim(novamaquina) + " where comanda=" + atrim(comanda)
   dbconsulta.Execute "update " + taulaplanificacio + " set ordre=" + passaradecimalpunt(atrim(ordre)) + " where comanda=" + atrim(comanda)
   'dbconsulta.Execute "update planificaciototes set maquina=" + atrim(novamaquina) + " where comanda=" + atrim(comanda)
   dbcomandes.Execute "update comandes set " + campmaquina + "=" + atrim(novamaquina) + " where comanda=" + atrim(comanda)
fi:
End Sub
Function campmaquina() As String
   Select Case taulaplanificacio
     Case "planificacioimp"
        campmaquina = "impressora"
     Case "planificaciolam"
        campmaquina = "laminadora"
     Case "planificacioreb"
        campmaquina = "rebobinadora"
     Case "planificaciosol"
        campmaquina = "soldadora"
 End Select

End Function
Private Sub canvidemaquina_Click()
  If Not botomaquina(0).visible Then Exit Sub
  escullirmaqdesti.visible = Not escullirmaqdesti.visible
End Sub
Sub carregar_comandesnoentregadesatemporal()
  Dim rstc As Recordset
  borrartaulanoentregades
  
  dbcomandes.Execute "SELECT comandes.*,comandes_extres.numpack, comandes_extres.clientvindraarevisarimpresio,comandes_extres.noplanificable,comandes_extres.passaraimpresores into comandesnoentregades IN '" + fitxertemp + "' FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda where proximaseccio<>'T' and isdate(dataactivacio) and not noplanificable "
  dbbaixes.Execute "DELETE comandes.proximaseccio, planificacio_reclamades.* FROM planificacio_reclamades LEFT JOIN comandes ON planificacio_reclamades.numcomanda = comandes.comanda WHERE (((comandes.proximaseccio)<>'I' And (comandes.proximaseccio)<>'E'));"
  
End Sub
Sub borrartaulanoentregades()
  On Error Resume Next
  dbconsulta.Execute "drop table comandesnoentregades"
End Sub
Sub carregaactualitzaciodelservidor()
         dbplanificacio.Close
         dbconsulta.Close
         Set dbplanificacio = Nothing
         Set dbconsulta = Nothing
         Set dbplanificacioalicia = Nothing
         Set dbplanificaciooperaris = Nothing
         Copiar_Fitxer rutadelfitxer(cami) + "planificaciotemporal.mdb", fitxertemp
         escriure_ini "Planificacio", "ultimaactualitzacio", Now, fitxerini
End Sub
Private Sub Command1_Click()
  Dim maq As Byte
  Dim vresp As Double
  Dim vdataservidor As Date
  
  If Command1.tag = "novaactualitzacio" Then
'     carregaactualitzaciodelservidor
     dbconsulta.Close
     Set dbconsulta = Nothing
     mirarsicopiaractualitzaciodelservidor
     Set dbconsulta = OpenDatabase(fitxertemp)
     carregarpestanyapredeterminada
     Command1.tag = ""
     Command1.BackColor = sortir.BackColor
     vdataservidor = CVDate(llegir_ini("Planificacio", "ultimaactualitzacio", rutadelfitxer(cami) + "\actualitzacioplanificacio.ini"))
     escriure_ini "Planificacio", "ultimaactualitzacio", atrim(vdataservidor), "comandes.ini"
     ultimaactualitzacio = vdataservidor
     Exit Sub
  End If
  If Not generarelfitxertemporal Then
     vresp = MsgBox("Vols que actualitzi el servidor?[Si] " + Chr(10) + "o actualització Local? [No]", vbInformation + vbYesNoCancel, "Actualització")
      Else: vresp = vbNo
  End If
  If vresp = vbNo Then
        fcontrols.Enabled = False
        factualitzant.visible = True
        factualitzant.Left = (planificacio.width / 2) - (factualitzant.width / 2)
        ratoli "espera"
        If existeix(fitxertemp) Then
          dbconsulta.Close
          borrartemps
          'Kill fitxertemp
        End If
        crearfitxertemp
        maq = nummaquina
        reixa.visible = True
        DoEvents
        carregar_comandesnoentregadesatemporal
        carregar_comandesatemporal
        treure_ordre_delsregistresnotriats
        possartotesleshoresprevistes
        'recarrearmaquinaseleccionada
        factualitzant.visible = False
        fcontrols.Enabled = True
        escriure_ini "Planificacio", "ultimaactualitzacio", Now, "comandes.ini"
        ultimaactualitzacio = Now
        ratoli "normal"
        'mgeneral_Click
        If Not generarelfitxertemporal Then carregarpestanyapredeterminada
  End If
  If vresp = vbYes Then
           escriure_ini "Planificacio", "forzaractualitzacio", "S", rutadelfitxer(cami) + "\actualitzacioplanificacio.ini"
           MsgBox "Espera que el botó d'actualitzar es possi vermell i fer-hi clic.", vbInformation, "Actualització"
  End If
End Sub
Sub barraprogres(actual As Double, gran As Double)
  Dim factor As Double
  
  factor = (actual * 100) / gran
  liniaprogres.width = factor * ((Shape1.width - 100) / 100)
  'If factor Mod 5 = 0 Then
  '   DoEvents
  'End If
  
End Sub
Sub ensenyarprogres(rsta As Recordset)
    barraprogres rsta.AbsolutePosition, rsta.RecordCount
End Sub
Sub possartotesleshoresprevistes()
  Dim rst As Recordset
  'impresores
   Set rst = dbcomandes.OpenRecordset("select * from maquines where maquina='I' and donadadebaixa =null")
   While Not rst.EOF
     taulaplanificacio = "planificacioimp"
     nummaquina = cadbl(rst!codi)
     carregar_ordre_correcte nummaquina
     possarhoraprevista
     rst.MoveNext
   Wend
  
  
  'laminadores
   Set rst = dbcomandes.OpenRecordset("select * from maquines where maquina='L' and donadadebaixa =null")
   While Not rst.EOF
     taulaplanificacio = "planificaciolam"
     nummaquina = cadbl(rst!codi)
     carregar_ordre_correcte nummaquina
     possarhoraprevista
     rst.MoveNext
   Wend
  'rebobinadores
  Set rst = dbcomandes.OpenRecordset("select * from maquines where maquina='R' and donadadebaixa =null")
   While Not rst.EOF
     taulaplanificacio = "planificacioreb"
     nummaquina = cadbl(rst!codi)
     carregar_ordre_correcte nummaquina
     possarhoraprevista
     rst.MoveNext
   Wend
  'soladores
  Set rst = dbcomandes.OpenRecordset("select * from maquines where maquina='S' and donadadebaixa =null")
   While Not rst.EOF
     taulaplanificacio = "planificaciosol"
     nummaquina = cadbl(rst!codi)
     carregar_ordre_correcte nummaquina
     possarhoraprevista
     rst.MoveNext
   Wend
End Sub
Sub recarrearmaquinaseleccionada(Optional boto As Integer)
    Dim i As Integer
    'Dim boto As Integer
    If cadbl(boto) = 0 Then boto = 254
    
     For i = 0 To 5
      If boto = 254 Then
       If botomaquina(i).BackColor = QBColor(10) Then boto = i
       If cadbl(botomaquina(i).tag) > 0 Then carregar_ordre_correcte cadbl(botomaquina(i).tag)
         Else
           If cadbl(botomaquina(i).tag) = boto Then boto = i: GoTo fi
      End If
      
     Next i
fi:
    If boto <> 254 Then botomaquina_Click boto
    
End Sub
Function tauladhoraris(t As String) As String
  Select Case t
     Case "planificacioimp"
        tauladhoraris = "horarisimpresores"
     Case "planificaciolam"
        tauladhoraris = "horarislaminadores"
     Case "planificacioreb"
        tauladhoraris = "horarisrebobinadores"
     Case "planificaciosol"
        tauladhoraris = "horarissoldadores"
 End Select

End Function
Sub possarhoraprevista()
  Dim rsttemps As Recordset
  Dim rst As Recordset
  Dim rstd As Recordset
  Dim dataihora As Date
  Dim datasql As String
  Dim dataanterior As Date
  Dim rstprog As Recordset
  Dim nouordre As String
  Dim reordenarperdata As Boolean
  If taulaplanificacio = "planificaciototes" Or taulaplanificacio = "reclamacionscomandes" Then Exit Sub
  dataihora = format(Now, "dd/mm/yy hh:00")
  datasql = format(Now, "mm/dd/yy hh:00")
  Set rsttemps = dbplanificacioalicia.OpenRecordset("select * from " + tauladhoraris(taulaplanificacio) + " where  dataihora>=#" + datasql + "# and maquina=" + atrim(nummaquina))
  Set rstprog = dbplanificacioalicia.OpenRecordset("select * from " + taulaplanificacio + " where dataprogramada<>null  and maquina=" + atrim(nummaquina) + " order by dataprogramada")
  Set rst = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + " where  horaprogramada=null and maquina=" + atrim(nummaquina) + " order by ordre,dataimpresio")
  If Not rstprog.EOF Then reordenarperdata = True
  dataanterior = dataihora
  While Not rst.EOF
    afegirtempsadataihora dataihora, cadbl(rst!tempsimpresio), rsttemps
    If Not rstprog.EOF Then
       While dataihora > rstprog!dataprogramada
          dbconsulta.Execute "update " + taulaplanificacio + " set dataimpresio=#" + format(dataanterior, "mm/dd/yy hh:nn") + "# where comanda=" + atrim(rstprog!comanda)
          dbplanificacio.Execute "update " + taulaplanificacio + " set dataprevista=#" + format(dataanterior, "mm/dd/yy hh:nn") + "# where comanda=" + atrim(rstprog!comanda)
          dbconsulta.Execute "update planificaciototes set " + nomcampmaquina + "='" + atrim(nummaquina) + "-" + format(dataanterior, "dd/mm") + "' where comanda=" + atrim(rstprog!comanda)
          nouordre = passaradecimalpunt(rst!ordre - 0.1)
          If rst!ordre = 999 Then nouordre = 999
          dbconsulta.Execute "update " + taulaplanificacio + " set ordre=" + nouordre + " where comanda=" + atrim(rstprog!comanda)
          dataanterior = dataihora
          Set rstd = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(rstprog!comanda))
          If Not rstd.EOF Then
           afegirtempsadataihora dataihora, cadbl(rstd!tempsimpresio), rsttemps
          End If
          rstprog.MoveNext
          If rstprog.EOF Then GoTo cont
       Wend
cont:
    End If
    dbconsulta.Execute "update " + taulaplanificacio + " set dataimpresio=#" + format(dataanterior, "mm/dd/yy hh:nn") + "# where comanda=" + atrim(rst!comanda)
    dbplanificacio.Execute "update " + taulaplanificacio + " set dataprevista=#" + format(dataanterior, "mm/dd/yy hh:nn") + "# where comanda=" + atrim(rst!comanda)
    dbconsulta.Execute "update planificaciototes set " + nomcampmaquina + "='" + atrim(nummaquina) + IIf(rst!ordre <> 999, "-" + format(dataanterior, "dd/mm"), "") + "' where comanda=" + atrim(rst!comanda)
    dataanterior = dataihora
    rst.MoveNext
  Wend
  'If reordenarperdata Then canviarordreperdata
End Sub
Function nomcampmaquina() As String
    If taulaplanificacio = "planificacioimp" Then nomcampmaquina = "impresora"
    If taulaplanificacio = "planificaciolam" Then nomcampmaquina = "laminadora"
    If taulaplanificacio = "planificacioreb" Then nomcampmaquina = "rebobinadora"
    If taulaplanificacio = "planificaciosol" Then nomcampmaquina = "soldadora"
End Function
Sub canviarordreperdata()
   Dim rst As Recordset
   Dim cont As Long
   Set rst = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + " where (horaprogramada<>null and maquina=" + atrim(nummaquina) + ") or  maquina=" + atrim(nummaquina) + "  order by dataimpresio")
   cont = 1
   While Not rst.EOF
    If rst!ordre < 999 Or IsDate(rst!horaprogramada) Then
      dbconsulta.Execute "update " + taulaplanificacio + " set ordre=" + atrim(cont) + " where id=" + atrim(rst!id)
    ' dbplanificacio.Execute "update planificacio set ordre=" + atrim(cont) + " where comanda=" + atrim(rst!comanda)
    ' dbplanificacio.Execute "update planificacio set maquina=" + atrim(nummaquina) + " where comanda=" + atrim(rst!comanda)
    End If
     cont = cont + 1
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub afegirtempsadataihora(data As Date, temps As Double, rsttemps As Recordset)
   Dim hores As Double
   Dim minuts As Double
   Dim minutsdata As Double
   Dim diferencialm As Double
   If rsttemps.EOF Or temps = 0 Then Exit Sub
   If temps < 0 Then temps = temps * -1
   minutsdata = data
   hores = Int(temps / 60)
   minuts = temps - (hores * 60)
   'minuts = minutsdata + minuts
   
   For i = 1 To hores
     If Not rsttemps.EOF Then
       rsttemps.MoveNext
       If Not rsttemps.EOF Then data = format(rsttemps!dataihora, "dd/mm/yy hh:" + format(minutsdata, "nn"))
     End If
   Next i
   
   If minuts + format(minutsdata, "nn") >= 60 And Not rsttemps.EOF Then
       rsttemps.MoveNext
       minuts = minuts - format(minutsdata, "nn")
       If minuts < 0 Then minuts = minuts * -1
       If Not rsttemps.EOF Then data = format(rsttemps!dataihora, "dd/mm/yy hh:00")
       data = DateAdd("n", minuts, data)
         Else: data = DateAdd("n", minuts, data)
   End If
   
   
   
   
End Sub
Sub afegirtempsadataihora2(data As Date, temps As Double, rsttemps As Recordset)
   Dim hores As Double
   Dim minuts As Double
   Dim minutsdata As Double
   If rsttemps.EOF Or temps = 0 Then Exit Sub
   minutsdata = format(data, "n")
   hores = Int(temps / 60)
   minuts = temps - (hores * 60)
   'minuts = minutsdata + minuts
   If minuts + minutsdata >= 60 Then
       minuts = minuts + minutsdata - 60
       rsttemps.MoveNext
       If Not rsttemps.EOF Then data = format(rsttemps!dataihora, "dd/mm/yy hh:00")
     Else: If minuts = 0 Then minuts = minutsdata
   End If
   
   For i = 1 To hores
     If Not rsttemps.EOF Then
       rsttemps.MoveNext
       If Not rsttemps.EOF Then data = format(rsttemps!dataihora, "dd/mm/yy hh:00")
     End If
   Next i
   
   data = DateAdd("n", minuts, data)
End Sub
Sub carregar_ordre_correcte(nummaq As Byte)
  Dim rst As Recordset
  Dim rsttotes As Recordset
  Dim vcomandesactualitzades As String
  
  Dim dataprog As String
  
  If taulaplanificacio = "planificaciototes" Or taulaplanificacio = "reclamacionscomandes" Then GoTo planificaciototes
 'posso lordre i la dataprogramada que es la que volem forçar
  Set rst = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where ordre>0 and maquina=" + atrim(nummaq) + " order by ordre")
  While Not rst.EOF
     dataprog = IIf(IsNull(rst!dataprogramada), "null", "#" + format(rst!dataprogramada, "mm/dd/yy hh:nn") + "#")
     If rst!maquina > 0 Then
           vnummaqprogramat = rst!maquina
              Else
               vnummaqprogramat = nummaq
     End If
     dbconsulta.Execute "update " + taulaplanificacio + " set ordre=" + passaradecimalpunt(rst!ordre) + ", maquina=" + atrim(vnummaqprogramat) + ",horaprogramada=" + dataprog + " where comanda=" + atrim(rst!comanda)
     vcomandesactualitzades = vcomandesactualitzades + IIf(vcomandesactualitzades <> "", ",", "") + atrim(rst!comanda)
     rst.MoveNext
  Wend
  If programaoperaris And UCase(taulaplanificacio) = "PLANIFICACIOIMP" Then
      If vcomandesactualitzades = "" Then vcomandesactualitzades = "0"
      dbconsulta.Execute "update planificacioimp set maquina=7 where tipusimpresio='N' and comanda not in (" + atrim(vcomandesactualitzades) + ")"
      dbconsulta.Execute "update planificacioimp set maquina=9 where tipusimpresio='T' and comanda not in (" + atrim(vcomandesactualitzades) + ")"
  End If
  
  'posso la data prevista de fabricacio segons lordre que ha donat l'operari
  Set rst = dbplanificaciooperaris.OpenRecordset("select * from " + taulaplanificacio + " where ordre>0 order by ordre")
  dbconsulta.Execute "update " + taulaplanificacio + " set dataoperari=null"
  While Not rst.EOF
     If rst!ordre < 999 Then
      dataprog = IIf(IsNull(rst!dataprevista), "null", "#" + format(rst!dataprevista, "mm/dd/yy hh:nn") + "#")
      dbconsulta.Execute "update " + taulaplanificacio + " set dataoperari=" + dataprog + " where comanda=" + atrim(rst!comanda)
     End If
     rst.MoveNext
  Wend
  Set rst = Nothing
  Exit Sub
planificaciototes:
  'posso la data prevista de fabricacio segons lordre que ha donat l'operari de rebobinadores a General
  ''Set rst = dbplanificaciooperaris.OpenRecordset("select * from planificacioreb where ordre>0  order by ordre")
  ''dbconsulta.Execute "update planificaciototes set dataoperari=null"
  ''While Not rst.EOF
  ''  If rst!ordre < 999 Then
  ''   dataprog = IIf(IsNull(rst!dataprevista), "null", "#" + format(rst!dataprevista, "mm/dd/yy hh:nn") + "#")
  ''   dbconsulta.Execute "update planificaciototes set dataoperari=" + dataprog + " where comanda=" + atrim(rst!comanda)
  ''  End If
  ''  rst.MoveNext
  ''Wend
  Set rst = Nothing
End Sub
Sub treure_ordre_delsregistresnotriats()
  Dim registresreactivats As Recordset
' On Error Resume Next
   dbplanificacio.Execute "update planificacioimp set ordre=0  where comanda not in (select comanda from planificacioimp IN '" + fitxertemp + "')"
   dbplanificacio.Execute "update planificaciolam set ordre=0  where comanda not in (select comanda from planificaciolam IN '" + fitxertemp + "')"
's'ha dactivar quan hi hagi rebobinadora i soldadora
   'dbplanificacio.Execute "update planificacioreb set ordre=0  where comanda not in (select comanda from planificacioreb IN '" + fitxertemp + "')"
   'dbplanificacio.Execute "update planificaciosol set ordre=0  where comanda not in (select comanda from planificaciosol IN '" + fitxertemp + "')"
   
   Set registresreactivats = dbplanificacio.OpenRecordset("select * from  historicplanificaciototes where comanda in (select comanda from  planificaciototes IN'" + fitxertemp + "')")
   dbplanificacio.Execute "delete * from  historicplanificaciototes where comanda not in (select comanda from planificaciototes IN '" + fitxertemp + "')"
   dbplanificacio.Execute "insert into historicplanificaciototes select * from planificaciototes where comanda not in (select comanda from planificaciototes IN '" + fitxertemp + "')"
   dbplanificacio.Execute "delete * from  planificaciototes where comanda not in (select comanda from planificaciototes IN '" + fitxertemp + "')"
   dbplanificacio.Execute "delete * from  planificacioimp where comanda not in (select comanda from planificacioimp IN '" + fitxertemp + "')"
   dbplanificacio.Execute "delete * from  planificaciolam where comanda not in (select comanda from planificaciolam IN '" + fitxertemp + "')"
   dbplanificacio.Execute "delete * from  planificacioreb where comanda not in (select comanda from planificacioreb IN '" + fitxertemp + "')"
   dbplanificacio.Execute "delete * from  planificaciosol where comanda not in (select comanda from planificaciosol IN '" + fitxertemp + "')"
   If Not generarelfitxertemporal Then ensenyarreactivats registresreactivats
End Sub
Sub ensenyarreactivats(rstr As Recordset)
   Dim msg As String * 1000
   
   msg = ""
   While Not rstr.EOF
     msg = "Comanda: " + atrim(rstr!comanda) + "    Data2:   " + format(rstr!data2, "dd/mm/yy") + Chr(10)
     rstr.Delete
     rstr.MoveNext
   Wend
   If atrim(msg) <> "" Then MsgBox atrim(msg), vbInformation, "Comandes Reactivades"
End Sub
Sub reordenarregistres()
   Dim rst As Recordset
   Dim cont As Long
   If taulaplanificacio = "planificaciototes" Or taulaplanificacio = "reclamacionscomandes" Then Exit Sub
   Set rst = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + " where maquina=" + atrim(nummaquina) + " and (ordre>0 and ordre<999) order by ordre", dbOpenSnapshot, dbReadOnly)
   cont = 1
   While Not rst.EOF
     dbconsulta.Execute "update " + taulaplanificacio + " set ordre=" + atrim(cont) + " where id=" + atrim(rst!id)
     dbplanificacio.Execute "update " + taulaplanificacio + " set ordre=" + atrim(cont) + " where comanda=" + atrim(rst!comanda)
     dbplanificacio.Execute "update " + taulaplanificacio + " set maquina=" + atrim(nummaquina) + " where comanda=" + atrim(rst!comanda)
     'dbplanificacio.Execute "update planificacio set imp_observacio='" + atrim(rst!observacions) + "' where comanda=" + atrim(rst!comanda)
     'dbplanificacio.Execute "update planificacio set imp_importancia=" + atrim(cadbl(rst!importancia)) + " where comanda=" + atrim(rst!comanda)
     'If IsDate(rst!Data2) Then dbplanificacio.Execute "update planificacio set imp_data2=#" + Format(rst!Data2, "mm/dd/yy") + "# where comanda=" + atrim(rst!comanda)
     cont = cont + 1
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Function esmaterialexacte(numc As Double) As Boolean
   Dim rstc As Recordset
   esmaterialexacte = False
   Set rstc = dbcomandes.OpenRecordset("SELECT materialexacte FROM comandes_extres Where comanda = " + atrim(numc), dbOpenSnapshot, dbReadOnly)
   If Not rstc.EOF Then
       If rstc!materialexacte Then esmaterialexacte = True
   End If
   Set rstc = Nothing
End Function
Sub copiar_dades_del_servidor()
    Dim vdataservidor As Date
    dbconsulta.Close
    Set dbconsulta = Nothing
    If existeix(rutadelfitxer(cami) + "\planificaciotemporal.mdb") Then
           vdataservidor = CVDate(llegir_ini("Planificacio", "ultimaactualitzacio", rutadelfitxer(cami) + "\actualitzacioplanificacio.ini"))
           If existeix(fitxertemp) Then Kill fitxertemp
           FileCopy rutadelfitxer(cami) + "\planificaciotemporal.mdb", fitxertemp
           escriure_ini "Planificacio", "ultimaactualitzacio", atrim(vdataservidor), "comandes.ini"
    End If
    Set dbconsulta = OpenDatabase(fitxertemp)
    
End Sub
Sub poblarlareixa(nummaquina As Byte, Optional were As String)
  Dim i As Byte
  Dim fila As Integer
  Dim col As Byte
  Dim rst As Recordset
  Dim rstexpedicions As Recordset
  Dim rstreclamades As Recordset
  Dim apuntxrimprimir As Double
  Dim tenimmaterial As Boolean
  Dim tenimclixes As Boolean
  Dim textetaula As String
  Dim vdataservidor As Date
  Dim rsttotes As Recordset
  Dim vTotalmetres As Double
  Dim vcoloradhesiu As String
  Dim vcolorceldaPC2 As Double
  Dim colorcelda As Double
  Dim vc As Long
  Dim vunitat As String
  Dim vrutaFT As String
  vrutaFT = "\\ord_copies\DadesProduccio\Arxius Produccio\DadesGenerals\FitxesTecniquesRefInplacsa"
inici:
  ratoli "espera"
  reixa.visible = False
  reixa.Clear
  reixa.Redraw = False
  reixa.BackColor = QBColor(15)
  configreixa IIf(were <> "", True, False)
 ' reixa.Rows = 0
  reixa.Rows = 1
  vhihanCdLtaronja = False
  Select Case taulaplanificacio
     Case "planificaciototes"
      Set rst = dbconsulta.OpenRecordset("select * from planificaciototes where comanda>1 " + were + ordrereixa, dbOpenSnapshot, dbReadOnly)
     Case "reclamacionscomandes"
      Set rst = dbconsulta.OpenRecordset("select * from reclamacionscomandes where comanda>1 " + were + ordrereixa, dbOpenSnapshot, dbReadOnly)
     Case "planificacioent"
'       ordrereixa = substituir(ordrereixa, "data1", "comanda")
'       ordrereixa = substituir(ordrereixa, "ordre", "comanda")
'       ordrereixa = substituir(ordrereixa, "dataimpresio", "comanda")
      Set rst = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + " where comanda>1 " + were + ordrereixa, dbOpenSnapshot, dbReadOnly)
     Case Else
      Set rst = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + " where maquina=" + atrim(nummaquina) + were + ordrereixa, dbOpenSnapshot, dbReadOnly)
      
  End Select
  Set rstCdL = dbconsulta.OpenRecordset("SELECT planificaciototes.numeroliniaimpresio, planificaciototes.estat, IIf([planificaciototes].[datacalloff],'S','N') AS tecalloff, planificaciototes.codiclient From planificaciototes WHERE (((Mid([numeroliniaimpresio],1,3)) In (SELECT Mid([numeroliniaimpresio],1,3) AS Expr1 From planificaciototes GROUP BY Mid([numeroliniaimpresio],1,3) HAVING (((Mid([numeroliniaimpresio],1,3)) Is Not Null And (Mid([numeroliniaimpresio],1,3))>'0') AND ((Count(planificaciototes.comanda))>1));))) ORDER BY planificaciototes.estat;")
  Set rstCdLestats = dbplanificacioalicia.OpenRecordset("select * from EstatsCdL")
  dbplanificacioalicia.Execute "delete * from EstatsCdL where comanda in (SELECT EstatsCdL.comanda FROM EstatsCdL LEFT JOIN comandes ON EstatsCdL.comandavinculada = comandes.comanda WHERE (((comandes.proximaseccio)<>'I'));)"
  Set rsttotes = dbconsulta.OpenRecordset("select * from planificaciototes", dbOpenSnapshot, dbReadOnly)
  Set rstexpedicions = dbplanificacio.OpenRecordset("select * from linies_expedicions")
  Set rstreclamades = dbbaixes.OpenRecordset("select * from planificacio_reclamades")
  If rst.EOF Then
     If MsgBox("No hi ha cap registre." + Chr(10) + "Vols recarregar les ultimes dades del servidor?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
        copiar_dades_del_servidor
        GoTo inici
     End If
     Exit Sub
  End If
  fila = 0
  reixa.tag = "poblant"
  
  
  While Not rst.EOF
   'Me.caption = atrim(rst!comanda): DoEvents
   fila = fila + 1
   reixa.Rows = fila + 1
   
   tenimmaterial = False
   tenimclixes = False
   col = 0
   If taulaplanificacio <> "planificaciosol" And taulaplanificacio <> "planificacioent" Then
       vTotalmetres = vTotalmetres + cadbl(rst!mts)
     Else: If taulaplanificacio <> "planificacioent" Then vTotalmetres = vTotalmetres + cadbl(rst!quantitatsol)
   End If
   For i = 0 To rst.Fields.Count - 1
    If camps(i + 1, 1) <> "" And camps(i + 1, 4) <> "N" Then
      
      reixa.TextMatrix(fila, col) = IIf(IsNull(rst.Fields(camps(i + 1, 1))), "", rst.Fields(camps(i + 1, 1)))
     
      
     'canvio el color si hi ha data a material
      If camps(i + 1, 1) = "material" Or camps(i + 1, 1) = "materialPC" Or camps(i + 1, 1) = "materialPC2" Then
       reixa.col = col
       reixa.Row = fila
       If Len(rst.Fields(camps(i + 1, 1))) = 8 Then
          reixa.CellBackColor = &H80C0FF 'taronja
         Else:
           If atrim(rst.Fields(camps(i + 1, 1))) <> "" Then
               reixa.CellBackColor = QBColor(10) 'verd
               tenimmaterial = True
           End If
       End If
       GoTo format
      End If
      If camps(i + 1, 1) = "capes" Then reixa.TextMatrix(fila, col) = IIf(cadbl(rst!capes) > 10, rst!capes - 10, rst!capes)
      If camps(i + 1, 1) = "tintesrevisades" Then
           If rst!tintesrevisades = "C" Then
             reixa.col = col
             reixa.Row = fila
             reixa.CellBackColor = &HFF& 'vermell
           End If
      End If
      If camps(i + 1, 1) = "preuclixes" Then
           If cadbl(rst!preuclixes) = 0 Then
               If hihaalbaransdefotogravador(rst!comanda) Then
                    reixa.col = col
                    reixa.Row = fila
                    reixa.CellBackColor = &HFF& 'vermell
               End If
           End If
      End If
      
      'canvio el color de estat si està acabada
      If camps(i + 1, 1) = "comanda" And Not programaoperaris Then
      
        If comandaacabada(rst.Fields("comanda"), Mid(eseccio, 1, 1)) Then
          reixa.col = col
          reixa.Row = fila
          reixa.CellBackColor = &HFF& 'vermell
        End If
        If baixasensefuncionament(rst.Fields("comanda"), Mid(eseccio, 1, 1)) Then
          reixa.col = col
          reixa.Row = fila
          reixa.CellBackColor = &HFF8080  'blau clar
        End If
        If taulaplanificacio = "planificaciototes" Then
            rstreclamades.FindFirst "numcomanda=" + atrim(rst!comanda)
            If Not rstreclamades.NoMatch Then
                If rst.Fields("material") = "A" Then
                        rstreclamades.Delete
                     Set rstreclamades = dbbaixes.OpenRecordset("select * from planificacio_reclamades")
                    Else: reixa.TextMatrix(fila, col) = reixa.TextMatrix(fila, col) + "R" + IIf(rstreclamades!reactivada, "a", "")
                End If
            End If
        End If
        GoTo format
      End If
      'si es refclient i es el d'operaris poso el numero de treball
      If camps(i + 1, 1) = "refclient" And programaoperaris And eseccio = "Impresores" Then reixa.TextMatrix(fila, col) = rst.Fields("numtreball")
      'canvio el color del material si es material exacte
      If camps(i + 1, 1) = "data1" And Not programaoperaris Then
          If esmaterialexacte(rst.Fields("comanda")) Then
           reixa.col = col
           reixa.Row = fila
           reixa.CellBackColor = &HFF8080 'blau clar
          End If
          GoTo format
       End If
       If camps(i + 1, 1) = "data2" And taulaplanificacio = "planificaciototes" Then
           If rst!estatclixes = "Nova" Or rst!estatclixes = "Modificada" Then
                If Not existeix(vrutaFT + "\FT-" + atrim(rst!refinplacsa) + ".pdf") Then
                    reixa.col = col
                    reixa.Row = fila
                    reixa.CellBackColor = &H80FFFF    'groc clar
                End If
           End If
      End If
      
      'canvio el color de estat si no es l'estat de comanda real
      If camps(i + 1, 1) = "estat" Then

        If Mid(rst.Fields("estat"), 2, 1) = Chr(255) Then
          reixa.col = col
          reixa.Row = fila
          reixa.CellBackColor = &HFF& 'vermell
          reixa.Text = Mid(rst.Fields("estat"), 1, 1)
        End If
        GoTo format
      End If
      
      'canvio el color de clientvindraarevisarimpresio
      If camps(i + 1, 1) = "clientvindraarevisarimpresio" Then
        If rst.Fields("clientvindraarevisarimpresio") <> "" And rst.Fields("clientvindraarevisarimpresio") <> "S" Then
          reixa.col = col
          reixa.Row = fila
          reixa.CellBackColor = &H5C31DD    'vermell xulu
        End If
        GoTo format
      End If
      If camps(i + 1, 1) = "numeroliniaimpresio" Then 'And taulaplanificacio = "planificaciototes" Then
          reixa.col = col
          reixa.Row = fila
          reixa.CellBackColor = colorLiniaImpresio(rst, rsttotes)
      End If
      
      'canvio el color de dataimpresio si hi ha data programada
      If camps(i + 1, 1) = "dataimpresio" Then
       If Not IsNull(rst.Fields("horaprogramada")) Then
        reixa.col = col
        reixa.Row = fila
        reixa.CellBackColor = &H80C0FF 'taronja
       End If
       GoTo format
      End If

      
      'canvio el la font de la filera impresio
      If camps(i + 1, 1) = "impresio" Then
       reixa.col = col
       reixa.Row = fila
       reixa.CellFontBold = True
       GoTo format
      End If
      
      'posso en vermell el texte si es reimpresió
      If camps(i + 1, 1) = "texteimpresio" Then
         reixa.col = col
         reixa.Row = fila
         If Mid(reixa.TextMatrix(fila, col), 1, 1) = "*" Then
             reixa.CellBackColor = QBColor(12) 'vermell clar
         End If
         If eseccio = "Impresores" Then If cadbl(rst!capes) > 10 Then reixa.CellForeColor = QBColor(13)
      End If
      
     
     'canvio la data si hi ha data de clixe
      If camps(i + 1, 1) = "clixes" Then
       If reixa.TextMatrix(fila, col) = "0:00:00" Then
          reixa.TextMatrix(fila, col) = ""
          reixa.CellBackColor = QBColor(15) 'blanc
        Else
          reixa.col = col
          reixa.Row = fila
          If Mid(reixa.TextMatrix(fila, col), 1, 1) = "*" Then
              reixa.CellBackColor = QBColor(10) 'verd   'vol dir que tenim el clixe
              tenimclixes = True
                Else:
                      If Len(atrim(reixa.TextMatrix(fila, col))) > 4 Then
                           reixa.CellBackColor = &H80C0FF 'taronja    vol dir que hiha data d'entrega del clixé per part del fotogravador
                          Else: reixa.CellBackColor = QBColor(15) 'blanc
                      End If
                      If Mid(reixa.TextMatrix(fila, col), 1, 1) = "#" Then
                        reixa.CellBackColor = QBColor(9) 'blau clar   vol dir que el clixe està apunt a l'espera de comanda
                        tenimclixes = True
                      End If
                      If Mid(reixa.TextMatrix(fila, col), 1, 1) = "!" Then
                        reixa.CellBackColor = QBColor(12) 'vermell clar   vol dir que el clixe els te el client
                        tenimclixes = False
                      End If
          End If
       End If
      End If
format:   ' apartir d'aqui aplico el format a la casella
      'posso el format del camp dataimpresio
      If camps(i + 1, 2) = "date" Then
        If camps(i + 1, 1) <> "dataexpedicio" And camps(i + 1, 1) <> "dataimpresio" And camps(i + 1, 1) <> "dataoperari" And camps(i + 1, 1) <> "datareclamacio" And camps(i + 1, 1) <> "datagestio" Then
          If reixa.TextMatrix(fila, col) = "0:00:00" Then
            reixa.TextMatrix(fila, col) = ""
                Else: reixa.TextMatrix(fila, col) = format(reixa.TextMatrix(fila, col), "dd/mm/yy")
          End If
            Else:
              If camps(i + 1, 1) = "dataexpedicio" Then
                    reixa.TextMatrix(fila, col) = format(reixa.TextMatrix(fila, col), "dd/mm/yy")
                    If reixa.TextMatrix(fila, col) <> "" Then
                        If DateDiff("d", Now, reixa.TextMatrix(fila, col)) < 1 Then
                            rstexpedicions.FindFirst "comanda=" + atrim(rst.Fields("comanda")) + " and data=#" + format(reixa.TextMatrix(fila, col), "mm/dd/yy") + "#"
                            reixa.col = col
                            reixa.Row = fila
                            If Not rstexpedicions.NoMatch Then
                                If Not rstexpedicions!enviat Then reixa.CellBackColor = QBColor(12) 'vermell clar   vol dir que no s'ha enviat encara
                                  Else: reixa.CellBackColor = QBColor(12) 'vermell clar   vol dir que no s'ha enviat encara
                            End If
                        End If
                    End If
                 Else: reixa.TextMatrix(fila, col) = format(reixa.TextMatrix(fila, col), "dd/mm hh:nn")
              End If
        End If
      End If
      
       'posso el format del camp tempsimpresio EXTRACOST I NEGATIUS
      If camps(i + 1, 1) <> "extracost" And camps(i + 1, 1) <> "Temps" And cadbl(reixa.TextMatrix(fila, col)) < 0 Then
           reixa.col = col
           reixa.Row = fila
    '       If reixa.col <> numcol("NºAlbarà") Then
              reixa.TextMatrix(fila, col) = cadbl(reixa.TextMatrix(fila, col)) * -1
              reixa.CellBackColor = &H80C0FF 'taronja
     '      End If
      End If
      If camps(i + 1, 1) = "kgimpost" Then
          If exempt_impost(rst!comanda) Then
           reixa.col = col
           reixa.CellBackColor = &HC78DFA
          End If
          If ImpostPlasticNoSurtAlbara(rst!comanda) Then
              reixa.col = col
             reixa.CellBackColor = &HFFC0FF
          End If
      End If
      If camps(i + 1, 1) = "quantitatTeorica" Or camps(i + 1, 1) = "quantitatEntregada" Then
           vunitat = Mid(rst!tipusunitat, InStr(1, rst!tipusunitat, "/") + 1)
           reixa.TextMatrix(fila, col) = atrim(cadbl(reixa.TextMatrix(fila, col))) + " " + vunitat
      End If
      
      If camps(i + 1, 1) = "preu" Then
           reixa.TextMatrix(fila, col) = reixa.TextMatrix(fila, col) + "€"
           If rst!Facturat = "N" Or rst!Facturat = "" Then
            If mirarpreucomanda(rst!comanda) <> rst!Preu Then
             If rst!Preu <> -1 Then
              reixa.col = col
              reixa.CellBackColor = &H80C0FF
             End If
            End If
              Else
                 If mirarpreufactura(rst!comanda, rst!dataalbara) <> mirarpreucomanda(rst!comanda) Then
                   reixa.col = col
                   reixa.CellBackColor = QBColor(12)
                 End If
           End If
      End If
      If camps(i + 1, 1) = "tanx100kgvs" Or camps(i + 1, 1) = "tanx100impostvs" Then
           reixa.TextMatrix(fila, col) = reixa.TextMatrix(fila, col) + "%"
      End If
      'posso color extracost si es negatiu (vol dir que té observació)
      If camps(i + 1, 1) = "extracost" Then
          If cadbl(reixa.TextMatrix(fila, col)) < 0 Then
           reixa.col = col
           reixa.Row = fila
           reixa.CellBackColor = &HFFC0FF
           reixa.TextMatrix(fila, col) = cadbl(reixa.TextMatrix(fila, col)) * -1
           If reixa.TextMatrix(fila, col) = 1 Then reixa.TextMatrix(fila, col) = 0
         End If
      End If
       'posso el format del camp tipusdeadhesius de laminadora
      If camps(i + 1, 1) = "tipuscola" Then
         vcoloradhesiu = 0
         colorcelda = 0
         vcolorceldaPC2 = 0
         If Mid(reixa.TextMatrix(fila, col), 1, 1) = "@" Then
           reixa.col = numcol("TipusCola")
           reixa.Row = fila
           vcoloradhesiu = (Mid(reixa.TextMatrix(fila, col), 2, InStr(2, reixa.TextMatrix(fila, col), " ") - 1))
           If InStr(1, vcoloradhesiu, "/") > 0 Then
                 vcolorceldaPC2 = Mid(vcoloradhesiu + "  ", InStr(1, vcoloradhesiu, "/") + 1)
                 vcoloradhesiu = Mid(vcoloradhesiu, 1, InStr(1, vcoloradhesiu + "  ", "/") - 1)
           End If
           colorcelda = cadbl(vcoloradhesiu)
           reixa.TextMatrix(fila, col) = Mid(reixa.TextMatrix(fila, col), InStr(2, reixa.TextMatrix(fila, col), " "))
           reixa.col = numcol("Nom Client")
           reixa.CellBackColor = cadbl(colorcelda)
           reixa.col = numcol("TipusCola")
           If vcolorceldaPC2 > 0 Then
                  vc = vcolorceldaPC2
                 Else: vc = colorcelda
           End If
           reixa.CellBackColor = vc
         End If
      End If
      If camps(i + 1, 1) = "descmat" Then   'subratllo el material si no hi ha material a LAMINADORA
         If faltamaterialdelcomplexa(rst!comanda) Then reixa.col = numcol("Desc. Material"): reixa.CellFontUnderline = True
     End If
      col = col + 1
       ' Else: Stop
    End If
   Next i
   If taulaplanificacio = "planificacioimp" Then
     If (rst!impresio = "R") And tenimmaterial Then apuntxrimprimir = apuntxrimprimir + 1
     If (rst!impresio = "M" Or rst!impresio = "N" Or rst!impresio = "F") And tenimmaterial And tenimclixes Then apuntxrimprimir = apuntxrimprimir + 1
     textetaula = "IMPRIMIR"
     
   End If
   If taulaplanificacio = "planificaciolam" Then
     If (rst!estat = "L") And tenimmaterial Then apuntxrimprimir = apuntxrimprimir + 1
     
     textetaula = "LAMINAR"
   End If
   If taulaplanificacio = "planificacioreb" Then
     If (rst!estat = "R") Then apuntxrimprimir = apuntxrimprimir + 1
     textetaula = "REBOBINAR"
   End If
   If taulaplanificacio = "planificaciosol" Then
     If (rst!estat = "S") Then apuntxrimprimir = apuntxrimprimir + 1
     textetaula = "SOLDADORES"
   End If
   rst.MoveNext
  Wend
 ' If taulaplanificacio = "planificaciototes" Then
  possar_comandes_parcials
  registres = atrim(rst.RecordCount) + " Comandes (" + atrim(format(vTotalmetres, "#,##0")) + IIf(taulaplanificacio <> "planificaciosol", " Metres) ", " Unitats) ") + IIf(apuntxrimprimir > 0, "  - Apunt per " + textetaula + ": " + atrim(apuntxrimprimir), "")
  Set rst = Nothing
  Set rstreclamades = Nothing
  Set rstCdL = Nothing
  reixa.tag = ""
  reixa.Redraw = True
  reixa.visible = True
  ratoli "normal"
  
End Sub
Function ImpostPlasticNoSurtAlbara(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select direnvio from comandes where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
       Set rst = dbcomandes.OpenRecordset("select impostinclosalpvp,pais from clients_envios where id=" + atrim(rst!direnvio))
       If Not rst.EOF Then
           If cabool(rst!impostinclosalpvp) Or rst!pais <> "ES" Then
                  ImpostPlasticNoSurtAlbara = True
           End If
       End If
   End If
   Set rst = Nothing
End Function
Function exempt_impost(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select refinplacsa from comandes_extres where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
       Set rst = dbcomandes.OpenRecordset("select * from tarifes_referencies where refinplacsa='" + atrim(rst!refinplacsa) + "'")
       If Not rst.EOF Then
           If atrim(rst!impost_regimenfiscal) <> "" Then
                  exempt_impost = True
           End If
       End If
   End If
   Set rst = Nothing
End Function
Function hihaalbaransdefotogravador(vnumc As Double) As Boolean
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
        Set rst = dbclixes.OpenRecordset("select * from clixes_albarans where id_treball=" + atrim(cadbl(rst!numtreball)) + " and ordremodificacio=" + atrim(cadbl(rst!numordremodificacio)) + "and facturat=false and import>0")
        If Not rst.EOF Then hihaalbaransdefotogravador = True
  End If
  Set rst = Nothing
End Function
Function faltamaterialdelcomplexa(vnumc As Double) As Boolean
  Dim rst As Recordset
  faltamaterialdelcomplexa = False
  'If vnumc = 221405 Then Stop
  Set rst = dbconsulta.OpenRecordset("select materialpc,materialpc2 from planificaciototes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
      If rst!materialpc <> "A" And rst!materialpc <> "R" And rst!materialpc <> "" Then faltamaterialdelcomplexa = True
      If rst!materialpc2 <> "A" And rst!materialpc2 <> "R" And rst!materialpc2 <> "" Then faltamaterialdelcomplexa = True
  End If
fi:
  Set rst = Nothing
End Function

Function mirarpreucomanda(vnumc As Double) As Double
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select pvp from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then mirarpreucomanda = Redondejar(cadbl(rst!pvp), 4)
  Set rst = Nothing
End Function
Sub possar_comandes_parcials()
  Dim rst As Recordset
  Dim i As Long
  Dim vcolcomanda As Double
  Dim vcolexpedicio As Double
  
  vcolexpedicio = numcol("Data_Expedició")
  If vcolexpedicio = 0 Then vcolexpedicio = numcol("Data_Exp.")
  If vcolexpedicio = 0 Then GoTo fi
  Set rst = dbplanificacio.OpenRecordset("select * from planificaciototes where entregaparcial")
  If rst.EOF Then GoTo fi
  vcolcomanda = numcol("NºLot")
  While Not rst.EOF
    For i = 1 To reixa.Rows - 1
      If reixa.TextMatrix(i, vcolcomanda) = rst!comanda Then
            reixa.Row = i: reixa.col = vcolexpedicio
            If reixa.CellBackColor = QBColor(12) Then
                  reixa.CellBackColor = &H80C0FF
                    Else: reixa.CellBackColor = QBColor(14)
            End If
            Exit For
      End If
    Next i
    rst.MoveNext
  Wend
fi:
  Set rst = Nothing
End Sub
Function colorLiniaImpresio(rst As Recordset, rsttotes As Recordset) As Double
  Dim vcamp As Field
  Dim vcolor As Double
  Dim vestat As String
  rsttotes.FindFirst "comanda=" + atrim(rst!comanda)
  If rsttotes.EOF Then Exit Function
  If rsttotes!estat <> "E" And rsttotes!estat <> "I" Then Exit Function
  Set vcamp = rst.Fields("numeroliniaimpresio")
'  If Mid(atrim(vcamp.Value), 1, 3) = "016" Then Stop
  If Mid(atrim(vcamp.Value) + " ", 1, 1) = "-" Then
           vcolor = QBColor(12)
         Else
            'If rst!comanda = 214508 Then Stop
            If Not rstCdL.EOF And atrim(vcamp.Value) <> "" Then
                  rstCdL.FindFirst "numeroliniaimpresio like '" + Mid(atrim(vcamp.Value), 1, 3) + "*' and estat='I'"
                  If Not rstCdL.NoMatch Then
                     rstCdL.FindFirst "numeroliniaimpresio like '" + Mid(atrim(vcamp.Value), 1, 3) + "*' and estat='E'"
                     If Not rstCdL.NoMatch Then
                         If rst!codiclient <> 6841 Then
                            If rsttotes!estat <> "E" Then
                                 vcolor = QBColor(10)   'verd  '&H80C0FF 'taronja
                                   Else: vcolor = &H80C0FF 'taronja
                            End If
                             Else
                                If rsttotes!datacalloff Then
                                      vcolor = &H80C0FF  'taronja  QBColor(10)  'verd
                                    Else: vcolor = QBColor(15)  'blanc
                                End If
                         End If
                         If vcolor = &H80C0FF And rsttotes!estat = "I" Then vcolor = QBColor(10) 'verd
                     End If
                  End If
                  
            End If
  End If
  If vcolor = &H80C0FF Then
    vhihanCdLtaronja = True
    rstCdLestats.FindFirst "comanda=" + atrim(rst!comanda)
    If Not rstCdLestats.NoMatch Then
        vestat = atrim(rstCdLestats!estat)
        If vestat = "A" Then vcolor = QBColor(10)  'verd
        If vestat = "R" Then vcolor = QBColor(12)    'vermell
    End If
      'Else: vhihanCdLtaronja = False
  End If
  colorLiniaImpresio = vcolor
End Function
Function comandaacabada(numc As Double, seccio As String) As Boolean
'   MsgBox atrim(numc) + "  ---   " + seccio
  Dim rst As Recordset
  Dim taula As String
  comandaacabada = False
  If seccio = "I" Then taula = "impressorestot"
  If seccio = "L" Then taula = "laminadorestot"
  If seccio = "R" Then taula = "rebobinadorestot"
  If seccio = "S" Then taula = "soldadorestot"
  If taula = "" Then Exit Function
  Set rst = dbbaixes.OpenRecordset("select acavada from " + taula + " where comanda=" + atrim(cadbl(numc)), dbOpenSnapshot, dbReadOnly)
    If Not rst.EOF Then
      If cadbl(rst!acavada) <> 0 Then comandaacabada = True
    End If
End Function
Function baixasensefuncionament(numc As Double, seccio As String) As Boolean
'   MsgBox atrim(numc) + "  ---   " + seccio
  Dim rst As Recordset
  Dim taula As String
  baixasensefuncionament = False
  If seccio = "I" Then taula = "impressorestot"
  If seccio = "L" Then taula = "laminadorestot"
  If seccio = "R" Then taula = "rebobinadorestot"
  If seccio = "S" Then taula = "soldadorestot"
  If taula = "" Then Exit Function
  Set rst = dbbaixes.OpenRecordset("select * from " + taula + " where  comanda=" + atrim(cadbl(numc)), dbOpenSnapshot, dbReadOnly)
  If Not rst.EOF Then
     'rst.FindFirst "tipus='F'"
     'If rst.NoMatch Then
     baixasensefuncionament = True
     If comandaacabada(numc, seccio) Then baixasensefuncionament = False
  End If
End Function
Function larutahiha(producte As String, seccio As String) As Boolean
   Dim rstp As Recordset
   Set rstp = dbcomandes.OpenRecordset("select ruta from productes where codi='" + atrim(producte) + "'", dbOpenSnapshot, dbReadOnly)
   If rstp.EOF Then Exit Function
   If InStr(1, rstp!ruta, seccio) > 0 Then
       larutahiha = True
      Else: larutahiha = False
   End If
End Function
Function posicioenlaruta(numc As Double) As String
  Dim rstp As Recordset
  Dim rstpr As Recordset
  Dim laruta As String
    Set rstp = dbbaixes.OpenRecordset("SELECT proximaseccio from comandes where comanda=" + atrim(numc))
    If Not rstp.EOF Then posicioenlaruta = rstp!proximaseccio
    Set rstp = Nothing
    Exit Function
       'POSSO AIXÓ PERQUÈ L'ALICIA DIU QUE HA DE SER PROXIMASECCIO SENSE TENIR EN COMPTE SI S'HA COMENÇAT EN UNA ALTRA
       'SI ES VOL CANVIAR NOMÉS ES TREU les linies de exit function amunt I JA ESTÀ
  
  'If InStr(1, "VPT", seccioactual) = 0 Then Exit Function
  'Set rstp = dbbaixes.OpenRecordset("SELECT comandes.comanda,comandes.proximaseccio,comandes.producte, soldadorestot.acavada as acavadas,rebobinadorestot.acavada as acavadar, laminadorestot.acavada as acavadal, impressorestot.acavada as acavadai FROM ((comandes LEFT JOIN rebobinadorestot ON comandes.comanda = rebobinadorestot.comanda) LEFT JOIN laminadorestot ON comandes.comanda = laminadorestot.comanda) LEFT JOIN impressorestot ON comandes.comanda = impressorestot.comanda WHERE (((comandes.comanda)=" + atrim(numc) + "));")
  'Clipboard.Clear
  'Clipboard.SetText "SELECT comandes.comanda, comandes.proximaseccio, comandes.producte, rebobinadorestot.acavada AS acavadar, laminadorestot.acavada AS acavadal, impressorestot.acavada AS acavadai, soldadorestot.acavada AS acavadas FROM (((comandes LEFT JOIN rebobinadorestot ON comandes.comanda = rebobinadorestot.comanda) LEFT JOIN laminadorestot ON comandes.comanda = laminadorestot.comanda) LEFT JOIN impressorestot ON comandes.comanda = impressorestot.comanda) LEFT JOIN soldadorestot ON comandes.comanda = soldadorestot.comanda WHERE (((comandes.comanda)=" + atrim(numc) + "));"
  Set rstp = dbbaixes.OpenRecordset("SELECT comandes.comanda, comandes.proximaseccio, comandes.producte, rebobinadorestot.acavada AS acavadar, laminadorestot.acavada AS acavadal, impressorestot.acavada AS acavadai, soldadorestot.acavada AS acavadas FROM (((comandes LEFT JOIN rebobinadorestot ON comandes.comanda = rebobinadorestot.comanda) LEFT JOIN laminadorestot ON comandes.comanda = laminadorestot.comanda) LEFT JOIN impressorestot ON comandes.comanda = impressorestot.comanda) LEFT JOIN soldadorestot ON comandes.comanda = soldadorestot.comanda WHERE (((comandes.comanda)=" + atrim(numc) + "));", dbOpenSnapshot, dbReadOnly)

  If Not rstp.EOF Then
     Set rstpr = dbcomandes.OpenRecordset("select ruta from productes where codi='" + atrim(rstp!producte) + "'", dbOpenSnapshot, dbReadOnly)
     If rstpr.EOF Then Exit Function
     laruta = atrim(rstpr!ruta)
     If InStr(1, laruta, "S") > 0 And cadblnull_1(rstp!acavadas) = 0 Then posicioenlaruta = "S"
     If InStr(1, laruta, "R") > 0 And cadblnull_1(rstp!acavadar) = 0 Then posicioenlaruta = "R"
     If InStr(1, laruta, "L") > 0 And cadblnull_1(rstp!acavadal) = 0 Then posicioenlaruta = "L"
     If InStr(1, laruta, "I") > 0 And cadblnull_1(rstp!acavadai) = 0 Then posicioenlaruta = "I"
  End If
  If posicioenlaruta = "" Or atrim(rstp!proximaseccio) = "E" Then posicioenlaruta = rstp!proximaseccio
  'If posicioenlaruta = "" Then posicioenlaruta = rstp!proximaseccio

  Set rstp = Nothing
  Set rstpr = Nothing
End Function
Function cadblnull_1(acabada As Variant) As Double
   If IsNull(acabada) Then cadblnull_1 = 0: Exit Function
   cadblnull_1 = cadbl(acabada)
End Function
Sub copiaregistreatemporalentregues(Optional vdies As String)
  Dim rstt As Recordset
  Dim rsttemp As Recordset
  Dim amplemax As Double
  Dim taulaplanificacio As String
  Dim tintes As Byte
  Dim rstalb As Recordset
  Dim dbvendes As Database
  Dim vsql As String
  Dim rstobs As Recordset
  Dim vunitat As String

  Set dbvendes = OpenDatabase(rutadelfitxer(cami) + "vendes.mdb")
  taulaplanificacio = "planificacioent"
  If vdies <> "" Then dbconsulta.Execute "delete * from planificacioent"
  Set rsttemp = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + "")
  'Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio, dbOpenSnapshot, dbReadOnly)
  vsql = "SELECT capcaleraalbara.*, liniesalbara.*, clients.nom, comandes.datacomanda, comandes.impressio,comandes.pvp, comandes.tubbaseext, comandes.numtreball, pressupostos.preu FROM (((capcaleraalbara RIGHT JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara) LEFT JOIN (Clients_envios LEFT JOIN clients ON Clients_envios.codi = clients.codi) ON capcaleraalbara.id_direnvio = Clients_envios.id) LEFT JOIN comandes ON liniesalbara.lotinplacsa = comandes.comanda) LEFT JOIN pressupostos ON liniesalbara.lotinplacsa = pressupostos.lotambelqueshafacturat "
  vsql = vsql + IIf(vdies = "", " WHERE (((DateDiff('d',[dataalbara],Now()))<=7));", " where " + vdies)
  Set rstalb = dbvendes.OpenRecordset(vsql)
  Set rstobs = dbcomandes.OpenRecordset("select * from comandes_observacioPVP")
  With rsttemp
  While Not rstalb.EOF
    vunitat = ""
    rstobs.FindFirst "comanda=" + atrim(rstalb!lotinplacsa) + " or comandesafectades like '*" + atrim(rstalb!lotinplacsa) + "*'"
    .AddNew
    !datacomanda = rstalb!datacomanda
    !dataalbara = rstalb!dataalbara
    !albara = rstalb![numalbaraSAP]
    If rstalb!albaravalorat Then !albara = !albara * -1
    !comanda = cadbl(rstalb!lotinplacsa)
    !numtreball = cadbl(rstalb!numtreball)
    !nomclient = atrim(rstalb!nom)
    !impresio = atrim(rstalb!impressio)
    !entregaToP = atrim(rstalb!tipusdeentrega) + hihaalgunaparcial(rstalb!lotinplacsa)
    !quantitatTeorica = cadbl(rstalb!tubbaseext)
    !quantitatEntregada = cadbl(rstalb!quantitat)
    !Preu = Redondejar(cadbl(rstalb!preuvenda), 4)
    '!preu = cadbl(rstalb!pvp)
    !kgentregats = rstalb!kgtotalsbruts
    
    If !kgentregats > 0 Then !eurokg = Redondejar((rstalb!preuvenda * rstalb!quantitat) / !kgentregats, 2)
    If !quantitatTeorica > 0 Then !tanx100kgvs = Redondejar((100 - (!quantitatEntregada * 100 / !quantitatTeorica)) * -1, 1)
    !tipusunitat = mesurainterna(atrim(rstalb!unitatmesura))
    !kgimpost = cadbl(rstalb!kgimpostenvasos)
    If !kgimpost > 0 Then !tanx100impostvs = Redondejar((100 - (!kgentregats * 100 / !kgimpost)), 1)
    !pvprevisat = IIf(firmesPVPcomanda(rstalb!lotinplacsa) > 1, "S", "N")
    !preuclixes = cadbl(rstalb!Preu) 'podria ser que si es un parcial aquest preu surti a totes les entregues d'aquesta comanda
                                     'per evitar-ho es podria agafar també la data de facturació a pressupostos i comparar-la amb la de pujada al sap
    !Facturat = IIf(cadbl(rstalb!numfacturasap) > 0, "S", "N")
    If Not rstobs.NoMatch Then
        !extracost = cadbl(rstobs!extracost)
        If atrim(rstobs!observacio) <> "" Then
            If !extracost = 0 Then !extracost = 1
            !extracost = !extracost * -1
        End If
    End If
    .Update
    rstalb.MoveNext
  Wend
  End With
  actualitzar_valorsUSUARI
  Set rstt = Nothing
  Set dbvendes = Nothing
  Set rstalb = Nothing
  Set rstt = Nothing
  Set rsttemp = Nothing
  Set rstobs = Nothing
End Sub
Function mesurainterna(vmesura As String) As String
   Dim rst As Recordset
   Dim i As Byte
   vmesura = UCase(vmesura)
   Set rst = dbcomandes.OpenRecordset("select * from mesures")
   While Not rst.EOF
      For i = 0 To rst.Fields.Count - 1
         If UCase(rst.Fields(i)) = vmesura Then GoTo fi
      Next i
      rst.MoveNext
   Wend
fi:
   If Not rst.EOF Then mesurainterna = atrim(rst!unitatinterna)
   Set rst = Nothing
End Function
Function hihaalgunaparcial(vnumc As Double) As String
  Dim rst As Recordset
  Set rst = dbplanificacio.OpenRecordset("select * from liniesalbara where lotinplacsa=" + atrim(vnumc))
  If Not rst.EOF Then
     rst.MoveLast
     If rst.RecordCount > 1 Then hihaalgunaparcial = "*"
  End If
  Set rst = Nothing
End Function
Sub actualitzar_valorsUSUARI()
  Dim rst As Recordset
  Dim rstp As Recordset
  Set rst = dbconsulta.OpenRecordset("select * from planificacioent")
  Set rstp = dbplanificacio.OpenRecordset("select * from planificacioent")
  While Not rst.EOF
    rstp.FindFirst "comanda=" + atrim(rst!comanda) + " and dataalbara=#" + format(rst!dataalbara, "mm/dd/yy") + "#"
    If Not rstp.NoMatch Then
        rst.Edit
        rst!Revisat = rstp!Revisat
        rst!okclixes = rstp!okclixes
        rst!observacio = rstp!observacio
        rst.Update
    End If
    rst.MoveNext
  Wend
  Set rst = Nothing
  Set rstp = Nothing
End Sub
Function firmesPVPcomanda(vnumc As Double) As Double
  Dim rst As Recordset
  firmesPVPcomanda = 0
  Set rst = dbcomandes.OpenRecordset("select comanda from comandes_firmes where comanda=" + atrim(vnumc) + " and tipus='PVP'")
  If Not rst.EOF Then
     rst.MoveLast
     firmesPVPcomanda = rst.RecordCount
  End If
  Set rst = Nothing
End Function
Sub carregar_comandesatemporal()
  Dim rstc As Recordset
  Dim rstc2  As Recordset
  Dim cont As Double
  Dim Data1 As String
  Dim data2 As String
  Dim proximaseccio As String
  dbconsulta.Execute "delete * from planificaciototes"
  dbconsulta.Execute "delete * from planificacioimp"
  dbconsulta.Execute "delete * from planificacioreb"
  dbconsulta.Execute "delete * from planificaciolam"
  dbconsulta.Execute "delete * from planificaciosol"
  dbconsulta.Execute "delete * from planificacioent"
  dbconsulta.Execute "delete * from reclamacionscomandes"
  
    'empleno la taula d'entregues
  copiaregistreatemporalentregues

  copiaregistreatemporalgeneral , True
  copiaregistreatemporallaminadora , True
  copiaregistreatemporalrebobinadora , True
  copiaregistreatemporalimpresora , True
  copiaregistreatemporalsoldadora , True

  'dbplanificacio.Execute "delete * from " + taulaplanificacio + " where ordre=0 or maquina=0"
  Set rstc = dbconsulta.OpenRecordset("select * from comandesnoentregades", 4, dbReadOnly) 'where proximaseccio<>'T' and (producte<>'PC' and producte<>'PC2'
  If Not rstc.EOF Then rstc.MoveLast: rstc.MoveFirst
  'dataactivacio<>null
  DoEvents
  While Not rstc.EOF
    'If rstc!comanda = 205760 Then Stop
    ensenyarprogres rstc
    cont = cont + 1
    proximaseccio = posicioenlaruta(rstc!comanda)
    'faig els generals
    If (rstc!producte <> "PC" And rstc!producte <> "PC2" And rstc!producte <> "PCP") Then
        copiaregistreatemporalgeneral rstc

    End If
    'faig les impresores
      'aqui apunto la maquina que vull que sigui per defecte
    nummaquina = 7
    If (proximaseccio = "E" Or proximaseccio = "I") And larutahiha(rstc!producte, "I") Then
        copiaregistreatemporalimpresora rstc
    End If
    
    'faig les laminadores
    'aqui apunto la maquina que vull que sigui per defecte
    nummaquina = 1
    If (proximaseccio = "E" Or proximaseccio = "I" Or proximaseccio = "L") And larutahiha(rstc!producte, "L") And (rstc!producte <> "PC" And rstc!producte <> "PC2" And rstc!producte <> "PCP") Then
        copiaregistreatemporallaminadora rstc
    End If
    
    'faig les rebobinadores
    'aqui apunto la maquina que vull que sigui per defecte
    nummaquina = 3
    If (proximaseccio = "E" Or proximaseccio = "I" Or proximaseccio = "L" Or proximaseccio = "R") And larutahiha(rstc!producte, "R") And (rstc!producte <> "PC" And rstc!producte <> "PC2" And rstc!producte <> "PCP") Then
        If atrim(rstc!microperforat) <> "L" Then   ' si el microperforat es Laser que no es planifiqui es farà fora
            copiaregistreatemporalrebobinadora rstc
        End If
    End If
    
        'faig les soldadores
    'aqui apunto la maquina que vull que sigui per defecte
    nummaquina = 5
    If (proximaseccio = "E" Or proximaseccio = "I" Or proximaseccio = "L" Or proximaseccio = "R" Or proximaseccio = "S") And larutahiha(rstc!producte, "S") And (rstc!producte <> "PC" And rstc!producte <> "PC2" And rstc!producte <> "PCP") Then
        copiaregistreatemporalsoldadora rstc
    End If
    
     rstc.MoveNext
     If cont Mod 100 = 0 Then
       DoEvents
     End If
     
  Wend

  
  Set rstc = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda>9000 and comanda<10000", dbOpenSnapshot, dbReadOnly)
  While Not rstc.EOF
    Data1 = "null"
    data2 = "null"
    If IsDate(rstc!Data1) Then Data1 = "#" + format(rstc!Data1, "mm/dd/yy") + "#"
    If IsDate(rstc!data2) Then data2 = "#" + format(rstc!data2, "mm/dd/yy") + "#"
    'afegeixo impresores
    Set rstc2 = dbplanificacioalicia.OpenRecordset("select * from planificacioimp where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
    If Not rstc2.EOF Then dbconsulta.Execute "insert into planificacioimp (ordre,comanda,maquina,observacions,tempsimpresio,data1,data2,importancia) values (" + atrim(cadbl(rstc2!ordre)) + "," + atrim(cadbl(rstc!comanda)) + "," + atrim(rstc2!maquina) + ",'" + atrim(rstc!observacio) + "',180," + Data1 + "," + data2 + "," + atrim(cadbl(rstc!importancia)) + ")"
    'afegeixo laminadora
    Set rstc2 = dbplanificacioalicia.OpenRecordset("select * from planificaciolam where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
    If Not rstc2.EOF Then dbconsulta.Execute "insert into planificaciolam (ordre,comanda,maquina,observacions,tempsimpresio,data1,data2,importancia) values (" + atrim(cadbl(rstc2!ordre)) + "," + atrim(cadbl(rstc!comanda)) + "," + atrim(rstc2!maquina) + ",'" + atrim(rstc!observacio) + "',180," + Data1 + "," + data2 + "," + atrim(cadbl(rstc!importancia)) + ")"
    'afegeixo rebobinadora
    Set rstc2 = dbplanificacioalicia.OpenRecordset("select * from planificacioreb where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
    If Not rstc2.EOF Then dbconsulta.Execute "insert into planificacioreb (ordre,comanda,maquina,observacions,tempsimpresio,data1,data2,importancia) values (" + atrim(cadbl(rstc2!ordre)) + "," + atrim(cadbl(rstc!comanda)) + "," + atrim(rstc2!maquina) + ",'" + atrim(rstc!observacio) + "',180," + Data1 + "," + data2 + "," + atrim(cadbl(rstc!importancia)) + ")"
    'aki faltaria el mateix per planificaciosol
    Set rstc2 = dbplanificacioalicia.OpenRecordset("select * from planificaciosol where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
    If Not rstc2.EOF Then dbconsulta.Execute "insert into planificaciosol (ordre,comanda,maquina,observacions,tempsimpresio,data1,data2,importancia) values (" + atrim(cadbl(rstc2!ordre)) + "," + atrim(cadbl(rstc!comanda)) + "," + atrim(rstc2!maquina) + ",'" + atrim(rstc!observacio) + "',180," + Data1 + "," + data2 + "," + atrim(cadbl(rstc!importancia)) + ")"
    
    'aki la seccio de totes General
    dbconsulta.Execute "insert into planificaciototes (comanda,observacions,data1,data2,importancia,observacioexpedicio,dataexpedicio) values (" + atrim(cadbl(rstc!comanda)) + ",'" + atrim(rstc!observacio) + "'," + Data1 + "," + data2 + "," + atrim(cadbl(rstc!importancia)) + ",'" + atrim(rstc!observacioexpedicio) + "'," + atrim(rstc!dataexpedicio) + ")"
    rstc.MoveNext
  Wend
fi:
  Set rstc = Nothing
  Set rstc2 = Nothing
End Sub

Function nomdelacola(id As Long, Optional vnumcPC2 As Double) As String
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim vcolor1 As String
   Dim vcolor2 As String
   Dim vidcolaPC2 As Double
   
   
   If cadbl(vnumcPC2) > 0 Then
       Set rstc = dbcomandes.OpenRecordset("select tipusadhesiu from comandes where comanda=" + atrim(cadbl(vnumcPC2)))
       vidcolaPC2 = cadbl(rstc!tipusadhesiu)
       Set rst = dbcomandes.OpenRecordset("select resina,color,predeterminada from adhesius where codi=" + atrim(cadbl(vidcolaPC2)), dbOpenSnapshot, dbReadOnly)
       If rst.EOF Then Set rst = dbcomandes.OpenRecordset("select resina,color,predeterminada from adhesius where predeterminada<>''", dbOpenSnapshot, dbReadOnly)
       If Not rst.EOF Then vcolor2 = rst!color
   End If
   Set rst = dbcomandes.OpenRecordset("select resina,color,predeterminada from adhesius where codi=" + atrim(cadbl(id)), dbOpenSnapshot, dbReadOnly)
   If rst.EOF Then Set rst = dbcomandes.OpenRecordset("select resina,color,predeterminada from adhesius where predeterminada<>''", dbOpenSnapshot, dbReadOnly)
   If Not rst.EOF Then
      vcolor1 = rst!color
      nomdelacola = IIf(atrim(rst!predeterminada) <> "", rst!resina, "@" + possarcoloradhesiu(vcolor1) + IIf(vidcolaPC2 > 0, "/" + possarcoloradhesiu(atrim(vcolor2)), "") + " " + atrim(rst!resina))
      'nomdelacola = "@" + possarcoloradhesiu(rst!color) + " " + atrim(rst!resina)
   End If
   Set rst = Nothing
   Set rstc = Nothing
End Function
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
  End Select
  possarcoloradhesiu = codicolor
End Function
Sub copiaregistreatemporallaminadora(Optional rstc As Recordset, Optional inicialitzar As Boolean)
  Dim rstt As Recordset
  Dim rsttotes As Recordset
  Dim taulaplanificacio As String
  Static rsttemp As Recordset
  Static primer As Boolean
  Dim tintes As Byte
  If inicialitzar Then primer = False: Exit Sub
  If rstc.EOF Then Exit Sub
  taulaplanificacio = "planificaciolam"
  If Not primer Then Set rsttemp = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + ""): primer = True
  Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  If rstt.EOF Then
    dbplanificacio.Execute "insert into " + taulaplanificacio + " (comanda) values (" + atrim(rstc!comanda) + ")" 'rstt.AddNew: rstt!comanda = rstc!comanda: rstt.Update
    Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  End If
  
  
  Set rsttotes = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  If rsttotes.EOF Then
    dbplanificacioalicia.Execute "insert into planificaciototes (comanda) values (" + atrim(rstc!comanda) + ")" 'rstt.AddNew: rstt!comanda = rstc!comanda: rstt.Update
    Set rsttotes = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  End If
  
  
  With rsttemp
  .AddNew
  !ordre = 999
  !codiclient = rstc!client
  !estat = posicioenlaruta(rstc!comanda) + IIf(rstc!proximaseccio <> posicioenlaruta(rstc!comanda), Chr(255), "") ' atrim(rstc!proximaseccio)
  'If InStr(1, rsttemp!estat, Chr(255)) > 0 Then Stop
  '!maquina = cadbl(rstc!impressora)
  !maquina = nummaquina ': dbcomandes.Execute "update comandes set impressora=" + atrim(nummaquina) + " where comanda=" + atrim(rstc!comanda)
  !nomclient = nomclient(rstc!client)
  !comanda = rstc!comanda
  !camisa = rstc!camisa
  !capes = capesdelacomanda(cadbl(rstc!comanda), cadbl(rstc!linkcomanda1), cadbl(rstc!linkcomanda2))
  !materialpc = estatdelmaterial(rstc!linkcomanda1, rstc!proximaseccio) 'estatdelmaterial(IIf(rstc!linkcomanda1 < rstc!comanda, rstc!linkcomanda2, rstc!linkcomanda1))
  !materialpc2 = estatdelmaterial(rstc!linkcomanda2, rstc!proximaseccio) 'estatdelmaterial(IIf(rstc!linkcomanda1 < rstc!comanda, rstc!comanda, rstc!linkcomanda2))
  !tipuscola = nomdelacola(cadbl(rstc!tipusadhesiu), cadbl(rstc!linkcomanda2))
  !mts = rstc!cantitatex
  !ample = rstc!ampleesq
  !descmat = nomdelmaterial(rstc!materialex)
  !refclient = rstc!refclient
  !texteimpresio = rstc!marcailinia
  If !texteimpresio = "" Or InStr(1, !texteimpresio, "NO HI HA LINIA") > 0 Then !texteimpresio = atrim(rstc!texteimpressio)
  !Data1 = rstc!dataentrega
  !data2 = IIf(IsDate(rsttotes!data2), rsttotes!data2, Null)
  !importancia = IIf(Not IsNull(rsttotes!importancia), rsttotes!importancia, "")
  !observacions = atrim(rsttotes!observacio)
  !tempsimpresio = quanestardara(1, cadbl(rstc!cantitatex), !maquina, "L") * (IIf(!capes = 3, 2, 1))
  .Update
  End With
  Set rstt = Nothing
  
End Sub
Sub copiaregistreatemporalrebobinadora(Optional rstc As Recordset, Optional inicialitzar As Boolean)
  Dim rstt As Recordset
  Dim rsttotes As Recordset
  Dim rstexp As Recordset
  Dim amplemax As Double
  Dim taulaplanificacio As String
  Static rsttemp As Recordset
  Static primer As Boolean
  Dim tintes As Byte
  If inicialitzar Then primer = False: Exit Sub
  If rstc.EOF Then Exit Sub
  
  taulaplanificacio = "planificacioreb"
  If Not primer Then Set rsttemp = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + ""): primer = True
  Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  If rstt.EOF Then
    dbplanificacio.Execute "insert into " + taulaplanificacio + " (comanda) values (" + atrim(rstc!comanda) + ")" 'rstt.AddNew: rstt!comanda = rstc!comanda: rstt.Update
    Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  End If
  Set rstexp = dbplanificacio.OpenRecordset("select data from linies_expedicions where comanda=" + atrim(rstc!comanda) + " order by id desc")
  
  Set rsttotes = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  If rsttotes.EOF Then
    dbplanificacioalicia.Execute "insert * into planificaciototes (comanda) values (" + atrim(rstc!comanda) + ")" 'rstt.AddNew: rstt!comanda = rstc!comanda: rstt.Update
    Set rsttotes = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  End If
  
  amplemax = amplemaxcomanda(rstc!comanda, rstc!linkcomanda1, rstc!linkcomanda2)
  With rsttemp
  .AddNew
  !ordre = 999
  !codiclient = rstc!client
  !estat = posicioenlaruta(rstc!comanda) + IIf(rstc!proximaseccio <> posicioenlaruta(rstc!comanda), Chr(255), "") 'atrim(rstc!proximaseccio)
  '!maquina = cadbl(rstc!impressora)
  !maquina = nummaquina ': dbcomandes.Execute "update comandes set impressora=" + atrim(nummaquina) + " where comanda=" + atrim(rstc!comanda)
  !nomclient = nomclient(rstc!client)
  !producte = rstc!producte
  !comanda = rstc!comanda
  !mandril = rstc!tubbase
  !mts = rstc!cantitatex
  !ample = rstc!amplereb
  !micromacro = IIf(atrim(rstc!microperforat) <> "N" And atrim(rstc!microperforat) <> "" And Not IsNull(atrim(rstc!microperforat)), "Micro", IIf(atrim(rstc!rebmacroperforat) = "S", "Macro", ""))
  !refclient = rstc!refclient
  !texteimpresio = rstc!marcailinia
  If !texteimpresio = "" Or InStr(1, !texteimpresio, "NO HI HA LINIA") > 0 Then !texteimpresio = atrim(rstc!texteimpressio)
  !bandes = rstc!simulteneitatreb
  If amplemax = 0 Then
      !merma = 0
    Else: !merma = amplemax - (cadbl(rstc!amplereb) * cadbl(rstc!simulteneitatreb))
  End If
  !Data1 = rstc!dataentrega
  !data2 = IIf(IsDate(rsttotes!data2), rsttotes!data2, Null)
  !importancia = IIf(Not IsNull(rsttotes!importancia), rsttotes!importancia, "")
  If Not rstexp.EOF Then !dataexpedicio = format(rstexp!data, "dd/mm/yy")
  !observacions = atrim(rsttotes!observacio)
  !tempsimpresio = quanestardarareb(cadbl(rstc!simulteneitatreb), cadbl(rstc!mtrslinbob), cadbl(rstc!cantitatex), cadbl(!maquina), IIf(CDbl(!merma) > 9, "merma", ""))
  .Update
  End With
  Set rstt = Nothing
  Set rsttotes = Nothing
  Set rstpexp = Nothing
End Sub
Sub copiaregistreatemporalsoldadora(Optional rstc As Recordset, Optional inicialitzar As Boolean)
  Dim rstt As Recordset
  Dim rsttotes As Recordset
  Dim amplemax As Double
  Dim taulaplanificacio As String
  Static rsttemp As Recordset
  Static primer As Boolean
  Dim tintes As Byte
  If inicialitzar Then primer = False: Exit Sub
  If rstc.EOF Then Exit Sub
  taulaplanificacio = "planificaciosol"
  If Not primer Then Set rsttemp = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + ""): primer = True
  Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  If rstt.EOF Then
    dbplanificacio.Execute "insert into " + taulaplanificacio + " (comanda) values (" + atrim(rstc!comanda) + ")" 'rstt.AddNew: rstt!comanda = rstc!comanda: rstt.Update
    Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  End If
  
  
  Set rsttotes = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  If rsttotes.EOF Then
    dbplanificacioalicia.Execute "insert into planificaciototes (comanda) values (" + atrim(rstc!comanda) + ")" 'rstt.AddNew: rstt!comanda = rstc!comanda: rstt.Update
    Set rsttotes = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  End If
  
  With rsttemp
  .AddNew
  !ordre = 999
  !codiclient = rstc!client
  !estat = posicioenlaruta(rstc!comanda) + IIf(rstc!proximaseccio <> posicioenlaruta(rstc!comanda), Chr(255), "") 'atrim(rstc!proximaseccio)
  '!maquina = cadbl(rstc!impressora)
  !maquina = nummaquina ': dbcomandes.Execute "update comandes set impressora=" + atrim(nummaquina) + " where comanda=" + atrim(rstc!comanda)
  !nomclient = nomclient(rstc!client)
  !comanda = rstc!comanda
  !quantitatsol = rstc!cantitatsol
  !amplesol = rstc!amplesol
  !longitud = rstc!longitudsol
  !refclient = rstc!refclient
  !texteimpresio = rstc!marcailinia
  If !texteimpresio = "" Or InStr(1, !texteimpresio, "NO HI HA LINIA") > 0 Then !texteimpresio = atrim(rstc!texteimpressio)
  !Data1 = rstc!dataentrega
  !data2 = IIf(IsDate(rsttotes!data2), rsttotes!data2, Null)
  !importancia = IIf(Not IsNull(rsttotes!importancia), rsttotes!importancia, "")
  !observacions = atrim(rsttotes!observacio)
  !tempsimpresio = quanestardara(1, cadbl(rstc!cantitatex), !maquina, "S")
  .Update
  End With
  Set rstt = Nothing
  
End Sub
Function quanestardarareb(bandes As Byte, metresbob As Double, metres As Double, maquina As Byte, merma As String)
     Dim rst As Recordset
     Dim rste As Recordset
     If metres = 0 Or metresbob = 0 Then quanestardarareb = 0: Exit Function
     quanestardarareb = 0
     Set rst = dbplanificacioalicia.OpenRecordset("select * from canvimaquinesreb where nummaquina=" + atrim(maquina) + " and bandes=" + atrim(bandes), dbOpenSnapshot, dbReadOnly)
     If rst.EOF Then Exit Function
     Set rste = dbplanificacioalicia.OpenRecordset("select * from escalatvelocitatsreb where idcanvireb=" + atrim(rst!id), dbOpenSnapshot, dbReadOnly)
     If Not rste.EOF Then
        rste.FindFirst "metres>" + atrim(metresbob)
        If Not rste.NoMatch Then
            If Not rste.BOF Then rste.MovePrevious
        End If
        If cadbl(rste.Fields("mtrsmin" + merma)) > 0 Then
          quanestardarareb = Redondejar(cadbl(rst.Fields("tempscanvi" + merma)) + (metres / cadbl(rste.Fields("mtrsmin" + merma))), 0)
        End If
     End If
End Function
Function amplemaxcomanda(numc, numc2, numc3) As Double
  Dim rstc As Recordset
  Dim vsql As String
  If numc = 0 Then numc = -1
  If numc2 = 0 Then numc2 = -1
  If numc3 = 0 Then numc3 = -1
  vsql = "SELECT Palets.Ample AS amplemax FROM Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet "
  vsql = vsql + " WHERE Parcials.comanda In ('" + atrim(numc) + "','" + atrim(numc2) + "','" + atrim(numc3) + "') ORDER BY Palets.Ample DESC;"
'  Set rstc = dbstocks.OpenRecordset("SELECT Max(Palets.Ample) AS amplemax FROM Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet GROUP BY CDbl([comanda]) HAVING (CDbl([comanda])=" + atrim(numc) + " or cdbl([comanda])=" + atrim(cadbl(numc3)) + " or cdbl([comanda])=" + atrim(cadbl(numc2)) + ");", dbOpenSnapshot, dbReadOnly)
  Set rstc = dbstocks.OpenRecordset(vsql, dbOpenSnapshot, dbReadOnly)
  If Not rstc.EOF Then amplemaxcomanda = Redondejar(rstc!amplemax, 1)
  
  Set rstc = Nothing
End Function

Sub copiaregistreatemporalimpresora(Optional rstc As Recordset, Optional inicialitzar As Boolean)
  Dim rstt As Recordset
  Dim rstl As Recordset
  Dim taulaplanificacio As String
  Dim rsttotes As Recordset
  Dim vesreprint As Boolean
  Dim vliniaimpresio As Double
  Static rsttemp As Recordset
  Static primer As Boolean
  Dim tintes As Byte
  
  If inicialitzar Then primer = False: Exit Sub
  If rstc.EOF Then Exit Sub
  taulaplanificacio = "planificacioimp"
  If Not primer Then Set rsttemp = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + ""): primer = True
  Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  Set rstl = dbclixes.OpenRecordset("select numerodelinia,codidelinia,codideliniav from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(cadbl(rstc!numordremodificacio)), dbOpenSnapshot, dbReadOnly)
  If rstt.EOF Then
    dbplanificacio.Execute "insert into " + taulaplanificacio + " (comanda) values (" + atrim(rstc!comanda) + ")" 'rstt.AddNew: rstt!comanda = rstc!comanda: rstt.Update
    Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  End If
  
  Set rsttotes = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  If rsttotes.EOF Then
    dbplanificacioalicia.Execute "insert into planificaciototes (comanda) values (" + atrim(rstc!comanda) + ")"
    Set rsttotes = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  End If
  
  With rsttemp
  .AddNew
  !ordre = 999
  '!maquina = cadbl(rstc!impressora)
  !maquina = nummaquina ': dbcomandes.Execute "update comandes set impressora=" + atrim(nummaquina) + " where comanda=" + atrim(rstc!comanda)
 ' If rstc!formaimp = "T" Then !maquina = 9 ' si es transparencia fer amb F2
 ' If rstc!formaimp = "N" Then !maquina = 7 ' si es normal fer am FW
  !dataactcomanda = rstc!dataactivacio
  !codiclient = rstc!client
  !nomclient = nomclient(rstc!client)
  !comanda = rstc!comanda
  !capes = capesdelacomanda(cadbl(rstc!comanda), cadbl(rstc!linkcomanda1), cadbl(rstc!linkcomanda2))
  !capes = !capes + IIf(rstc!microperforat <> "" And rstc!microperforat <> "N", 10, 0)
  !tipusimpresio = rstc!formaimp
  !material = estatdelmaterial(rstc!comanda, rstc!proximaseccio)
  !mts = rstc!cantitatex
  !ample = rstc!ampleesq
  !espesor = micresmaterial(cadbl(rstc!mesuraesp), cadbl(rstc!espessor), atrim(rstc!tubolam)) + " " + r
  !descmat = nomdelmaterial(rstc!materialex)
  !refclient = rstc!refclient
  !numtreball = atrim(cadbl(rstc!numtreball)) + "/" + atrim(cadbl(rstc!numordremodificacio))
  !texteimpresio = atrim(rstc!marcailinia)
  If !texteimpresio = "" Or InStr(1, !texteimpresio, "NO HI HA LINIA") > 0 Then !texteimpresio = atrim(rstc!texteimpressio)
  If cadbl(rstl!codidelinia) > 0 Then !numeroliniaimpresio = format(rstl!codidelinia, "000") + "#" + atrim(rstl!codideliniav)  'cadbl(rstl!numerodelinia)
  vliniaimpresio = atrim(cadbl(rstl!numerodelinia))
  !impresio = rstc!impressio
  !clixes = "     "
  !clixes = buscadatadelclixenous(cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), cadbl(rstc!direnvio), IIf(rstc!impressio = "R", atrim(rstc!marques), "No"), vesreprint) '!clixes = buscadatadelclixe(cadbl(rstc!numtreball)) '
  If vesreprint Then !texteimpresio = "*" + !texteimpresio
  'If rstc!impressio = "R" And Mid(!clixes, 1, 1) = "*" Then !clixes = ""
  !Data1 = rstc!dataentrega
  !data2 = IIf(IsDate(rsttotes!data2), rsttotes!data2, Null)
  !tintesrevisades = mirarsitintesrevisades(rstc!comanda)
  !clientvindraarevisarimpresio = IIf(rstc!clientvindraarevisarimpresio, "S", "")
  !gruixclixes = cadbl(rstc!gruixpol)
  !cilindre = cadbl(rstc!cilindres)
  !muntat = IIf(InStr(1, "EI", rstc!proximaseccio) <> 0, estanmuntatselsclixes(rstc!comanda), "")
  !importancia = IIf(Not IsNull(rsttotes!importancia), rsttotes!importancia, "")
  !standbyimpresio = IIf(cadbl(rstc!passaraimpresores) = 0, "S", "")
  !observacions = atrim(rsttotes!observacio)
  If cadbl(rstc!numerotintes) = 0 Then
     tintes = 6
    Else: tintes = cadbl(rstc!numerotintes)
  End If
  !tempsimpresio = quanestardara(tintes, cadbl(rstc!cantitatex), !maquina, "I")
  If cadbl(rstc!numerotintes) = 0 Then !tempsimpresio = !tempsimpresio * -1
  .Update
  dbconsulta.Execute "update reclamacionscomandes set liniaimpresio=" + atrim(vliniaimpresio) + " where comanda=" + atrim(rstc!comanda)
  End With
  Set rstt = Nothing
  Set rstl = Nothing
End Sub
Function MirarPackComandes(vnumc As Double, vnumpressupost As String, vnumpack As String) As String
  Dim rst As Recordset
  'If vnumc = 226412 Then Stop
  If vnumpressupost = "" Or vnumpack = "" Then Exit Function
  Set rst = dbcomandes.OpenRecordset("select * from comandesmesextres where proximaseccio<>'T' and numpack='" + atrim(vnumpack) + "'")
  If Not rst.EOF Then
       rst.MoveLast
       rst.MoveFirst
       MirarPackComandes = "P:" + atrim(rst.RecordCount) + "c(" + atrim(vnumpack) + ") "
  End If
  Set rst = Nothing
End Function
Function estanmuntatselsclixes(comanda As Double) As String
  Dim rstc As Recordset
  estanmuntatselsclixes = " "
  Set rstc = dbbaixes.OpenRecordset("select comanda from muntadora_ordremuntatge where comanda=" + atrim(comanda), , ReadOnly)
  If Not rstc.EOF Then
    estanmuntatselsclixes = "M"
     Else: GoTo fi
  End If
  
  Set rstc = dbbaixes.OpenRecordset("select * from muntadoratot where comanda=" + atrim(comanda), dbOpenSnapshot, dbReadOnly)
  If Not rstc.EOF Then
     estanmuntatselsclixes = IIf(rstc!acabada, "S", "M")
  End If
fi:
  Set rstc = Nothing
End Function
Sub copiaregistreatemporalgeneral(Optional rstc As Recordset, Optional inicialitzar As Boolean)
  Dim rstt As Recordset
  Dim rstextres As Recordset
  Dim taulaplanificacio As String
  Dim rsttotes As Recordset
  Dim Data1 As String
  Dim vesreprint As Boolean
  Static rsttemp As Recordset
  Static primer As Boolean
  Dim tintes As Byte
  If inicialitzar Then primer = False: Exit Sub
  If rstc.EOF Then Exit Sub
  Set rstextres = dbplanificacio.OpenRecordset("select * from comandes_extres")
  Data1 = "null"
  If IsDate(rstc!dataentrega) Then Data1 = "#" + format(rstc!dataentrega, "mm/dd/yy") + "#"
  taulaplanificacio = "planificaciototes"
  If Not primer Then Set rsttemp = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + ""): primer = True
  Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  If rstt.EOF Then
    dbplanificacio.Execute "insert into " + taulaplanificacio + " (comanda,data1) values (" + atrim(rstc!comanda) + "," + Data1 + ")"
    Set rstt = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
      Else
       dbplanificacio.Execute "update " + taulaplanificacio + " set data1=" + Data1 + " where comanda=" + atrim(rstc!comanda)
  End If
  
  
  Set rsttotes = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  If rsttotes.EOF Then
    dbplanificacioalicia.Execute "insert into planificaciototes (comanda,data1) values (" + atrim(rstc!comanda) + "," + Data1 + ")"
    Set rsttotes = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(rstc!comanda), dbOpenSnapshot, dbReadOnly)
  End If
  
  With rsttemp
   rstextres.FindFirst "comanda=" + atrim(rstc!comanda)
  .AddNew
  !Data1 = rstc!dataentrega
  !data2 = IIf(IsDate(rsttotes!data2), rsttotes!data2, Null)
  !dataexpedicio = IIf(IsDate(rsttotes!dataexpedicio), rsttotes!dataexpedicio, Null)
  !observacioexpedicio = atrim(rsttotes!observacioexpedicio)
  !datacalloff = buscarsitecalloff(rstc!comanda, atrim(rstc!refclient))
  !importancia = IIf(Not IsNull(rsttotes!importancia), rsttotes!importancia, "")
  !observacions = atrim(rsttotes!observacio)
  !comandaclient = atrim(rstc!comandaclient)
  !standbyimpresio = IIf(cadbl(rstc!passaraimpresores) = 0, "S", "")
  !material = estatdelmaterial(rstc!comanda, rstc!proximaseccio)
  !materialpc = estatdelmaterial(cadbl(rstc!linkcomanda1), rstc!proximaseccio)
  !materialpc2 = estatdelmaterial(cadbl(rstc!linkcomanda2), rstc!proximaseccio)
  !micromacro = IIf(atrim(rstc!microperforat) <> "N" And atrim(rstc!microperforat) <> "" And Not IsNull(atrim(rstc!microperforat)), "Micro", IIf(atrim(rstc!rebmacroperforat) = "S", "Macro", ""))
  !mts = rstc!cantitatex
  !muntat = IIf(InStr(1, "EI", rstc!proximaseccio) <> 0, estanmuntatselsclixes(rstc!comanda), "")
  !refclient = rstc!refclient
  !refinplacsa = rstextres!refinplacsa
  !codiclient = rstc!client
  If larutahiha(rstc!producte, "I") Then
    !texteimpresio = rstc!marcailinia
    If !texteimpresio = "" Or InStr(1, !texteimpresio, "NO HI HA LINIA") > 0 Then !texteimpresio = atrim(rstc!texteimpressio)
    Set rstl = dbclixes.OpenRecordset("select codidelinia,codideliniav from modificacions where id_treball=" + atrim(rstc!numtreball) + " and ordre=" + atrim(rstc!numordremodificacio))
    If Not rstl.EOF Then If cadbl(rstl!codidelinia) > 0 Then !numeroliniaimpresio = format(rstl!codidelinia, "000") + "#" + atrim(rstl!codideliniav) 'cadbl(rstl!numerodelinia)
    !estatclixes = possarestatclixes(atrim(rstc!impressio))
    'If rstc!impressio <> "R" Then
    !clixes = buscadatadelclixenous(cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), cadbl(rstc!direnvio), IIf(rstc!impressio = "R", atrim(rstc!marques), "No"), vesreprint)  '!clixes = buscadatadelclixe(cadbl(rstc!numtreball))
    If vesreprint Then !texteimpresio = "*" + !texteimpresio
    If rstc!impressio = "R" And Mid(!clixes, 1, 1) = "*" Then !clixes = ""
  End If
  If larutahiha(rstc!producte, "L") Then !tipuscola = nomdelacola(cadbl(rstc!tipusadhesiu), cadbl(rstc!linkcomanda2))
  !producte = rstc!producte
  !tintesrevisades = mirarsitintesrevisades(rstc!comanda)
  !clientvindraarevisarimpresio = MirarPackComandes(rstc!comanda, atrim(rstc!numpressupost), atrim(rstc!numpack))
  !clientvindraarevisarimpresio = IIf(rstc!clientvindraarevisarimpresio, !clientvindraarevisarimpresio + "S", !clientvindraarevisarimpresio)
  !estat = posicioenlaruta(rstc!comanda) + IIf(rstc!proximaseccio <> posicioenlaruta(rstc!comanda), Chr(255), "") 'rstc!proximaseccio
  !dataactcomanda = rstc!dataactivacio
  !direnvioclient = direnvioclient(cadbl(rstc!direnvio))
  !nomclient = nomclient(rstc!client)
  !comanda = rstc!comanda
  .Update
  End With
  Set rstt = Nothing
  rsttemp.MoveLast
  copiarregistresalataula_reclamacionscomandes rsttemp
  
End Sub
Sub copiarregistresalataula_reclamacionscomandes(rstc As Recordset)
   Dim rstreclam As Recordset
   Dim rst As Recordset
   Dim rstp As Recordset
   Dim i As Byte
   'Set rst = dbconsulta.OpenRecordset("select * from planificaciototes")
   Set rstreclam = dbconsulta.OpenRecordset("select * from reclamacionscomandes")
   Set rstp = dbplanificacio.OpenRecordset("select * from reclamacionscomandes")
'   On Error Resume Next
   
   If Not rstc.EOF Then
     
     rstreclam.AddNew
     On Error Resume Next
     For i = 0 To rstc.Fields.Count - 1
        rstreclam.Fields(rstc.Fields(i).Name) = rstc.Fields(i)
     Next i
     On Error GoTo 0
     rstp.FindFirst "comanda=" + atrim(rstc!comanda)
     If Not rstp.NoMatch Then
            For i = 3 To rstp.Fields.Count - 1
                 rstreclam.Fields(rstp.Fields(i).Name) = rstp.Fields(i)
            Next i
     End If
     rstreclam.Update
   End If
   On Error GoTo 0
   Set rst = Nothing
   Set rstreclam = Nothing
   Set rstp = Nothing
End Sub
Function mirarsitintesrevisades(numc As Double) As String
  Dim rsttintes As Recordset
   Set rsttintes = dbtintes.OpenRecordset("select * from comandesrevisadesatintes where comanda=" + atrim(numc), dbOpenSnapshot, dbReadOnly)
  If Not rsttintes.EOF Then mirarsitintesrevisades = atrim(rsttintes!estatgestio)
  Set rsttintes = Nothing
  If mirarsitintesrevisades = "N" Then mirarsitintesrevisades = ""
  
End Function
Function buscarsitecalloff(vnumc As Double, vitem As String) As Date
   Dim rstent As Recordset
   Dim rst As Recordset
   
   Set rstent = dbbaixes.OpenRecordset("select distinct numcalloff,entregat from bobinesent where (entregat<>'S' or entregat=null or entregat='') and (numcalloff<>'' and numcalloff<>null) and comanda=" + atrim(vnumc) + " order by entregat", dbOpenSnapshot, dbReadOnly)
   If rstent.EOF Then
      Set rst = dbcomandes.OpenRecordset("select * from calloffs_detall where comanda=" + atrim(vnumc), dbOpenSnapshot, dbReadOnly)
      If Not rst.EOF Then
         'Set rst = dbcomandes.OpenRecordset("select * from calloffs where numcalloff='" + atrim(rst!numcalloff) + "' and item='" + atrim(vitem) + "'")
         buscarsitecalloff = buscardatadelcalloff(atrim(rst!numcalloff), vitem)
      End If
       Else
          'si ja està entregat ha de tornar sense data
          If Not rstent.EOF Then
             buscarsitecalloff = buscardatadelcalloff(atrim(rstent!numcalloff), vitem)
          End If
   End If
End Function

Function buscardatadelcalloff(vnumcalloff As String, vitem As String) As Date
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select * from calloffs where numcalloff='" + atrim(vnumcalloff) + "'" + IIf(vitem <> "", " and item='" + vitem + "'", ""), dbOpenSnapshot, dbReadOnly)
   If Not rst.EOF Then buscardatadelcalloff = CVDate(rst!data)
End Function
Function possarestatclixes(tipusclixes As String) As String
possarestatclixes = ""
If tipusclixes = "R" Then possarestatclixes = "Repetida"
If tipusclixes = "N" Then possarestatclixes = "Nova"
If tipusclixes = "M" Then possarestatclixes = "Modificada"
If tipusclixes = "F" Then possarestatclixes = "Falta Aut."
End Function
Function quanestardara(tintes As Byte, metres As Double, maquina As Byte, seccio As String)
     Dim rst As Recordset
     If metres = 0 Then quanestardara = 0: Exit Function
     Set rst = dbplanificacioalicia.OpenRecordset("select * from canvismaquines where seccio='" + seccio + "' and nummaquina=" + atrim(maquina) + " and tintes=" + atrim(tintes), dbOpenSnapshot, dbReadOnly)
     quanestardara = 0
     If Not rst.EOF Then
        If cadbl(rst!mtrsmin) > 0 Then
          quanestardara = Redondejar(cadbl(rst!tempscanvi) + (metres / cadbl(rst!mtrsmin)), 0)
        End If
     End If
     Set rst = Nothing
End Function
Function estatdelmaterial(numc As Long, vproximaseccio As String) As String
  Dim rstt As Recordset
  If numc = 0 Then Exit Function
  Set rstt = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, capcalera.dataentrega as dataent, liniescompra.totentregat as entregat, comandesxlinia.numcomanda FROM (capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra) RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra WHERE (((comandesxlinia.numcomanda)=" + atrim(numc) + "));", dbOpenSnapshot, dbReadOnly)
  If Not rstt.EOF Then
     estatdelmaterial = format(rstt!dataent, "dd/mm/yy")
     If cabool(rstt!entregat) Then estatdelmaterial = "E" + format(rstt!dataent, "dd/mm/yy")
  End If
  Set rstt = dbstocks.OpenRecordset("select * from percomandaoclient where  numcomanda=" + atrim(numc))
  If Not rstt.EOF Then estatdelmaterial = "R"
  Set rstt = dbstocks.OpenRecordset("select * from parcials where  comanda='" + atrim(numc) + "'")
  If Not rstt.EOF Then estatdelmaterial = "A"
  Set rstt = dbcomandes.OpenRecordset("select assignarstock,materialexacte from comandes_extres where comanda=" + atrim(numc))
 ' Set rstt = dbcomandes.OpenRecordset("SELECT comandes_extres.assignarstock, comandes_extres.materialexacte, comandes.proximaseccio FROM comandes_extres INNER JOIN comandes ON comandes_extres.comanda = comandes.comanda where comandes_extres.comanda=" + atrim(numc), dbOpenSnapshot, dbReadOnly)
 ' Set rstt = dbcomandes.OpenRecordset("SELECT comandes_extres.assignarstock, comandes_extres.materialexacte, comandes.proximaseccio FROM comandes_extres INNER JOIN comandes ON comandes_extres.comanda = comandes.comanda where comandes_extres.comanda=" + atrim(numc), dbOpenSnapshot, dbReadOnly)
  If Not rstt.EOF Then
     If rstt!assignarstock Then estatdelmaterial = "A"
     If cabool(rstt!materialexacte) And atrim(vproximaseccio) = "E" And estatdelmaterial = "A" Then estatdelmaterial = "ESP"
  End If
  Set rstt = Nothing
End Function
Function buscar_Data_Recepcio_Fotogravador(vtreball As Long, vordre As Integer) As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from comandesfotogravador where id_Treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vordre))
   If Not rst.EOF Then
         buscar_Data_Recepcio_Fotogravador = atrim(rst!datarecepcio)
   End If
   Set rst = Nothing
End Function
Function buscadatadelclixenous(treball As Long, ordremodificacio As Integer, direnvio As Double, canvidesti As String, vesreprint As Boolean) As String
  Dim rst As Recordset
  Dim rstv As Recordset
  Dim rstm As Recordset
  If ordremodificacio = 0 Then ordremodificacio = 1
  buscadatadelclixenous = "     "
  Set rst = dbclixes.OpenRecordset("SELECT clixes_modifi.id_estatclixe,Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, CLIXES_MODIFI.data_prevista,Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE Clixes_modifi.id_treball=" + atrim(treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + " AND clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + ");", dbOpenSnapshot, dbReadOnly)
  Set rstv = dbclixes.OpenRecordset("select arxiuimp from clientsvinculats where id_treball=" + atrim(treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " and direnvio=" + atrim(direnvio), dbOpenSnapshot, dbReadOnly)
  Set rstm = dbclixes.OpenRecordset("select reimpres from modificacions where id_treball=" + atrim(treball) + " and ordre=" + atrim(ordremodificacio), dbOpenSnapshot, dbReadOnly)
  ' CONSULTA ULTIMA MODIFICACIO AMB DATA FI  VINCULANT "SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE (((Clixes_modifi.id_treball)=" + atrim(id_treball) + ") AND ((Clixes_modifi.ordremodificacio)=" + atrim(ordremodificacio) + ") AND ((Clixes_estats.vinculant)=True and isdate(clixes_modifi.data_fi)) and clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + "));"
  ' CONSULTA ULTIMA MODIFICACIO AMB DATA FI SENSE VINCULANT "SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE (((Clixes_modifi.id_treball)=" + atrim(id_treball) + ") AND ((Clixes_modifi.ordremodificacio)=" + atrim(ordremodificacio) + ") AND (isdate(clixes_modifi.data_fi)) and clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + "));"
  If Not rst.EOF Then
      '8 clixes entrats
     If rst!id_estatclixe = 8 Then
          vdatarebutsfotogravador = buscar_Data_Recepcio_Fotogravador(treball, ordremodificacio)
          If vdatarebutsfotogravador = "" Then
               buscadatadelclixenous = "*" + format(rst!data_fi, "dd/mm/yy")
               Else: buscadatadelclixenous = "*" + format(vdatarebutsfotogravador, "dd/mm/yy")
          End If
     End If
      '17 ESPERA REBRE COMANDA
     If rst!id_estatclixe = 17 Then buscadatadelclixenous = "#NOCOMANDA"
      '15 POLIMERS O CLIXES    22 REPOSICIÓ DEL CLIXE
     If rst!id_estatclixe = 15 Or rst!id_estatclixe = 22 Then buscadatadelclixenous = format(rst!data_prevista, "dd/mm/yy")
      '19 RETORNEM CLIXES
     If rst!id_estatclixe = 19 Then buscadatadelclixenous = "!TORNEM"
      '20 CLIXES REBUTS
     If rst!id_estatclixe = 20 Then buscadatadelclixenous = "REBUTS"
     If rstv.EOF Then
          If canvidesti = "Si" Then
           buscadatadelclixenous = "!NO_IMPS"
          End If
          Else:
             If Not rstv.arxiuimp And canvidesti = "Si" Then
                buscadatadelclixenous = "!NO_IMPS"
             End If
     End If
     vesreprint = rstm!reimpres
    ' If vesreprint Then Stop
     
  End If
  Set rstm = Nothing
  Set rst = Nothing
  Set rstv = Nothing
End Function
Function buscadatadelclixe(treball As Long) As String
  Dim rst As Recordset
  buscadatadelclixe = "     "
  Set rst = dbclixes.OpenRecordset("select dataprevclixes,dataentrega,id_estatclixe from clixes where id_treball=" + atrim(treball), dbOpenSnapshot, dbReadOnly)
  If Not rst.EOF Then
     If IsDate(rst!dataprevclixes) Then buscadatadelclixe = format(rst!dataprevclixes, "dd/mm/yy")
     If IsDate(rst!dataentrega) Then buscadatadelclixe = "*" + format(rst!dataentrega, "dd/mm/yy")
     If cadbl(rst!id_estatclixe) = 17 Then
        buscadatadelclixe = "#"
     End If
  End If
  Set rst = Nothing
End Function
Function nomdelmaterial(codimat As Long) As String
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("SELECT materials.codi, familiesmaterials.descripcio as descfammat, familiescolorants.descripcio as descfamcol FROM familiescolorants INNER JOIN (familiesmaterials INNER JOIN materials ON familiesmaterials.codi = materials.familia) ON familiescolorants.codi = materials.familiacol WHERE (((materials.codi)=" + atrim(codimat) + "));", dbOpenSnapshot, dbReadOnly)
  If Not rst.EOF Then nomdelmaterial = atrim(rst!descfammat) + " - " + atrim(rst!descfamcol)
  Set rst = Nothing
End Function
Function capesdelacomanda(numc As Double, numc2 As Double, numc3 As Double) As Byte
   If numc3 > 0 Then capesdelacomanda = 3: Exit Function
   If numc2 > 0 Then capesdelacomanda = 2: Exit Function
   capesdelacomanda = 1: Exit Function
End Function
Function nomclient(codi As Long) As String
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select nom from clients where codi=" + atrim(codi), dbOpenSnapshot, dbReadOnly)
  If Not rst.EOF Then
     nomclient = atrim(rst!nom)
    Else: nomclient = ""
  End If
  Set rst = Nothing
  
End Function

Function direnvioclient(direnvio As Long) As String
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select nome,poblacioe from clients_envios where id=" + atrim(direnvio), dbOpenSnapshot, dbReadOnly)
  If Not rst.EOF Then
     direnvioclient = atrim(rst!nome) + " -> " + atrim(rst!poblacioe)
    Else: direnvioclient = ""
  End If
  Set rst = Nothing
  
End Function

Sub bxrcontrolagafafocus(i As Integer)
  Dim cntrl As Control
  Set cntrl = Screen.ActiveControl
  If cntrl.Text <> "" Then
     If cntrl.Text = camps(cadbl(filtre(i).tag), 3) Then cntrl.Text = ""
     cntrl.ForeColor = QBColor(0)
   Else:
       cntrl.Text = camps(cadbl(filtre(i).tag), 3)
       cntrl.ForeColor = &H808080
  End If
End Sub

Private Sub Command3_Click()
  Dim ultimregistre As Double
  If taulaplanificacio = "planificacioent" Then Exit Sub
  dbconsulta.Execute "delete * from llistat" + taulaplanificacio
  If (exportarapdf.tag = "exportar" Or exportaraxls.tag = "exportar") And Command3.tag <> "" Then
       dbconsulta.Execute "insert into llistat" + taulaplanificacio + " select * from " + taulaplanificacio + " where " + Command3.tag + " " + ordrereixa
    Else
     If taulaplanificacio <> "planificaciototes" Then
      If taulaplanificacio <> "planificaciosol" And multiseleccio = 0 Then
        ultimregistre = cadbl(InputBox("Entra l'ultim numero d'ordre que vols que s'imprimeixi.", "Impresió", "30"))
          Else: ultimregistre = 999
      End If
      If ultimregistre = 0 Then Exit Sub
      dbconsulta.Execute "insert into llistat" + taulaplanificacio + " select * from " + taulaplanificacio + " where ordre<=" + atrim(ultimregistre) + " and maquina=" + atrim(nummaquina) + ""
      If taulaplanificacio = "planificacioreb" Then afegircampsextres
        Else: dbconsulta.Execute "insert into llistat" + taulaplanificacio + "  select * from planificaciototes "
    End If
  End If
  ratoli "espera"
  wait 5
  ratoli "normal"
  If multiseleccio = 1 Then
     borrarelsnoseleccionats
     ultimregistre = 1000
  End If
  wait 2
  imprimirllistat ultimregistre
End Sub
Sub afegircampsextres()
   Dim rsttotes As Recordset
   Dim rstreb As Recordset
   Dim vmaterialpc As String
   Dim vmaterialpc2 As String
   Set rsttotes = dbconsulta.OpenRecordset("select * from planificaciototes", dbOpenSnapshot, dbReadOnly)
   Set rstreb = dbconsulta.OpenRecordset("select * from planificacioreb", dbOpenSnapshot, dbReadOnly)
   While Not rstreb.EOF
      rsttotes.FindFirst "comanda=" + atrim(rstreb!comanda)
      If Not rsttotes.NoMatch Then
          vmaterialpc = IIf(Len(rsttotes!materialpc) > 3, Mid(rsttotes!materialpc, 1, 5), "")
          vmaterialpc2 = IIf(Len(rsttotes!materialpc2) > 3, Mid(rsttotes!materialpc2, 1, 5), "")
          If vmaterialpc <> "" Or vmaterialpc2 <> "" Then dbconsulta.Execute "update llistatplanificacioreb set materialPC='" + vmaterialpc + "',materialPC2='" + vmaterialpc2 + "' where comanda=" + atrim(rstreb!comanda)
      End If
      rstreb.MoveNext
   Wend
   Set rstreb = Nothing
   Set rsttotes = Nothing
End Sub
Sub borrarelsnoseleccionats()
   Dim fila As Long
   Dim filainici As Long
   Dim filafi As Long
   filainici = seleccionats("inici")
   filafi = seleccionats("fi")
   fila = 0
   While fila <= reixa.Rows - 1
     If fila < filainici Or fila > filafi Then
         dbconsulta.Execute "delete * from llistat" + taulaplanificacio + " where comanda=" + atrim(cadbl(reixa.TextMatrix(fila, numcol("NºLot"))))
     End If
     fila = fila + 1
   Wend
   
   
End Sub
Function nommaquina(maq As Byte) As String
  Dim rst As Recordset
  Dim seccio As String
  seccio = UCase(Mid(campmaquina, 1, 1))
  Set rst = dbcomandes.OpenRecordset("select * from maquines where maquina='" + seccio + "' and codi=" + atrim(maq), dbOpenSnapshot, dbReadOnly)
  If Not rst.EOF Then
    nommaquina = rst!descripcio
  End If
  Set rst = Nothing
End Function
Sub imprimirllistat(ultimregistre As Double)
  Dim vordre As String
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Dim vdireccioordre As CRSortDirection
  
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistat" + taulaplanificacio + ".rpt", 1)
  oreport.Database.Tables.Item(1).Location = fitxertemp

  
  oreport.DiscardSavedData
  vordre = substituir(atrim(ordrereixa), " order by", "")
  If InStr(1, vordre, ",") > 0 Then vordre = Mid(vordre, 1, InStr(1, vordre, ",") - 1)
  If taulaplanificacio <> "planificaciototes" Then
     If Command3.tag = "" Then oreport.RecordSelectionFormula = "{" + taulaplanificacio + ".ordre}<=" + atrim(ultimregistre) + " and {" + taulaplanificacio + ".maquina}=" + atrim(nummaquina)
     If ultimregistre <> 1000 Then
        If vordre = "" Then vordre = "ordre"
        vdireccioordre = crAscendingOrder
        If InStr(1, UCase(vordre), "DESC") Then vdireccioordre = crDescendingOrder
        vordre = substituir(atrim((vordre)), "DESC", ""): vordre = substituir(atrim((vordre)), "ASC", "")
        oreport.RecordSortFields.Add oreport.Database.Tables(1).Fields.GetItemByName((atrim(vordre))), vdireccioordre
       Else: oreport.RecordSortFields.Add oreport.Database.Tables(1).Fields.GetItemByName("data1"), crAscendingOrder
     End If
     oreport.FormulaFields.GetItemByName("titol").Text = "'Llistat de la planificacio de la maquina: " + atrim(nummaquina) + " - " + nommaquina(nummaquina) + "'"
       Else: ' oreport.RecordSortFields(0).Field = "{planificaciototes.data1}"
       oreport.FormulaFields.GetItemByName("titol").Text = "'Llistat de la planificacio General'"
  End If
  
  If exportarapdf.tag = "exportar" Then
   oreport.ExportOptions.DiskFileName = "c:\temp\llistatexportat.pdf"
   oreport.ExportOptions.PDFExportAllPages = True
   oreport.ExportOptions.FormatType = crEFTPortableDocFormat
   oreport.ExportOptions.DestinationType = crEDTDiskFile
   oreport.Export False
   exportarapdf.tag = ""
   GoTo fi
  End If
  If exportaraxls.tag = "exportar" Then
   oreport.ExportOptions.DiskFileName = "c:\temp\llistatexportat.xls"
   oreport.ExportOptions.PDFExportAllPages = True
   oreport.ExportOptions.FormatType = crEFTExcel80Tabular
   oreport.ExportOptions.DestinationType = crEDTDiskFile
   oreport.Export False
   exportaraxls.tag = ""
   GoTo fi
  End If
  oreport.PageEngine.ValueFormatOptions = crIncludeFieldValues
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
    Else
      oreport.PrintOut False, 1
  End If
fi:
  planificacio.SetFocus

End Sub
Private Sub filtre_DblClick(Index As Integer)
 ' If Index <> -1 Then
 '  triar_ordre camps(filtre(Index).tag, 1)
 ' End If
 If camps(cadbl(filtre(Index).tag), 1) = "numeroliniaimpresio" Then
  If vhihanCdLtaronja And eseccio = "General - Totes les seccions" Then
            filtre(Index) = "TARONJA"
            filtre_LostFocus Index
  End If
 End If
End Sub

Private Sub filtre_GotFocus(Index As Integer)
  bxrcontrolagafafocus Index
  ultimfiltre = Index
  If camps(cadbl(filtre(Index).tag), 1) = "clixes" Then postitfiltre "Pots posar BLANC TARONJA VERD i a davant NO si no el vols", True
  If camps(cadbl(filtre(Index).tag), 1) = "nomclient" Then postitfiltre "També pots posar els codis dels clients separats per comes o fer alias amb (). EX. (ardo) 6320,6415   La proxima vegada (ardo) serà aquests codis", True
  If camps(cadbl(filtre(Index).tag), 1) = "materialPC" Then postitfiltre "Pots posar BLANC TARONJA VERD i a davant NO si no el vols", True
  If camps(cadbl(filtre(Index).tag), 1) = "materialPC2" Then postitfiltre "Pots posar BLANC TARONJA VERD i a davant NO si no el vols", True
  If camps(cadbl(filtre(Index).tag), 1) = "numeroliniaimpresio" Then postitfiltre "Pots posar TARONJA VERD VERMELL per filtrar aquests colors.", True
  
End Sub
Sub postitfiltre(msg As String, ensenyar As Boolean)
   
   postit.Text = msg
   postit.Left = Screen.ActiveControl.Left
   postit.Top = Screen.ActiveControl.Top + Screen.ActiveControl.Height + 20
   postit.width = Len(msg) * 80
   postit.visible = ensenyar
   
End Sub
Sub buscarcomanda(numc As Double)
  Dim i As Integer
  Dim trobat As Boolean
  Dim ultimafila As Integer
  Dim vnumcolumna As Double

  trobat = False
  vnumcolumna = numcol("NºLot")
  i = 1
  While Not trobat And i < reixa.Rows
    reixa.Row = i
    If InStr(1, reixa.TextMatrix(reixa.Row, vnumcolumna), atrim(numc)) > 0 Then trobat = True
    i = i + 1
  Wend
  If trobat And numc > 0 Then

    reixa.SetFocus
    ultimafila = reixa.Row - 17
    If ultimafila > reixa.Rows Or ultimafila < 1 Then ultimafila = 1
    reixa.TopRow = ultimafila
    reixa.col = 1
    reixa_Click
  End If
End Sub
Sub creariposaralias(Index As Integer)
   Dim alias As String
   Dim v As String
   Dim guardat As String
   If Index > filtre.Count Then Exit Sub
   v = filtre(Index)
   If Mid(v, 1, 1) = "(" And InStr(1, v, ")") > 0 Then
     alias = UCase(Mid(v, 2, InStr(1, v, ")") - 2))
     v = Mid(v, InStr(1, v, ")") + 1)
     If atrim(v) = "" Then
        guardat = llegir_ini("Planificacio", "alias_" + alias, "comandes.ini")
        If guardat = "{[}]" Then MsgBox "No hi ha res guardat amb aquest alias" + Chr(10) + "Despres dels parentesis has de posar els valors que vols que substitueixin.", vbInformation, "Atenció": Exit Sub
        filtre(Index) = guardat
          Else:
             escriure_ini "Planificacio", "alias_" + alias, atrim(v), "comandes.ini"
             filtre(Index) = v
     End If
   End If
End Sub

Sub ensenyar_nomes_aquestscolors(vcamp As String, vcolor As String)
   Dim j As Double
   Dim vcol As Long
   Dim vcolorn As Double
   reixa.Redraw = False
   vcol = numcol(vcamp)
   vcolorn = IIf(vcolor = "TARONJA", &H80C0FF, IIf(vcolor = "VERD", QBColor(10), IIf(vcolor = "VERMELL", QBColor(12), 0)))
   For j = 1 To reixa.Rows - 1
     reixa.Row = j
     reixa.col = vcol
     If reixa.CellBackColor <> vcolorn Then reixa.RowHeight(j) = 1
   Next j
   reixa.Redraw = True
End Sub
Private Sub filtre_LostFocus(Index As Integer)
  Static vjaesticdins As Boolean
  Dim noufiltre As String
  If vjaesticdins Then Exit Sub
  vjaesticdins = True
  creariposaralias Index
  postitfiltre "", False
  If Index = 998 Then whereultimfiltre = "": vjaesticdins = False: Exit Sub
  If camps(cadbl(filtre(ultimfiltre).tag), 1) = "comanda" Then
    If cadbl(filtre(ultimfiltre)) > 0 Then
      buscarcomanda cadbl(filtre(ultimfiltre))
      vjaesticdins = False
      Exit Sub
    End If
  End If
  'si es filtre per numeroliniaimpresió només volen veure estats I i E (ho ha dit l'Esaú data 29/06/2023
  If camps(cadbl(filtre(ultimfiltre).tag), 1) = "numeroliniaimpresio" Then
        If taulaplanificacio = "planificaciototes" Then filtre(numcol("Estat")) = IIf(filtre(ultimfiltre) <> "", "E,I", "")
        If UCase(filtre(ultimfiltre)) = "VERD" Or UCase(filtre(ultimfiltre)) = "TARONJA" Or UCase(filtre(ultimfiltre)) = "VERMELL" Then
             ensenyar_nomes_aquestscolors "NºLinia Imp", UCase(filtre(ultimfiltre))
             filtre(ultimfiltre) = ""
             If UCase(filtre(ultimfiltre)) <> "TARONJA" Then vhihanCdLtaronja = False
             GoTo fi
        End If
  End If
  noufiltre = crearfiltre
  If filtre(ultimfiltre).Text = "" Then
    filtre(ultimfiltre).Text = camps(cadbl(filtre(ultimfiltre).tag), 3)
    filtre(ultimfiltre).ForeColor = &H808080
  End If
  If noufiltre <> whereultimfiltre Or Index = 999 Then
     If noufiltre <> "" Then poblarlareixa nummaquina, " and " + noufiltre
     If noufiltre = "" Then borrarelfiltre
  End If
  
  If Index = 999 And noufiltre = "" Then
     poblarlareixa nummaquina
  End If
fi:
  ratoli "normal"
  reixa.visible = True
  reixa.SetFocus
  DoEvents
  whereultimfiltre = noufiltre
  Command3.tag = noufiltre ' el guardo pel llistat
  vjaesticdins = False
  
End Sub
Function crearfiltre() As String
  Dim i As Integer
  Dim were As String
  Dim w As String
  For i = 0 To filtre.Count - 1
    If filtre(i).Text <> camps(cadbl(filtre(i).tag), 3) And Not (camps(cadbl(filtre(i).tag), 1) = "comanda" And cadbl(filtre(i)) > 0) Then
      w = crearwere(i)
      If were = "" Then
         were = w
        Else: If w <> "" Then were = were + " and " + w
      End If
    End If
  Next i
  If taulaplanificacio = "planificacioent" And Checknomesclixes.Value = 1 Then were = were + IIf(were <> "", " and ", were) + " impresio<>'R' "
  crearfiltre = were
End Function

Function generarwerecolorsMATERIALPC(camp As String, filtre As String) As String
  Dim re As String
  Dim partfiltre As String
  Dim negat As String
'camps(j, 1) + " LIKE '*" + treure_apostruf(filtre(i)) + "*'"
  filtre = filtre + ","
  'If camp <> "clixes" Then generarwerecolorsdates = ""
  While InStr(1, filtre, ",") > 0 And filtre <> ""
    partfiltre = UCase(Mid(filtre, 1, InStr(1, filtre, ",") - 1))
    re = IIf(re <> "", re + " or ", "") + camp
    If Mid(partfiltre, 1, 2) = "NO" Then negat = "!": partfiltre = Trim(Mid(partfiltre, 3))
    Select Case UCase(partfiltre)
          Case "VERD"
             re = IIf(negat = "!", "Not ", "") + " (len(" + re + ")>0 and len(" + re + ")<8)"
          Case "TARONJA"
             re = " LEN(" + re + ")" + IIf(negat = "!", "<8 ", "=8 ")
          Case "BLANC"
             re = re + IIf(negat = "!", " <>''", " =''")
             
         Case Else
             re = re + " like '*" + partfiltre + "*'"
    End Select
    filtre = Mid(filtre, InStr(1, filtre, ",") + 1)
    negat = ""
  Wend
  If re <> "" Then re = "(" + re + ")"
  generarwerecolorsMATERIALPC = re
End Function
Function generarwerecolorsdates(camp As String, filtre As String) As String
  Dim re As String
  Dim partfiltre As String
  Dim negat As String
'camps(j, 1) + " LIKE '*" + treure_apostruf(filtre(i)) + "*'"
  filtre = filtre + ","
  If camp <> "clixes" Then generarwerecolorsdates = ""
  While InStr(1, filtre, ",") > 0 And filtre <> ""
    partfiltre = UCase(Mid(filtre, 1, InStr(1, filtre, ",") - 1))
    re = IIf(re <> "", re + " or ", "") + camp
    If Mid(partfiltre, 1, 2) = "NO" Then negat = "!": partfiltre = Trim(Mid(partfiltre, 3))
    Select Case UCase(partfiltre)
          Case "VERD"
             re = re + " like '[" + negat + "*]*'"
          Case "TARONJA"
             re = re + " like '[" + negat + "0-9]*'"
          Case "BLANC"
             re = re + IIf(negat = "!", "<>", "=") + " ''"
         Case Else
             re = re + " like '*" + partfiltre + "*'"
    End Select
    filtre = Mid(filtre, InStr(1, filtre, ",") + 1)
    negat = ""
  Wend
  If re <> "" Then re = "(" + re + ")"
  generarwerecolorsdates = re
End Function
Function crearwere(i As Integer) As String
   Dim w As String
   Dim j As Integer
   If filtre(i) = "" Then Exit Function
   j = cadbl(filtre(i).tag)
   If camps(j, 2) = "date" Then
      If IsDate(filtre(i)) Then
         crearwere = camps(j, 1) + "=#" + format(filtre(i), "mm/dd/yy") + "# "
      End If
      Exit Function
   End If
   If InStr(1, camps(j, 2), "string") > 0 Or camps(j, 1) = "comanda" Then
       If camps(j, 1) <> "clixes" And camps(j, 1) <> "materialPC2" And camps(j, 1) <> "materialPC" Then
         crearwere = possarweres(camps(j, 1), "LIKE", treure_apostruf(filtre(i)))
         Else
            If camps(j, 1) = "clixes" Then crearwere = generarwerecolorsdates(camps(j, 1), filtre(i))
            If camps(j, 1) = "materialPC2" Or camps(j, 1) = "materialPC" Then crearwere = generarwerecolorsMATERIALPC(camps(j, 1), filtre(i))
       End If
       Exit Function
   End If
   crearwere = camps(j, 1) + "=" + passaradecimalpunt(atrim(cadbl(filtre(i))))
   

End Function
Function possarweres(ByVal camp As String, condicio As String, ByVal filtre As String) As String
  Dim re As String
'camps(j, 1) + " LIKE '*" + treure_apostruf(filtre(i)) + "*'"
  filtre = filtre + ","
  If camp = "nomclient" And cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1)) > 0 Then camp = "codiclient"
  While InStr(1, filtre, ",") > 0 And filtre <> ""
    If camp <> "codiclient" Then
       If Mid(filtre, 1, InStr(1, filtre, ",") - 1) <> " " Then  ' si es espai sol busco nomes espai sense semblants per poder filtrar els camps sense res
          re = IIf(re <> "", re + " or ", "") + camp + " like '*" + Mid(filtre, 1, InStr(1, filtre, ",") - 1) + "*'"
           Else: re = IIf(re <> "", re + " or ", "") + camp + " like '" + Mid(filtre, 1, InStr(1, filtre, ",") - 1) + "'"
       End If
      Else: re = IIf(re <> "", re + " or ", "") + camp + " =" + atrim(cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1))) + ""
    End If
    filtre = Mid(filtre, InStr(1, filtre, ",") + 1)
  Wend
  If re <> "" Then re = "(" + re + ")"
  possarweres = re
End Function
Private Sub Form_Load()
  Dim arguments As Variant
  If App.PrevInstance Then End
  
  arguments = ObtenerLíneaComando
 ' arguments(1) = "OPERARIS"
 ' arguments(2) = "NOMESLECTURA"
  If UCase(atrim(arguments(1))) = "OPERARIS" Then
      programaoperaris = True
  End If
  
  If UCase(atrim(arguments(2))) = "NOMESLECTURA" Then
      programanomeslectura = True
  End If
      
  If UCase(atrim(arguments(3))) <> "" Then primerapestanya = UCase(atrim(arguments(3)))
  'programaoperaris = True
  cami = llegir_ini("General", "cami", "comandes.ini")
  If llegir_ini("General", "usuari", "comandes.ini") = "Usr_Reb" Then
      AcroPDF1.tag = "no"
  End If
  Set dbcomandes = OpenDatabase(llegir_ini("General", "cami", "comandes.ini"))
  Set dbbaixes = OpenDatabase(rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "baixes.mdb")
  Set dbtintes = OpenDatabase(rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "tintes.mdb")
  Set dbplanificaciooperaris = OpenDatabase(rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "planificaciooperaris.mdb")
  Set dbplanificacioalicia = OpenDatabase(rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "planificacio.mdb")
  If Not programaoperaris Then
     Set dbplanificacio = OpenDatabase(rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "planificacio.mdb")
     fitxertemp = "c:\temp\planificaciotmp.mdb"
    Else:
           Set dbplanificacio = OpenDatabase(rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "planificaciooperaris.mdb")
           fitxertemp = "c:\temp\planificaciooperaristmp.mdb"
  End If
  
  Set dbclixes = OpenDatabase(rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "clixesnous.mdb")
  Set dbcompres = OpenDatabase(rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "compres.mdb")
  Set dbstocks = OpenDatabase(rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "palets.mdb")
  
  mirarsicopiaractualitzaciodelservidor
  If llegir_ini("Planificacio", "ultimaactualitzacio", "comandes.ini") <> "{[}]" Then
     v = llegir_ini("Planificacio", "ultimaactualitzacio", "comandes.ini")
     If v = "" Or v = "{[}]" Then v = Now
     ultimaactualitzacio = CVDate(v)
     etultimaactualitzacio = format(ultimaactualitzacio, "dd/mm hh:nn")
  End If
    'canviant el numero de versió obligo a actualitzar la  base de dades... també s'ha de canviar al escriure_ini seguent per ferho nomes un cop
  If llegir_ini("Planificacio", "versiobd", "comandes.ini") <> "56" Then
   If existeix(fitxertemp) Then Kill fitxertemp
   crearfitxertemp
   escriure_ini "Planificacio", "versiobd", "56", "comandes.ini"
  End If
  '&HFFC0C0
  If programaoperaris Then fcontrols.BackColor = &HDCB8FC: etiquetaoperaris = "Operaris": etiquetaoperaris.visible = True
  If programanomeslectura Then cNomesLectura.visible = True
  poteditar = IIf(llegir_ini("Planificacio", "poteditar", "comandes.ini") = "Si", True, False)
  If Not poteditar Then escriure_ini "Planificacio", "poteditar", "No", "comandes.ini"
  If llegir_ini("General", "usuari", fitxerini) = "Usr_A" Then poteditar = True
  taulaplanificacio = "planificacioimp"
  iniconfigreixa = "reixaplanificacio.ini"
  
  
  
  carregarllistadecampstemporals "I"
  crearfitxertemp
  carregartamanyform
  configreixa
  carregarmaquines "I"
  ordrereixa = triar_ordre_reixa
  If UCase(atrim(arguments(1))) <> "GENERARFITXERTEMPORAL" Then
      carregarpestanyapredeterminada
       Else: generarelfitxertemporal = True
  End If
  
  'borro l'historial de programacio de comandes de mes de tres mesos
  dbplanificacio.Execute "delete * from llistadeplanificacions where dataordreassignat<#" + atrim(format(DateAdd("m", -3, Now), "mm/dd/yy hh:nn") + "#")
  treure_comandesdataexpedicioerronees
End Sub
Sub treure_comandesdataexpedicioerronees()
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim vdates As String
  
  Set rst = dbplanificacioalicia.OpenRecordset("SELECT capcaleraalbara.dataenvioasap, * FROM linies_expedicions LEFT JOIN capcaleraalbara ON linies_expedicions.albara = capcaleraalbara.numalbara WHERE (((linies_expedicions.enviat)=False) AND ((capcaleraalbara.dataenvioasap) Is Null));")
  While Not rst.EOF
    Set rst2 = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where dataexpedicio=#" + atrim(format(rst!data, "mm/dd/yy")) + "# and comanda=" + atrim(rst!comanda))
    If rst2.EOF Then dbplanificacioalicia.Execute "delete * from linies_expedicions where not enviat and comanda=" + atrim(rst!comanda) + " and data=#" + atrim(format(rst!data, "mm/dd/yy")) + "#"
    rst.MoveNext
  Wend
  
  Set rst = Nothing
End Sub
Sub mirarsicopiaractualitzaciodelservidor()
   Dim vdatalocal As Date
   Dim vdataservidor As Date
   Dim v As String
   
   v = llegir_ini("Planificacio", "ultimaactualitzacio", rutadelfitxer(cami) + "\actualitzacioplanificacio.ini")
   If v = "{[}]" Or v = "" Then Exit Sub
   v = llegir_ini("Planificacio", "ultimaactualitzacio", "comandes.ini")
   If v = "{[}]" Or v = "" Then Exit Sub
   vdatalocal = CVDate(llegir_ini("Planificacio", "ultimaactualitzacio", "comandes.ini"))
   vdataservidor = CVDate(llegir_ini("Planificacio", "ultimaactualitzacio", rutadelfitxer(cami) + "\actualitzacioplanificacio.ini"))
   If DateDiff("s", vdatalocal, vdataservidor) > 0 Then
      If existeix(rutadelfitxer(cami) + "\planificaciotemporal.mdb") Then
         If existeix(fitxertemp) Then Kill fitxertemp
         FileCopy rutadelfitxer(cami) + "\planificaciotemporal.mdb", fitxertemp
         escriure_ini "Planificacio", "ultimaactualitzacio", atrim(vdataservidor), "comandes.ini"
      End If
   End If
End Sub
Sub carregarpestanyapredeterminada()
  If primerapestanya = "" Then mgeneral_Click
  If primerapestanya = "IMPRESORES" Then mimpresores_Click
  If primerapestanya = "LAMINADORES" Then mlaminadores_Click
  If primerapestanya = "REBOBINADORES" Then mrebobinadores_Click
  If primerapestanya = "SOLDADORES" Then msoldadores_Click
End Sub
Sub carregarmaquines(seccio As String)
  Dim rst As Recordset
  Dim i As Byte
  AcroPDF1.visible = False
  botoreclamar.visible = False
  Command67(11).visible = False
  If seccio = "I" Then Command67(11).visible = True 'botoreclamar.Visible = True:
  Set rst = dbcomandes.OpenRecordset("select * from maquines where maquina='" + seccio + "' and donadadebaixa =null", dbOpenSnapshot, dbReadOnly)
  For i = 0 To 5
    botomaquina(i).visible = False
    botomaquina(i).tag = ""
  Next i
  i = 0
  If Not rst.EOF And nummaquina = 0 Then nummaquina = rst!codi
  While Not rst.EOF
   If i = 6 Then GoTo cont
    botomaquina(i).caption = atrim(rst!codi) + "-" + atrim(rst!descripcio)
    botomaquina(i).visible = True
    botomaquina(i).tag = atrim(rst!codi)
    i = i + 1
    rst.MoveNext
  Wend
cont:
   Set rst = Nothing
   botomaquina(0).BackColor = QBColor(10)
End Sub
Sub descarregarfiltres()
  Dim i As Byte
  For i = 1 To filtre.Count - 1
   Unload filtre.Item(i)
  Next i
End Sub
Sub configreixa(Optional nocarregaramples As Boolean)
  Dim rst As Recordset
  Dim col As Long
  Dim enes As Byte
  If Not nocarregaramples Then descarregarfiltres
  reixa.LeftCol = 0
  'reixa.Redraw = True
  If reixa.Rows > 1 Then reixa.TopRow = 1
  Set rst = dbconsulta.OpenRecordset("select * from " + taulaplanificacio, dbOpenSnapshot, dbReadOnly)
  col = 0
  enes = 0
  reixa.Cols = rst.Fields.Count
  For i = 0 To rst.Fields.Count - 1
    If camps(i + 1, 4) <> "N" Then
       reixa.ColAlignment(col) = 2
       reixa.TextMatrix(0, col) = camps(i + 1, 3)
       If Not nocarregaramples Then colocarfiltre col, i + 1
       col = col + 1
       'If camps(i + 1, 1) = "" Then reixa.Cols = reixa.Cols - 1
        Else: enes = enes + 1
    End If
  Next i
  If enes = 0 Then If taulaplanificacio = "planificacioent" Then reixa.Cols = reixa.Cols - 1
     
  If enes > 0 Then reixa.Cols = reixa.Cols - (enes + 1)
  If Not nocarregaramples Then carregar_amples_reixa
  reixa.Row = 0
  'For i = 0 To reixa.Cols - 1
  '  reixa.col = i
  '  reixa.ColSel = i
  '  reixa.CellBackColor = QBColor(8)
  'Next i
  Set rst = Nothing
End Sub
Sub colocarfiltre(col As Long, i As Long)
  If filtre.Count <= col Then Load filtre(col)
  filtre(col).Text = camps(i, 3)
  filtre(col).tag = i
'  Load filtre(col + 1)
End Sub
Sub carregarllistadecampstemporals(seccio As String)
  Dim i As Byte
  
  mgeneral.tag = ""
  mimpresores.tag = ""
  mlaminadores.tag = ""
  mrebobinadores.tag = ""
  
  For i = 1 To 50
     camps(i, 1) = "": camps(i, 2) = "": camps(i, 3) = "": camps(i, 4) = ""
  Next i
  i = 1
  If seccio = "I" Then
    camps(i, 1) = "ordre": camps(i, 2) = "double": camps(i, 3) = "Ordre": i = i + 1
    camps(i, 1) = "codiclient": camps(i, 2) = "long": camps(i, 3) = "Codiclient": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "maquina": camps(i, 2) = "byte": camps(i, 3) = "Maq": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "dataimpresio": camps(i, 2) = "date": camps(i, 3) = "DataImp.": i = i + 1
    camps(i, 1) = "nomclient": camps(i, 2) = "string": camps(i, 3) = "Nom Client": i = i + 1
    camps(i, 1) = "dataactcomanda": camps(i, 2) = "date": camps(i, 3) = "DataAc": i = i + 1
    camps(i, 1) = "comanda": camps(i, 2) = "long": camps(i, 3) = "NºLot": i = i + 1
    camps(i, 1) = "capes": camps(i, 2) = "byte": camps(i, 3) = "NºCapes": i = i + 1
    camps(i, 1) = "tipusimpresio": camps(i, 2) = "string(1)": camps(i, 3) = "T.Imp": i = i + 1
    camps(i, 1) = "material": camps(i, 2) = "string(10)": camps(i, 3) = "Material": i = i + 1
    camps(i, 1) = "mts": camps(i, 2) = "double": camps(i, 3) = "Metres": i = i + 1
    camps(i, 1) = "ample": camps(i, 2) = "double": camps(i, 3) = "Ample": i = i + 1
    camps(i, 1) = "espesor": camps(i, 2) = "string(10)": camps(i, 3) = "Espesor": i = i + 1
    camps(i, 1) = "descmat": camps(i, 2) = "string": camps(i, 3) = "Desc. Material": i = i + 1
    camps(i, 1) = "refclient": camps(i, 2) = "string(50)": camps(i, 3) = "RefClient": i = i + 1
    camps(i, 1) = "numeroliniaimpresio": camps(i, 2) = "string": camps(i, 3) = "NºLinia Imp": i = i + 1
    camps(i, 1) = "texteimpresio": camps(i, 2) = "string": camps(i, 3) = "Texte Impresio": i = i + 1
    camps(i, 1) = "impresio": camps(i, 2) = "string(1)": camps(i, 3) = "Impresio": i = i + 1
    camps(i, 1) = "clixes": camps(i, 2) = "string(10)": camps(i, 3) = "Clixes": i = i + 1
    camps(i, 1) = "gruixclixes": camps(i, 2) = "double": camps(i, 3) = "Gruix Clixes": i = i + 1
    camps(i, 1) = "cilindre": camps(i, 2) = "double": camps(i, 3) = "Cilindre": i = i + 1
    camps(i, 1) = "Muntat": camps(i, 2) = "string(1)": camps(i, 3) = "Muntat?": i = i + 1
    camps(i, 1) = "tintesrevisades": camps(i, 2) = "string": camps(i, 3) = "Tintes revisades": i = i + 1
    camps(i, 1) = "data1": camps(i, 2) = "date": camps(i, 3) = "Data1": i = i + 1
    camps(i, 1) = "data2": camps(i, 2) = "date": camps(i, 3) = "Data2": i = i + 1
    camps(i, 1) = "clientvindraarevisarimpresio": camps(i, 2) = "string": camps(i, 3) = "Cli.Vindrà": i = i + 1
    camps(i, 1) = "dataoperari": camps(i, 2) = "date": camps(i, 3) = "DataOp.": i = i + 1
    camps(i, 1) = "importancia": camps(i, 2) = "byte": camps(i, 3) = "Impor.": i = i + 1
    camps(i, 1) = "standbyimpresio": camps(i, 2) = "string(2)": camps(i, 3) = "StandBy": i = i + 1
    camps(i, 1) = "tempsimpresio": camps(i, 2) = "double": camps(i, 3) = "Temps": i = i + 1
    camps(i, 1) = "observacions": camps(i, 2) = "string": camps(i, 3) = "Observacio": i = i + 1
    camps(i, 1) = "horaprogramada": camps(i, 2) = "date": camps(i, 3) = "Hora Programada": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "numtreball": camps(i, 2) = "string": camps(i, 3) = "NºTreball": camps(i, 4) = "N": i = i + 1
  End If
  If seccio = "L" Then
    camps(i, 1) = "ordre": camps(i, 2) = "double": camps(i, 3) = "Ordre": i = i + 1
    camps(i, 1) = "codiclient": camps(i, 2) = "long": camps(i, 3) = "Codiclient": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "maquina": camps(i, 2) = "byte": camps(i, 3) = "Maq": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "estat": camps(i, 2) = "string(2)": camps(i, 3) = "Estat": i = i + 1
    camps(i, 1) = "camisa": camps(i, 2) = "double": camps(i, 3) = "Camisa": i = i + 1
    camps(i, 1) = "dataimpresio": camps(i, 2) = "date": camps(i, 3) = "DataLam.": i = i + 1
    camps(i, 1) = "nomclient": camps(i, 2) = "string": camps(i, 3) = "Nom Client": i = i + 1
    camps(i, 1) = "comanda": camps(i, 2) = "long": camps(i, 3) = "NºLot": i = i + 1
    camps(i, 1) = "capes": camps(i, 2) = "byte": camps(i, 3) = "NºCapes": i = i + 1
    camps(i, 1) = "materialPC": camps(i, 2) = "string(10)": camps(i, 3) = "MaterialPC": i = i + 1
    camps(i, 1) = "materialPC2": camps(i, 2) = "string(10)": camps(i, 3) = "MaterialPC2": i = i + 1
    camps(i, 1) = "tipuscola": camps(i, 2) = "string(40)": camps(i, 3) = "TipusCola": i = i + 1
    camps(i, 1) = "mts": camps(i, 2) = "double": camps(i, 3) = "Metres": i = i + 1
    camps(i, 1) = "ample": camps(i, 2) = "double": camps(i, 3) = "Ample": i = i + 1
    camps(i, 1) = "descmat": camps(i, 2) = "string": camps(i, 3) = "Desc. Material": i = i + 1
    camps(i, 1) = "refclient": camps(i, 2) = "string(50)": camps(i, 3) = "RefClient": i = i + 1
    camps(i, 1) = "texteimpresio": camps(i, 2) = "string": camps(i, 3) = "Texte Impresio": i = i + 1
    camps(i, 1) = "data1": camps(i, 2) = "date": camps(i, 3) = "Data1": i = i + 1
    camps(i, 1) = "data2": camps(i, 2) = "date": camps(i, 3) = "Data2": i = i + 1
    camps(i, 1) = "dataoperari": camps(i, 2) = "date": camps(i, 3) = "DataOp.": i = i + 1
    camps(i, 1) = "importancia": camps(i, 2) = "byte": camps(i, 3) = "Impor.": i = i + 1
    camps(i, 1) = "tempsimpresio": camps(i, 2) = "double": camps(i, 3) = "Temps": i = i + 1
    camps(i, 1) = "observacions": camps(i, 2) = "string": camps(i, 3) = "Observacio": i = i + 1
    camps(i, 1) = "horaprogramada": camps(i, 2) = "date": camps(i, 3) = "Hora Programada": camps(i, 4) = "N": i = i + 1
  End If
  
  If seccio = "R" Then
    camps(i, 1) = "ordre": camps(i, 2) = "double": camps(i, 3) = "Ordre": i = i + 1
    camps(i, 1) = "codiclient": camps(i, 2) = "long": camps(i, 3) = "Codiclient": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "maquina": camps(i, 2) = "byte": camps(i, 3) = "Maq": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "dataimpresio": camps(i, 2) = "date": camps(i, 3) = "DataReb.": i = i + 1
    camps(i, 1) = "mandril": camps(i, 2) = "double": camps(i, 3) = "Mandril": camps(i, 4) = "": i = i + 1
    camps(i, 1) = "comanda": camps(i, 2) = "long": camps(i, 3) = "NºLot": i = i + 1
    camps(i, 1) = "producte": camps(i, 2) = "string(10)": camps(i, 3) = "Producte": i = i + 1
    camps(i, 1) = "estat": camps(i, 2) = "string(2)": camps(i, 3) = "Estat": i = i + 1
    camps(i, 1) = "nomclient": camps(i, 2) = "string": camps(i, 3) = "Nom Client": i = i + 1
    camps(i, 1) = "refclient": camps(i, 2) = "string(50)": camps(i, 3) = "RefClient": i = i + 1
    camps(i, 1) = "texteimpresio": camps(i, 2) = "string": camps(i, 3) = "Texte Impresio": i = i + 1
    camps(i, 1) = "mts": camps(i, 2) = "double": camps(i, 3) = "Metres": i = i + 1
    camps(i, 1) = "ample": camps(i, 2) = "double": camps(i, 3) = "Ample": i = i + 1
    camps(i, 1) = "bandes": camps(i, 2) = "BYTE": camps(i, 3) = "Bandes": i = i + 1
    camps(i, 1) = "merma": camps(i, 2) = "double": camps(i, 3) = "Merma": i = i + 1
    camps(i, 1) = "micromacro": camps(i, 2) = "string(5)": camps(i, 3) = "MIcro/MAcro": i = i + 1
    camps(i, 1) = "data1": camps(i, 2) = "date": camps(i, 3) = "Data1": i = i + 1
    camps(i, 1) = "data2": camps(i, 2) = "date": camps(i, 3) = "Data2": i = i + 1
    camps(i, 1) = "dataoperari": camps(i, 2) = "date": camps(i, 3) = "DataOp.": i = i + 1
    camps(i, 1) = "importancia": camps(i, 2) = "byte": camps(i, 3) = "Impor.": i = i + 1
    camps(i, 1) = "dataexpedicio": camps(i, 2) = "date": camps(i, 3) = "Data_Exp.": i = i + 1
    camps(i, 1) = "tempsimpresio": camps(i, 2) = "double": camps(i, 3) = "Temps": i = i + 1
    camps(i, 1) = "observacions": camps(i, 2) = "string": camps(i, 3) = "Observacio": i = i + 1
    camps(i, 1) = "horaprogramada": camps(i, 2) = "date": camps(i, 3) = "Hora Programada": camps(i, 4) = "N": i = i + 1
    
  End If
  
  If seccio = "S" Then
    camps(i, 1) = "ordre": camps(i, 2) = "double": camps(i, 3) = "Ordre": i = i + 1
    camps(i, 1) = "codiclient": camps(i, 2) = "long": camps(i, 3) = "Codiclient": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "maquina": camps(i, 2) = "byte": camps(i, 3) = "Maq": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "estat": camps(i, 2) = "string(2)": camps(i, 3) = "Estat": i = i + 1
    camps(i, 1) = "dataimpresio": camps(i, 2) = "date": camps(i, 3) = "DataImp.": i = i + 1
    camps(i, 1) = "nomclient": camps(i, 2) = "string": camps(i, 3) = "Nom Client": i = i + 1
    camps(i, 1) = "comanda": camps(i, 2) = "long": camps(i, 3) = "NºLot": i = i + 1
    camps(i, 1) = "producte": camps(i, 2) = "string": camps(i, 3) = "Nom Client": i = i + 1
    camps(i, 1) = "quantitatsol": camps(i, 2) = "double": camps(i, 3) = "Quantitat": i = i + 1
    camps(i, 1) = "amplesol": camps(i, 2) = "double": camps(i, 3) = "AmpleSol": i = i + 1
    camps(i, 1) = "longitud": camps(i, 2) = "double": camps(i, 3) = "Longitud": i = i + 1
    camps(i, 1) = "refclient": camps(i, 2) = "string(50)": camps(i, 3) = "RefClient": i = i + 1
    camps(i, 1) = "texteimpresio": camps(i, 2) = "string": camps(i, 3) = "Texte Impresio": i = i + 1
    camps(i, 1) = "data1": camps(i, 2) = "date": camps(i, 3) = "Data1": i = i + 1
    camps(i, 1) = "data2": camps(i, 2) = "date": camps(i, 3) = "Data2": i = i + 1
    camps(i, 1) = "dataoperari": camps(i, 2) = "date": camps(i, 3) = "DataOp.": i = i + 1
    camps(i, 1) = "importancia": camps(i, 2) = "byte": camps(i, 3) = "Impor.": i = i + 1
    camps(i, 1) = "tempsimpresio": camps(i, 2) = "double": camps(i, 3) = "Temps": i = i + 1
    camps(i, 1) = "observacions": camps(i, 2) = "string": camps(i, 3) = "Observacio": i = i + 1
    camps(i, 1) = "horaprogramada": camps(i, 2) = "date": camps(i, 3) = "Hora Programada": camps(i, 4) = "N": i = i + 1
  End If
  
  If seccio = "C" Then
    camps(i, 1) = "comanda": camps(i, 2) = "long": camps(i, 3) = "NºLot": i = i + 1
    camps(i, 1) = "codiclient": camps(i, 2) = "long": camps(i, 3) = "Codiclient": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "dataactcomanda": camps(i, 2) = "date": camps(i, 3) = "DataAc": i = i + 1
    camps(i, 1) = "estat": camps(i, 2) = "string(10)": camps(i, 3) = "Estat": i = i + 1
    camps(i, 1) = "nomclient": camps(i, 2) = "string": camps(i, 3) = "Nom Client": i = i + 1
    camps(i, 1) = "refclient": camps(i, 2) = "string(50)": camps(i, 3) = "RefClient": i = i + 1
    camps(i, 1) = "comandaclient": camps(i, 2) = "string(50)": camps(i, 3) = "ComandaClient": i = i + 1
    camps(i, 1) = "texteimpresio": camps(i, 2) = "string": camps(i, 3) = "Texte Impresio": i = i + 1
    camps(i, 1) = "liniaimpresio": camps(i, 2) = "long": camps(i, 3) = "LiniaImp":  i = i + 1
    camps(i, 1) = "Estatclixes": camps(i, 2) = "string(10)": camps(i, 3) = "EstatClixes": i = i + 1
    camps(i, 1) = "clixes": camps(i, 2) = "string(10)": camps(i, 3) = "Clixes": i = i + 1
    camps(i, 1) = "material": camps(i, 2) = "string(10)": camps(i, 3) = "Material": i = i + 1
    camps(i, 1) = "materialPC": camps(i, 2) = "string(10)": camps(i, 3) = "MaterialPC": i = i + 1
    camps(i, 1) = "materialPC2": camps(i, 2) = "string(10)": camps(i, 3) = "MaterialPC2": i = i + 1
    camps(i, 1) = "mts": camps(i, 2) = "double": camps(i, 3) = "Metres": i = i + 1
    camps(i, 1) = "tintesrevisades": camps(i, 2) = "string": camps(i, 3) = "Tintes revisades": i = i + 1
    camps(i, 1) = "data1": camps(i, 2) = "date": camps(i, 3) = "Data1": i = i + 1
    camps(i, 1) = "data2": camps(i, 2) = "date": camps(i, 3) = "Data2": i = i + 1
    camps(i, 1) = "clientvindraarevisarimpresio": camps(i, 2) = "string": camps(i, 3) = "Cli.Vindrà": i = i + 1
    camps(i, 1) = "datacalloff": camps(i, 2) = "date": camps(i, 3) = "DataCalloff": i = i + 1
    camps(i, 1) = "importancia": camps(i, 2) = "byte": camps(i, 3) = "Impor.": i = i + 1
    camps(i, 1) = "standbyimpresio": camps(i, 2) = "string(2)": camps(i, 3) = "StandBy": i = i + 1
    camps(i, 1) = "observacions": camps(i, 2) = "string": camps(i, 3) = "Observacio": i = i + 1
    camps(i, 1) = "tipusreclamacio": camps(i, 2) = "string": camps(i, 3) = "Tipus_R": i = i + 1
    camps(i, 1) = "datareclamacio": camps(i, 2) = "date": camps(i, 3) = "Data_R": i = i + 1
    camps(i, 1) = "contestaoficina": camps(i, 2) = "string": camps(i, 3) = "Gestió_Of": i = i + 1
    camps(i, 1) = "datagestio": camps(i, 2) = "date": camps(i, 3) = "Data_Gestió": i = i + 1
    camps(i, 1) = "dataentradafabrica": camps(i, 2) = "byte": camps(i, 3) = "Data_EntradaFàbrica": i = i + 1
    camps(i, 1) = "observacionsoficina": camps(i, 2) = "string": camps(i, 3) = "ObservacioOficina": i = i + 1
  End If
  
   If seccio = "T" Then
    camps(i, 1) = "comanda": camps(i, 2) = "long": camps(i, 3) = "NºLot": i = i + 1
    camps(i, 1) = "codiclient": camps(i, 2) = "long": camps(i, 3) = "Codiclient": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "dataactcomanda": camps(i, 2) = "date": camps(i, 3) = "DataAc": i = i + 1
    camps(i, 1) = "estat": camps(i, 2) = "string(10)": camps(i, 3) = "Estat": i = i + 1
    camps(i, 1) = "tipuscola": camps(i, 2) = "string(40)": camps(i, 3) = "TipusCola": i = i + 1
    camps(i, 1) = "producte": camps(i, 2) = "string(10)": camps(i, 3) = "Producte": i = i + 1
    camps(i, 1) = "nomclient": camps(i, 2) = "string": camps(i, 3) = "Nom Client": i = i + 1
    camps(i, 1) = "direnvioclient": camps(i, 2) = "string": camps(i, 3) = "DirEnvio": i = i + 1
    camps(i, 1) = "refclient": camps(i, 2) = "string(50)": camps(i, 3) = "RefClient": i = i + 1
    camps(i, 1) = "comandaclient": camps(i, 2) = "string(50)": camps(i, 3) = "ComandaClient": i = i + 1
    camps(i, 1) = "texteimpresio": camps(i, 2) = "string": camps(i, 3) = "Texte Impresio": i = i + 1
    camps(i, 1) = "numeroliniaimpresio": camps(i, 2) = "string": camps(i, 3) = "NºLinia Imp": i = i + 1
    camps(i, 1) = "Estatclixes": camps(i, 2) = "string(10)": camps(i, 3) = "EstatClixes": i = i + 1
    camps(i, 1) = "clixes": camps(i, 2) = "string(10)": camps(i, 3) = "Clixes": i = i + 1
    camps(i, 1) = "material": camps(i, 2) = "string(10)": camps(i, 3) = "Material": i = i + 1
    camps(i, 1) = "materialPC": camps(i, 2) = "string(10)": camps(i, 3) = "MaterialPC": i = i + 1
    camps(i, 1) = "materialPC2": camps(i, 2) = "string(10)": camps(i, 3) = "MaterialPC2": i = i + 1
    camps(i, 1) = "mts": camps(i, 2) = "double": camps(i, 3) = "Metres": i = i + 1
    camps(i, 1) = "micromacro": camps(i, 2) = "string(5)": camps(i, 3) = "MIcro/MAcro": i = i + 1
    camps(i, 1) = "Muntat": camps(i, 2) = "string(1)": camps(i, 3) = "Muntat?": i = i + 1
    camps(i, 1) = "tintesrevisades": camps(i, 2) = "string": camps(i, 3) = "Tintes revisades": i = i + 1
    camps(i, 1) = "data1": camps(i, 2) = "date": camps(i, 3) = "Data1": i = i + 1
    camps(i, 1) = "data2": camps(i, 2) = "date": camps(i, 3) = "Data2": i = i + 1
    camps(i, 1) = "clientvindraarevisarimpresio": camps(i, 2) = "string": camps(i, 3) = "Pack/Cli.Vindrà": i = i + 1
    camps(i, 1) = "datacalloff": camps(i, 2) = "date": camps(i, 3) = "DataCalloff": i = i + 1
    camps(i, 1) = "dataoperari": camps(i, 2) = "date": camps(i, 3) = "DataOpReb.": i = i + 1
    camps(i, 1) = "importancia": camps(i, 2) = "byte": camps(i, 3) = "Impor.": i = i + 1
    camps(i, 1) = "standbyimpresio": camps(i, 2) = "string(2)": camps(i, 3) = "StandBy": i = i + 1
    camps(i, 1) = "impresora": camps(i, 2) = "string(10)": camps(i, 3) = "Imp.": i = i + 1
    camps(i, 1) = "laminadora": camps(i, 2) = "string(10)": camps(i, 3) = "Lam.": i = i + 1
    camps(i, 1) = "rebobinadora": camps(i, 2) = "string(10)": camps(i, 3) = "Reb.": i = i + 1
    camps(i, 1) = "soldadora": camps(i, 2) = "string(10)": camps(i, 3) = "Sol.": i = i + 1
    camps(i, 1) = "observacions": camps(i, 2) = "string": camps(i, 3) = "Observacio": i = i + 1
    camps(i, 1) = "dataexpedicio": camps(i, 2) = "date": camps(i, 3) = "Data_Expedició": camps(i, 4) = "S": i = i + 1
    camps(i, 1) = "observacioexpedicio": camps(i, 2) = "string": camps(i, 3) = "Obs_Expedició": camps(i, 4) = "S": i = i + 1
    camps(i, 1) = "refinplacsa": camps(i, 2) = "string": camps(i, 3) = "Ref_inplacsa": camps(i, 4) = "N": i = i + 1
  End If
  
  If seccio = "E" Then
    camps(i, 1) = "datacomanda": camps(i, 2) = "date": camps(i, 3) = "Data_comanda": i = i + 1
    camps(i, 1) = "dataalbara": camps(i, 2) = "date": camps(i, 3) = "Data_albarà": i = i + 1
    camps(i, 1) = "albara": camps(i, 2) = "double": camps(i, 3) = "NºAlbarà": i = i + 1
    camps(i, 1) = "comanda": camps(i, 2) = "long": camps(i, 3) = "NºLot": i = i + 1
    camps(i, 1) = "numtreball": camps(i, 2) = "long": camps(i, 3) = "NºTreball": i = i + 1
    camps(i, 1) = "nomclient": camps(i, 2) = "string": camps(i, 3) = "Nom Client": i = i + 1
    camps(i, 1) = "impresio": camps(i, 2) = "string(1)": camps(i, 3) = "Impresio": i = i + 1
    camps(i, 1) = "entregaToP": camps(i, 2) = "string(2)": camps(i, 3) = "Entrega_ToP": i = i + 1
    camps(i, 1) = "quantitatTeorica": camps(i, 2) = "double": camps(i, 3) = "Quant.Teòrica": i = i + 1
    camps(i, 1) = "quantitatEntregada": camps(i, 2) = "double": camps(i, 3) = "Quant.Entregada": i = i + 1
    camps(i, 1) = "tanx100kgvs": camps(i, 2) = "double": camps(i, 3) = "%Kg_Desv": i = i + 1
    camps(i, 1) = "preu": camps(i, 2) = "double": camps(i, 3) = "Preu_Albarà": i = i + 1
    camps(i, 1) = "tipusunitat": camps(i, 2) = "string(10)": camps(i, 3) = "Unitat": camps(i, 4) = "N": i = i + 1
    camps(i, 1) = "kgentregats": camps(i, 2) = "double": camps(i, 3) = "Kg_Entregats": i = i + 1
    camps(i, 1) = "kgimpost": camps(i, 2) = "double": camps(i, 3) = "Kg_Impost": i = i + 1
    camps(i, 1) = "tanx100impostvs": camps(i, 2) = "double": camps(i, 3) = "%Impost_Desv": i = i + 1
    camps(i, 1) = "eurokg": camps(i, 2) = "double": camps(i, 3) = "€/Kg_Calculat": i = i + 1
    camps(i, 1) = "pvprevisat": camps(i, 2) = "string(1)": camps(i, 3) = "PVP_Revisat": i = i + 1
    camps(i, 1) = "extracost": camps(i, 2) = "double": camps(i, 3) = "Extra_Cost": i = i + 1
    camps(i, 1) = "preuclixes": camps(i, 2) = "double": camps(i, 3) = "Preu_Clixes": i = i + 1
    camps(i, 1) = "facturat": camps(i, 2) = "string(1)": camps(i, 3) = "Facturat": i = i + 1
    camps(i, 1) = "revisat": camps(i, 2) = "string(1)": camps(i, 3) = "Revisat": i = i + 1
    camps(i, 1) = "okclixes": camps(i, 2) = "string(1)": camps(i, 3) = "Ok_Clixes": i = i + 1
    camps(i, 1) = "observacio": camps(i, 2) = "string": camps(i, 3) = "Observació_Entrega":  i = i + 1
  End If

  
End Sub
Sub crearfitxertemp(Optional obrint As Boolean)
     
  '   If Not existeix("c:\ordprog.ini") Then
  '      If existeix(fitxertemp) And obrint Then Kill fitxertemp
  '   End If
    If Not existeix(fitxertemp) Then
       crearfitxertemporal
    End If
   Set dbconsulta = DBEngine.OpenDatabase(fitxertemp)
  'creant taula TOTES
   taulaplanificacio = "planificaciototes"
   carregarllistadecampstemporals "T"
   creartaula
   taulaplanificacio = "llistatplanificaciototes"
   creartaula
   
  'creant taula RECLAMACIONS
   taulaplanificacio = "reclamacionscomandes"
   carregarllistadecampstemporals "C"
   creartaula
   taulaplanificacio = "llistatreclamacionscomandes"
   creartaula
   
  'creant taula LAM
   taulaplanificacio = "planificaciolam"
   carregarllistadecampstemporals "L"
   creartaula
   taulaplanificacio = "llistatplanificaciolam"
   creartaula
  
  'creant taula REB
   taulaplanificacio = "planificacioreb"
   carregarllistadecampstemporals "R"
   creartaula
   taulaplanificacio = "llistatplanificacioreb"
   creartaula
   On Error Resume Next
   dbconsulta.Execute "alter table llistatplanificacioreb add column materialPC string"
   dbconsulta.Execute "alter table llistatplanificacioreb add column materialPC2 string"
   On Error GoTo 0

  
  'creant taula SOL
   taulaplanificacio = "planificaciosol"
   carregarllistadecampstemporals "S"
   creartaula
   taulaplanificacio = "llistatplanificaciosol"
   creartaula
  
  'creant taula IMP
   taulaplanificacio = "planificacioimp"
   carregarllistadecampstemporals "I"
   creartaula
   taulaplanificacio = "llistatplanificacioimp"
   creartaula
   
  'creant taula ENT (Entregues)
   taulaplanificacio = "planificacioent"
   carregarllistadecampstemporals "E"
   creartaula
   taulaplanificacio = "llistatplanificacioent"
   creartaula
  
  taulaplanificacio = "planificacioimp"
    Set dbconsulta = OpenDatabase(fitxertemp)
    SetAllowZeroLength dbconsulta
End Sub

Sub imprimirllistatestat(Optional vnomllistat As String)
  
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  If vnomllistat = "" Then vnomllistat = "llistatcomandesseccions.rpt"
  
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + vnomllistat, 1)
  oreport.Database.Tables.Item(1).Location = fitxertemp

  
  oreport.DiscardSavedData
  
 ' If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1, Me
  '  Else
 '     oreport.PrintOut False, 1
'  End If
fi:
  planificacio.SetFocus

End Sub
Sub possarregistres(vllistat As String)
  Dim rstt As Recordset
  
  If vllistat = "2-Impresores per LAMINADORES (Ordre impressió)" Then
    Me.caption = "Seleccionant 1-1"
    Set rstt = dbbaixes.OpenRecordset("SELECT comandes.*, comandes.proximaseccio as comandes_proximaseccio, impresores_ordreimpresio.ordre FROM impresores_ordreimpresio LEFT JOIN comandes ON impresores_ordreimpresio.comanda = comandes.comanda ORDER BY impresores_ordreimpresio.ordre;", dbOpenSnapshot, dbReadOnly)
    passarelsregistresaltemporal rstt, vllistat
  End If
  If vllistat = "2-Impresores (Ordre impressió)" Then
    Me.caption = "Seleccionant 1"
    'Set rstt = dbbaixes.OpenRecordset("SELECT comandes.*, comandes.proximaseccio as comandes_proximaseccio,muntadoratot.acabada, comandes.proximaseccio FROM muntadoratot INNER JOIN comandes ON muntadoratot.comanda = comandes.comanda WHERE (((muntadoratot.acabada)=True) AND ((comandes.proximaseccio)='I'));", dbOpenSnapshot, dbReadOnly)
    Set rstt = dbbaixes.OpenRecordset("SELECT comandes.*, comandes.proximaseccio as comandes_proximaseccio, impresores_ordreimpresio.ordre FROM impresores_ordreimpresio LEFT JOIN comandes ON impresores_ordreimpresio.comanda = comandes.comanda ORDER BY impresores_ordreimpresio.ordre;", dbOpenSnapshot, dbReadOnly)
    passarelsregistresaltemporal rstt, vllistat
  End If
  
  If vllistat = "3-Laminadores" Then
    Me.caption = "Seleccionant 2"
    Set rstt = dbcomandes.OpenRecordset("select * ,comandes.proximaseccio as comandes_proximaseccio from comandes where proximaseccio='L'", dbOpenSnapshot, dbReadOnly)
    passarelsregistresaltemporal rstt, vllistat
  End If
  
  If vllistat = "1-Muntadores (Ordre pendent de muntar)" Then
    Me.caption = "Seleccionant 3"
    Set rstt = dbbaixes.OpenRecordset("SELECT comandes.*, comandes.proximaseccio AS comandes_proximaseccio, muntadora_ordremuntatge.muntada FROM comandes RIGHT JOIN muntadora_ordremuntatge ON comandes.comanda = muntadora_ordremuntatge.comanda Where (((muntadora_ordremuntatge.muntada) = False))ORDER BY muntadora_ordremuntatge.ordre;", dbOpenSnapshot, dbReadOnly)
    passarelsregistresaltemporal rstt, vllistat
  End If
  
  Set rstt = Nothing
End Sub
Sub passarelsregistresaltemporal(rstt As Recordset, tipuscomanda As String)
  Dim rstnc As Recordset
  Dim rstc As Recordset
  Dim rstp As Recordset
  Dim nomclient As String
  Set rstc = dbconsulta.OpenRecordset("llistatestat")
  With rstc
  While Not rstt.EOF
     nomclient = ""
     Set rstc = dbcomandes.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(rstt!client)), dbOpenSnapshot, dbReadOnly)
     Set rstp = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(cadbl(rstt!comanda)))
'     MsgBox dbplanificacio.Name
     If Not rstc.EOF Then nomclient = rstc!nom
     If cadbl(rstt!comanda) > 0 Then
        .AddNew
        !comanda = atrim(cadbl(rstt!comanda))
        !comandavisual = atrim(cadbl(rstt!comanda))
        !producte = atrim(rstt!producte)
        !client = atrim(rstt!client) + " - " + nomclient
        !Texte = atrim(rstt!marcailinia)
        !estat = posicioenlaruta(cadbl(rstt!comanda)) ' + IIf(rstt![comandes_proximaseccio] <> posicioenlaruta(rstt!comanda), Chr(255), "") 'atrim(rstt![comandes_proximaseccio])
        !detall = estatdelacompra(rstt!comanda)
        !amplelam = ampleminimpackinglist(rstt!comanda, rstt!linkcomanda1, rstt!linkcomanda2)
        If !amplelam = 0 Then !amplelam = cadbl(rstt!ampleesq)
        !bandes = cadbl(rstt!simulteneitatreb)
        !camisa = cadbl(rstt!camisa)
        If Not rstp.EOF Then
           !importancia = rstp!importancia
           !dataentrega = format(IIf(atrim(rstp!dataexpedicio) = "", atrim(rstp!data2), atrim(rstp!dataexpedicio)), "dd/mm/yy")
        End If
        !metres = cadbl(rstt!cantitatex)
        !tipuscomanda = tipuscomanda
        .Update
        'If atrim(rstt![comandes_proximaseccio]) = "I" Or atrim(rstt![comandes_proximaseccio]) = "E" Then
          If cadbl(rstt!linkcomanda1) > 0 Then
            .AddNew
            !comanda = rstt!linkcomanda1
            !comandavisual = "    " + atrim(rstt!linkcomanda1)
            !detall = estatdelacompra(rstt!linkcomanda1)
            !tipuscomanda = tipuscomanda
            .Update
          End If
        
          If cadbl(rstt!linkcomanda2) > 0 Then
            .AddNew
            !comanda = rstt!linkcomanda2
            !comandavisual = "    " + atrim(rstt!linkcomanda2)
            !detall = estatdelacompra(rstt!linkcomanda2)
            !tipuscomanda = tipuscomanda
            .Update
          End If
        'End If
     End If
     rstt.MoveNext
  Wend
  End With
  Set rstp = Nothing
  Set rstc = Nothing
  Set rstnc = Nothing
End Sub
Function ampleminimpackinglist(vnumc As Double, vnumc1 As Double, vnumc2 As Double) As Double
  Dim rst As Recordset
  
  Set rst = dbstocks.OpenRecordset("SELECT Parcials.comanda, Min(Palets.Ample) AS MinimDeAmple FROM Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet Where (((Parcials.orcomassignacio) <> '500')) GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + atrim(vnumc) + "'));")
  If Not rst.EOF Then ampleminimpackinglist = rst!MinimDeAmple
  
  If vnumc1 > 0 Then
    Set rst = dbstocks.OpenRecordset("SELECT Parcials.comanda, Min(Palets.Ample) AS MinimDeAmple FROM Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet Where (((Parcials.orcomassignacio) <> '500')) GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + atrim(vnumc1) + "'));")
    If Not rst.EOF Then ampleminimpackinglist = IIf(rst!MinimDeAmple < ampleminimpackinglist, rst!MinimDeAmple, ampleminimpackinglist)
  End If
  If vnumc2 > 0 Then
    Set rst = dbstocks.OpenRecordset("SELECT Parcials.comanda, Min(Palets.Ample) AS MinimDeAmple FROM Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet Where (((Parcials.orcomassignacio) <> '500')) GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + atrim(vnumc2) + "'));")
    If Not rst.EOF Then ampleminimpackinglist = IIf(rst!MinimDeAmple < ampleminimpackinglist, rst!MinimDeAmple, ampleminimpackinglist)
  End If
  
  Set rst = Nothing
End Function
Function estatdelacompra(numc As Double) As String
   Dim rstc As Recordset
   Set rstc = dbstocks.OpenRecordset("select * from parcials where comanda='" + atrim(numc) + "'", dbOpenSnapshot, dbReadOnly)
   If Not rstc.EOF Then
      estatdelacompra = "Packing-List"
     Else
        Set rstc = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(numc), dbOpenSnapshot, dbReadOnly)
        If Not rstc.EOF Then If rstc!assignarstock Then estatdelacompra = "Estoc"
   End If
   If estatdelacompra = "" Then
       Set rstc = dbstocks.OpenRecordset("select * from percomandaoclient where numcomanda=" + atrim(numc), dbOpenSnapshot, dbReadOnly)
       If Not rstc.EOF Then estatdelacompra = "Reservat"
   End If
   If estatdelacompra = "" Then
    Set rstc = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, capcalera.nomprov,capcalera.dataentrega, comandesxlinia.numcomanda FROM (capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra) RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra WHERE (((comandesxlinia.numcomanda)=" + atrim(numc) + "));", dbOpenSnapshot, dbReadOnly)
    If Not rstc.EOF Then
       estatdelacompra = "Compra: " + atrim(rstc![capcalera.numcomanda]) + " Entrega: " + atrim(rstc!dataentrega) + "  " + atrim(rstc!nomprov)
    End If
   End If
   
   Set rstc = Nothing
End Function
Sub creartaulallistatestat()
  Dim i As Integer
  Dim campsestat(100, 2) As String
  i = 1
  campsestat(i, 1) = "comanda": campsestat(i, 2) = "double": i = i + 1
  campsestat(i, 1) = "comandavisual": campsestat(i, 2) = "string": i = i + 1
  campsestat(i, 1) = "producte": campsestat(i, 2) = "string": i = i + 1
  campsestat(i, 1) = "client": campsestat(i, 2) = "string": i = i + 1
  campsestat(i, 1) = "texte": campsestat(i, 2) = "string": i = i + 1
  campsestat(i, 1) = "estat": campsestat(i, 2) = "string": i = i + 1
  campsestat(i, 1) = "detall": campsestat(i, 2) = "string": i = i + 1
  campsestat(i, 1) = "tipuscomanda": campsestat(i, 2) = "string": i = i + 1
  campsestat(i, 1) = "metres": campsestat(i, 2) = "double": i = i + 1
  campsestat(i, 1) = "dataentrega": campsestat(i, 2) = "string": i = i + 1
  campsestat(i, 1) = "importancia": campsestat(i, 2) = "byte": i = i + 1
  campsestat(i, 1) = "camisa": campsestat(i, 2) = "double": i = i + 1
  campsestat(i, 1) = "amplelam": campsestat(i, 2) = "double": i = i + 1
  campsestat(i, 1) = "bandes": campsestat(i, 2) = "double": i = i + 1
  On Error GoTo jaexisteix
 ' dbconsulta.Execute "drop table llistatestat"
  dbconsulta.Execute ("create table llistatestat (id counter)")
  On Error GoTo 0


  For i = 1 To 100
    If campsestat(i, 1) <> "" Then
       dbconsulta.Execute ("alter table llistatestat add column " + campsestat(i, 1) + " " + campsestat(i, 2))
      ' camps(i, 1) = ""
       
        Else: i = 1000
    End If
    
  Next i
  SetAllowZeroLength dbconsulta
jaexisteix:
  On Error Resume Next
  dbconsulta.Execute "delete * from llistatestat"
  wait 1
End Sub

Private Sub mcomananoreal_Click()
   Dim resp As String
   Dim rst As Recordset
   Dim numc As Double
   Dim ordre As Double
   Set rst = dbplanificacio.OpenRecordset("select max(comanda) as gran from planificaciototes where comanda<10000", dbOpenSnapshot, dbReadOnly)
   If Not rst.EOF Then
        If cadbl(rst!gran) = 0 Then
            numc = 9001
           Else
             numc = cadbl(rst!gran) + 1
        End If
      Else: numc = 9001
   End If
   resp = InputBox("Entra una descripcio per la COMANDA NO REAL que vols crear.", "Comanda NO REAL")
   If resp <> "" Then
      If taulaplanificacio <> "planificaciototes" Then
        ordre = cadbl(InputBox("Vols col.locar-lo en algun ordre concret? escriu-lo, sino serà 999", "Atenció", 999))
        If ordre > 999 Or ordre < 1 Then
           ordre = 999
          Else: ordre = ordre - 0.1
        End If
      End If
      If taulaplanificacio <> "planificaciototes" Then
       dbconsulta.Execute "insert into " + taulaplanificacio + " (ordre,comanda,maquina,observacions,tempsimpresio) values (" + passaradecimalpunt(atrim(ordre)) + "," + atrim(numc) + "," + atrim(nummaquina) + ",'" + resp + "',180)"
       dbplanificacio.Execute "insert into " + taulaplanificacio + " (ordre,comanda,maquina) values (" + passaradecimalpunt(atrim(ordre)) + "," + atrim(numc) + "," + atrim(nummaquina) + ")"
      End If
      dbconsulta.Execute "insert into planificaciototes (comanda,observacions) values (" + atrim(numc) + ",'" + treure_apostruf(resp) + "')"
      dbplanificacio.Execute "insert into planificaciototes (comanda,observacio) values (" + atrim(numc) + ",'" + treure_apostruf(resp) + "')"
      reordenarregistres
      filtre_LostFocus 999
   End If
   Set rst = Nothing
End Sub

Private Sub menviat_Click()
  actualitzar_enviats_expedicio
   Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "planificacio.mdb"
  formseleccio.Data1.RecordSource = "select data as [Data_Expedició],observaciogeneral as [Observació] from Expedicions where enviat"
  formseleccio.refrescar
  formseleccio.width = 13000
  formseleccio.DBGrid2.Columns(0).width = 2000
  formseleccio.DBGrid2.Columns(1).width = 9500
  formseleccio.DBGrid2.width = formseleccio.width - 3500
  formseleccio.Left = (Screen.width / 2) - (formseleccio.width / 2)
  
  If formseleccio.Data1.Recordset.EOF Then MsgBox "No hi ha Enviaments fets.": Exit Sub
  formseleccio.Data1.Recordset.MoveLast
  formseleccio.Data1.Recordset.MoveFirst
  formseleccio.Show 1
  If seleccioret = 1 Then
     ensenyar_expedicionsdia formseleccio.Data1.Recordset.Fields("Data_Expedició")
  End If
  Unload formseleccio
End Sub

Private Sub mgeneral_Click()
  Set dbsap = Nothing
   eseccio = "General - Totes les seccions"
   Frameentregues.visible = False
  carregarmaquines "T"
  taulaplanificacio = "planificaciototes"
  carregarllistadecampstemporals "T"
  carregar_ordre_correcte 0
  
  ratoli "espera"
  reixa.visible = False
  ordrereixa = " order by data1"
  Command3.tag = ""
   configreixa
   'reordenarregistres
   poblarlareixa nummaquina
   ratoli "normal"
   reixa.visible = True
   buscarcomanda cadbl(mgeneral.tag)
End Sub

Private Sub mhoraris_Click()
   horarismaquines.Show 1
End Sub
Sub crearfitxertemporal()
    borrartemps
    'fitxertemp = "c:\temp\comprestmp.mdb"
    '"c:\temp\~compres" + Format(Now, "ddmmhhnnss") + ".mdb"
    If Not existeix(fitxertemp) Then
       DBEngine.CreateDatabase fitxertemp, dbLangGeneral
    End If
    Set dbconsulta = DBEngine.OpenDatabase(fitxertemp)
    
End Sub

Sub borrartemps()
   On Error GoTo fi
   'Kill "c:\temp\~compres*.*"
   If existeix("c:\temp\planificaciotmp.mdb") Then Kill "c:\temp\planificaciotmp.mdb"
   Exit Sub
fi:
 msgboxEx "No es pot borrar el fitxer temporal." + Chr(13) + "Mira que no hi hagi una altra planificació oberta", 0, "Error", 5
 End
End Sub
Sub creartaula()
  Dim i As Integer
  On Error GoTo jaexisteix
  dbconsulta.Execute ("create table " + taulaplanificacio + " (id counter)")
  On Error GoTo 0
  dbconsulta.Execute "CREATE INDEX principal ON " + taulaplanificacio + " ([id]) witH PRIMARY;"


  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbconsulta.Execute ("alter table " + taulaplanificacio + " add column " + camps(i, 1) + " " + camps(i, 2))
      ' camps(i, 1) = ""
       
        Else: i = 1000
    End If
    
  Next i
    dbconsulta.Execute "CREATE INDEX icomanda ON " + taulaplanificacio + " ([comanda]);"
  
  SetAllowZeroLength dbconsulta
jaexisteix:
End Sub
Function SetAllowZeroLength(db As Database)
    Dim i As Integer, j As Integer
    Dim td As TableDef, fld As Field

    
    'The following line prevents the code from stopping if you do not
    'have permissions to modify particular tables, such as system
    'tables.
    On Error Resume Next
    For i = 0 To db.TableDefs.Count - 1
       Set td = db(i)
       For j = 0 To td.Fields.Count - 1
          Set fld = td(j)
          If (fld.Type = 10) And Not _
            fld.AllowZeroLength Then
             fld.AllowZeroLength = True
          End If
       Next j
    Next i
    
End Function

Private Sub mimpperlam_Click()
   Me.caption = " Creant la taula temporal"
   creartaulallistatestat
   Me.caption = " Posant els registres"
   possarregistres "2-Impresores per LAMINADORES (Ordre impressió)"
   Me.caption = "Imprimint el resultat."
   imprimirllistatestat "llistatcomandesseccions_ImpXrLam.rpt"
   Me.caption = "Planificació"
End Sub

Private Sub mimpresores_Click()
  Dim comanda As Double
  Set dbsap = Nothing
  Set dbtmp = dbcomandes
  Set dbtmpb = dbbaixes
  Frameentregues.visible = False
  ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  nummaquina = cadbl(mimpresores.tag)
  comanda = cadbl(Mid(mgeneral.tag, 1, IIf(InStr(1, mgeneral.tag, "R"), InStr(1, mgeneral.tag, "R") - 1, 10)))
  'If taulaplanificacio = "planificaciototes" Then
  '  nummaquina = cadbl(Mid(reixa.TextMatrix(reixa.Row, numcol("Imp.")), 1, 1))
  '  comanda = cadbl(reixa.TextMatrix(reixa.Row, numcol("NºLot")))
  'End If
  
  eseccio = "Impresores"
  carregarmaquines "I"
  taulaplanificacio = "planificacioimp"
  carregarllistadecampstemporals "I"
  If comanda > 0 Then
     carregarcomandaescullida cadbl(nummaquina), comanda
      Else: recarrearmaquinaseleccionada 7
  End If
  
End Sub
Sub carregarcomandaescullida(nummaquina As Integer, comanda As Double)
   recarrearmaquinaseleccionada nummaquina
   buscarcomanda comanda
End Sub

Private Sub mimprimirveurecomanda_Click()
  Dim numc As String
  numc = InputBox("Entra la comanda que vols visualitzar.", "Visualitzar/Imprimir comanda")
  If cadbl(numc) = 0 Then Exit Sub
  escriure_ini "Baixes", "imprimircomanda", cadbl(numc), "comandes.ini"
  Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", "comandes.ini")) + "comandes.exe - imprimir", vbHide
  missatgevist.Show 1
End Sub

Private Sub mlaminadores_Click()
 Dim comanda As Double
 Set dbsap = Nothing
 Frameentregues.visible = False
 nummaquina = cadbl(mlaminadores.tag)
 comanda = cadbl(mgeneral.tag)
  'If taulaplanificacio = "planificaciototes" Then
  '  nummaquina = cadbl(Mid(reixa.TextMatrix(reixa.Row, numcol("Lam.")), 1, 1))
  '  comanda = cadbl(reixa.TextMatrix(reixa.Row, numcol("NºLot")))
  'End If

eseccio = "Laminadores"
carregarmaquines "L"
taulaplanificacio = "planificaciolam"
carregarllistadecampstemporals "L"
carregarcomandaescullida cadbl(nummaquina), comanda

End Sub

Private Sub MouseWheel1_WheelMove(bDown As Boolean)
  Dim v As Byte
  v = 3
  If reixa.Rows < 2 Then Exit Sub
  If bDown Then
     If reixa.TopRow + v < reixa.Rows Then
        reixa.TopRow = reixa.TopRow + v
       Else: reixa.TopRow = reixa.Rows - 1
     End If
    Else:
        If reixa.TopRow - v > 1 Then
           reixa.TopRow = reixa.TopRow - v
          Else: reixa.TopRow = 1
        End If
  End If
  postit.visible = False
End Sub
Function justificar(v As String, longitut As Integer, DoE As String) As String
    v = Mid(v, 1, longitut)
    If DoE = "E" Then
       v = v + Space(longitut - Len(v))
      Else: v = Space(longitut - Len(v)) + v
    End If
    justificar = v
End Function
Function borrarvfitxer(vfitxer As String) As Boolean
  On Error GoTo etErrors
  If existeix(vfitxer) Then Kill vfitxer
  borrarvfitxer = True
  Exit Function
  borrarvfitxer = False
  Exit Function
etErrors:
  MsgBox "Error al crear el fitxer temporal, assegura que no el tinguis obert.", vbCritical, "Error"
  
End Function

Private Sub mllimpresores_Click()
  Me.caption = " Creant la taula temporal"
   creartaulallistatestat
   Me.caption = " Posant els registres"
   possarregistres "2-Impresores (Ordre impressió)"
   Me.caption = "Imprimint el resultat."
   imprimirllistatestat
   Me.caption = "Planificació"
End Sub

Private Sub mllistathoradeprogramacio_Click()
   Dim vseccio As String
   Dim vlletraseccio As String
   Dim rstp As Recordset
   Dim vfitxer As String
   vfitxer = "c:\temp\llistathoresdeprogramacio.txt"
   If Not borrarvfitxer(vfitxer) Then Exit Sub
   vseccio = InputBox("Entra la secció que vols veure (I,L,R,S):", "Llista d'hora de programació", "L")
   vlletraseccio = vseccio
   Select Case vseccio
       Case "I"
         vseccio = "planificacioimp"
       Case "L"
         vseccio = "planificaciolam"
       Case "S"
         vseccio = "planificaciosol"
       Case "R"
         vseccio = "planificacioreb"
       Case Else
          Exit Sub
   End Select
   Set rstp = dbplanificacio.OpenRecordset("Select * from llistadeplanificacions where seccio='" + atrim(vseccio) + "' order by dataordreassignat Desc")
  ' Clipboard.Clear
  ' Clipboard.SetText "Select * from llistadeplanificacions where seccio='" + atrim(vseccio) + "' order by dataordreassignat Desc"
   If Not rstp.EOF Then
         Open vfitxer For Output As #3
         Print #3, "     "
         Print #3, "   Llistat hores de programació de la secció ( " + vlletraseccio + " )"
         Print #3, "     "
         Print #3, "     "
         
        While Not rstp.EOF
           vlinia = justificar("Lot: " + atrim(rstp!comanda), 15, "E") + justificar("Maq: " + atrim(rstp!maquina), 8, "E") + justificar("Ordre: " + atrim(rstp!ordre), 15, "E") + justificar("H: " + format(rstp!dataordreassignat, "dd/mm hh:nn"), 20, "E")
           Print #3, vlinia
           rstp.MoveNext
        Wend
       Close #3
       Shell "notepad.exe " + vfitxer, vbMaximizedFocus
         Else: MsgBox "No hi ha res a llistar", vbInformation, "Atenció"
   End If
   
fi:
   Set rstp = Nothing
   
End Sub

Private Sub mllistattotes_Click()
   Me.caption = " Creant la taula temporal"
   creartaulallistatestat
   Me.caption = " Posant els registres"
   possarregistres "1-Muntadores (Ordre pendent de muntar)"
   Me.caption = " Posant els registres"
   possarregistres "2-Impresores (Ordre impressió)"
   Me.caption = " Posant els registres"
   possarregistres "3-Laminadores"
   Me.caption = "Imprimint el resultat."
   imprimirllistatestat
   Me.caption = "Planificació"
End Sub

Private Sub mlllaminadores_Click()
Me.caption = " Creant la taula temporal"
   creartaulallistatestat
   Me.caption = " Posant els registres"
   possarregistres "3-Laminadores"
   Me.caption = "Imprimint el resultat."
   imprimirllistatestat
   Me.caption = "Planificació"
End Sub

Private Sub mllmuntadores_Click()
Me.caption = " Creant la taula temporal"
   creartaulallistatestat
   Me.caption = " Posant els registres"
   possarregistres "1-Muntadores (Ordre pendent de muntar)"
   Me.caption = "Imprimint el resultat."
   imprimirllistatestat
   Me.caption = "Planificació"
End Sub

Private Sub mnoenviat_Click()
  actualitzar_enviats_expedicio
   Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "planificacio.mdb"
  formseleccio.Data1.RecordSource = "select data as [Data_Expedició],observaciogeneral as [Observació] from Expedicions where not enviat"
  formseleccio.refrescar
  formseleccio.width = 13000
  formseleccio.DBGrid2.Columns(0).width = 2000
  formseleccio.DBGrid2.Columns(1).width = 9500
  formseleccio.DBGrid2.width = formseleccio.width - 3500
  formseleccio.Left = (Screen.width / 2) - (formseleccio.width / 2)
  
  If formseleccio.Data1.Recordset.EOF Then MsgBox "No hi ha Enviaments pendents": Exit Sub
  formseleccio.Data1.Recordset.MoveLast
  formseleccio.Data1.Recordset.MoveFirst
  formseleccio.Show 1
  If seleccioret = 1 Then
      ensenyar_expedicionsdia formseleccio.Data1.Recordset.Fields("Data_Expedició")
  End If
  Unload formseleccio
 
End Sub
Sub ensenyar_expedicionsdia(vdia As String)
    imprimir_llistat_expedicionspendents vdia
End Sub
Sub actualitzar_enviats_expedicio(Optional venviat As Boolean)
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim rstc As Recordset
   'elimino totes les linies_expedicons on pot haver una comanda a expedicions que hagi sigut eliminada
   dbplanificacio.Execute "DELETE  linies_expedicions.* FROM linies_expedicions LEFT JOIN comandes ON linies_expedicions.comanda = comandes.comanda WHERE (((comandes.comanda) Is Null));"
   

   Set rst = dbplanificacio.OpenRecordset("select * from expedicions where datediff('d',now,data)>-30")
   While Not rst.EOF
      Set rst2 = dbplanificacio.OpenRecordset("select * from linies_expedicions where data=#" + format(rst!data, "mm/dd/yy") + "# order by enviat desc ")
      If rst2.EOF Then GoTo proxima
      rst.Edit
      If Not rst2!enviat Then
          rst!enviat = False
           Else: rst!enviat = True
      End If
      venviat = rst!enviat
      rst.Update
proxima:
      rst.MoveNext
   Wend
   dbplanificacio.Execute "UPDATE linies_expedicions LEFT JOIN planificaciototes ON linies_expedicions.comanda = planificaciototes.comanda SET linies_expedicions.parcial = [planificaciototes].[entregaparcial] WHERE (((linies_expedicions.enviat)=False));"

   Set rst = Nothing
   Set rst2 = Nothing
   wait 1
End Sub

Private Sub mpujaraexpedicions_Click()
   Dim vdia As String
   Dim rst As Recordset
   Dim vobs As String
   Dim venviat As Boolean
   venviat = False
   treure_comandesdataexpedicioerronees
   vdia = InputBox("Entra el dia que vols passar a expedicions.", "Passar a Expedició", format(proximdianatural, "dd/mm/yy"))
   If Not IsDate(vdia) Then MsgBox "Dia no vàlid.", vbCritical, "Error": Exit Sub
   If DateDiff("d", Now, vdia) < 0 Then If MsgBox("Aquesta data ja està passada es correcte?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   Set rst = dbplanificacio.OpenRecordset("select * from Expedicions where data=#" + format(vdia, "mm/dd/yy") + "#")
   If Not rst.EOF Then vobs = atrim(rst!observaciogeneral)
   vobs = InputBox("Escriu la observació que vols possar per expedicions de tots els enviaments d'aquest dia.", "Observacions generals", vobs)
   If StrPtr(vobs) = 0 Then Exit Sub
   If Not rst.EOF Then
      venviat = rst!enviat
      rst.Delete
   End If
   rst.AddNew
   rst!data = vdia
   rst!observaciogeneral = vobs
   rst!enviat = venviat
   rst.Update
   Set rst = Nothing
   refrescarnumerosdalbara
   actualitzar_enviats_expedicio venviat
   wait 2
   imprimir_llistat_expedicionspendents vdia
   Set rst = Nothing
   If MsgBox("Vols enviar per E-Mail aquest llistat a Expedicions?", vbInformation + vbDefaultButton2 + vbYesNo, "Enviar llistat") = vbNo Then Exit Sub
   enviaremailgenericambadjunt "EnviamentLlistatExpedicionsPlanificacio", "Llistat expedicio pel dia " + atrim(vdia), "Llistat Expedició.", "c:\temp\llistatexpedicions.pdf"
End Sub
Sub refrescarnumerosdalbara()
  Dim rstc As Recordset
  Dim vcont As Long
  Set rstc = dbplanificacio.OpenRecordset("SELECT expedicions.enviat as enviatexp,linies_expedicions.comanda, linies_expedicions.enviat as enviatlinies,linies_expedicions.comanda as [Comanda], linies_expedicions.data as [Data_Ex],linies_expedicions.albara as [Albarà], linies_expedicions.observacio as [Obs_Exp], Expedicions.observaciogeneral as [Obs_General] FROM Expedicions RIGHT JOIN linies_expedicions ON Expedicions.data = linies_expedicions.data where expedicions.enviat=false order by linies_expedicions.data desc;")
  While Not rstc.EOF And vcont < 200
     'If rstc!comanda = 202108 Then Stop
     Set rst = dbplanificacio.OpenRecordset("SELECT capcaleraalbara.numalbara, capcaleraalbara.dataalbara, liniesalbara.lotinplacsa FROM capcaleraalbara INNER JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara where liniesalbara.lotinplacsa=" + atrim(rstc!comanda) + " order by capcaleraalbara.dataalbara Desc")
     If Not rst.EOF Then
        If rst!dataalbara = rstc!data_Ex Then
          If cadbl(rstc![Albarà]) <> cadbl(rst!numalbara) Then
            rstc.Edit
            rstc![Albarà] = rst!numalbara
            rstc.Update
          End If
        End If
     End If
     If (cadbl(rstc![Albarà]) = 0 And rstc!enviatlinies) Or rst.EOF Then
         rstc.Edit
         rstc!enviatlinies = False
         rstc.Update
         dbplanificacio.Execute "update expedicions set enviat=true where data=#" + format(rstc!data_Ex, "mm/dd/yy") + "#"
     End If
     rstc.MoveNext
     vcont = vcont + 1
  Wend
  Set rstc = Nothing
End Sub
Sub imprimir_llistat_expedicionspendents(vdia As String)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatexpedicionspendents.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "planificacio.mdb"
  oreport.Database.Tables.Item(2).Location = rutadelfitxer(cami) + "VENDES.mdb"
  
    

  
  oreport.DiscardSavedData
  
  
  oreport.RecordSelectionFormula = "{expedicions.data}=#" + format(vdia, "mm/dd/yy") + "#"
  oreport.FormulaFields.GetItemByName("data").Text = "'" + vdia + "'"
  
  If existeix("c:\temp\llistatexpedicions.pdf") Then Kill "c:\temp\llistatexpedicions.pdf"
   oreport.ExportOptions.DiskFileName = "c:\temp\llistatexpedicions.pdf"
   oreport.ExportOptions.PDFExportAllPages = True
   oreport.ExportOptions.FormatType = crEFTPortableDocFormat
   oreport.ExportOptions.DestinationType = crEDTDiskFile
   oreport.Export False
  
 
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.CRViewer.Zoom (2)
   veurereport.width = Screen.width / 2
   veurereport.Left = (Screen.width / 2) - (veurereport.width / 2)
   veurereport.Top = 1000
   veurereport.Height = Screen.Height - 2000
   veurereport.Show 1, Me
   
End Sub
Private Sub mrebobinadores_Click()
Dim comanda As Double
Set dbsap = Nothing
Frameentregues.visible = False
 nummaquina = cadbl(mrebobinadores.tag)
 comanda = cadbl(mgeneral.tag)

eseccio = "Rebobinadores"
carregarmaquines "R"
taulaplanificacio = "planificacioreb"
carregarllistadecampstemporals "R"
If comanda > 0 Then
   carregarcomandaescullida cadbl(nummaquina), comanda
   Else: recarrearmaquinaseleccionada 3
End If
 
End Sub
Function potfermicromacro(numm As Byte) As String
   Dim rstm As Recordset
   Set rstm = dbcomandes.OpenRecordset("select rebmicromacro from maquines where maquina='R' and codi=" + atrim(numm))
   If Not rstm.EOF Then potfermicromacro = atrim(rstm!rebmicromacro)
   Set rstm = Nothing
End Function
Sub comprovarcomandesquenopodenfersealarebobinadoraescullida()
   Dim rst As Recordset
   Dim rstm As Recordset
   Dim vmicromacro As String
   Dim vcomandesnoespodenfer As String
   Dim vnovamaquina As String
   Dim vmicromacromaquinaassignada As String
   
   
   vmicromacro = potfermicromacro(nummaquina)
   Set rst = dbconsulta.OpenRecordset("select comanda,micromacro,maquina from " + taulaplanificacio + " where maquina=" + atrim(nummaquina) + " and micromacro<>''")
   While Not rst.EOF
      Set rstm = dbplanificacio.OpenRecordset("select maquina from planificacioreb where comanda=" + atrim(rst!comanda))
      vmicromacromaquinaassignada = "": vmicromacromaquina = ""
      If Not rstm.EOF Then vmicromacromaquinaassignada = potfermicromacro(rstm!maquina)
      If vmicromacromaquinaassignada <> "Tots" And InStr(1, vmicromacro, atrim(rst!micromacro)) = 0 And InStr(1, vmicromacromaquinaassignada, atrim(rst!micromacro)) = 0 Then vcomandesnoespodenfer = vcomandesnoespodenfer + IIf(vcomandesnoespodenfer = "", "", ",") + atrim(rst!comanda)
      rst.MoveNext
   Wend
   If vcomandesnoespodenfer <> "" Then
       If MsgBox("Les comandes " + vcomandesnoespodenfer + " NO es poden fer amb aquesta màquina pel Micro o Macroperforat" + Chr(10) + Chr(10) + Chr(10) + "Vols moure-les totes a una altra màquina?", vbCritical + vbDefaultButton2 + vbYesNo, "Error") = vbYes Then
           vnovamaquina = InputBox("Escriu el número de màquina que vols moure-les.", "Nova màquina")
           If cadbl(vnovamaquina) = 0 Then Exit Sub
           vmicromacro = potfermicromacro(cadbl(vnovamaquina))
           If Trim(vmicromacro) = "" Then MsgBox "Aquesta màquina tampoc pot fer Micro/Macro", vbCritical, "Atenció": GoTo fi
           'Set rst = dbconsulta.OpenRecordset("select comanda,micromacro,ordre from " + taulaplanificacio + " where maquina=" + atrim(nummaquina) + " and micromacro<>'' order by ordre desc")
           Set rst = dbconsulta.OpenRecordset("select comanda,micromacro,ordre from " + taulaplanificacio + " where comanda in(" + vcomandesnoespodenfer + ") order by ordre desc")
           While Not rst.EOF
                'If InStr(1, vmicromacro, atrim(rst!micromacro)) = 0 Or vmicromacro = "Tots" Then
                    canviarmaquinaalacomanda cadbl(rst!comanda), cadbl(vnovamaquina), 999 ' IIf(rst!ordre = 999, 999, 998 - rst.AbsolutePosition)
                'End If
                rst.MoveNext
           Wend
           MsgBox "Les comandes " + vcomandesnoespodenfer + " ja estan a la màquina " + vnovamaquina, vbInformation, "Moure comandes Micro/Macro"
       End If
   End If
fi:
   Set rst = Nothing
   Set rstm = Nothing
End Sub

Private Sub mreclamacions_Click()
  eseccio = "-Reclamació de comandes"
  Frameentregues.visible = False
  carregarmaquines "T"
  taulaplanificacio = "reclamacionscomandes"
  carregarllistadecampstemporals "C"
  carregar_ordre_correcte 0
  crearregistresalataulademodificacions
  ratoli "espera"
  reixa.visible = False
  ordrereixa = " order by datareclamacio desc"
  Command3.tag = ""
   configreixa
   'reordenarregistres
   poblarlareixa nummaquina
   ratoli "normal"
   reixa.visible = True
   buscarcomanda cadbl(mgeneral.tag)
End Sub
Sub crearregistresalataulademodificacions()
   dbplanificacio.Execute "insert into reclamacionscomandes select comanda from planificaciototes"
End Sub

Private Sub msoldadores_Click()
Dim comanda As Double
Set dbsap = Nothing
Frameentregues.visible = False
 nummaquina = cadbl(msoldadores.tag)
 comanda = cadbl(mgeneral.tag)

eseccio = "Soldadores"
carregarmaquines "S"
taulaplanificacio = "planificaciosol"
carregarllistadecampstemporals "S"
If comanda > 0 Then
   carregarcomandaescullida cadbl(nummaquina), comanda
   Else: recarrearmaquinaseleccionada 5
End If
End Sub

Private Sub multiseleccio_Click()
   If multiseleccio.Value = 0 Then
      reixa.RowSel = reixa.Row
      reixa.ColSel = reixa.Cols - 1
      canvidemaquina.Height = 525
        Else
          canvidemaquina.Height = 285
   End If
End Sub
Sub veureelpdf(ncomanda As String)
  Dim rstc As Recordset
  Set rstc = dbcomandes.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(cadbl(ncomanda)))
  obrir_pdf_treball cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio)
  
End Sub

Sub obrir_pdf_treball(treball As Double, modificacio As Double)
   Dim generarfitxer_pdf As String
   Dim generarfitxer_pdf_SC As String
   Dim ruta_documentacio_clixes As String
   
    ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
   If modificacio = 0 Then modificacio = 1
   'generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "_SC.pdf"
   'If Not existeix(generarfitxer_pdf) Then
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + format(treball, "00000") + "\pdf" + format(treball, "00000") + "-" + format(modificacio, "000") + ".pdf"
   generarfitxer_pdf_SC = ruta_documentacio_clixes + "\" + format(treball, "00000") + "\pdf" + format(treball, "00000") + "-" + format(modificacio, "000") + "_SC.pdf"
   'If existeix(generarfitxer_pdf) And existeix(generarfitxer_pdf_SC) Then
   '   If MsgBox("Hi ha el pdf per capes disponible." + Chr(10) + "VOLS VEURE'L?", vbInformation + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbYes Then
   '      generarfitxer_pdf = generarfitxer_pdf_SC
   '   End If
   'End If
   If existeix(generarfitxer_pdf) Then
     If existeix("c:\temp\pdftemp.pdf") Then Kill "c:\temp\pdftemp.pdf"
     FileCopy generarfitxer_pdf, "c:\temp\pdftemp.pdf"
     AcroPDF1.LoadFile "c:\temp\pdftemp.pdf"
     AcroPDF1.Height = 4665
     AcroPDF1.Left = 1245
     AcroPDF1.Top = 3180

       '  Else: MsgBox "No he trobat el fitxer" + Chr(10) + generarfitxer_pdf + Chr(10) + " i tampoc el de separació de colors.", vbCritical, "Error"
  End If
End Sub
Sub carregar_comanda(vnumc As Double)
   Dim rst As Recordset
   Dim vnumcol As Double
   'If Not vensenya_ultima_comanda Then Exit Sub
   Set rst = dbcomandes.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
   If rst.EOF Then GoTo fi
   Set rst = dbcomandes.OpenRecordset("select comanda from comandes where numtreball=" + atrim(rst!numtreball) + " and numordremodificacio=" + atrim(rst!numordremodificacio) + " order by comanda desc")
   If rst.EOF Then
      Unload formannex
      formannex.Show
      Exit Sub
   End If
   If cadbl(rst!comanda) = 0 Then
      Unload formannex
      formannex.Show
   End If
   formannex.carregarcomanda rst!comanda
   formannex.Show
   vnumcol = numcol("NºLot")
   formannex.Left = fcontrols.Left + reixa.ColPos(vnumcol) + reixa.Left + reixa.ColWidth(vnumcol) + 100
   formannex.Top = planificacio.Top + 200
fi:
   Set rst = Nothing
End Sub

Private Sub registres_DblClick()
  Clipboard.Clear
  Clipboard.SetText registres.caption
  MsgBox "Informació copiada al portapapers.", vbInformation, "Atenció"
End Sub

Private Sub reixa_Click()
  Dim col As Integer
  Dim v As String
  Dim numc As Double
  Dim vdataalbara As String
  Unload formannex
  col = reixa.col
  
  v = reixa.TextMatrix(reixa.Row, numcol("NºLot"))
  numc = cadbl(v)
  If InStr(1, v, "R") > 2 Then numc = cadbl(Mid(atrim(v), 1, InStr(1, v, "R") - 1))
  
  If multiseleccio.Value = 0 Then
    reixa.RowSel = reixa.Row
    reixa.ColSel = reixa.Cols - 1
  End If
  If reixa.TextMatrix(0, reixa.col) = "NºLot" And eseccio = "Impresores" Then carregar_comanda cadbl(v)
  If reixa.TextMatrix(0, reixa.col) = "Texte Impresio" Then
  End If
  If reixa.TextMatrix(0, reixa.col) = "Texte Impresio" Then
      If AcroPDF1.tag <> "no" And eseccio = "Impresores" Then
        AcroPDF1.visible = True
        veureelpdf atrim(numc)
      End If
      Else:
       If AcroPDF1.tag <> "no" Then
         AcroPDF1.LoadFile "a": AcroPDF1.visible = False
       End If
  End If
  If reixa.TextMatrix(0, reixa.col) = "Texte Impresio" Then
     postit.visible = True
     postit.Text = reixa.Text
     postit.Left = reixa.CellLeft + reixa.Left
     postit.Top = reixa.CellTop + reixa.Top
     postit.Height = reixa.CellHeight
     postit.width = Len(reixa.Text) * 100
  End If

  If (reixa.TextMatrix(0, reixa.col) = "Revisat" Or reixa.TextMatrix(0, reixa.col) = "Ok_Clixes") And reixa.Text <> "" Then
     vdataalbara = reixa.TextMatrix(reixa.Row, numcol("Data_albarà"))
     postit.visible = True
     postit.Text = nomordinadorrepassador(numc, vdataalbara, IIf(reixa.TextMatrix(0, reixa.col) = "Revisat", "usuari_revisat", "usuari_okclixes"))
     postit.Left = reixa.CellLeft + reixa.Left
     postit.Top = reixa.CellTop + reixa.Top + reixa.CellHeight
     postit.Height = reixa.CellHeight
     postit.width = Len(postit.Text) * 120
  End If
  If taulaplanificacio = "planificaciototes" Then
     mgeneral.tag = atrim(numc)
     mimpresores.tag = cadbl(Mid(reixa.TextMatrix(reixa.Row, numcol("Imp.")), 1, 1))
     mlaminadores.tag = cadbl(Mid(reixa.TextMatrix(reixa.Row, numcol("Lam.")), 1, 1))
     mrebobinadores.tag = cadbl(Mid(reixa.TextMatrix(reixa.Row, numcol("Reb.")), 1, 1))
  End If
  If reixa.BackColorFixed <> treurefiltre.BackColor Then
      ordenarlareixa reixa.col
  End If
End Sub
Function nomordinadorrepassador(vnumc As Double, vdataalbara As String, vnomcamp) As String
   Dim rst As Recordset
   Set rst = dbplanificacio.OpenRecordset("select * from planificacioent where comanda=" + atrim(vnumc) + " and dataalbara=#" + format(vdataalbara, "mm/dd/yy") + "#")
   If Not rst.EOF Then nomordinadorrepassador = atrim(rst.Fields(vnomcamp))
   Set rst = Nothing
End Function
Sub ordenarlareixa(vcol As Double)
    triar_ordre camps(filtre(vcol).tag, 1)
End Sub
Sub ordenarlareixa2()
'vordre = camps(reixa.col + 1, 1)
      vordre = nomdelcamp(reixa.TextMatrix(0, reixa.col))
     ' If vordre = "dataobertura" Then vordre = "cvdate(dataobertura)"
     If vordre = "dataexpedicio" And InStr(1, bordre.tag, vordre) = 0 Then bordre.tag = "dataexpedicio ASC"
      If InStr(1, bordre.tag, vordre) > 0 Then
          If InStr(1, bordre.tag, "ASC") > 0 Then
                bordre.tag = " DESC"
              Else: bordre.tag = " ASC"
          End If
           Else
              bordre.tag = " ASC"
      End If
      etordre = camps(reixa.col + 1, 3) + " " + bordre.tag
      bordre.tag = vordre + bordre.tag
      etmsgajuda.visible = False
      bordre.BackColor = treurefiltre.BackColor
      reixa.BackColorFixed = treurefiltre.BackColor
      If atrim(bordre.tag) <> "" Then ordrereixa = " order by " + atrim(bordre.tag)
      poblarlareixa nummaquina, IIf(whereultimfiltre <> "", " and ", "") + whereultimfiltre
End Sub
Sub triar_ordre(nom As String)
  'nom = nomdelcamp(nom)
  Dim ascodesc As String
  Static ultimordre As String
  ratoli "espera"
  If ultimordre = nom Then
     ascodesc = " DESC"
     ultimordre = ""
       Else: ultimordre = nom
  End If
  If Len(nom) > 2 Then
    ordrereixa = " order by " + nom + ascodesc
    
   Else: ordrereixa = triar_ordre_reixa
  End If
  If taulaplanificacio <> "planificacioent" Then
    carregar_ordre_correcte nummaquina
    possarhoraprevista
  End If
   configreixa
   poblarlareixa nummaquina
   ratoli "normal"
   
End Sub
Function triar_ordre_reixa() As String
   Dim rst As Recordset
   triar_ordre_reixa = " order by ordre,dataimpresio"
   If primerapestanya = "IMPRESORES" Then triar_ordre_reixa = " order by data1 desc,importancia desc"
   Exit Function
   Set rst = dbplanificacio.OpenRecordset("select * from " + taulaplanificacio + " where dataprogramada<>null  and maquina=" + atrim(nummaquina) + " order by dataprogramada")
   If rst.EOF Then
      triar_ordre_reixa = " order by ordre"
     Else: triar_ordre_reixa = " order by horaprogramada"
   End If
     
End Function
Function nomdelcamp(nomcapcalera As String) As String
  Dim i As Byte
  For i = 1 To 100
    If camps(i, 3) = nomcapcalera Then nomdelcamp = camps(i, 1)
    
  Next i
End Function

Sub cridarcomandescompra(comanda As Double)
  Dim rstcompres As Recordset
  Set rstcompres = dbcompres.OpenRecordset("SELECT capcalera.numcomanda as numc , comandesxlinia.numcomanda , liniescompra.kgentregats AS kilos FROM capcalera RIGHT JOIN (liniescompra RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra) ON capcalera.id = liniescompra.idcompra WHERE (((comandesxlinia.numcomanda)=" + atrim(comanda) + "));")
  If Not rstcompres.EOF Then
       comanda = rstcompres!numc
      Else: Exit Sub
  End If
On Error GoTo obrircomandes
  escriure_ini "Planificacio", "comandacompraxrobrir", atrim(comanda), "comandes.ini"
  AppActivate "Comandes de compra.", True
  
  
  On Error Resume Next
  Exit Sub
obrircomandes:
  On Error Resume Next
   Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", "comandes.ini")) + "compres.exe", vbNormalFocus
End Sub


Sub cridarcomandes(comanda As Double)
 On Error GoTo obrircomandes
  escriure_ini "Planificacio", "comandaxrobrir", atrim(comanda), "comandes.ini"
  AppActivate "Manteniment de Comandes"
  
  On Error Resume Next
  Exit Sub
obrircomandes:
  On Error Resume Next
   Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", "comandes.ini")) + "comandes.exe - comandes", vbNormalFocus
End Sub

Sub passarinventadaareal(ordre As Integer, comanda As Double, Data1 As String, data2 As String, importancia As Byte)
     Dim rst As Recordset
     Dim resp As String
     Dim novacomanda As Double
     resp = InputBox("Entra la comanda que vols que es possi en aquesta posició." + Chr(10) + "O escriu [ELIMINAR] si vols treure aquesta reserva d'aquesta posició.", "Col.locar comanda real en aquesta posició.")
     If cadbl(resp) > 10000 Then
         novacomanda = cadbl(resp)
         Set rst = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(novacomanda))
         If Not rst.EOF Then
             If IsDate(Data1) Then
               dbcomandes.Execute "update comandes set dataentrega=#" + format(Data1, "mm/dd/yy") + "# where comanda=" + atrim(novacomanda)
               dbconsulta.Execute "update " + taulaplanificacio + " set data1=#" + format(Data1, "mm/dd/yy") + "# where comanda=" + atrim(novacomanda)
               dbconsulta.Execute "update planificaciototes set data1=#" + format(Data1, "mm/dd/yy") + "# where comanda=" + atrim(novacomanda)
               dbplanificacio.Execute "update planificaciototes set data1=#" + format(Data1, "mm/dd/yy") + "# where comanda=" + atrim(novacomanda)
             End If
             If IsDate(data2) Then
               dbconsulta.Execute "update " + taulaplanificacio + " set data2=#" + format(data2, "mm/dd/yy") + "# where comanda=" + atrim(novacomanda)
               dbconsulta.Execute "update planificaciototes set data2=#" + format(data2, "mm/dd/yy") + "# where comanda=" + atrim(novacomanda)
               dbplanificacio.Execute "update planificaciototes set data2=#" + format(data2, "mm/dd/yy") + "# where comanda=" + atrim(novacomanda)
             End If
             dbconsulta.Execute "update planificaciototes set importancia=" + atrim(importancia) + " where comanda=" + atrim(novacomanda)
             dbplanificacio.Execute "update planificaciototes set importancia=" + atrim(importancia) + " where comanda=" + atrim(novacomanda)
             dbplanificacio.Execute "update " + taulaplanificacio + " set comanda=9999 where comanda=" + atrim(novacomanda)
             dbplanificacio.Execute "update " + taulaplanificacio + " set comanda=" + atrim(novacomanda) + " where comanda=" + atrim(comanda)
             dbplanificacio.Execute "delete * from " + taulaplanificacio + " where comanda=9999"
             If taulaplanificacio <> "planificaciototes" Then canviarmaquinaalacomanda novacomanda, cadbl(nummaquina), cadbl(ordre)
             dbconsulta.Execute "delete * from " + taulaplanificacio + " where comanda=" + atrim(comanda)
             dbconsulta.Execute "delete * from planificaciototes where comanda=" + atrim(comanda)
             dbplanificacio.Execute "delete * from " + taulaplanificacio + " where comanda=" + atrim(comanda)
             dbplanificacio.Execute "delete * from planificaciototes where comanda=" + atrim(comanda)
             buscarcomanda novacomanda
                  Else: MsgBox "No trobo aquesta comanda com a pendent de planificar.", vbInformation, "Atenció"
         End If
          Else
            If resp = "ELIMINAR" Then
                dbconsulta.Execute "delete * from " + taulaplanificacio + " where comanda=" + atrim(comanda)
                dbplanificacio.Execute "delete * from " + taulaplanificacio + " where comanda=" + atrim(comanda)
            End If
     End If
End Sub
Sub canviarordrecomanda(canvi As Boolean, comanda As Double, Optional ordre As String)
    If eseccio = "Laminadores" And programaoperaris Then If Not comprovar_siporta_Zipper_i_potferla(comanda, cadbl(nummaquina)) Then MsgBox "Aquesta màquina no pot fer ZIPPER primer canvia de màquina.", vbCritical, "Error": Exit Sub
    If cadbl(ordre) = 0 Then ordre = InputBox("Entra el numero d'ordre que vols per aquesta comanda.", "Atenció")
    If cadbl(ordre) <= 999 And cadbl(ordre) > 0 Then
       canvi = True
       dbconsulta.Execute "update " + taulaplanificacio + " set maquina=" + atrim(nummaquina) + ",ordre=" + atrim(cadbl(ordre)) + IIf(cadbl(ordre) <> 999, "-0.1", "") + " where comanda=" + atrim(comanda)
       dbplanificacio.Execute "update " + taulaplanificacio + " set maquina=" + atrim(nummaquina) + ",ordre=" + atrim(cadbl(ordre)) + IIf(cadbl(ordre) <> 999, "-0.1", "") + " where comanda=" + atrim(comanda)
       dbplanificacio.Execute "update " + taulaplanificacio + " set maquina=" + atrim(nummaquina) + ",dataordreassignat=" + IIf(cadbl(ordre) <> 999, "Now", "Null") + " where comanda=" + atrim(comanda)
       dbplanificacio.Execute "delete * from llistadeplanificacions where comanda=" + atrim(comanda)
       If cadbl(ordre) <> 999 Then dbplanificacio.Execute "insert into llistadeplanificacions (ordre,comanda,maquina,dataordreassignat,seccio) values(" + atrim(cadbl(ordre)) + "," + atrim(comanda) + "," + atrim(nummaquina) + ",now,'" + taulaplanificacio + "')"
    End If
End Sub
Sub mirar_si_avisar_client_vindra()
   If reixa.TextMatrix(reixa.Row, numcol("Cli.Vindrà")) = "S" Then MsgBox "Atenció aquest client vol venir a revisar la impresió, s'ha de tenir en compte.", vbExclamation, "Atenció"
End Sub
Private Sub reixa_DblClick()
  Dim ordre As String
  Dim obs As String
  Dim canvi As Boolean
  Dim comanda As Double
  Dim dataprog As String
  Dim horaprog As String
  Dim rstt As Recordset
  Dim vr As Integer
  Dim v As String
  Dim resp As String
  Dim vdata As Date
  Dim rstobs As Recordset
  Dim vdataalbara As String
  Dim vnumalbara As String
  Dim vsumaunitats As Double
  Dim vmesura As String
  Dim vp As Double
  If programanomeslectura Then Exit Sub
  If reixa.Row = 0 Then Exit Sub
  If framereclamar.visible = True Then
    If atrim(reixa.TextMatrix(reixa.Row, numcol("Muntat?"))) <> "" Then
       MsgBox "Aquesta comanda no pots reclamar-la ja està a muntadora", vbCritical, "A muntadora"
        Else: afegir_comandaperreclamar
    End If
    Exit Sub
  End If
  v = reixa.TextMatrix(reixa.Row, numcol("NºLot"))
  vdataalbara = reixa.TextMatrix(reixa.Row, numcol("Data_albarà"))
  If IsDate(vdataalbara) Then vdataalbara = format(vdataalbara, "mm/dd/yy")
  comanda = cadbl(v)
  If InStr(1, v, "R") > 2 Then comanda = Mid(atrim(v), 1, InStr(1, v, "R") - 1)
  If Not poteditar Then MsgBox "No tens permis per fer canvis a planificacio", vbCritical, "Atenció": Exit Sub
  mirar_si_avisar_client_vindra
  If reixa.TextMatrix(0, reixa.col) = "Ordre" Then
    canviarordrecomanda canvi, comanda
     Else
      If taulaplanificacio <> "planificaciototes" And taulaplanificacio <> "reclamacionscomandes" And taulaplanificacio <> "planificacioent" Then
        dbconsulta.Execute ("update " + taulaplanificacio + " set maquina=" + atrim(nummaquina) + ",ordre=" + atrim(cadbl(reixa.TextMatrix(reixa.Row, numcol("Ordre")))) + " where comanda=" + atrim(comanda))
        dbplanificacio.Execute "update " + taulaplanificacio + " set maquina=" + atrim(nummaquina) + ",ordre=" + reixa.TextMatrix(reixa.Row, numcol("Ordre")) + " where comanda=" + atrim(comanda)
        
      End If
  End If
  If reixa.TextMatrix(0, reixa.col) = "Extra_Cost" Then
       Set rstobs = dbcomandes.OpenRecordset("select * from comandes_observacioPVP where comanda=" + atrim(comanda))
       If Not rstobs.EOF Then MsgBox atrim(rstobs!observacio) + vbNewLine + IIf(cadbl(rstobs!extracost) > 0, "Extracost: " + atrim(rstobs!extracost) + "€", ""), , "Observacio PVP"
       Set rstobs = Nothing
  End If
  If reixa.TextMatrix(0, reixa.col) = "NºAlbarà" Then
     vnumalbara = reixa.TextMatrix(reixa.Row, reixa.col)
     'Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", "comandes.ini")) + "vendes.exe obriralbara " + atrim(vnumalbara), vbNormalFocus
     ensenyar_albara_SAP cadbl(vnumalbara)
  End If
  If reixa.TextMatrix(0, reixa.col) = "NºLinia Imp" Then
        v = reixa.TextMatrix(reixa.Row, reixa.col)
        borrarelfiltre
        filtre(numcol("NºLinia Imp")).Text = Mid(v, 1, 3)
        If numcol("Estat") > 0 Then filtre(numcol("Estat")).Text = "E,I"
        filtre(0).SetFocus
        filtre_LostFocus 0
        GoTo fi
  End If
  If programaoperaris Then GoTo fi
  If reixa.TextMatrix(0, reixa.col) = "Data_Expedició" Then
     'Me.m_opcions_reixa.WindowList = True
       Me.PopupMenu mexpedicio
  End If
   
  
  If reixa.TextMatrix(0, reixa.col) = "DataImp." Or reixa.TextMatrix(0, reixa.col) = "DataLam." Then
     Set rstt = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(comanda))
     If Not rstt.EOF Then
       If IsDate(rstt!horaprogramada) Then
          dataprog = InputBox("Ja hi ha una data programada si vols canvia-la o escriu [ELIMINAR] per eliminar-la.", "Programació de la data", format(rstt!horaprogramada, "dd/mm/yy"))
         Else: dataprog = InputBox("Entra la data que vols programar." + Chr(10) + " Ex:01/08/" + format(Now, "yy"), "Programació de la data")
       End If
       If UCase(dataprog) = "ELIMINAR" Then
           dbplanificacio.Execute "update " + taulaplanificacio + " set dataprogramada=null where comanda=" + atrim(comanda)
           dbconsulta.Execute "update " + taulaplanificacio + " set horaprogramada=null where comanda=" + atrim(comanda)
           
           canvi = True
           GoTo fi
       End If
        If Not IsDate(dataprog) Then Exit Sub
        horaprog = InputBox("Entra la hora que vols programar." + Chr(10) + " Ex:10:30", "Programació de la hora", IIf(Not rstt.EOF, format(rstt!horaprogramada, "hh:nn"), ""))
        If Len(horaprog) < 4 Then Exit Sub
        If Not IsDate(dataprog + " " + horaprog) Then Exit Sub
        dbplanificacio.Execute "update  " + taulaplanificacio + " set dataprogramada=#" + format(dataprog + " " + horaprog, "mm/dd/yy hh:nn") + "# where comanda=" + atrim(comanda)
        dbplanificacio.Execute "update " + taulaplanificacio + " set maquina=" + atrim(nummaquina) + ",ordre=" + reixa.TextMatrix(reixa.Row, numcol("Ordre")) + " where comanda=" + atrim(comanda)
        canvi = True
      
      GoTo fi
     End If
  End If
  
  ' canvia de maquina a seccio a general
  If reixa.TextMatrix(0, reixa.col) = "Imp." Or reixa.TextMatrix(0, reixa.col) = "Lam." Or reixa.TextMatrix(0, reixa.col) = "Reb." Or reixa.TextMatrix(0, reixa.col) = "Sol." Then
     canviardemaquinageneral comanda, Mid(reixa.TextMatrix(0, reixa.col), 1, 1)
  End If
  
  
  If reixa.TextMatrix(0, reixa.col) = "NºLot" Then
    If comanda < 10000 Then
        passarinventadaareal reixa.TextMatrix(reixa.Row, numcol("Ordre")), comanda, reixa.TextMatrix(reixa.Row, numcol("Data1")), reixa.TextMatrix(reixa.Row, numcol("Data2")), reixa.TextMatrix(reixa.Row, numcol("Impor."))
      Else
         cridarcomandes comanda
         If taulaplanificacio = "planificacioent" Then GoTo fi
    End If
    canvi = True
  End If
  
 ' If comanda < 10000 Then GoTo fi
  
  If reixa.TextMatrix(0, reixa.col) = "Material" Or reixa.TextMatrix(0, reixa.col) = "MaterialPC" Or reixa.TextMatrix(0, reixa.col) = "MaterialPC2" Then
    cridarcomandescompra comanda + IIf(reixa.TextMatrix(0, reixa.col) = "MaterialPC", 1, IIf(reixa.TextMatrix(0, reixa.col) = "MaterialPC2", 2, 0))
  End If
  
  If reixa.TextMatrix(0, reixa.col) = "Observació_Entrega" Then
    obs = InputBox("Entra la observació", "Entrada", reixa.TextMatrix(reixa.Row, reixa.col))
    If obs <> "" Then
      dbplanificacio.Execute "insert into planificacioent (comanda,dataalbara) values (" + atrim(comanda) + ",#" + vdataalbara + "#)"
      dbconsulta.Execute "update " + taulaplanificacio + " set observacio='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda) + " and dataalbara=#" + vdataalbara + "#"
      dbplanificacio.Execute "update " + taulaplanificacio + " set observacio='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda) + " and dataalbara=#" + vdataalbara + "#"
      reixa.Text = obs
    End If
  End If
  
  If reixa.TextMatrix(0, reixa.col) = "Revisat" Then
   vp = 0
   If reixa.TextMatrix(reixa.Row, numcol("PVP_Revisat")) = "S" Then
    reixa.col = numcol("Preu_Albarà")
    If reixa.TextMatrix(reixa.Row, numcol("Revisat")) <> "P" And reixa.TextMatrix(reixa.Row, numcol("Revisat")) <> "F" Then obs = "P"
    If reixa.TextMatrix(reixa.Row, numcol("Revisat")) = "P" Then
      If reixa.TextMatrix(reixa.Row, numcol("Facturat")) = "S" Then
       obs = "F"
       If reixa.CellBackColor <> 0 Then MsgBox "No pots validar una entrega si el PVP no està correcte", vbCritical, "Error": vp = 1
         Else: MsgBox "No pots validar una Facturació si no està facturat.", vbCritical, "Error": obs = "F": vp = 1
      End If
    End If
validarlinia:
    If vp = 0 Then
     If MsgBox("Has revisat aquesta comanda " + atrim(comanda) + "?", vbQuestion + vbDefaultButton2 + vbYesNo, "Revisat") <> vbYes Then obs = "N"
    End If
    reixa.col = numcol("Revisat")
       Else: MsgBox "No pots marcar revisat si no hi ha el PVP revisat abans.", vbCritical, "Error"
   End If
   If vp = 1 Then If UCase(InputBox("Vols validar igualment aquesta Entrega?" + vbNewLine + "Escriu [VALIDAR] per fer-ho.", "Validar igualment")) = "VALIDAR" Then vp = 0: GoTo validarlinia
    If obs <> "" And obs <> reixa.TextMatrix(reixa.Row, reixa.col) Then
        dbplanificacio.Execute "insert into planificacioent (comanda,dataalbara) values (" + atrim(comanda) + ",#" + vdataalbara + "#)"
        dbconsulta.Execute "update " + taulaplanificacio + " set revisat='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda) + " and dataalbara=#" + vdataalbara + "#"
        dbplanificacio.Execute "update " + taulaplanificacio + " set revisat='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda) + " and dataalbara=#" + vdataalbara + "#"
        dbplanificacio.Execute "update " + taulaplanificacio + " set usuari_revisat='" + IIf(obs = "P" Or obs = "F", nomordinador, "") + "' where comanda=" + atrim(comanda) + " and dataalbara=#" + vdataalbara + "#"
        reixa.Text = obs
    End If
  End If
  
  If reixa.TextMatrix(0, reixa.col) = "Ok_Clixes" Then
   'If cadbl(reixa.TextMatrix(reixa.Row, numcol("Preu_Clixes"))) > 0 Then
    vmsg = missatge_revisióclixes(comanda)
    If MsgBox(vmsg + "?", vbQuestion + vbDefaultButton2 + vbYesNo, "Revisat clixes") = vbYes Then
         obs = "S"
          Else: obs = "N"
    End If
    If obs <> "" And obs <> reixa.TextMatrix(reixa.Row, reixa.col) Then
        dbplanificacio.Execute "insert into planificacioent (comanda,dataalbara) values (" + atrim(comanda) + ",#" + vdataalbara + "#)"
        dbconsulta.Execute "update " + taulaplanificacio + " set okclixes='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda) + " and dataalbara=#" + vdataalbara + "#"
        dbplanificacio.Execute "update " + taulaplanificacio + " set okclixes='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda) + " and dataalbara=#" + vdataalbara + "#"
        dbplanificacio.Execute "update " + taulaplanificacio + " set usuari_okclixes='" + IIf(obs = "S", nomordinador, "") + "' where comanda=" + atrim(comanda) + " and dataalbara=#" + vdataalbara + "#"
        reixa.Text = obs
    End If
   'End If
  End If
  
  If reixa.TextMatrix(0, reixa.col) = "Obs_Expedició" Then
    obs = InputBox("Entra la observació", "Entrada", reixa.TextMatrix(reixa.Row, reixa.col))
    If obs <> "" Then
'      canvi = True
      dbconsulta.Execute "update " + taulaplanificacio + " set observacioexpedicio='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda)
      dbplanificacio.Execute "update planificaciototes set observacioexpedicio='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda)
      modificarexpedicio comanda, reixa.TextMatrix(reixa.Row, numcol("Data_Expedició")), reixa.TextMatrix(reixa.Row, numcol("Data_Expedició")), obs, atrim(reixa.TextMatrix(reixa.Row, numcol("Nom Client")))
      reixa.Text = obs
    End If
  End If
  If reixa.TextMatrix(0, reixa.col) = "Entrega_ToP" Then
     vsumaunitats = sumaunitatsentregades(comanda, vmesura)
     If InStr(1, reixa.Text, "*") Then MsgBox "Unitat entregades totals: " + atrim(vsumaunitats) + " " + vmesura
  End If
  'If reixa.TextMatrix(0, reixa.col) = "Preu_Albarà" And reixa.CellBackColor <> 0 Then
  If reixa.TextMatrix(0, reixa.col) = "Preu_Albarà" Then
    vpreufactura = mirarpreufactura(comanda, reixa.TextMatrix(reixa.Row, numcol("Data_albarà")))
    vp = mirarpreucomanda(comanda)
    MsgBox "El preu de la comanda es de: " + IIf(vp = -1, "Sense Cost", atrim(vp) + "€") + vbNewLine + "i el preu de l'albarà SAP es de " + atrim(reixa.Text) + vbNewLine + "El preu de la FACTURA es de: " + atrim(vpreufactura) + "€", vbCritical, "Atenció al Preu"
  End If
  If reixa.TextMatrix(0, reixa.col) = "Observacio" Then
    obs = InputBox("Entra la observació", "Entrada", reixa.TextMatrix(reixa.Row, reixa.col))
    If obs <> "" Then
      canvi = True
      dbconsulta.Execute "update " + taulaplanificacio + " set observacions='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda)
      dbplanificacio.Execute "update planificaciototes set observacio='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda)
    End If
  End If
  If reixa.TextMatrix(0, reixa.col) = "ObservacioOficina" Then
    obs = InputBox("Entra la observació", "Entrada", reixa.TextMatrix(reixa.Row, reixa.col))
    If obs <> "" Then
      'canvi = True
      reixa.TextMatrix(reixa.Row, numcol("ObservacioOficina")) = treure_apostruf(obs)
      dbconsulta.Execute "update " + taulaplanificacio + " set observacionsoficina='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda)
      dbplanificacio.Execute "update reclamacionscomandes set observacionsoficina='" + treure_apostruf(obs) + "' where comanda=" + atrim(comanda)
    End If
  End If
  
  If reixa.TextMatrix(0, reixa.col) = "Data1" Then
    obs = InputBox("Entra la primera data prevista" + Chr(10) + "Escriu [B] per borrar la data.", "Entrada", reixa.TextMatrix(reixa.Row, reixa.col))
    If IsDate(obs) Then
      If DateDiff("d", Now, obs) < 1 Then MsgBox "Aquesta data ja està passada no es pot possar.", vbCritical, "Error": Exit Sub
      If MsgBox("Aquesta data es d'aqui a  " + atrim(DateDiff("d", Now, obs)) + " DIES" + Chr(10) + "Es correcte?", vbDefaultButton2 + vbYesNo, "Canvi de data") = vbNo Then Exit Sub
      canvi = True
      dbconsulta.Execute "update " + taulaplanificacio + " set data1=#" + format(obs, "mm/dd/yy") + "# where comanda=" + atrim(comanda)
      dbconsulta.Execute "update planificaciototes set data1=#" + format(obs, "mm/dd/yy") + "# where comanda=" + atrim(comanda)
      dbplanificacio.Execute "update planificaciototes set data1=#" + format(obs, "mm/dd/yy") + "# where comanda=" + atrim(comanda)
      If comanda > 10000 Then
          dbcomandes.Execute "update comandes set dataentrega=#" + format(obs, "mm/dd/yy") + "# where comanda=" + atrim(comanda)
         Else
            dbplanificacio.Execute "update planificaciototes set data1=#" + format(obs, "mm/dd/yy") + "# where comanda=" + atrim(comanda)
      End If
       Else
        If UCase(obs) = "B" Then
         canvi = True
         dbconsulta.Execute "update " + taulaplanificacio + " set data1=null where comanda=" + atrim(comanda)
         dbconsulta.Execute "update planificaciototes set data1=null where comanda=" + atrim(comanda)
         dbplanificacio.Execute "update planificaciototes set data1=null where comanda=" + atrim(comanda)
         If comanda > 10000 Then
            dbcomandes.Execute "update comandes set dataentrega=null where comanda=" + atrim(comanda)
         End If
        End If
    End If
  End If
  
  If reixa.TextMatrix(0, reixa.col) = "Data2" Then
    obs = InputBox("Entra la segona data prevista" + Chr(10) + "Escriu [B] per borrar la data.", "Entrada", reixa.TextMatrix(reixa.Row, reixa.col))
    If IsDate(obs) Then
      If DateDiff("d", Now, obs) < 1 Then MsgBox "Aquesta data ja està passada no es pot possar.", vbCritical, "Error": Exit Sub
      If MsgBox("Aquesta data es d'aqui a  " + atrim(DateDiff("d", Now, obs)) + " DIES" + Chr(10) + "Es correcte?", vbDefaultButton2 + vbYesNo, "Canvi de data") = vbNo Then Exit Sub
      canvi = True
      dbplanificacio.Execute "update planificaciototes set data2=#" + format(obs, "mm/dd/yy") + "# where comanda=" + atrim(comanda)
      dbconsulta.Execute "update " + taulaplanificacio + " set data2=#" + format(obs, "mm/dd/yy") + "# where comanda=" + atrim(comanda)
      dbconsulta.Execute "update planificaciototes set data2=#" + format(obs, "mm/dd/yy") + "# where comanda=" + atrim(comanda)
        Else
          If UCase(obs) = "B" Then
           canvi = True
           dbplanificacio.Execute "update planificaciototes set data2=null where comanda=" + atrim(comanda)
           dbconsulta.Execute "update " + taulaplanificacio + " set data2=null where comanda=" + atrim(comanda)
           dbconsulta.Execute "update planificaciototes set data2=null where comanda=" + atrim(comanda)
         End If
    End If
  End If
  
  If reixa.TextMatrix(0, reixa.col) = "Impor." Then
    obs = InputBox("Entra la Importancia de la comanda ", "Entrada", reixa.TextMatrix(reixa.Row, reixa.col))
    If IsNumeric(obs) Then
      canvi = True
      dbconsulta.Execute "update " + taulaplanificacio + " set importancia=" + atrim(cadbl(obs)) + " where comanda=" + atrim(comanda)
      dbplanificacio.Execute "update planificaciototes set importancia=" + atrim(cadbl(obs)) + " where comanda=" + atrim(comanda)
      If cadbl(obs) = 3 And elclientescrops Then treurestandby comanda
    End If
  End If
  
  If reixa.TextMatrix(0, reixa.col) = "Gestió_Of" And reixa.TextMatrix(reixa.Row, numcol("Tipus_R")) <> "" Then
    resp = ""
    vr = MsgBox("Creus que es podrà complir la reclamació?", vbExclamation + vbDefaultButton2 + vbYesNoCancel, "Gestió Oficina")
    If vr = vbYes Then resp = "Sí"
    If vr = vbNo Then resp = "No"
    vdata = nul
    If resp <> "" Then vdata = Now
     reixa.TextMatrix(reixa.Row, numcol("Gestió_Of")) = resp
        reixa.TextMatrix(reixa.Row, numcol("Data_Gestió")) = IIf(vdata <> "0:00:00", format(vdata, "dd/mm"), "")
        dbconsulta.Execute "update reclamacionscomandes set tipusreclamacio='" + resp + "' where comanda=" + atrim(comanda)
        dbconsulta.Execute "update reclamacionscomandes set datareclamacio=" + IIf(vdata <> "0:00:00", "now", "null") + " where comanda=" + atrim(comanda)
        dbplanificacioalicia.Execute "update reclamacionscomandes set contestaoficina='" + resp + "' where comanda=" + atrim(comanda)
        dbplanificacioalicia.Execute "update reclamacionscomandes set datagestio=" + IIf(vdata <> "0:00:00", "now", "null") + " where comanda=" + atrim(comanda)
        dbplanificaciooperaris.Execute "update reclamacionscomandes set contestaoficina='" + resp + "' where comanda=" + atrim(comanda)
        dbplanificaciooperaris.Execute "update reclamacionscomandes set datagestio=" + IIf(vdata <> "0:00:00", "now", "null") + " where comanda=" + atrim(comanda)
  End If
  If reixa.TextMatrix(0, reixa.col) = "Data_EntradaFàbrica" Then
    resp = InputBox("Escriu la data d'entrada a fàbrica.", "Data entrada", format(Now, "dd/mm/yy"))
    If StrPtr(resp) <> 0 Then
      If resp = "" Then
         vdata = "0:00:00"
          Else: vdata = resp
      End If
      If vdata = "0:00:00" Then
         reixa.TextMatrix(reixa.Row, numcol("Data_EntradaFàbrica")) = ""
          Else:  reixa.TextMatrix(reixa.Row, numcol("Data_EntradaFàbrica")) = format(vr, "dd/mm")
      End If
      dbconsulta.Execute "update reclamacionscomandes set dataentradafabrica=" + IIf(vdata <> "0:00:00", "now", "null") + " where comanda=" + atrim(comanda)
      dbplanificacioalicia.Execute "update reclamacionscomandes set dataentradafabrica='" + IIf(vdata <> "0:00:00", "now", "null") + "' where comanda=" + atrim(comanda)
      dbplanificaciooperaris.Execute "update reclamacionscomandes set dataentradafabrica='" + IIf(vdata <> "0:00:00", "now", "null") + "' where comanda=" + atrim(comanda)
      dbplanificaciooperaris.Execute "update reclamacionscomandes set dataentradafabrica=" + IIf(vdata <> "0:00:00", "now", "null") + " where comanda=" + atrim(comanda)
    End If
  End If
  If reixa.TextMatrix(0, reixa.col) = "Tipus_R" Then Me.PopupMenu menureclam
fi:
  If canvi Then
    borrarlacomandaabuscar
    actualitzarlareixa comanda
    ' poblarlareixa nummaquina
  End If
End Sub
Function mirarpreufactura(vnumc As Double, vdataalb As String) As Double
   Dim rst As Recordset
   Set rst = dbsap.OpenRecordset("select * from Importada_LiniesFacturesSAP_Inplacsa where U_GSP_INFABLOTE='" + atrim(vnumc) + "' and Dataalbara=#" + atrim(format(vdataalb, "mm/dd/yy")) + "# and ItemCode<>'IMP_ENV' and ItemCode<>'PLATES'")
   If Not rst.EOF Then
     While rst!Tipusdelinia = "A" Or rst.EOF
          Set rst = dbsap.OpenRecordset("select * from Importada_LiniesFacturesSAP_Inplacsa where U_GSP_INFABLOTE='" + atrim(vnumc) + "' and ItemCode<>'IMP_ENV' and ItemCode<>'PLATES' and numfact>" + atrim(rst!numfact) + " order by numfact")
          If rst.EOF Then GoTo surtwhile
     Wend
surtwhile:
   End If
   If Not rst.EOF Then mirarpreufactura = Redondejar(cadbl(rst!Price), 4)
   Set rst = Nothing
End Function
Sub ensenyar_albara_SAP(vnumalbaraSAP As Double)
   Dim vpdftmp As String
   Dim vnomfitxer As String
   vnomfitxer = "\\ord_copies\AlbaransSAPClients\" + atrim(vnumalbaraSAP) + ".pdf"
   If Not existeix(vnomfitxer) Then MsgBox "No he trobat el PDF", vbCritical, "Error": Exit Sub
   vpdftmp = "c:\temp\pdfalbaratmp.pdf"
   
    AcroPDF1.Height = reixa.Height
    AcroPDF1.Left = reixa.CellLeft + reixa.CellWidth + 200
    AcroPDF1.Top = reixa.Top
    AcroPDF1.LoadFile ""
    If existeix(vpdftmp) Then Kill vpdftmp
    FileCopy vnomfitxer, vpdftmp
    If UCase(nomordinador) <> "ORD_RSIMON" Then
        AcroPDF1.LoadFile vpdftmp
        AcroPDF1.visible = True
         Else: obrir_document vpdftmp
    End If
   
   
End Sub
Function sumaunitatsentregades(vnumc As Double, vmesura As String) As Double
  Dim rst As Recordset
  Set rst = dbplanificacio.OpenRecordset("select sum(quantitat) as squantitat,first(unitatmesura) as fmesura from liniesalbara where lotinplacsa=" + atrim(vnumc))
  If Not rst.EOF Then
      vmesura = rst!fmesura
      sumaunitatsentregades = cadbl(rst!squantitat)
      
  End If
  Set rst = Nothing
End Function
Function missatge_revisióclixes(vnumc As Double) As String
  Dim rst As Recordset
  Dim rstc As Recordset
  Dim vmsg As String
  Set rst = dbclixes.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
      vmsg = "DETALL PRESSUPOST CLIXES" + vbNewLine + vbNewLine + "Treball: " + atrim(rst!numtreball) + " Versió:" + atrim(rst!numordremodificacio) + " Comanda: " + atrim(vnumc) + vbNewLine
      Set rst = dbclixes.OpenRecordset("SELECT Clixes.id_treball, Modificacions.ordre, Clixes.nomclienttemporal, Modificacions.codiclientfactclixes, pressupostos.preu FROM (Clixes LEFT JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) LEFT JOIN pressupostos ON (Modificacions.ordre = pressupostos.ordremodificacio) AND (Modificacions.id_treball = pressupostos.id_treball) where clixes.id_treball=" + atrim(rst!numtreball) + " and modificacions.ordre=" + atrim(rst!numordremodificacio))
      If Not rst.EOF Then
          If cadbl(rst!codiclientfactclixes) > 0 Then
           Set rstc = dbcomandes.OpenRecordset("select * from clients_codissap where codisap=" + atrim(rst!codiclientfactclixes))
           vmsg = vmsg + "Client clixes: " + atrim(rst!nomclienttemporal) + vbNewLine
           If Not rstc.EOF Then vmsg = vmsg + "Client on facturar: " + atrim(rstc!codisap) + "-" + atrim(rstc!nomclient) + vbNewLine
           vmsg = vmsg + "Preu pressupost: " + atrim(rst!Preu) + "€" + vbNewLine
           vmsg = vmsg + "Preu facturat: " + atrim(reixa.TextMatrix(reixa.Row, numcol("Preu_Clixes"))) + "€" + vbNewLine + vbNewLine
          End If
           vmsg = vmsg + "Es correcte?"
            
      End If
  End If
  missatge_revisióclixes = vmsg
  Set rst = Nothing
  Set rstc = Nothing
End Function
Function elclientescrops() As Boolean
  If InStr(1, UCase(reixa.TextMatrix(reixa.Row, numcol("Nom Client"))), "CROP´S") > 0 Then
       elclientescrops = True
  End If
End Function
Sub treurestandby(numc As Double)
   reixa.TextMatrix(reixa.Row, numcol("StandBy")) = ""
   dbcomandes.Execute "update comandes_extres set passaraimpresores=1 where comanda=" + atrim(numc)
   dbconsulta.Execute "update planificaciototes set standbyimpresio='' where comanda=" + atrim(numc)
   'si es treu d'standby comprovar si hi ha comanda amb el mateix treball a muntadora pendent per avisar
   comprovarsihihaamuntadoraeltreballiavisar cadbl(reixa.TextMatrix(reixa.Row, numcol("NºLot")))
End Sub
Sub borrarlacomandaabuscar()
  Dim i As Byte
  For i = 0 To filtre.Count - 1
    If camps(cadbl(filtre(i).tag), 1) = "comanda" Then filtre(i).Text = ""
  Next i
  
End Sub
Function actualitzarlareixa(comanda As Double)
   If taulaplanificacio = "planificaciototes" Then
      passarcanvisatemporals comanda
     Else:
        passarcanvisdeseccioatotes comanda
        passarcanvisatemporals comanda
   End If
   carregar_ordre_correcte nummaquina
   possarhoraprevista
'   configreixa
    
    reordenarregistres
    filtre_LostFocus 999
    buscarcomanda comanda
End Function
Sub canviardemaquinageneral(comanda As Double, seccio As String)
  Dim rstm As Recordset
  Dim maq As String
  Dim maquinanova As String
  Set rstm = dbcomandes.OpenRecordset("select * from maquines where maquina='" + seccio + "' and donadadebaixa =null")
  While Not rstm.EOF
        maq = maq + " [" + atrim(rstm!codi) + "] "
        rstm.MoveNext
  Wend
  Set rstm = Nothing
  maquinanova = InputBox("Entra la màquina que vols." + Chr(10) + " Opcions:  " + maq, "Canvi de màquina")
  If InStr(1, maq, "[" + maquinanova + "]") = 0 Then
       MsgBox "Aquesta màquina no existeix.", vbCritical, "Atenció": Exit Sub
  End If
  
  If seccio = "I" Then
      crearlanoreal "planificacioimp", comanda, cadbl(maquinanova)
      taulaplanificacio = "planificacioimp"
      canviarmaquinaalacomanda comanda, cadbl(maquinanova)
      dbconsulta.Execute "update planificaciototes set impresora=" + atrim(maquinanova) + " where comanda=" + atrim(comanda)
  End If
  If seccio = "L" Then
      crearlanoreal "planificaciolam", comanda, cadbl(maquinanova)
      taulaplanificacio = "planificaciolam"
      canviarmaquinaalacomanda comanda, cadbl(maquinanova)
      dbconsulta.Execute "update planificaciototes set laminadora=" + atrim(maquinanova) + " where comanda=" + atrim(comanda)
  End If
  If seccio = "R" Then
      crearlanoreal "planificacioreb", comanda, cadbl(maquinanova)
      taulaplanificacio = "planificacioreb"
      canviarmaquinaalacomanda comanda, cadbl(maquinanova)
      dbconsulta.Execute "update planificaciototes set rebobinadora=" + atrim(maquinanova) + " where comanda=" + atrim(comanda)
  End If
  If seccio = "S" Then
      crearlanoreal "planificaciosol", comanda, cadbl(maquinanova)
      taulaplanificacio = "planificaciosol"
      canviarmaquinaalacomanda comanda, cadbl(maquinanova)
      dbconsulta.Execute "update planificaciototes set soldadora=" + atrim(maquinanova) + " where comanda=" + atrim(comanda)
  End If
  
  taulaplanificacio = "planificaciototes"
  mgeneral_Click
End Sub
Sub crearlanoreal(taula As String, numc As Double, nummaq As Double)
   dbconsulta.Execute "insert into " + taula + " (ordre,comanda,maquina,tempsimpresio) values (999," + atrim(numc) + "," + atrim(nummaq) + ",180)"
   dbplanificacio.Execute "insert into " + taula + " (ordre,comanda,maquina) values (999," + atrim(numc) + "," + atrim(nummaq) + ")"
   passarcanvisatemporals numc
End Sub
Sub passarcanvisatemporals(comanda As Double)
   Dim rsttotes As Recordset
   Dim campsmodificats As String
   
   
   Set rsttotes = dbplanificacioalicia.OpenRecordset("select * from planificaciototes where comanda=" + atrim(comanda))
   If Not rsttotes.EOF Then
       campsmodificats = "observacions='" + atrim(rsttotes!observacio) + "',data1=" + IIf(IsDate(rsttotes!Data1), "#" + format(rsttotes!Data1, "mm/dd/yy") + "#", "Null") + ",data2=" + IIf(IsDate(rsttotes!data2), "#" + format(rsttotes!data2, "mm/dd/yy") + "#", "Null") + ","
       campsmodificats = campsmodificats + "importancia=" + atrim(cadbl(rsttotes!importancia))
       dbconsulta.Execute "update planificacioimp set " + campsmodificats + " where comanda=" + atrim(comanda)
       dbconsulta.Execute "update planificaciolam set " + campsmodificats + " where comanda=" + atrim(comanda)
       dbconsulta.Execute "update planificacioreb set " + campsmodificats + " where comanda=" + atrim(comanda)
       dbconsulta.Execute "update planificaciosol set " + campsmodificats + " where comanda=" + atrim(comanda)
   End If
End Sub
Sub passarcanvisdeseccioatotes(comanda As Double)
   Dim rsttotes As Recordset
   Dim campsmodificats As String
   
   
   Set rsttotes = dbconsulta.OpenRecordset("select * from " + taulaplanificacio + " where comanda=" + atrim(comanda))
   If Not rsttotes.EOF Then
       campsmodificats = "observacions='" + atrim(rsttotes!observacions) + "',data1=" + IIf(IsDate(rsttotes!Data1), "#" + format(rsttotes!Data1, "mm/dd/yy") + "#", "Null") + ",data2=" + IIf(IsDate(rsttotes!data2), "#" + format(rsttotes!data2, "mm/dd/yy") + "#", "Null") + ","
       campsmodificats = campsmodificats + "importancia=" + atrim(cadbl(rsttotes!importancia))
       dbconsulta.Execute "update planificaciototes set " + campsmodificats + " where comanda=" + atrim(comanda)
       dbconsulta.Execute "update planificaciototes set " + campsmodificats + " where comanda=" + atrim(comanda)
       'dbconsulta.Execute "update planificacioreb set " + campsmodificats + " where comanda=" + atrim(comanda)
       'dbconsulta.Execute "update planificaciosol set observacions='" + atrim(rsttotes!observacio) + "' where comanda=" + atrim(comanda)
   End If
End Sub


Private Sub reixa_LostFocus()


 postit.visible = False
 guardar_amples_reixa
End Sub
Function numfilaonestaelpunter(y As Single) As String
   Dim i As Integer
   Dim n As Double
   For i = 0 To reixa.Rows - 1
     If y > reixa.RowPos(i) Then n = i ' IIf(i = 0, 0, i - 1)
   Next i
   numfilaonestaelpunter = n
End Function

Function numcolumnaonestaelpunter(x As Single) As String
   Dim i As Byte
   Dim n As Double
   For i = 0 To reixa.Cols - 1
     If x > reixa.ColPos(i) Then n = i ' IIf(i = 0, 0, i - 1)
   Next i
   numcolumnaonestaelpunter = n
End Function
Private Sub reixa_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim numc As Double
  Dim v As String
  If taulaplanificacio = "planificacioent" And Button = 2 Then Exit Sub
  v = reixa.TextMatrix(reixa.Row, numcol("NºLot"))
  numc = cadbl(v)
  If InStr(1, v, "R") > 2 Then numc = cadbl(Mid(atrim(v), 1, InStr(1, v, "R") - 1))
  
 vnumcolumnaclickreixa = numcolumnaonestaelpunter(x)
 reixa.col = numcolumnaonestaelpunter(x): reixa.Row = numfilaonestaelpunter(y)
 If y < reixa.RowHeight(0) Then ordenarlareixa reixa.col
 If Button = 2 And vnumcolumnaclickreixa = numcol("Data_Expedició") And Not programaoperaris Then
     'Me.m_opcions  _reixa.WindowList = True
       Me.PopupMenu mexpedicio
  End If
 If Button = 2 And vnumcolumnaclickreixa = numcol("StandBy") And Not programaoperaris Then
     'Me.m_opcions_reixa.WindowList = True
       Me.PopupMenu mreixa1
  End If
  If Button = 2 And vnumcolumnaclickreixa = numcol("NºLinia Imp") And Not programaoperaris Then
      If reixa.TextMatrix(reixa.Row, numcol("Estat")) = "E" Then
       'If reixa.CellBackColor <> QBColor(15) And reixa.CellBackColor <> 0 Then
       Me.PopupMenu macceptarCdL, , x, y + reixa.Top + reixa.RowHeight(reixa.Row)
      End If
  End If
  If Button = 2 And vnumcolumnaclickreixa = numcol("Tipus_R") And numcol("Tipus_R") > 0 Then
     'Me.m_opcions_reixa.WindowList = True
       Me.PopupMenu menureclam
  End If
  If Button = 2 And vnumcolumnaclickreixa = numcol("NºLot") And Not programaoperaris And taulaplanificacio <> "planificacioent" Then
     'Me.m_opcions_reixa.WindowList = True
       If InStr(1, reixa.TextMatrix(reixa.Row, vnumcolumnaclickreixa), "Ra") > 0 Then
           If MsgBox("Aquesta comanda està Reclamada i reactivada" + Chr(10) + "Vols desactivar-la un altra cop?", vbInformation + vbDefaultButton2 + vbYesNo, "Desactivar") = vbYes Then
             dbbaixes.Execute "update planificacio_reclamades set reactivada=false where numcomanda=" + atrim(numc)
           End If
          Else
            If InStr(1, reixa.TextMatrix(reixa.Row, vnumcolumnaclickreixa), "R") > 0 Then
             If MsgBox("Aquesta comanda està Reclamada" + Chr(10) + "Vols reactivar-la?", vbInformation + vbDefaultButton2 + vbYesNo, "Reactivar") = vbYes Then
                  dbbaixes.Execute "update planificacio_reclamades set reactivada=true where numcomanda=" + atrim(numc)
                     Else
                        If MsgBox("Aquesta comanda està Reclamada" + Chr(10) + "Vols anular la reclamació?", vbInformation + vbDefaultButton2 + vbYesNo, "Reclamada") = vbYes Then
                          dbbaixes.Execute "delete * from planificacio_reclamades where numcomanda=" + atrim(numc)
                        End If
             End If
            End If
       End If
    borrarlacomandaabuscar
    actualitzarlareixa numc
  End If
End Sub



Private Sub reixa_RowColChange()
 If postit.visible And reixa.TextMatrix(0, reixa.col) = "Texte Impresio" Then
    reixa_Click
   Else
    postit.visible = False
 End If
 btarifaipressupost.visible = False
 If taulaplanificacio = "planificacioent" Then
     If reixa.TextMatrix(0, reixa.col) = "Preu_Albarà" Then
         btarifaipressupost.Left = reixa.CellLeft + reixa.CellWidth - btarifaipressupost.width + reixa.Left
         btarifaipressupost.Top = reixa.CellTop + reixa.Top
         btarifaipressupost.visible = True
     End If
 End If
End Sub

Private Sub smno_Click()
    Dim numc As Double
    Dim v As String
    If Screen.ActiveControl.Name <> "reixa" Then MsgBox "Escull una fila primer": Exit Sub
    v = reixa.TextMatrix(reixa.Row, numcol("NºLot"))
    numc = cadbl(v)
    If InStr(1, v, "R") > 2 Then numc = cadbl(Mid(atrim(v), 1, InStr(1, v, "R") - 1))
    If MsgBox("Vols treure la reclamació d'aquesta comanda?", vbInformation + vbYesNo + vbDefaultButton2, "Reclam de comanda") = vbYes Then
        reixa.TextMatrix(reixa.Row, numcol("Tipus_R")) = ""
        reixa.TextMatrix(reixa.Row, numcol("Data_R")) = ""
        dbconsulta.Execute "update reclamacionscomandes set tipusreclamacio='' where comanda=" + atrim(numc)
        dbconsulta.Execute "update reclamacionscomandes set datareclamacio=null where comanda=" + atrim(numc)
        dbplanificacioalicia.Execute "update reclamacionscomandes set tipusreclamacio='' where comanda=" + atrim(numc)
        dbplanificacioalicia.Execute "update reclamacionscomandes set datareclamacio=null where comanda=" + atrim(numc)
        dbplanificaciooperaris.Execute "update reclamacionscomandes set tipusreclamacio='' where comanda=" + atrim(numc)
        dbplanificaciooperaris.Execute "update reclamacionscomandes set datareclamacio=null where comanda=" + atrim(numc)
    End If
End Sub

Private Sub sortir_Click()
  tancartaules
 End
End Sub
Sub tancartaules()
  Unload Formagrupartreballs
  Unload formannex
  Set dbtmp = Nothing
  Set dbtmpb = Nothing
  Set dbcomandes = Nothing
  Set dbplanificacio = Nothing
  Set dbclixes = Nothing
  Set dbcompres = Nothing
  Set dbstocks = Nothing

End Sub

Private Sub Timer1_Timer()
  Dim diftemps As Long
  Static vcont As Long
  Dim vdatalocal As Date
  Dim vdataservidor As Date
  generareltemporalsical
  diftemps = DateDiff("n", ultimaactualitzacio, Now)
  If diftemps > 300 Then reixa.visible = False
  If IsDate(ultimaactualitzacio) And diftemps < 400 Then
     tempsproximrefresc = format(DateAdd("n", diftemps * -1, "05:00"), "hh:nn")
       Else: tempsproximrefresc = ""
  End If
  If vcont > 10 Then
     v = llegir_ini("Planificacio", "ultimaactualitzacio", "comandes.ini")
     If v = "{[}]" Or v = "" Then v = Now
     vdatalocal = CVDate(v)
     v = llegir_ini("Planificacio", "ultimaactualitzacio", rutadelfitxer(cami) + "\actualitzacioplanificacio.ini")
     If v = "" Or v = "{[}]" Then v = Now
     vdataservidor = CVDate(v)
     If DateDiff("s", vdatalocal, vdataservidor) > 0 Then
        Command1.tag = "novaactualitzacio"
        Command1.BackColor = QBColor(12)
        etultimaactualitzacio = format(vdataservidor, "dd/mm hh:nn")
     End If
     cont = 0
    Else: vcont = vcont + 1
  End If
End Sub

Private Sub Timer2_Timer()
  If vhihanCdLtaronja And eseccio = "General - Totes les seccions" Then
     If filtre(10).BackColor = &HFFC0FF Then
        filtre(10).BackColor = QBColor(14)
          Else: filtre(10).BackColor = &HFFC0FF
     End If
       Else: filtre(10).BackColor = &HFFC0FF
  End If
End Sub

Private Sub timercontrol_Timer()
  Static cont As Integer
  cont = cont + 1
  If cont = 350 Then End
End Sub

Private Sub treurefiltre_Click()
   borrarelfiltre
End Sub
Sub borrarelfiltre()
 configreixa False
  poblarlareixa nummaquina
  mgeneral.tag = ""
     mimpresores.tag = ""
     mlaminadores.tag = ""
     mrebobinadores.tag = ""
     Command3.tag = ""
  filtre_LostFocus 998
  triar_ordre_reixa
End Sub
Function proximdianatural() As Date
   proximdianatural = DateAdd("d", 1, Now)
   While format(proximdianatural, "w", vbMonday) > 5
      proximdianatural = DateAdd("d", 1, proximdianatural)
   Wend
End Function
Sub mentregaparcialtotal_click()
   Dim rst As Recordset
   Dim vcomanda As String
   vcomanda = cadbl(reixa.TextMatrix(reixa.Row, numcol("NºLot")))
   If vcomanda = 0 Then Exit Sub
   Set rst = dbplanificacio.OpenRecordset("select entregaparcial from planificaciototes where comanda=" + atrim(vcomanda))
   If rst!entregaparcial Then
        If MsgBox("Vols passar aquesta entrega a ENTREGA TOTAL?", vbQuestion + vbDefaultButton2 + vbYesNo, "ENTREGA TOTAL?") = vbNo Then Exit Sub
        dbplanificacio.Execute "update planificaciototes set entregaparcial=FALSE where comanda=" + atrim(vcomanda)
        reixa.CellBackColor = QBColor(15)
         Else
         If MsgBox("Vols passar aquesta entrega a ENTREGA PARCIAL?", vbQuestion + vbDefaultButton2 + vbYesNo, "ENTREGA PARCIAL?") = vbNo Then Exit Sub
         dbplanificacio.Execute "update planificaciototes set entregaparcial=TRUE where comanda=" + atrim(vcomanda)
         reixa.CellBackColor = QBColor(14)
   End If
   Set rst = Nothing
End Sub
Sub mnoenviarencara_click()
   Dim vr As String
   Dim vo As String
   Dim vcomanda As Double
   Dim vdataantiga As String
   If reixa.Text = "" Then Exit Sub
   If MsgBox("Segur que vols borra aquesta data d'Expedició?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   vo = reixa.TextMatrix(reixa.Row, numcol("NºLot"))
   If Len(vo) > 2 Then
    If Not IsNumeric(Mid(vo, Len(vo), 1)) Then vo = Mid(vo, 1, Len(vo) - 1)
   End If
   vcomanda = vo
   vdataantiga = reixa.TextMatrix(reixa.Row, numcol("Data_Expedició"))
   If cadbl(vcomanda) = 0 Then Exit Sub
   dbconsulta.Execute "update " + taulaplanificacio + " set dataexpedicio=null where comanda=" + atrim(vcomanda)
   dbplanificacio.Execute "update planificaciototes set dataexpedicio=null where comanda=" + atrim(vcomanda)
   reixa.Text = vr
   modificarexpedicio vcomanda, vdataantiga, "", "", ""

End Sub
Sub menviarja_click()
   Dim vr As String
   Dim vcomanda As Double
   Dim vdataantiga As String
   Dim vobs As String
   Dim vnomclient As String
   
   vr = reixa.TextMatrix(reixa.Row, numcol("NºLot"))
   If Len(vr) > 2 Then
    If Not IsNumeric(Mid(vr, Len(vr), 1)) Then vr = Mid(vr, 1, Len(vr) - 1)
   End If
   vcomanda = vr
   vdataantiga = reixa.TextMatrix(reixa.Row, numcol("Data_Expedició"))
   vobs = reixa.TextMatrix(reixa.Row, numcol("Obs_Expedició"))
   vnomclient = reixa.TextMatrix(reixa.Row, numcol("Nom Client"))
   If cadbl(vcomanda) = 0 Then Exit Sub
   vr = InputBox("Entra la data que vols que surti aquesta comanda.", "Data Expedició", format(proximdianatural, "dd/mm/yy"))
   If vr = "" Then Exit Sub
   If Not IsDate(vr) Then MsgBox "La data d'expedició no es vàlida", vbCritical, "Error": Exit Sub
   
   dbconsulta.Execute "update " + taulaplanificacio + " set dataexpedicio=#" + format(vr, "mm/dd/yy") + "# where comanda=" + atrim(vcomanda)
   dbplanificacio.Execute "update planificaciototes set dataexpedicio=#" + format(vr, "mm/dd/yy") + "# where comanda=" + atrim(vcomanda)
   reixa.Text = vr
   modificarexpedicio vcomanda, vdataantiga, vr, vobs, vnomclient
End Sub
Sub modificarexpedicio(vcomanda As Double, vdataantiga As String, vdatanova As String, vobs As String, vnomclient As String)
   Dim rst As Recordset
   dbplanificacio.Execute "delete * from linies_expedicions where comanda=" + atrim(vcomanda) + " and not enviat"
   If vdatanova = "" Then GoTo fi
   Set rst = dbplanificacio.OpenRecordset("select * from linies_expedicions where comanda=" + atrim(vcomanda) + " order by enviat desc ") '+ IIf(vdataantiga <> "", " and data=#" + format(vdataantiga, "mm/dd/yy") + "#", ""))
   If Not rst.EOF Then
       If rst!enviat Then
          If MsgBox("Aquesta comanda ja s'ha enviat anteriorment." + Chr(10) + " Vols possar una nova data d'enviament?", vbExclamation + vbDefaultButton2 + vbYesNo, "Comanda amb un enviament.") = vbNo Then GoTo fi
          
          GoTo afegirnou
       End If
       If rst!enviat = False Then
          If vdatanova <> "" Then
                rst.Edit
                rst!data = vdatanova
                rst!nomclient = vnomclient
                rst!observacio = vobs
                rst.Update
                 Else: rst.Delete
          End If
       End If
        Else
          If vdatanova <> "" Then
afegirnou:
            rst.AddNew
            rst!data = vdatanova
            rst!comanda = vcomanda
            rst!observacio = vobs
            rst!nomclient = vnomclient
            rst.Update
          End If
       End If
fi:
   
   Set rst = Nothing
End Sub

Sub sms1ja_click()
  Dim numc As Double
    Dim v As String
    If Screen.ActiveControl.Name <> "reixa" Then MsgBox "Escull una fila primer": Exit Sub
    v = reixa.TextMatrix(reixa.Row, numcol("NºLot"))
    numc = cadbl(v)
    If InStr(1, v, "R") > 2 Then numc = cadbl(Mid(atrim(v), 1, InStr(1, v, "R") - 1))
    If MsgBox("Vols reclamar aquesta comanda per Jà?", vbInformation + vbYesNo + vbDefaultButton2, "Reclam de comanda") = vbYes Then
        reixa.TextMatrix(reixa.Row, numcol("Tipus_R")) = "S1-Jà"
        reixa.TextMatrix(reixa.Row, numcol("Data_R")) = format(Now, "dd/mm hh:nn")
        dbconsulta.Execute "update reclamacionscomandes set tipusreclamacio='S1-Jà' where comanda=" + atrim(numc)
        dbconsulta.Execute "update reclamacionscomandes set datareclamacio=now where comanda=" + atrim(numc)
        dbplanificacioalicia.Execute "update reclamacionscomandes set tipusreclamacio='S1-Jà' where comanda=" + atrim(numc)
        dbplanificacioalicia.Execute "update reclamacionscomandes set datareclamacio=now where comanda=" + atrim(numc)
        dbplanificaciooperaris.Execute "update reclamacionscomandes set tipusreclamacio='S1-Jà' where comanda=" + atrim(numc)
        dbplanificaciooperaris.Execute "update reclamacionscomandes set datareclamacio=now where comanda=" + atrim(numc)
    End If
End Sub
Sub sms2_click()
Dim numc As Double
    Dim v As String
    If Screen.ActiveControl.Name <> "reixa" Then MsgBox "Escull una fila primer": Exit Sub
    v = reixa.TextMatrix(reixa.Row, numcol("NºLot"))
    numc = cadbl(v)
    If InStr(1, v, "R") > 2 Then numc = cadbl(Mid(atrim(v), 1, InStr(1, v, "R") - 1))
    If MsgBox("Vols reclamar aquesta comanda per d'aqui 12 Hores?", vbInformation + vbYesNo + vbDefaultButton2, "Reclam de comanda") = vbYes Then
        reixa.TextMatrix(reixa.Row, numcol("Tipus_R")) = "S2-12H"
        reixa.TextMatrix(reixa.Row, numcol("Data_R")) = format(Now, "dd/mm hh:nn")
        dbconsulta.Execute "update reclamacionscomandes set tipusreclamacio='S2-12H' where comanda=" + atrim(numc)
        dbconsulta.Execute "update reclamacionscomandes set datareclamacio=now where comanda=" + atrim(numc)
        dbplanificacioalicia.Execute "update reclamacionscomandes set tipusreclamacio='S2-12H' where comanda=" + atrim(numc)
        dbplanificacioalicia.Execute "update reclamacionscomandes set datareclamacio=now where comanda=" + atrim(numc)
        dbplanificaciooperaris.Execute "update reclamacionscomandes set tipusreclamacio='S2-12H' where comanda=" + atrim(numc)
        dbplanificaciooperaris.Execute "update reclamacionscomandes set datareclamacio=now where comanda=" + atrim(numc)
    End If

End Sub
Sub enviaravisreclamacio(vnumc As Double, vavis As String)
    enviaremailgeneric "liniesimpresio@inplacsa.com", "S'ha afegit la comanda " + atrim(vnumc) + " a reclamades de planificació", vavis
End Sub
Sub sms3_click()
Dim numc As Double
    Dim v As String
    If Screen.ActiveControl.Name <> "reixa" Then MsgBox "Escull una fila primer": Exit Sub
    v = reixa.TextMatrix(reixa.Row, numcol("NºLot"))
    numc = cadbl(v)
    If InStr(1, v, "R") > 2 Then numc = cadbl(Mid(atrim(v), 1, InStr(1, v, "R") - 1))
    If MsgBox("Vols reclamar aquesta comanda per d'aqui 24 Hores?", vbInformation + vbYesNo + vbDefaultButton2, "Reclam de comanda") = vbYes Then
        reixa.TextMatrix(reixa.Row, numcol("Tipus_R")) = "S3-24H"
        reixa.TextMatrix(reixa.Row, numcol("Data_R")) = format(Now, "dd/mm hh:nn")
        dbconsulta.Execute "update reclamacionscomandes set tipusreclamacio='S3-25H' where comanda=" + atrim(numc)
        dbconsulta.Execute "update reclamacionscomandes set datareclamacio=now where comanda=" + atrim(numc)
        dbplanificacioalicia.Execute "update reclamacionscomandes set tipusreclamacio='S3-25H' where comanda=" + atrim(numc)
        dbplanificacioalicia.Execute "update reclamacionscomandes set datareclamacio=now where comanda=" + atrim(numc)
        dbplanificaciooperaris.Execute "update reclamacionscomandes set tipusreclamacio='S3-25H' where comanda=" + atrim(numc)
        dbplanificaciooperaris.Execute "update reclamacionscomandes set datareclamacio=now where comanda=" + atrim(numc)
        enviaravisreclamacio numc, reixa.TextMatrix(reixa.Row, numcol("Tipus_R")) = "S3-24H"
    End If

End Sub
