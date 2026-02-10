VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desembolicar bobines i fer canutus."
   ClientHeight    =   11970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19200
   Icon            =   "Form desembolicar bobines.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11970
   ScaleWidth      =   19200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Framepassword 
      BackColor       =   &H00EAD9CE&
      Height          =   9165
      Left            =   11145
      TabIndex        =   32
      Top             =   10920
      Visible         =   0   'False
      Width           =   7230
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
         TabIndex        =   46
         Top             =   8070
         Width           =   5505
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
         TabIndex        =   45
         Top             =   945
         Width           =   1365
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
         TabIndex        =   44
         Top             =   4545
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
         TabIndex        =   43
         Top             =   2745
         Width           =   1770
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
         TabIndex        =   42
         Top             =   945
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
         TabIndex        =   41
         Top             =   6375
         Width           =   3630
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
         TabIndex        =   40
         Top             =   4545
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
         TabIndex        =   39
         Top             =   4545
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
         TabIndex        =   38
         Top             =   2745
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
         TabIndex        =   37
         Top             =   2745
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
         TabIndex        =   36
         Top             =   945
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
         TabIndex        =   35
         Top             =   945
         Width           =   1770
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
         TabIndex        =   34
         Top             =   6390
         Width           =   1770
      End
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
         Picture         =   "Form desembolicar bobines.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   8085
         Width           =   1275
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
         TabIndex        =   47
         Top             =   165
         Width           =   7020
      End
   End
   Begin VB.CommandButton botoensenyarpacking 
      BackColor       =   &H00FFFF00&
      Caption         =   "Command7"
      Height          =   195
      Left            =   -75
      TabIndex        =   26
      Top             =   1950
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Frame framellistattubos 
      Height          =   9765
      Left            =   3945
      TabIndex        =   15
      Top             =   3915
      Visible         =   0   'False
      Width           =   18645
      Begin VB.CommandButton Command9 
         BackColor       =   &H00989FF8&
         Caption         =   "Canutus Standard"
         Height          =   570
         Left            =   3465
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   1125
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF80FF&
         Caption         =   "Passar lot a NO FET"
         Height          =   270
         Left            =   16380
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   105
         Width           =   1680
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FF80FF&
         Caption         =   "Estat Comanda"
         Height          =   570
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   1125
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FDDECE&
         Caption         =   "Imprimir Etiquetes"
         Height          =   885
         Left            =   16335
         Picture         =   "Form desembolicar bobines.frx":1754
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Imprimir etiquetes pels tubs  i passar a TALLATS"
         Top             =   375
         Width           =   1725
      End
      Begin VB.CheckBox checknoactualitzar 
         Caption         =   "No actualitzar"
         Height          =   225
         Left            =   13980
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton btubstallats 
         BackColor       =   &H006BEBB1&
         Caption         =   "Tubs TALLATS"
         Height          =   570
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   180
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.CommandButton btubsnotallats 
         BackColor       =   &H00DCB8FC&
         Caption         =   "               ACTUALITZAR            (Tubs NO TALLATS)"
         Height          =   555
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   135
         Width           =   2265
      End
      Begin VB.CommandButton btubs 
         BackColor       =   &H00F1B75F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   3
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   735
         Visible         =   0   'False
         Width           =   2500
      End
      Begin VB.CommandButton btubs 
         BackColor       =   &H00F1B75F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   2
         Left            =   5490
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   735
         Visible         =   0   'False
         Width           =   2500
      End
      Begin VB.CommandButton btubs 
         BackColor       =   &H00F1B75F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   2955
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   735
         Visible         =   0   'False
         Width           =   2500
      End
      Begin VB.CommandButton btubs 
         BackColor       =   &H00F1B75F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   735
         Visible         =   0   'False
         Width           =   2500
      End
      Begin MSFlexGridLib.MSFlexGrid reixatubos 
         Height          =   8070
         Left            =   375
         TabIndex        =   17
         Top             =   1275
         Width           =   17715
         _ExtentX        =   31247
         _ExtentY        =   14235
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label etdataactualitzacio 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4665
         TabIndex        =   30
         Top             =   165
         Width           =   11880
      End
      Begin VB.Label etestat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1020
         Left            =   4650
         TabIndex        =   16
         Top             =   225
         Width           =   11565
      End
   End
   Begin VB.Frame Framebobinescomandes 
      Height          =   9780
      Left            =   105
      TabIndex        =   4
      Top             =   1935
      Width           =   18645
      Begin VB.Frame frameinfobobina 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Informació de la bobina seleccionada"
         Height          =   5175
         Left            =   10000
         TabIndex        =   6
         Top             =   1000
         Visible         =   0   'False
         Width           =   7605
         Begin VB.CommandButton Command6 
            Caption         =   "Treure de DESEMBOLICADA"
            Height          =   495
            Left            =   5775
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   4605
            Width           =   1725
         End
         Begin VB.Label etinformaciobobina 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4500
            Left            =   165
            TabIndex        =   7
            Top             =   390
            Width           =   7050
         End
      End
      Begin MSFlexGridLib.MSFlexGrid reixa 
         Height          =   9360
         Left            =   870
         TabIndex        =   5
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   16510
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frameordrecomandes 
         BackColor       =   &H00EAD9CE&
         Height          =   9315
         Left            =   5805
         TabIndex        =   8
         Top             =   225
         Width           =   12405
         Begin MSFlexGridLib.MSFlexGrid reixabobines 
            Height          =   9090
            Left            =   165
            TabIndex        =   9
            Top             =   150
            Width           =   3780
            _ExtentX        =   6668
            _ExtentY        =   16034
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipologia bobines: P-Parcial   J-Bobina entera    R-Restu"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4230
            TabIndex        =   50
            Top             =   480
            Width           =   5295
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "<----->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8865
            TabIndex        =   28
            Top             =   135
            Width           =   660
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "<----->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   5970
            TabIndex        =   27
            Top             =   135
            Width           =   660
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DESEMBOLICADA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   9585
            TabIndex        =   13
            Top             =   135
            Width           =   1830
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NO DESEMBOLICADA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   6660
            TabIndex        =   12
            Top             =   120
            Width           =   2145
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NO LAM o REB"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4230
            TabIndex        =   11
            Top             =   135
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   18615
      Begin VB.CommandButton Command10 
         Caption         =   "Escanejar Etiquetes"
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
         Left            =   14250
         Picture         =   "Form desembolicar bobines.frx":1A8F
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   210
         Width           =   2010
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Passar a SALA"
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
         Left            =   2265
         Picture         =   "Form desembolicar bobines.frx":1CD6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Escanejar les bobines per possar-les a SALA"
         Top             =   195
         Width           =   1950
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Llistat tubos"
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
         Left            =   16410
         Picture         =   "Form desembolicar bobines.frx":2463
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   195
         Width           =   2010
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pujar LAM"
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
         Left            =   4440
         Picture         =   "Form desembolicar bobines.frx":2CCC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   195
         Width           =   1950
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ordre Comandes"
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
         Left            =   105
         Picture         =   "Form desembolicar bobines.frx":33A5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   195
         Width           =   1950
      End
      Begin VB.Label etpaletasala 
         BackStyle       =   0  'Transparent
         Caption         =   "Ult. Palet a SALA: 45485/1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   960
         Left            =   6795
         TabIndex        =   29
         Top             =   525
         Visible         =   0   'False
         Width           =   7410
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arguments As Variant
Dim vultimbotoconsultatubos As Long
Dim vllistacomandes As String
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

Private Sub bcanutusstd_Click()

End Sub

Private Sub btubs_Click(Index As Integer)
   config_reixatubos
   If btubs(Index).Caption <> "" Then
    vultimbotoconsultatubos = Index
    carregar_dadesreixatubos btubs(Index).Tag, IIf(InStr(1, btubs(Index).Caption, "Cartró") > 0, "C", "P")
   End If
   If btubs(Index).Caption <> "" Then etestat.Caption = btubs(Index).Caption
End Sub

Private Sub btubsnotallats_Click()
   If checknoactualitzar.Value = 0 Then
      'llistat_tubos False
      llistat_tubos_servidor
      etestat.Caption = "Carregant les dades..."
      wait 2
   End If
    etestat.Caption = ""
    ensenyar_tubos
    'bcanutusstd.Enabled = True
    Command1.Enabled = True
    
    
    
End Sub
Sub llistat_tubos_servidor()
  escriure_ini "Llistattubos", "horainici", Now, rutadelfitxer(cami) + "valorsprograma.ini"
  etdataactualitzacio = ""
  etestat.Caption = "Seleccionant registres... AQUEST PROCÉS POT TRIGAR UNA ESTONA"
  ratoli "espera"
  DoEvents
  vinici = Now
  vtitol = etestat.Caption
  While DateDiff("s", vinici, Now) < 30 And llegir_ini("Llistattubos", "horainici", rutadelfitxer(cami) + "valorsprograma.ini") <> ""
     etestat.Caption = vtitol + " " + atrim(30 - DateDiff("s", vinici, Now)) + "s"
     DoEvents
  Wend
  etestat.Caption = ""
  ratoli "normal"
  If DateDiff("s", vinici, Now) >= 30 Then
      If MsgBox("Sembla que el proces d'actualització del servidor ha fallat, vols que la tablet el generi? aixó trigarà uns minuts.", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
          llistat_tubos False
      End If
  End If
End Sub
Private Sub btubstallats_Click()
   If checknoactualitzar.Value = 0 Then llistat_tubos False
    ensenyar_tubos
    bcanutusstd.Enabled = False
    Command1.Enabled = False
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

Private Sub Command1_Click()
    Dim printerx As Printer
    Dim vlot As String
    Dim vllarg As String
    Dim vdiametre As String
    Dim vquant As String
    If reixatubos.row = 0 And reixatubos.col = 0 Then Exit Sub
    For Each printerx In Printers
        If InStr(1, UCase(printerx.DeviceName), "RP-12N") > 0 Then Set Printer = printerx
    Next
    vlot = reixatubos.TextMatrix(reixatubos.row, 0)
    vquant = reixatubos.TextMatrix(reixatubos.row, 1)
    vllarg = reixatubos.TextMatrix(reixatubos.row, 2)
    vdiametre = Mid(atrim(Mid(etestat, 5)), 1, 4)
    
    If cadbl(vlot) = 0 Then Exit Sub
    Printer.FontSize = 20
    Printer.FontBold = True
    Printer.Font = "Courier New"
    Printer.Print "================="
    Printer.Print "LOT:" + justificar(vlot, 12, "D") + vbNewLine
    Printer.Print "QUANTITAT:" + justificar(vquant, 6, "D") + vbNewLine
    Printer.Print "LLARG:" + justificar(vllarg, 10, "D") + vbNewLine
    Printer.Print "DIÀMETRE:" + justificar(vdiametre, 8, "D") + vbNewLine
    Printer.Print vbNewLine + "================="
    Printer.EndDoc
    If MsgBox("Vols marcar aquest LOT com a canutus tallats?", vbExclamation + vbDefaultButton2 + vbYesNo, "CANUTUS TALLATS?") = vbYes Then
        dbbaixes.Execute "delete * from tmp_canutuspertallar where comanda=" + atrim(vlot)
        If reixatubos.Rows = 2 Then
             reixatubos.Rows = 1
              Else: reixatubos.RemoveItem reixatubos.row
        End If
        dbbaixes.Execute "delete * from canutusjatallats where comanda=" + atrim(vlot)
        dbbaixes.Execute "insert into canutusjatallats (comanda) values (" + atrim(vlot) + ")"
    End If
End Sub

Function justificar(v As String, longitut As Integer, Optional DoE As String) As String
    v = Mid(v, 1, longitut)
    If DoE <> "D" Then
       v = v + Space(longitut - Len(v))
      Else: v = Space(longitut - Len(v)) + v
    End If
    justificar = v
End Function

Private Sub Command10_Click()
   Formescanejaretiqueta.Show 1
End Sub

Private Sub Command14_Click()
  If Len(cpassword.Tag) = 0 Then Exit Sub
  cpassword.Tag = Mid(cpassword.Tag, 1, Len(cpassword.Tag) - 1)
  If Framepassword.Tag = "password" Then
     cpassword = Mid(cpassword, 1, Len(cpassword) - 1)
      Else: cpassword = Mid(cpassword, 1, Len(cpassword) - 1)
  End If
End Sub

Private Sub Command2_Click()
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim rst3 As Recordset
  Dim vrow As Long
  reixa.Tag = ""
  reixa.Font.Size = 18
  Framebobinescomandes.Visible = True
  framellistattubos.Visible = False
  frameinfobobina.Visible = False
  Frameordrecomandes.Visible = True
  vrow = 1
  reixa.Clear
  reixa.Rows = 1
  reixabobines.Clear
  DoEvents
  reixa.Redraw = False
  reixabobines.Rows = 1
  frameinfobobina.BackColor = &HEAD9CE
  reixa.Cols = 3
  reixa.ColWidth(0) = 900
  reixa.ColWidth(1) = 1400
  reixa.ColWidth(2) = 2000
  reixa.Width = 4885
  reixa.TextMatrix(0, 0) = "Maq.": reixa.TextMatrix(0, 1) = "Ordre": reixa.TextMatrix(0, 2) = "Comanda"
  reixabobines.TextMatrix(0, 0) = "PackingList"
  DoEvents
  vllistacomandes = ""
  Set rst2 = dbcomandes.OpenRecordset("select * from productes")
  Set rst = dbplanificacio.OpenRecordset("SELECT planificaciolam.*,comandes.producte, comandes.proximaseccio,comandes.linkcomanda1,comandes.linkcomanda2 FROM planificaciolam INNER JOIN comandes ON planificaciolam.comanda = comandes.comanda WHERE (((comandes.proximaseccio)='I' Or (comandes.proximaseccio)='L' Or (comandes.proximaseccio)='E') ) and ordre>0 and ordre<999 order by maquina,ordre")
  While Not rst.EOF
    reixa.Rows = vrow + 1
    reixa.TextMatrix(vrow, 0) = rst!maquina
    reixa.TextMatrix(vrow, 1) = rst!ordre
    reixa.TextMatrix(vrow, 2) = rst!comanda
    reixa.row = vrow
    For i = 0 To reixa.Cols - 2: reixa.col = i: reixa.CellBackColor = IIf(rst!maquina = 1, QBColor(13), QBColor(11)): Next i
    reixa.col = 2
    reixa.CellBackColor = color_comanda(rst!comanda)
    guardar_vllistacomandes rst2, Trim(rst!comanda) + "," + atrim(rst!linkcomanda1) + "," + atrim(rst!linkcomanda2)
    rst.MoveNext
    vrow = vrow + 1
  Wend
    'rebobinadores
  Set rst = dbplanificacio.OpenRecordset("SELECT planificacioreb.*, comandes.producte, comandes.proximaseccio, comandes.linkcomanda1, comandes.linkcomanda2, InStr(1,[ruta],'I') AS HihaIMP, productes.ruta FROM (planificacioreb INNER JOIN comandes ON planificacioreb.comanda = comandes.comanda) LEFT JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.proximaseccio)='I' Or (comandes.proximaseccio)='R' Or (comandes.proximaseccio)='E') AND ((InStr(1,[ruta],'I'))=0) AND ((planificacioreb.ordre)>0 And (planificacioreb.ordre)<999)) ORDER BY planificacioreb.maquina, planificacioreb.ordre;")
  While Not rst.EOF
    Set rst3 = dbstocks.OpenRecordset("SELECT bobines.*, Parcials.*, Parcials.idpalet, Parcials.idbobina, * FROM bobines INNER JOIN Parcials ON (bobines.Idbobina = Parcials.idbobina) AND (bobines.Idpalet = Parcials.idpalet) where metres>0 and comanda='" + Trim(rst!comanda) + "' order by PARCIALS.idpalet,parcials.idbobina")
    If Not rst3.EOF Then
        reixa.Rows = vrow + 1
        reixa.TextMatrix(vrow, 0) = rst!maquina
        reixa.TextMatrix(vrow, 1) = rst!ordre
        reixa.TextMatrix(vrow, 2) = "R" + atrim(rst!comanda)
        reixa.row = vrow
        For i = 0 To reixa.Cols - 2: reixa.col = i: reixa.CellBackColor = IIf(rst!maquina = 1, QBColor(9), QBColor(14)): Next i
        reixa.col = 2
        reixa.CellBackColor = color_comanda(rst!comanda)
        guardar_vllistacomandes rst2, Trim(rst!comanda) + "," + atrim(rst!linkcomanda1) + "," + atrim(rst!linkcomanda2)
    End If
    rst.MoveNext
    vrow = vrow + 1
  Wend
  Set rst = Nothing
  Set rst2 = Nothing
  Set rst3 = Nothing
  reixa.Redraw = True
  reixa.Tag = "ordrecomandes"
  reixa.row = reixa.Rows - 1
End Sub
Sub passar_seccioSiR_comagrups(vnomtmp As String, dbvnomtmp As Database)
   Dim vsubsql As String
   'Rebobinadora
   vsubsql = "SELECT id FROM productes RIGHT JOIN (Parcials_DBL LEFT JOIN comandes ON Parcials_DBL.comandaDBL = comandes.comanda) ON productes.codi = comandes.producte WHERE (((comandes.proximaseccio)<>'T' And (comandes.proximaseccio)<>'V') AND ((Parcials_DBL.utilitzada)=False) AND ((Mid([ruta]+' ',2,1))='R'));"
   dbtmp.Execute "SELECT * from parcials where metres>0 and parcials.id in (" + vsubsql + ")"
   
End Sub

Sub guardar_vllistacomandes(rstproductes As Recordset, vcomandes As String)
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select producte,comanda from comandes where comanda in (" + atrim(vcomandes) + ")")
  While Not rst.EOF
    rstproductes.FindFirst ("codi='" + atrim(rst!producte) + "'")
    If rstproductes.NoMatch Then GoTo proxima
    If InStr(1, rstproductes!ruta, "I") = 0 Then
         vllistacomandes = vllistacomandes + IIf(vllistacomandes = "", "", ",") + "'" + atrim(rst!comanda) + "'"
    End If
proxima:
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Private Sub Command3_Click()
   Dim rst As Recordset
   Dim rstp As Recordset
   Dim mtrs As Double
   Dim vrow As Long
   ratoli "espera"
   Framebobinescomandes.Visible = True
  framellistattubos.Visible = False
  Command6.Visible = False
   vrow = 0
   reixa.Clear
   reixa.FixedRows = 1
   reixa.Cols = 1
   reixa.ColWidth(0) = 3000
   reixa.Tag = "pujarlam"
   reixa.Font.Size = 30
   frameinfobobina.Visible = True
   frameinfobobina.Top = 400
   frameinfobobina.Left = 6000
   Frameordrecomandes.Visible = False
   reixa.TextMatrix(0, 0) = "  Palet/Bob"
   etinformaciobobina = "ACTUALITZANT BOBINES PER PUJAR DE LAM..." + vbNewLine + vbNewLine + vbNewLine + "AQUESTA CONSULTA TRIGA UNA ESTONA..."
   DoEvents
   Set rst = dbstocks.OpenRecordset("SELECT Parcials.Idpalet, Parcials.Idbobina, First(Bobines.Sit) AS ample, First(Bobines.disponible) AS metres,'' as sit, cdbl(first(parcials.comanda)) as comanda FROM Bobines INNER JOIN Parcials ON (Bobines.Idpalet = Parcials.idpalet) AND (Bobines.Idbobina = Parcials.idbobina) GROUP BY Parcials.Idpalet, Parcials.Idbobina HAVING (((First(Bobines.Sit)) Like '*LAM*') AND ((First(Bobines.disponible))>150));")
   While Not rst.EOF
    Set rstp = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(rst!idpalet) + " and idbobina=" + atrim(rst!idbobina) + " and utilitzada=false  ")
    If Not rstp.EOF Then GoTo proxima
    Set rstp = dbstocks.OpenRecordset("select * from palets where idpalet=" + atrim(rst!idpalet))
    If rstp.EOF Then GoTo proxima
    mtrs = bobinesdentrada.calcular_mtrsdispreals(rst!idpalet, rst!idbobina)
    If mtrs < 500 Then GoTo proxima
    vrow = vrow + 1
    reixa.Rows = vrow + 1
    reixa.TextMatrix(vrow, 0) = atrim(rst!idpalet) + "/" + atrim(rst!idbobina)
    
proxima:
    rst.MoveNext
   Wend
   etinformaciobobina = ""
   Set rst = Nothing
   Set rstp = Nothing
   ratoli "normal"
End Sub

Private Sub Command4_Click()
  Framebobinescomandes.Visible = False
  framellistattubos.Visible = True
  framellistattubos.Top = Framebobinescomandes.Top
  framellistattubos.Left = Framebobinescomandes.Left
  DoEvents
  'dbbaixes.Execute "delete * from  tmp_canutuspertallar where comanda not in (select comanda from canutusjatallats)"
  ensenyar_tubos
  btubs_Click 0
End Sub
Sub ensenyar_tubos()
  config_reixatubos
  possar_botons_tubsbase
  If Not btubs(0).Visible Then etestat.Caption = "NO HI HA TUBS PER TALLAR."
End Sub
Sub possar_botons_tubsbase()
  Dim rst As Recordset
  Dim i As Byte
  For i = 0 To btubs.Count - 1
      btubs(i).Visible = False
  Next i
  
  Set rst = dbbaixes.OpenRecordset("SELECT First(tmp_canutuspertallar.datapreu) AS dataconsulta,tmp_canutuspertallar.tubbase, tmp_canutuspertallar.seccioactual From tmp_canutuspertallar GROUP BY tmp_canutuspertallar.tubbase, tmp_canutuspertallar.seccioactual; ")
  i = 0
  If Not rst.EOF Then etdataactualitzacio = "Data actualització: " + Format(rst!dataconsulta, "dd/mm hh:nn")
  If Not rst.EOF Then rst.MoveLast: rst.MoveFirst
  If rst.RecordCount > 4 Then GoTo fi
  While Not rst.EOF
     btubs(i).Visible = True
     btubs(i).Caption = "Tub: " + Format(rst!tubbase, "#.0") + " " + IIf(rst!seccioactual = "C", "Cartró", "PVC")
     btubs(i).Tag = atrim(rst!tubbase)
     rst.MoveNext
     i = i + 1
  Wend
fi:
  Set rst = Nothing
End Sub
Sub carregar_dadesreixatubos(vtubbase As Double, vtipusmaterial As String)
  Dim rst As Recordset
  Dim vrow As Double
  Dim vnomcli As String
  Dim rstcli As Recordset
  Set rstcli = dbbaixes.OpenRecordset("select * from clients")
  Set rst = dbbaixes.OpenRecordset("select * from tmp_canutuspertallar where tubbase=" + atrim(passaradecimalpunt(atrim(vtubbase))) + " and seccioactual='" + atrim(vtipusmaterial) + "' order by amplereb desc")
  reixatubos.Redraw = False
  vrow = 1
  While Not rst.EOF
     If reixatubos.Rows = vrow Then reixatubos.Rows = reixatubos.Rows + 1: reixatubos.RowHeight(reixatubos.Rows - 1) = 800
     rstcli.FindFirst "codi=" + atrim(rst!client)
     If rstcli.NoMatch Then vnomcli = "" Else vnomcli = atrim(rstcli!nom)
     reixatubos.TextMatrix(vrow, 0) = atrim(rst!comanda)
     reixatubos.col = 1: reixatubos.row = vrow: reixatubos.CellFontBold = True: reixatubos.CellFontSize = 18: reixatubos.CellForeColor = QBColor(13)
     reixatubos.TextMatrix(vrow, 1) = atrim(rst!rebobinadora)
     reixatubos.col = 2: reixatubos.row = vrow: reixatubos.CellFontBold = True: reixatubos.CellFontSize = 18
     reixatubos.TextMatrix(vrow, 2) = atrim(rst!amplereb)
     reixatubos.TextMatrix(vrow, 3) = atrim(rst!proximaseccio)
     reixatubos.TextMatrix(vrow, 4) = atrim(vnomcli)
     reixatubos.TextMatrix(vrow, 5) = atrim(rst!simulteneitatreb)
     
     reixatubos.TextMatrix(vrow, 6) = atrim(rst!mtrslinbob)
     reixatubos.TextMatrix(vrow, 7) = atrim(rst!rebmtrs)
     reixatubos.TextMatrix(vrow, 8) = atrim(rst!dataentrega)
     rst.MoveNext
     vrow = vrow + 1
  Wend
  Set rst = Nothing
  agrupar_tubosiguals_ambcolors
  reixatubos.row = 0: reixatubos.col = 0
  reixatubos.Redraw = True
End Sub
Sub agrupar_tubosiguals_ambcolors()
  Dim i As Long
  Dim vcolor As Double
  i = 2
  reixatubos.col = 2
  vcolor = 7
  While i < reixatubos.Rows - 1
     If reixatubos.TextMatrix(i, 2) = reixatubos.TextMatrix(i - 1, 2) Then
         reixatubos.row = i - 1: reixatubos.CellBackColor = QBColor(vcolor)
         reixatubos.row = i: reixatubos.CellBackColor = QBColor(vcolor)
           Else: vcolor = IIf(vcolor = 7, 8, 7)
     End If
     i = i + 1
  Wend
End Sub

Sub config_reixatubos()
  reixatubos.Clear
  reixatubos.RowHeight(0) = 600
  reixatubos.Rows = 1
  reixatubos.Cols = 9
  reixatubos.ColWidth(0) = 1200
  reixatubos.ColWidth(1) = 1100
  reixatubos.ColWidth(2) = 1500
  reixatubos.ColWidth(3) = 900
  reixatubos.ColWidth(4) = 6000
  reixatubos.ColWidth(5) = 800
  reixatubos.ColWidth(6) = 1200
  reixatubos.ColWidth(7) = 1200
  reixatubos.ColWidth(8) = 1800
  reixatubos.TextMatrix(0, 0) = "Comanda"
  reixatubos.TextMatrix(0, 1) = "Q.Tubs"
  reixatubos.TextMatrix(0, 2) = "AmpleREB"
  reixatubos.TextMatrix(0, 3) = "Estat"
  reixatubos.TextMatrix(0, 4) = "Nom Client"
  reixatubos.TextMatrix(0, 5) = "Sim."
  reixatubos.TextMatrix(0, 6) = "Mtr.Bob"
  reixatubos.TextMatrix(0, 7) = "Mtrs/Tot"
  reixatubos.TextMatrix(0, 8) = "DataEntrega"
  reixatubos.ColAlignment(2) = 3
  reixatubos.ColAlignment(1) = 3
End Sub
Sub borrartaulatmp_canutuspertallar()
  dbbaixes.Execute "delete * from tmp_canutuspertallar"
End Sub

Sub llistat_tubos(vtallats As Boolean)
   Dim sql As String
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim rst3 As Recordset
   Dim vcomandaanterior As Double
   ratoli "espera"
   For i = 0 To btubs.Count - 1
     btubs(i).Visible = False
   Next i
   borrartaulatmp_canutuspertallar
   etestat.Caption = "Seleccionant registres... AQUEST PROCÉS TRIGAR UNA ESTONA"
      DoEvents
   wait 2
   sql = "insert INTO tmp_canutuspertallar SELECT comandes.* FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.proximaseccio)<>'E' And (comandes.proximaseccio)<>'V' And (comandes.proximaseccio)<>'P' And (comandes.proximaseccio)<>'T' And (comandes.proximaseccio)<>'I' And (comandes.proximaseccio)<>'S') AND ((InStr(1,[productes].[ruta],'R'))>0)) OR (((comandes.proximaseccio)='I') AND ((InStr(1,[productes].[ruta],'R'))>0) AND ((comandes.comanda) In (select comanda from muntadora_ordremuntatge))) OR (((comandes.proximaseccio)='I') AND ((InStr(1,[productes].[ruta],'R'))>0) AND ((comandes.comanda) In (select comanda from muntadoratot where acabada)));"
   dbbaixes.Execute sql
   
   Set rst = dbbaixes.OpenRecordset("select * from comandes where proximaseccio<>'T' and dataactivacio<>null")
   While Not rst.EOF
      Set rst3 = dbbaixes.OpenRecordset("select tipusmaterialcanutureb from comandes_extres where comanda=" + atrim(rst!comanda))
      Set rst2 = dbbaixes.OpenRecordset("select InStr(1,[productes].[ruta],'R') as tereb from productes where codi='" + atrim(rst!producte) + "'")
      If Not rst2.EOF Then
        If rst2!tereb > 0 And atrim(rst3!tipusmaterialcanutureb) = "P" Then
           dbbaixes.Execute "insert into tmp_canutuspertallar select * from comandes where comanda=" + atrim(rst!comanda)
        End If
      End If
      rst.MoveNext
   Wend
   'On Error GoTo fi
   dbbaixes.Execute "delete * from  canutusjatallats where comanda in (SELECT comandes.comanda FROM comandes RIGHT JOIN canutusjatallats ON comandes.comanda = canutusjatallats.comanda WHERE (((comandes.proximaseccio)='T')))"
   etestat.Caption = "Preparant el llistat..."
      DoEvents
  ' wait 2
   
   'dbbaixesb.Execute "update tmp_canutuspertallar set seccioactual='*' where comanda in (select comanda from canutusjatallats)"
   If Not vtallats Then
       dbbaixes.Execute "delete * from  tmp_canutuspertallar where comanda in (select comanda from canutusjatallats)"
         Else: dbbaixes.Execute "delete * from  tmp_canutuspertallar where comanda not in (select comanda from canutusjatallats)"
   End If
   dbbaixes.Execute "delete * From tmp_canutuspertallar WHERE (((tmp_canutuspertallar.tubbase) Is Null)) OR (((tmp_canutuspertallar.tubbase)=0))"
   dbbaixes.Execute "delete * from tmp_canutuspertallar as t1 where amplereb in (select ample_Canutu from canutusestandard where mida_canutu=t1.tubbase)"
   Set rst = dbbaixes.OpenRecordset("select * from tmp_canutuspertallar order by comanda")
   dbbaixes.Execute "update tmp_canutuspertallar set seccioactual=''" 'borro tot el continut de seccioactual per mes avall utilitzarlo per passar el material del canutu
   While Not rst.EOF
     If vcomandaanterior = rst!comanda Then
       rst.Delete
       GoTo cont
     End If
     Set rst2 = dbbaixes.OpenRecordset("select tipusmaterialcanutureb from comandes_Extres where comanda=" + atrim(rst!comanda))
     rst.Edit
     'aquest dos camps son de la taula temporal no de la PRINCIPAL'aprofito el camp seccioactual per guardar el tipus de canutu PVC o Cartró
     rst!seccioactual = atrim(rst2!tipusmaterialcanutureb)
     rst!rebobinadora = calcularcanutosnecessaris(rst)
     rst!datapreu = Now 'posso el datapreu amb data actual per utilitzar-ho per saber l'hora d'actualització
     rst.Update
     vcomandaanterior = rst!comanda
cont:
     rst.MoveNext
   Wend
   etestat.Caption = ""
  ' wait 4
   DoEvents
  ' dbbaixes.Execute "update tmp_canutuspertallar set datapreu=now" 'posso el datapreu amb data actual per utilitzar-ho per saber l'hora d'actualització
fi:
   ratoli "normal"
   Set rst = Nothing
   Set rst2 = Nothing
   Set rst3 = Nothing
 End Sub
 Function calcularcanutosnecessaris(rst As Recordset) As Integer
  Dim vmetresdelabobina As Double
  Dim vmicres As Double
  
  vmetresdelabobina = cadbl(rst!mtrslinbob)
  'si no hi ha metres per la bobina calculo sobre diametre de 50cm
  If vmetresdelabobina = 0 Then
    vmicres = espesordelmaterial(rst)
    vmetresdelabobina = calcular_diametre(50, cadbl(rst!tubbase), vmicres)
  End If
  
  If vmetresdelabobina > 0 Then
   If ((cadbl(rst!rebmtrs) / vmetresdelabobina) - Int((cadbl(rst!rebmtrs) / vmetresdelabobina))) * 10 > 0 Then
      calcularcanutosnecessaris = Redondejar(((cadbl(rst!rebmtrs) / vmetresdelabobina) + 1) / cadbl(rst!simulteneitatreb), 0) * cadbl(rst!simulteneitatreb)
     Else: calcularcanutosnecessaris = Redondejar((cadbl(rst!rebmtrs) / vmetresdelabobina) / cadbl(rst!simulteneitatreb), 0) * cadbl(rst!simulteneitatreb)
   End If
     Else
  End If

 End Function
 Function micresmaterial(codimesuralineal As Byte, espesor As Double, tubolam As String) As Double
  Dim rstmesural As Recordset
  Set rstmesural = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(codimesuralineal))
  r = ""
  If rstmesural.EOF Then Exit Function
  r = espesor
  If rstmesural!descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = Redondejar(espesor / 4, 0)
                  Else: r = Redondejar(espesor / 2, 0)
            End If
  End If
  If InStr(1, rstmesural!descripcio, "GR/") > 0 Then
    r = espesor * -1
  End If
  micresmaterial = r
End Function
 
 Function calcular_diametre(diametreext As Double, canutu As Double, micres As Double) As Double
    Dim metres As Double
    diametreext = diametreext * 10 ' paso a milimetres
    canutu = canutu * 10 'paso a milimetres
    'calcul
    metres = ((diametreext * diametreext) - (canutu * canutu) / micres) * 0.746
    calcular_diametre = Redondejar(metres, 0)
 End Function


 
 Function espesordelmaterial(rstc As Recordset) As Double
  Dim rst As Recordset
  Dim vespesor As Double
  Dim vmesura As String
  Set rst = dbtmp.OpenRecordset("select comanda,mesuraesp,espessor,tubolam from comandes where comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2))
  While Not rst.EOF
    If rst!comanda > 0 Then
      vespesor = vespesor + micresmaterial(cadbl(rst!mesuraesp), rst!espessor, rst!tubolam)
    End If
    rst.MoveNext
  Wend
  espesordelmaterial = vespesor
End Function
 

Private Sub Command5_Click()
   Dim v As String
   Dim vpalet As String
   Dim vbobina As String
   etpaletasala = ""
   etpaletasala.Visible = True
   Framebobinescomandes.Visible = True
  framellistattubos.Visible = False
   v = " "
   While v <> ""
    v = InputBox("Escanejar bobina per passar-la a SALA LAM o REB.", "Passar a SALA")
    If atrim(v) = "" Then GoTo fi
    v = substituir(v, "-", "/")
    'If MsgBox("Bobina: " + v + vbNewLine + "   Ès correcte?", vbYesNo + vbDefaultButton2 + vbInformation, "Bobina") = vbNo Then GoTo fi
    separarpaletibobina v, vpalet, vbobina
    etpaletasala = "Ult. Palet a SALA: " + v
    dbbaixes.Execute "insert into laminadores_bobinesdesembolicades (idpalet,idbobina) values (" + atrim(vpalet) + "," + atrim(vbobina) + ")"
   Wend
fi:
   If cadbl(vpalet) > 0 Then Command2_Click
   etpaletasala.Visible = False
   etpaletasala = ""
End Sub

Private Sub Command6_Click()
   Dim v As String
   Dim vpalet As String
   Dim vbobina As String
   Dim vpos As Long
   
   v = reixabobines.Text
   If atrim(v) = "" Then Exit Sub
   separarpaletibobina v, vpalet, vbobina
   dbbaixes.Execute "delete * from  laminadores_bobinesdesembolicades where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbobina)
   vpos = reixabobines.row
   reixa_RowColChange
   reixabobines.row = vpos
End Sub
Function demanar_comanda() As String
   cpassword.Tag = ""
       cpassword = ""
       etmissatgepassword = "Escriu el numero de comanda:"
       etmissatgepassword.FontSize = 24
       Framepassword.Tag = ""
       Framepassword.Visible = True
       Framepassword.Top = 600
       Framepassword.Left = 6000
       While Framepassword.Visible
         DoEvents
       Wend
       v = atrim(cpassword.Tag)
       Framepassword.Visible = False
       demanar_comanda = v
End Function
Private Sub Command7_Click()
   Dim v As String
   Dim rst As Recordset
   v = demanar_comanda
   If cadbl(v) = 0 Then Exit Sub
   Set rst = dbcomandes.OpenRecordset("Select proximaseccio from comandes where comanda=" + atrim(v))
   If Not rst.EOF Then
       Load Formmsgbox
       Formmsgbox.etmissatge = "La comanda " + v + " està actualment en " + vbNewLine + " estat de " + atrim(evaluar_situacio_estat(rst!proximaseccio))
       Formmsgbox.Show 1
   End If
End Sub
Function evaluar_situacio_estat(vestat As String) As String
   evaluar_situacio_estat = "(" + vestat + ") EN PRODUCCIÓ"
   If vestat = "T" Or vestat = "V" Then evaluar_situacio_estat = "(" + vestat + ") ENTREGADA"
   If vestat = "P" Then evaluar_situacio_estat = "(" + vestat + ") PARCIALMENT ENTREGADA"
End Function

Private Sub Command8_Click()
  Dim vnumc As Double
  vnumc = cadbl(demanar_comanda)
  If vnumc = 0 Then Exit Sub
  If MsgBox("Estas segur que vols passar la comanda " + atrim(vnumc) + " a TUBOS NO FETS?", vbExclamation + vbDefaultButton2 + vbYesNo, "A T E N C I Ó") = vbYes Then
      dbbaixes.Execute "delete * from canutusjatallats where comanda=" + atrim(vnumc)
      MsgBox "COMANDA PASSADA A TUBOS NO FETS PERQUÈ ET SURTI A LA LLISTA HAS DE ACTUALITZAR LA LLISTA DE TUBS NO FETS.", vbInformation, "TUBS NO FETS"
  End If
End Sub

Private Sub Command9_Click()
   Dim v As String
   Dim rst As Recordset
   v = demanar_comanda
   If cadbl(v) = 0 Then Exit Sub
   dbbaixes.Execute "delete * from canutusjatallats where comanda=" + atrim(cadbl(v))
   dbbaixes.Execute "insert into canutusjatallats (comanda,agafarstd) values (" + atrim(v) + ",True)"
   MsgBox "La Comanda " + atrim(v) + " marcada com a canutus Tallats agafar Standards.", vbInformation, "Info"
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
 Static vdins As Boolean
 'If existeix("c:\ordprog.ini") And Not vdins Then vdins = True: Formescanejaretiqueta.Show 1
End Sub

Private Sub Form_Load()
Dim vinici As Date
assignardecimalipunt
arguments = ObtenerLíneaComando
fitxerini = "comandes.ini"
  cami = llegir_ini("General", "cami", fitxerini)
  camistock = rutadelfitxer(cami) + "palets.mdb"
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  If existeix("c:\ordprog.ini") Then
     cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
  End If
  centerscreen Me
  Set dbcomandes = DBEngine.OpenDatabase(cami)
  Set dbtmp = DBEngine.OpenDatabase(cami) 'la utilitzo per bobinesdentrada que fa servir aquesta BD
  Set dbplanificacio = DBEngine.OpenDatabase(rutadelfitxer(cami) + "planificaciooperaris.mdb")
  Set dbstocks = DBEngine.OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
  Set dbbaixes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
 If LCase(atrim(arguments(1))) = "escanejarbobines" Then
       Formescanejaretiqueta.Tag = "escanejarbobines"
       Formescanejaretiqueta.Show 1
       End
 End If
 If atrim(arguments(1)) = "llistattubos" Then
   llistat_tubos False
   escriure_ini "Llistattubos", "horainici", "", rutadelfitxer(cami) + "valorsprograma.ini"
   End
 End If
End Sub

Private Sub reixa_Click()
  If reixa.Tag = "pujarlam" Then
      informaciodelabobina reixa.Text
  End If
 
End Sub
Sub carregar_toteslesbobines(vrst As Recordset)
   Dim vsql As String
   If vllistacomandes = "" Then vllistacomandes = "0"
   vsql = "SELECT bobines.*, Parcials.*, Parcials.idpalet, Parcials.idbobina, * FROM bobines INNER JOIN Parcials ON (bobines.Idbobina = Parcials.idbobina) AND (bobines.Idpalet = Parcials.idpalet) where metres>0 and comanda in (" + Trim(vllistacomandes) + ") order by PARCIALS.idpalet,parcials.idbobina"
   Set vrst = dbstocks.OpenRecordset(vsql)
End Sub
Sub carregar_bobinespackinglist(vnumc As Double)
   Dim rst As Recordset
   Dim rstcom As Recordset
   Dim vrow As Double
   Dim vrowactual As Double
   Dim rstbobsasala As Recordset
   Dim rsttoteslesbobines As Recordset
   Dim vtipusb As Byte
   
   reixabobines.Visible = True
   vrowactual = reixa.row
   reixabobines.Clear
   reixabobines.TextMatrix(0, 0) = "PackingList"
   reixabobines.Redraw = False
   reixabobines.Rows = 1
   reixabobines.Cols = 5
   carregar_toteslesbobines rsttoteslesbobines
   reixabobines.col = 0 'ordena per tipus de bobina
   reixabobines.Sort = 0
   Set rstbobsasala = dbbaixes.OpenRecordset("select * from laminadores_bobinesdesembolicades")
   Set rst = dbstocks.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(vnumc))
   If rst.EOF Then Exit Sub
   Set rstcom = dbstocks.OpenRecordset("SELECT comandes.comanda, productes.ruta FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi where (comanda=" + atrim(rst!comanda) + " or comanda=" + atrim(rst!linkcomanda1) + " or comanda=" + atrim(rst!linkcomanda2) + ") order by comanda")
   i = 10
   While Not rstcom.EOF
      If InStr(1, rstcom!ruta, "I") = 0 Then
        Set rst = dbstocks.OpenRecordset("SELECT bobines.*, Parcials.*, Parcials.idpalet, Parcials.idbobina, * FROM bobines INNER JOIN Parcials ON (bobines.Idbobina = Parcials.idbobina) AND (bobines.Idpalet = Parcials.idpalet) where metres>0 and comanda='" + Trim(rstcom!comanda) + "' order by PARCIALS.idpalet,parcials.idbobina")
        While Not rst.EOF
           rstbobsasala.FindFirst "idbobina=" + atrim(rst![parcials.idbobina]) + " and idpalet=" + atrim(rst![parcials.idpalet])
           If cadbl(rst![parcials.idbobina]) > 0 Then
            reixabobines.Rows = reixabobines.Rows + 1
            reixabobines.row = reixabobines.Rows - 1
            reixabobines.Text = Trim(rst![parcials.idpalet]) + "/" + Trim(rst![parcials.idbobina])
              'faig lo de i+ per poder ordenar diferent per cada lot, sino agrupa parcials i restos de lots diferents
            vtipusb = tipusbobina(cadbl(rst![parcials.idpalet]), cadbl(rst![parcials.idbobina]))
            reixabobines.TextMatrix(reixabobines.row, 1) = IIf(vtipusb = 1, "R", IIf(vtipusb = 3, "P", "J"))
            reixabobines.TextMatrix(reixabobines.row, 2) = cadbl(i + vtipusb)
            reixabobines.TextMatrix(reixabobines.row, 3) = Trim(rst![parcials.idpalet])
            reixabobines.TextMatrix(reixabobines.row, 4) = Trim(rst![parcials.idbobina])
            reixabobines.col = 0
           End If
           If Not rstbobsasala.NoMatch Then
                reixabobines.CellBackColor = QBColor(10)
                   Else: If rst!sit = "LAM" Then reixabobines.CellBackColor = QBColor(14) Else reixabobines.CellBackColor = QBColor(12)
           End If
           rsttoteslesbobines.FindFirst "comanda<>'" + atrim(rstcom!comanda) + "' and bobines.idbobina=" + atrim(rst![parcials.idbobina]) + " and bobines.idpalet=" + atrim(rst![parcials.idpalet])
           If Not rsttoteslesbobines.NoMatch Then reixabobines.CellFontUnderline = True
           rst.MoveNext
        Wend
        i = i + 10
      End If
      rstcom.MoveNext
      
   Wend
fi:
   reixabobines.col = 2 'ordena per tipus de bobina
   reixabobines.ColSel = 4
   reixabobines.Sort = flexSortGenericAscending
   reixabobines.Redraw = True
   reixabobines.ColWidth(1) = 500
   reixabobines.ColWidth(2) = 0
   reixabobines.ColWidth(3) = 0
   reixabobines.ColWidth(4) = 0
   Set rst = Nothing
   Set rstcom = Nothing
End Sub
Function tipusbobina(vpalet As Double, vbobina As Double) As Double
   Dim vt As Double
   vt = 2
   If bobinesdentrada.esrestu(vpalet, vbobina) Then vt = 1
   If bobinesdentrada.esparcial(vpalet, vbobina) Then vt = 3
   tipusbobina = vt
End Function
Function color_comanda(vnumc As Double) As Double
   Dim rst As Recordset
   Dim rstcom As Recordset
   Dim vrow As Double
   Dim vrowactual As Double
   Dim rstbobsasala As Recordset
   Dim vcolor As Double
   Dim valgungroc As Boolean
   Dim valgunaverda As Boolean
   
   Set rstbobsasala = dbbaixes.OpenRecordset("select * from laminadores_bobinesdesembolicades")
   Set rst = dbstocks.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(vnumc))
   If rst.EOF Then GoTo fi
   If rst!linkcomanda1 = 0 And rst!linkcomanda2 = 0 Then
        Set rstcom = dbstocks.OpenRecordset("select comanda from comandes where comanda=" + atrim(rst!comanda) + " order by comanda")
       Else
         Set rstcom = dbstocks.OpenRecordset("select comanda from comandes where (producte='PC' or producte='PC2') and (comanda=" + atrim(rst!comanda) + " or comanda=" + atrim(rst!linkcomanda1) + " or comanda=" + atrim(rst!linkcomanda2) + ") order by comanda")
   End If
   
   While Not rstcom.EOF
        Set rst = dbstocks.OpenRecordset("SELECT bobines.*, Parcials.*, Parcials.idpalet, Parcials.idbobina, * FROM bobines INNER JOIN Parcials ON (bobines.Idbobina = Parcials.idbobina) AND (bobines.Idpalet = Parcials.idpalet) where metres>0 and comanda='" + Trim(rstcom!comanda) + "' order by PARCIALS.idpalet,parcials.idbobina")
        vcolor = QBColor(10) 'verd
        While Not rst.EOF
            rstbobsasala.FindFirst "idbobina=" + atrim(rst![parcials.idbobina]) + " and idpalet=" + atrim(rst![parcials.idpalet])
            If rstbobsasala.NoMatch Then
                    vcolor = QBColor(12) 'vermell
                    Else: valgunaverda = True
            End If
            If (rst!sit = "LAM" Or rst!sit = "REB") And rstbobsasala.NoMatch Then vcolor = QBColor(14): valgungroc = True 'groc
            If vcolor = QBColor(12) Then GoTo fi  ' si es vermell ja surt directament
            rst.MoveNext
        Wend
        
        rstcom.MoveNext
   Wend
fi:
   If vcolor = QBColor(12) And valgungroc Then vcolor = &H80FF& 'taronja
   If vcolor = QBColor(14) And valgunaverda Then vcolor = &HFFFF00               'verd clarissim
   Set rst = Nothing
   Set rstcom = Nothing
   color_comanda = vcolor
End Function

Sub separarpaletibobina(vnumbob As String, vpalet As String, vbob As String)

    If vnumbob = "" Then Exit Sub
    If InStr(1, vnumbob, "/") = 0 Then Exit Sub
    vpalet = cadbl(Mid(vnumbob, 1, InStr(1, vnumbob, "/") - 1))
    vbob = cadbl(substituirtot(vnumbob, vpalet + "/", ""))
End Sub

Sub informaciodelabobina(vbobina As String)
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim vpalet As String
  Dim vmetres As Double
  Dim vbob As String
  Dim vsumametres As Double
  Dim vcomandes As String
  Dim rstmat As Recordset
  Dim vdescmaterial As String
'  cbobina = "55129/10"
  etinformaciobobina = ""
  separarpaletibobina vbobina, vpalet, vbob
  Command6.Visible = False
  If cadbl(vpalet) = 0 Then Exit Sub
  Set rst = dbstocks.OpenRecordset("SELECT bobines.*, Palets.*, * FROM Palets LEFT JOIN bobines ON Palets.Idpalet = bobines.Idpalet Where bobines.idpalet = " + vpalet + " And bobines.idbobina = " + vbob)
  If Not rst.EOF Then
      Set rstmat = dbstocks.OpenRecordset("select * from materials where codi=" + atrim(rst!codimatprognou))
      vsumametres = 0
      vcomandes = ""
      vdescmaterial = atrim(rstmat!descripcio) + Chr(13) + Chr(10) + "Ample:" + atrim(rst!ample) + " Esp:" + atrim(IIf(cadbl(rst!micres) > 0, rst!micres, cadbl(rst!grmsm2)))
      vmetres = rst![bobines.disponible]
      'Set rst2 = dbcomandes.OpenRecordset("SELECT Parcials.idpalet, Parcials.idbobina, Sum(Parcials.metres) AS SumaDemetres From Parcials Where (((Parcials.utilitzada) = False)) GROUP BY Parcials.idpalet, Parcials.idbobina HAVING (((Parcials.idpalet)=" + vpalet + ") AND ((Parcials.idbobina)=" + vbob + "))")
      Set rst2 = dbstocks.OpenRecordset("SELECT Parcials.idpalet, Parcials.idbobina,parcials.comanda, Parcials.metres From Parcials Where (((Parcials.utilitzada) = False)) and  (((Parcials.idpalet)=" + vpalet + ") AND ((Parcials.idbobina)=" + vbob + "))")
      
      While Not rst2.EOF
         vsumametres = vsumametres + cadbl(rst2!metres)
         vcomandes = vcomandes + " [" + atrim(rst2!comanda) + "]"
         rst2.MoveNext
      Wend
      vmetres = vmetres + vsumametres
      vdiametre = calculardiametre(IIf(cadbl(rst!micres) > 0, rst!micres, cadbl(rst!grmsm2)), vmetres, rst!tamanycanutu)
       If vcomandes = "" Then vcomandes = "Cap."
      etinformaciobobina = "Informació de la bobina " + vbobina + vbNewLine + " Canutu: " + atrim(rst!tamanycanutu) + vbNewLine + "té " + Trim(vmetres) + " metres (Ø" + atrim(vdiametre) + "mm)" + vbNewLine + vdescmaterial + vbNewLine + "Comandes assignades: " + vbNewLine + vcomandes
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

Private Sub reixa_RowColChange()
  Dim vnumc As Double
  If reixa.Tag = "ordrecomandes" Then
     vnumc = cadbl(reixa.TextMatrix(reixa.row, 2))
     If vnumc = 0 Then vnumc = cadbl(Mid(reixa.TextMatrix(reixa.row, 2) + " ", 2))
      carregar_bobinespackinglist vnumc
  End If
End Sub

Private Sub reixabobines_Click()
  frameinfobobina.Top = 1000
  frameinfobobina.Left = 10000
  frameinfobobina.Visible = True
  frameinfobobina.BackColor = reixabobines.CellBackColor
  informaciodelabobina reixabobines.Text
  Command6.Visible = False
  If reixabobines.CellBackColor = QBColor(10) Then Command6.Visible = True
End Sub

