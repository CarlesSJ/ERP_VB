VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Treballs Dibuix"
      Height          =   7080
      Left            =   150
      TabIndex        =   8
      Top             =   750
      Width           =   10455
      Begin VB.Frame Frame5 
         Caption         =   "Tintes"
         Height          =   2340
         Left            =   300
         TabIndex        =   57
         Top             =   4650
         Width           =   8940
         Begin VB.TextBox txtFields 
            DataField       =   "t1"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   24
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   66
            Top             =   225
            Width           =   2400
         End
         Begin VB.TextBox txtFields 
            DataField       =   "t2"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   25
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   65
            Top             =   465
            Width           =   2400
         End
         Begin VB.TextBox txtFields 
            DataField       =   "t3"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   26
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   64
            Top             =   720
            Width           =   2400
         End
         Begin VB.TextBox txtFields 
            DataField       =   "t4"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   27
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   63
            Top             =   960
            Width           =   2400
         End
         Begin VB.TextBox txtFields 
            DataField       =   "t5"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   28
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   62
            Top             =   1200
            Width           =   2400
         End
         Begin VB.TextBox txtFields 
            DataField       =   "t6"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   29
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   61
            Top             =   1455
            Width           =   2400
         End
         Begin VB.TextBox txtFields 
            DataField       =   "t7"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   30
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   60
            Top             =   1695
            Width           =   2400
         End
         Begin VB.TextBox txtFields 
            DataField       =   "numtintes"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   22
            Left            =   225
            TabIndex        =   58
            Top             =   525
            Width           =   510
         End
         Begin VB.TextBox txtFields 
            DataField       =   "t8"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   31
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   74
            Top             =   1950
            Width           =   2400
         End
         Begin VB.Label lblLabels 
            Caption         =   "T8:"
            Height          =   255
            Index           =   31
            Left            =   1575
            TabIndex        =   75
            Top             =   1950
            Width           =   315
         End
         Begin VB.Label lblLabels 
            Caption         =   "T1:"
            Height          =   255
            Index           =   24
            Left            =   1575
            TabIndex        =   73
            Top             =   225
            Width           =   315
         End
         Begin VB.Label lblLabels 
            Caption         =   "T2:"
            Height          =   255
            Index           =   25
            Left            =   1575
            TabIndex        =   72
            Top             =   465
            Width           =   315
         End
         Begin VB.Label lblLabels 
            Caption         =   "T3:"
            Height          =   255
            Index           =   26
            Left            =   1575
            TabIndex        =   71
            Top             =   720
            Width           =   315
         End
         Begin VB.Label lblLabels 
            Caption         =   "T4:"
            Height          =   255
            Index           =   27
            Left            =   1575
            TabIndex        =   70
            Top             =   960
            Width           =   315
         End
         Begin VB.Label lblLabels 
            Caption         =   "T5:"
            Height          =   255
            Index           =   28
            Left            =   1575
            TabIndex        =   69
            Top             =   1200
            Width           =   315
         End
         Begin VB.Label lblLabels 
            Caption         =   "T6:"
            Height          =   255
            Index           =   29
            Left            =   1575
            TabIndex        =   68
            Top             =   1455
            Width           =   315
         End
         Begin VB.Label lblLabels 
            Caption         =   "T7:"
            Height          =   255
            Index           =   30
            Left            =   1575
            TabIndex        =   67
            Top             =   1695
            Width           =   315
         End
         Begin VB.Label lblLabels 
            Caption         =   "Núm. Tintes:"
            Height          =   255
            Index           =   22
            Left            =   150
            TabIndex        =   59
            Top             =   300
            Width           =   990
         End
      End
      Begin VB.TextBox txtFields 
         DataField       =   "tipusimpresio"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   23
         Left            =   1575
         MaxLength       =   15
         TabIndex        =   55
         Top             =   4350
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "mattipus"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   21
         Left            =   1575
         MaxLength       =   50
         TabIndex        =   53
         Top             =   4050
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "llarg"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   20
         Left            =   5100
         TabIndex        =   51
         Top             =   3675
         Width           =   810
      End
      Begin VB.TextBox txtFields 
         DataField       =   "plegat"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   19
         Left            =   2925
         TabIndex        =   50
         Top             =   3675
         Width           =   660
      End
      Begin VB.TextBox txtFields 
         DataField       =   "ample"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   18
         Left            =   750
         TabIndex        =   48
         Top             =   3675
         Width           =   510
      End
      Begin VB.TextBox txtFields 
         DataField       =   "numtreball"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   1500
         TabIndex        =   24
         Top             =   225
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "data"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   4575
         TabIndex        =   23
         Top             =   225
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "proveidor"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   7575
         TabIndex        =   22
         Top             =   225
         Width           =   285
      End
      Begin VB.TextBox txtFields 
         DataField       =   "treball"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   6
         Left            =   1425
         MaxLength       =   150
         TabIndex        =   21
         Top             =   975
         Width           =   8850
      End
      Begin VB.TextBox txtFields 
         DataField       =   "client"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   5
         Left            =   6450
         TabIndex        =   20
         Top             =   600
         Width           =   3810
      End
      Begin VB.TextBox txtFields 
         DataField       =   "numcomanda"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   4575
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         DataField       =   "dataentrega"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   1500
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sol.licitud de:"
         Height          =   990
         Left            =   225
         TabIndex        =   10
         Top             =   1425
         Width           =   10065
         Begin VB.TextBox txtFields 
            DataField       =   "solprovacolor"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   10
            Left            =   1650
            TabIndex        =   34
            Top             =   525
            Width           =   285
         End
         Begin VB.TextBox txtFields 
            DataField       =   "solcromablanc"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   9
            Left            =   8925
            TabIndex        =   33
            Top             =   525
            Width           =   285
         End
         Begin VB.TextBox txtFields 
            DataField       =   "solcromatransp"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   8
            Left            =   8925
            TabIndex        =   32
            Top             =   150
            Width           =   285
         End
         Begin VB.TextBox txtFields 
            DataField       =   "soldibuix"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   7
            Left            =   3900
            TabIndex        =   12
            Top             =   150
            Width           =   285
         End
         Begin VB.TextBox txtFields 
            DataField       =   "altres"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   11
            Left            =   2970
            MaxLength       =   100
            TabIndex        =   11
            Top             =   600
            Width           =   4800
         End
         Begin VB.Label Label1 
            Caption         =   "CROMALÍN"
            Height          =   165
            Left            =   6750
            TabIndex        =   45
            Top             =   300
            Width           =   915
         End
         Begin VB.Line Line4 
            X1              =   7650
            X2              =   7800
            Y1              =   375
            Y2              =   375
         End
         Begin VB.Line Line3 
            X1              =   7950
            X2              =   7800
            Y1              =   300
            Y2              =   300
         End
         Begin VB.Line Line2 
            X1              =   7800
            X2              =   7800
            Y1              =   300
            Y2              =   750
         End
         Begin VB.Line Line1 
            X1              =   7950
            X2              =   7800
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label lblLabels 
            Caption         =   "Dibuix:"
            Height          =   255
            Index           =   7
            Left            =   3300
            TabIndex        =   17
            Top             =   225
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Altres:"
            Height          =   255
            Index           =   11
            Left            =   2325
            TabIndex        =   16
            Top             =   600
            Width           =   690
         End
         Begin VB.Label lblLabels 
            Caption         =   "Transparent:"
            Height          =   255
            Index           =   8
            Left            =   7950
            TabIndex        =   15
            Top             =   225
            Width           =   1065
         End
         Begin VB.Label lblLabels 
            Caption         =   "Blanc:"
            Height          =   255
            Index           =   12
            Left            =   7950
            TabIndex        =   14
            Top             =   600
            Width           =   915
         End
         Begin VB.Label lblLabels 
            Caption         =   "Prova color per Imp:"
            Height          =   255
            Index           =   10
            Left            =   150
            TabIndex        =   13
            Top             =   600
            Width           =   1515
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Material Entregat:"
         Height          =   915
         Left            =   225
         TabIndex        =   9
         Top             =   2475
         Width           =   10065
         Begin VB.TextBox txtFields 
            DataField       =   "matentaltres"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   16
            Left            =   2925
            MaxLength       =   150
            TabIndex        =   44
            Top             =   525
            Width           =   7050
         End
         Begin VB.CheckBox chkFields 
            DataField       =   "matentprovacolor"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   14
            Left            =   9375
            TabIndex        =   41
            Top             =   150
            Width           =   375
         End
         Begin VB.CheckBox chkFields 
            DataField       =   "matentcdr"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   13
            Left            =   6000
            TabIndex        =   39
            Top             =   150
            Width           =   300
         End
         Begin VB.CheckBox chkFields 
            DataField       =   "matentnegatiu"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   12
            Left            =   3150
            TabIndex        =   38
            Top             =   150
            Width           =   300
         End
         Begin VB.CheckBox chkFields 
            DataField       =   "matentmostra"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   15
            Left            =   1350
            TabIndex        =   35
            Top             =   525
            Width           =   300
         End
         Begin VB.Label lblLabels 
            Caption         =   "Altres:"
            Height          =   255
            Index           =   16
            Left            =   2325
            TabIndex        =   43
            Top             =   525
            Width           =   690
         End
         Begin VB.Label lblLabels 
            Caption         =   "Prova Color:"
            Height          =   255
            Index           =   14
            Left            =   8325
            TabIndex        =   42
            Top             =   150
            Width           =   1065
         End
         Begin VB.Label lblLabels 
            Caption         =   "CD-R:"
            Height          =   255
            Index           =   13
            Left            =   5400
            TabIndex        =   40
            Top             =   150
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Negatiu:"
            Height          =   255
            Index           =   9
            Left            =   2325
            TabIndex        =   37
            Top             =   150
            Width           =   765
         End
         Begin VB.Label lblLabels 
            Caption         =   "Mostra Impresa:"
            Height          =   255
            Index           =   15
            Left            =   75
            TabIndex        =   36
            Top             =   525
            Width           =   1215
         End
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipus Impressió:"
         Height          =   255
         Index           =   23
         Left            =   225
         TabIndex        =   56
         Top             =   4350
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipus Material:"
         Height          =   255
         Index           =   21
         Left            =   225
         TabIndex        =   54
         Top             =   4050
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         Caption         =   "Llarg:"
         Height          =   255
         Index           =   20
         Left            =   4650
         TabIndex        =   52
         Top             =   3750
         Width           =   540
      End
      Begin VB.Label lblLabels 
         Caption         =   "Plegat:"
         Height          =   255
         Index           =   19
         Left            =   2400
         TabIndex        =   49
         Top             =   3750
         Width           =   540
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ample:"
         Height          =   255
         Index           =   18
         Left            =   225
         TabIndex        =   47
         Top             =   3750
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Producte:"
         Height          =   255
         Index           =   17
         Left            =   225
         TabIndex        =   46
         Top             =   3450
         Width           =   840
      End
      Begin VB.Label lblLabels 
         Caption         =   "Treball:"
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   31
         Top             =   225
         Width           =   990
      End
      Begin VB.Label lblLabels 
         Caption         =   "Data:"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   30
         Top             =   225
         Width           =   540
      End
      Begin VB.Label lblLabels 
         Caption         =   "Proveidor:"
         Height          =   255
         Index           =   2
         Left            =   6675
         TabIndex        =   29
         Top             =   225
         Width           =   810
      End
      Begin VB.Label lblLabels 
         Caption         =   "Desc.Treball:"
         Height          =   255
         Index           =   6
         Left            =   225
         TabIndex        =   28
         Top             =   975
         Width           =   1065
      End
      Begin VB.Label lblLabels 
         Caption         =   "Client:"
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   27
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Comanda:"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   26
         Top             =   600
         Width           =   915
      End
      Begin VB.Label lblLabels 
         Caption         =   "Data Entrega:"
         Height          =   255
         Index           =   3
         Left            =   225
         TabIndex        =   25
         Top             =   600
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   135
      TabIndex        =   0
      Tag             =   "100"
      Top             =   -60
      Width           =   10455
      Begin VB.CommandButton sortir 
         Height          =   525
         Left            =   9780
         Picture         =   "treballs.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Sortir a Menú"
         Top             =   165
         Width           =   570
      End
      Begin VB.CommandButton consultar 
         Height          =   450
         Left            =   1245
         Picture         =   "treballs.frx":0502
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Buscar Registres"
         Top             =   225
         Width           =   765
      End
      Begin VB.CommandButton alta 
         Height          =   450
         Left            =   75
         Picture         =   "treballs.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   225
         Width           =   450
      End
      Begin VB.CommandButton eliminar 
         Height          =   450
         Left            =   2010
         Picture         =   "treballs.frx":0CB6
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   225
         Width           =   465
      End
      Begin VB.CommandButton modificar 
         Height          =   450
         Left            =   525
         Picture         =   "treballs.frx":0FC8
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Modificar Registres"
         Top             =   225
         Width           =   720
      End
      Begin VB.Data datPrimaryRS 
         Caption         =   "                     Comandes"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   390
         Left            =   3975
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   255
         Width           =   3210
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Imprimir"
         Height          =   525
         Index           =   0
         Left            =   8775
         Picture         =   "treballs.frx":1316
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   150
         Width           =   900
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
         Left            =   3810
         TabIndex        =   7
         Top             =   300
         Width           =   105
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblLabels_Click(Index As Integer)

End Sub

Private Sub txtFields_Change(Index As Integer)

End Sub
