VERSION 5.00
Begin VB.Form formconsumtintes 
   Caption         =   "Consum de tintes"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6630
   Icon            =   "formconsumtintes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data imppantones 
      Caption         =   "imppantones"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   495
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame framepantones 
      Caption         =   "Pantones"
      Height          =   3315
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   6450
      Begin VB.TextBox pantone 
         BackColor       =   &H00F3B378&
         DataField       =   "pantone1"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   0
         Left            =   255
         MaxLength       =   40
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   375
         Width           =   2850
      End
      Begin VB.TextBox pantone 
         BackColor       =   &H00F3B378&
         DataField       =   "pantone2"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   1
         Left            =   255
         MaxLength       =   40
         TabIndex        =   31
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   660
         Width           =   2850
      End
      Begin VB.TextBox pantone 
         BackColor       =   &H00F3B378&
         DataField       =   "pantone3"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   2
         Left            =   255
         MaxLength       =   40
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   930
         Width           =   2850
      End
      Begin VB.TextBox pantone 
         BackColor       =   &H00F3B378&
         DataField       =   "pantone4"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   3
         Left            =   255
         MaxLength       =   40
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   1200
         Width           =   2850
      End
      Begin VB.TextBox pantone 
         BackColor       =   &H00F3B378&
         DataField       =   "pantone5"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   4
         Left            =   255
         MaxLength       =   40
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   1470
         Width           =   2850
      End
      Begin VB.TextBox pantone 
         BackColor       =   &H00F3B378&
         DataField       =   "pantone6"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   5
         Left            =   255
         MaxLength       =   40
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   1755
         Width           =   2850
      End
      Begin VB.TextBox pantone 
         BackColor       =   &H00F3B378&
         DataField       =   "pantone7"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   6
         Left            =   255
         MaxLength       =   40
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   2025
         Width           =   2850
      End
      Begin VB.TextBox pantone 
         BackColor       =   &H00F3B378&
         DataField       =   "pantone8"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   7
         Left            =   255
         MaxLength       =   40
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   2310
         Width           =   2850
      End
      Begin VB.TextBox pantone 
         BackColor       =   &H00F3B378&
         DataField       =   "pantone9"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   8
         Left            =   255
         MaxLength       =   40
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   2580
         Width           =   2850
      End
      Begin VB.TextBox pantone 
         BackColor       =   &H00F3B378&
         DataField       =   "pantone10"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   9
         Left            =   255
         MaxLength       =   40
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   2850
         Width           =   2850
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot1"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   0
         Left            =   4065
         MaxLength       =   12
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   405
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg1"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   0
         Left            =   5730
         MaxLength       =   8
         TabIndex        =   1
         Tag             =   "1"
         Top             =   405
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot2"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   1
         Left            =   4065
         MaxLength       =   12
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   680
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg2"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   1
         Left            =   5730
         MaxLength       =   8
         TabIndex        =   2
         Tag             =   "1"
         Top             =   690
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot3"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   2
         Left            =   4065
         MaxLength       =   12
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   955
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg3"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   2
         Left            =   5730
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "1"
         Top             =   960
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot4"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   3
         Left            =   4065
         MaxLength       =   12
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   1230
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg4"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   3
         Left            =   5730
         MaxLength       =   8
         TabIndex        =   4
         Tag             =   "1"
         Top             =   1230
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot5"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   4
         Left            =   4065
         MaxLength       =   12
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   1505
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg5"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   4
         Left            =   5730
         MaxLength       =   8
         TabIndex        =   5
         Tag             =   "1"
         Top             =   1500
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot6"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   5
         Left            =   4065
         MaxLength       =   12
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   1780
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg6"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   5
         Left            =   5730
         MaxLength       =   8
         TabIndex        =   6
         Tag             =   "1"
         Top             =   1785
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot7"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   6
         Left            =   4065
         MaxLength       =   12
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   2055
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg7"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   6
         Left            =   5730
         MaxLength       =   8
         TabIndex        =   7
         Tag             =   "1"
         Top             =   2055
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot8"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   7
         Left            =   4065
         MaxLength       =   12
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   2330
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg8"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   7
         Left            =   5730
         MaxLength       =   8
         TabIndex        =   8
         Tag             =   "1"
         Top             =   2340
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot9"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   8
         Left            =   4065
         MaxLength       =   12
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   2605
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg9"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   8
         Left            =   5730
         MaxLength       =   8
         TabIndex        =   9
         Tag             =   "1"
         Top             =   2610
         Width           =   550
      End
      Begin VB.TextBox compantone 
         BackColor       =   &H00E0E0E0&
         DataField       =   "lot10"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   9
         Left            =   4065
         MaxLength       =   12
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "888"
         Top             =   2880
         Width           =   1680
      End
      Begin VB.TextBox kbpantone 
         DataField       =   "kg10"
         DataSource      =   "imppantones"
         Height          =   285
         Index           =   9
         Left            =   5730
         MaxLength       =   8
         TabIndex        =   10
         Tag             =   "1"
         Top             =   2880
         Width           =   550
      End
      Begin VB.Label Label3 
         Caption         =   "Kg"
         Height          =   255
         Left            =   5895
         TabIndex        =   34
         Top             =   180
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Lot o llauna de traçabilitat"
         Height          =   255
         Left            =   3300
         TabIndex        =   33
         Top             =   180
         Width           =   2340
      End
      Begin VB.Label Label2 
         Caption         =   "Nom de la tinta"
         Height          =   255
         Left            =   855
         TabIndex        =   22
         Top             =   150
         Width           =   1665
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1 2 3 4 5 6 7 8 9 10"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   45
         TabIndex        =   21
         Top             =   390
         Width           =   195
      End
   End
End
Attribute VB_Name = "formconsumtintes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 imppantones.DatabaseName = rutadelfitxer(cami) + "baixes.mdb"
 imppantones.RecordSource = "select * from impresorespantones where comanda=" + atrim(cadbl(Form1.comanda))
 imppantones.Refresh
End Sub

Private Sub kbpantone_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 And Index < kbpantone.Count - 1 Then kbpantone(Index + 1).SetFocus
  If KeyCode = 38 And Index > 0 Then kbpantone(Index - 1).SetFocus
End Sub
