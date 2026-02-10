VERSION 5.00
Begin VB.Form formllegirdeltaibarres 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Llegir valor"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8370
   Icon            =   "formllegirdeltaibarres.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Begin VB.Frame framedelta 
      Caption         =   "Valor Delta"
      Height          =   5790
      Left            =   105
      TabIndex        =   0
      Top             =   330
      Visible         =   0   'False
      Width           =   8130
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00EAD9CE&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   3
         Left            =   6105
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   330
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00FDDECE&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   7
         Left            =   6105
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1470
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00F3B378&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   11
         Left            =   6105
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2610
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00ED823A&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   15
         Left            =   6105
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3750
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00FF80FF&
         Caption         =   "No llegeix"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   16
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4950
         Width           =   7680
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00ED823A&
         Caption         =   "3,75"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   14
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3750
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00ED823A&
         Caption         =   "3,50"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   13
         Left            =   2175
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3750
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00ED823A&
         Caption         =   "3,25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   12
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3765
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00F3B378&
         Caption         =   "2,75"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   10
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2610
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00F3B378&
         Caption         =   "2,50"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   9
         Left            =   2175
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2610
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00F3B378&
         Caption         =   "2,25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   8
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2625
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00FDDECE&
         Caption         =   "1,75"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   6
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1470
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00FDDECE&
         Caption         =   "1,50"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   5
         Left            =   2175
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1470
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00FDDECE&
         Caption         =   "1,25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   4
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1485
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00EAD9CE&
         Caption         =   "0,75"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   2
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   330
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00EAD9CE&
         Caption         =   "0,50"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   1
         Left            =   2175
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   330
         Width           =   1815
      End
      Begin VB.CommandButton botodelta 
         BackColor       =   &H00EAD9CE&
         Caption         =   "0,25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   0
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   345
         Width           =   1815
      End
   End
   Begin VB.Frame frameCB 
      Caption         =   "Valors Codi de barres"
      Height          =   5820
      Left            =   390
      TabIndex        =   14
      Top             =   315
      Visible         =   0   'False
      Width           =   7500
      Begin VB.CommandButton botoCB 
         BackColor       =   &H00FF00FF&
         Caption         =   "No en té"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   5
         Left            =   3435
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4995
         Width           =   3870
      End
      Begin VB.CommandButton botoCB 
         BackColor       =   &H00FF80FF&
         Caption         =   "No Llegeix"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4995
         Width           =   3180
      End
      Begin VB.CommandButton botoCB 
         BackColor       =   &H00EAD9CE&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Index           =   1
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   210
         Width           =   7080
      End
      Begin VB.CommandButton botoCB 
         BackColor       =   &H00FDDECE&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Index           =   2
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1425
         Width           =   7080
      End
      Begin VB.CommandButton botoCB 
         BackColor       =   &H00F3B378&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Index           =   3
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2640
         Width           =   7080
      End
      Begin VB.CommandButton botoCB 
         BackColor       =   &H00ED823A&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Index           =   4
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3855
         Width           =   7080
      End
   End
   Begin VB.Label etdeltamaxicolor 
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
      Height          =   360
      Left            =   15
      TabIndex        =   21
      Top             =   0
      Width           =   4560
   End
End
Attribute VB_Name = "formllegirdeltaibarres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const HWND_NOTOPMOST = -2





Private Sub botoCB_Click(Index As Integer)
  formllegirdeltaibarres.tag = botoCB(Index).caption
  formllegirdeltaibarres.Hide
End Sub

Private Sub botodelta_Click(Index As Integer)
  formllegirdeltaibarres.tag = botodelta(Index).caption
  formllegirdeltaibarres.Hide
End Sub

Private Sub Form_Activate()
   SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

