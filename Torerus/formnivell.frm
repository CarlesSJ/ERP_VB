VERSION 5.00
Begin VB.Form formnivell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Escullir nivell(Pis)"
   ClientHeight    =   10740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10740
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEE4D7&
      Height          =   10695
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   4485
      Begin VB.Image Image11 
         Height          =   1500
         Index           =   0
         Left            =   15
         Picture         =   "formnivell.frx":0000
         Stretch         =   -1  'True
         Top             =   9165
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Image Image11 
         Height          =   1545
         Index           =   1
         Left            =   15
         Picture         =   "formnivell.frx":139F
         Stretch         =   -1  'True
         Top             =   7635
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Image Image11 
         Height          =   1500
         Index           =   6
         Left            =   15
         Picture         =   "formnivell.frx":2F55
         Stretch         =   -1  'True
         Top             =   135
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Image Image11 
         Height          =   1500
         Index           =   5
         Left            =   15
         Picture         =   "formnivell.frx":4B0B
         Stretch         =   -1  'True
         Top             =   1635
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Image Image11 
         Height          =   1500
         Index           =   4
         Left            =   15
         Picture         =   "formnivell.frx":66C1
         Stretch         =   -1  'True
         Top             =   3135
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Image Image11 
         Height          =   1500
         Index           =   3
         Left            =   15
         Picture         =   "formnivell.frx":8277
         Stretch         =   -1  'True
         Top             =   4635
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Image Image11 
         Height          =   1500
         Index           =   2
         Left            =   15
         Picture         =   "formnivell.frx":9E2D
         Stretch         =   -1  'True
         Top             =   6135
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label foratocupat 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Index           =   6
         Left            =   840
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.Label foratocupat 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Index           =   5
         Left            =   825
         TabIndex        =   6
         Top             =   1590
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.Label foratocupat 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   4
         Left            =   840
         TabIndex        =   5
         Top             =   3105
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.Label foratocupat 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   3
         Left            =   840
         TabIndex        =   4
         Top             =   4605
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.Label foratocupat 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   2
         Left            =   870
         TabIndex        =   3
         Top             =   6090
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.Label foratocupat 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   1
         Left            =   840
         TabIndex        =   2
         Top             =   7605
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.Label foratocupat 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   9135
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Height          =   10530
         Left            =   855
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   2760
      End
   End
End
Attribute VB_Name = "formnivell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image7_Click()
End Sub

Private Sub Image6_Click()
End Sub

Private Sub Form_Load()
   Dim i As Byte
   For i = 0 To 6
      foratocupat(i).BackColor = &H80FF80
      Image11(i).Visible = True
   Next i
End Sub

Private Sub Image11_Click(Index As Integer)

  If foratocupat(Index).Visible Then
     If foratocupat(Index).BackColor = &H8080FF And Frame1.Tag = "" Then
       If MsgBox("Aquest forat ja hi ha un palet." + Chr(10) + "VOLS POSAR-LO IGUALMENT?", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then GoTo fi
     End If
     formnivell.Tag = atrim(Index + 1)
  End If
'   Unload formnivell
fi:
   Me.Hide
End Sub

