VERSION 5.00
Begin VB.Form form_escullircondicionant 
   BackColor       =   &H00EAD9CE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Escull el condicionant"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5355
   ClipControls    =   0   'False
   Icon            =   "Escullir_condicionant.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox combocondicionants 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Escullir_condicionant.frx":058A
      Left            =   330
      List            =   "Escullir_condicionant.frx":05B8
      TabIndex        =   2
      Top             =   300
      Width           =   4770
   End
   Begin VB.Frame Frame1 
      Height          =   2550
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   5145
      Begin VB.Frame Frame5 
         BackColor       =   &H00FDDECE&
         Height          =   1110
         Left            =   1140
         TabIndex        =   3
         Top             =   765
         Width           =   3060
         Begin VB.OptionButton csumaresta 
            BackColor       =   &H00FDDECE&
            Caption         =   "+ Augment de preu."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   105
            TabIndex        =   6
            Top             =   150
            Width           =   2475
         End
         Begin VB.OptionButton csumaresta 
            BackColor       =   &H00FDDECE&
            Caption         =   "-  Reducció de preu."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   105
            TabIndex        =   5
            Top             =   420
            Width           =   2475
         End
         Begin VB.OptionButton csumaresta 
            BackColor       =   &H00FDDECE&
            Caption         =   "Preu Fix."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   105
            TabIndex        =   4
            Top             =   690
            Width           =   2910
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Acceptar"
         Height          =   465
         Left            =   1500
         TabIndex        =   1
         Top             =   1950
         Width           =   2190
      End
   End
End
Attribute VB_Name = "form_escullircondicionant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim vopcio As String
   If csumaresta(0) = True Then vopcio = "+"
   If csumaresta(1) = True Then vopcio = "-"
   If csumaresta(2) = True Then vopcio = "F"
   If vopcio = "" Then
        MsgBox "Primer escull un valors de Suma, Resta o Fix.", vbExclamation, "Escull"
        Exit Sub
   End If
   form_tarifes.bcondicionantbarem.tag = vopcio + combocondicionants
   Unload Me
End Sub

Private Sub Form_Activate()
   combocondicionants.SetFocus
   SendKeys "%{DOWN}"
End Sub

