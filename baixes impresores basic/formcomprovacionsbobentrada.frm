VERSION 5.00
Begin VB.Form formcomprovacionsbobentrada 
   Caption         =   "Comprovacions bobina entrada"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "D'acord"
      Height          =   615
      Left            =   1335
      TabIndex        =   7
      Top             =   2250
      Width           =   1905
   End
   Begin VB.CheckBox cverificaciotractat 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1950
      TabIndex        =   6
      Top             =   1665
      Width           =   210
   End
   Begin VB.ComboBox ccolormaterial 
      Height          =   315
      Left            =   1890
      TabIndex        =   4
      Top             =   1005
      Width           =   2505
   End
   Begin VB.TextBox cespessor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1890
      TabIndex        =   1
      Text            =   "0"
      Top             =   330
      Width           =   705
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Verificació Tractat:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   5
      Top             =   1725
      Width           =   1785
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Color del material."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   3
      Top             =   1065
      Width           =   1785
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Micres"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2685
      TabIndex        =   2
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Espessor material."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   0
      Top             =   420
      Width           =   1785
   End
End
Attribute VB_Name = "formcomprovacionsbobentrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If cadbl(cespessor) = 0 Then MsgBox "Has d'entrar espessor", vbCritical, "Error": Exit Sub
 If ccolormaterial.Text = "" Then MsgBox "Has d'entrar color del material", vbCritical, "Error": Exit Sub
 If cverificaciotractat.Value = 0 Then MsgBox "Has de comprovar el tractat", vbCritical, "Error": Exit Sub
 formcomprovacionsbobentrada.Hide
End Sub

Private Sub Form_Activate()
  cespessor.SetFocus
End Sub

