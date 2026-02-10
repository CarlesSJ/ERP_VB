VERSION 5.00
Begin VB.Form triarcomandastocopre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Triar Estoc, Precomanda o Comanda"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3900
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2310
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   3765
      Begin VB.CommandButton Command3 
         Height          =   540
         Left            =   2700
         Picture         =   "triarcomandastocopre.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Acceptar canvis"
         Top             =   1680
         Width           =   840
      End
      Begin VB.TextBox comanda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1770
         TabIndex        =   3
         Top             =   1140
         Width           =   1800
      End
      Begin VB.CommandButton precomanda 
         BackColor       =   &H00D29F7D&
         Caption         =   "Precomanda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   270
         Width           =   1890
      End
      Begin VB.CommandButton estoc 
         BackColor       =   &H00D29F7D&
         Caption         =   "Estoc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Comanda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   255
         TabIndex        =   4
         Top             =   1215
         Width           =   1500
      End
   End
End
Attribute VB_Name = "triarcomandastocopre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   
End Sub

Private Sub comanda_Change()
estoc.BackColor = &HD29F7D
   precomanda.BackColor = &HD29F7D
End Sub

Private Sub Command3_Click()
  If existeixcomanda Then triarcomandastocopre.Hide
End Sub
Function existeixcomanda() As Boolean
   Dim rstc As Recordset
   existeixcomanda = True
   If cadbl(comanda) > 0 Then
      Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(comanda)))
      If rstc.EOF Then
         existeixcomanda = False
         MsgBox "Aquesta comanda no existeix a la base de dades de comandes", vbCritical, "Error"
      End If
   End If
End Function
Private Sub estoc_Click()
   estoc.BackColor = &H9AA6FA
   precomanda.BackColor = &HD29F7D
   comanda = ""
End Sub

Private Sub Form_Activate()
 comanda.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then Command3_Click
End Sub

Private Sub precomanda_Click()
   estoc.BackColor = &HD29F7D
   precomanda.BackColor = &H9AA6FA
   comanda = ""
End Sub
