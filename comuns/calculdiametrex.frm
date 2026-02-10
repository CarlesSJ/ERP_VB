VERSION 5.00
Begin VB.Form calculdiametre 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calcular Diametre"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3180
   Icon            =   "calculdiametrex.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3270
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   2970
      Begin VB.TextBox metres17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2685
         Width           =   1095
      End
      Begin VB.TextBox metres15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2205
         Width           =   1095
      End
      Begin VB.TextBox metres7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1695
         Width           =   1095
      End
      Begin VB.TextBox cdiametre 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1800
         TabIndex        =   3
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox micres 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1800
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Canutu 17 cm"
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
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   2760
         Width           =   1680
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Metres"
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
         Height          =   285
         Left            =   1785
         TabIndex        =   9
         Top             =   1380
         Width           =   1110
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Canutu 15,5 cm"
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
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   2280
         Width           =   1680
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   150
         X2              =   2850
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Canutu 7,5 cm"
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
         Height          =   285
         Left            =   165
         TabIndex        =   6
         Top             =   1770
         Width           =   1680
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Diàmetre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   765
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Micres"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   165
         TabIndex        =   2
         Top             =   300
         Width           =   945
      End
   End
End
Attribute VB_Name = "calculdiametre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub calculardiametre()
   Dim pi As Double
   Dim canut As Double
   Dim diam As Double
   Dim mtrs_7 As Double
   Dim mtrs_15 As Double
   Dim h10_7 As Double
   Dim h9 As Double
   Dim h7 As Double
   Dim h8_7 As Double
   Dim h8_15 As Double
   Dim h10_15 As Double
   Dim h8_17 As Double
   Dim h10_17 As Double
    pi = 4 * Atn(1)
    
    h7 = (cadbl(micres) * 0.0001) / 100
    If h7 = 0 Then GoTo fi
    h8_7 = (7.5 / 2) / 100
    h8_15 = (15.5 / 2) / 100
    h8_17 = (17 / 2) / 100
    h9 = (cadbl(cdiametre) / 2) / 100
    h10_7 = (h9 - h8_7) / h7
    h10_15 = (h9 - h8_15) / h7
    h10_17 = (h9 - h8_17) / h7
    mtrs_7 = (2 * pi * h8_7 * h10_7) + ((pi * (h10_7 * h10_7)) * h7)
    mtrs_15 = (2 * pi * h8_15 * h10_15) + ((pi * (h10_15 * h10_15)) * h7)
    mtrs_17 = (2 * pi * h8_17 * h10_17) + ((pi * (h10_17 * h10_17)) * h7)
fi:
    metres7 = "0"
    metres15 = "0"
    metres17 = "0"
    If mtrs_7 > 0 Then metres7 = Format(mtrs_7, "#,##0")
    If mtrs_15 > 0 Then metres15 = Format(mtrs_15, "#,##0")
    If mtrs_17 > 0 Then metres17 = Format(mtrs_17, "#,##0")
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cdiametre_Change()
calculardiametre
End Sub

Private Sub metres_Change()

End Sub

Private Sub Form_Activate()
cdiametre.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 110 Then KeyCode = 188
End Sub

Private Sub micres_Change()
  calculardiametre
End Sub

Private Sub Text1_Change()

End Sub

