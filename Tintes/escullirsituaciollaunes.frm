VERSION 5.00
Begin VB.Form formescullirsituaciollaunes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Situació"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   2145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton sortirs 
      Height          =   480
      Left            =   1350
      Picture         =   "escullirsituaciollaunes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Sortir sense canvis"
      Top             =   3000
      Width           =   585
   End
   Begin VB.CommandButton Command1 
      Height          =   480
      Left            =   615
      Picture         =   "escullirsituaciollaunes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Acceptar canvis"
      Top             =   3000
      Width           =   585
   End
   Begin VB.Frame Frame1 
      Caption         =   "Situació"
      Height          =   2880
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1995
      Begin VB.CommandButton situacio 
         Caption         =   "Impresores"
         Height          =   540
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Tag             =   "IMP"
         Top             =   270
         Width           =   1665
      End
      Begin VB.CommandButton situacio 
         Caption         =   "Sala"
         Height          =   540
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Tag             =   "SALA"
         Top             =   885
         Width           =   1665
      End
      Begin VB.ComboBox combosituacio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   165
         TabIndex        =   1
         Top             =   1605
         Width           =   1650
      End
      Begin VB.Label etmaxim 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   60
         TabIndex        =   6
         Top             =   1965
         Width           =   1875
      End
   End
End
Attribute VB_Name = "formescullirsituaciollaunes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub combosituacio_Click()
  comprovar_maximpersituacio
End Sub

Private Sub combosituacio_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
  comprovar_maximpersituacio
End Sub
Sub comprovar_maximpersituacio()
   Dim rst As Recordset
   
   Set rst = dbtintes.OpenRecordset("SELECT Count(Llaunes.id) AS quanteshiha From Llaunes WHERE (((Llaunes.activa)=True) AND ((Llaunes.situacio)='" + combosituacio + "'));", , ReadOnly)
   If rst.EOF Then GoTo fi
   etmaxim = "Q: " + atrim(cadbl(rst!quanteshiha)) + " Max: " + atrim(cadbl(combosituacio.ItemData(combosituacio.ListIndex)))
   If cadbl(rst!quanteshiha) >= cadbl(combosituacio.ItemData(combosituacio.ListIndex)) And cadbl(combosituacio.ItemData(combosituacio.ListIndex)) > 0 Then
      combosituacio.BackColor = QBColor(12)
     Else: combosituacio.BackColor = QBColor(15)
   End If
fi:
   Set rst = Nothing
End Sub

Private Sub Command1_Click()
   If Not comprovar_situacio_existeix Then MsgBox "La situació de llauna escullida no existeix", vbCritical, "Error": Exit Sub
   If combosituacio.BackColor = QBColor(12) Then If UCase(InputBox("Aquesta UBICACIÓ ja està plena," + vbNewLine + "VOLS POSSAR IGUALMENT AQUEST BIDÓ EN AQUESTA UBICACIÓ? " + vbNewLine + "ESCRIU [SI] per acceptar-ho igualment", "CANVI UBICACIÓ")) <> "SI" Then Exit Sub
   formescullirsituaciollaunes.Hide
End Sub
Function comprovar_situacio_existeix() As Boolean
  Dim rst As Recordset
  comprovar_situacio_existeix = True
  Set rst = dbtintes.OpenRecordset("Select * from situacionsllaunes where situacio='" + atrim(combosituacio) + "'")
  If rst.EOF Then comprovar_situacio_existeix = False
  Set rst = Nothing
End Function

Private Sub Form_Load()
carregarsituacionsalcombo
End Sub
Sub carregarsituacionsalcombo()
    Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from situacionsllaunes order by situacio")
   While Not rst.EOF
     combosituacio.AddItem UCase(rst!situacio)
     combosituacio.ItemData(combosituacio.NewIndex) = cadbl(rst!unitatsmaxim)
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub situacio_Click(Index As Integer)
  combosituacio.Text = situacio(Index).tag
End Sub

Private Sub sortirs_Click()
 combosituacio = ""
 formescullirsituaciollaunes.Hide
End Sub
