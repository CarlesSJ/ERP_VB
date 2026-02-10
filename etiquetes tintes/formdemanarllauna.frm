VERSION 5.00
Begin VB.Form formdemanarllauna 
   BackColor       =   &H00EAD9CE&
   Caption         =   "Impresio Etiqueta | Entra el Nº de Llauna que utilitzes."
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5565
   Icon            =   "formdemanarllauna.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox Llaunes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   2610
   End
   Begin VB.CommandButton Command1 
      Caption         =   "D'acord"
      Height          =   675
      Left            =   3300
      Picture         =   "formdemanarllauna.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1455
      Width           =   1935
   End
   Begin VB.TextBox numerollauna 
      Height          =   405
      Left            =   3420
      TabIndex        =   1
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label eformula 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2670
      TabIndex        =   4
      Top             =   45
      Width           =   2910
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº llauna Nou:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3435
      TabIndex        =   0
      Top             =   540
      Width           =   1815
   End
End
Attribute VB_Name = "formdemanarllauna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Me.Hide
End Sub
Function comprovarsiexisteixformula() As Boolean
  Dim rst As Recordset
  comprovarsiexisteixformula = True
  Set rst = dbtintes.OpenRecordset("SELECT * from formules where codiformula='" + atrim(eformula) + "'")
  If rst.EOF Then
      comprovarsiexisteixformula = False
      formdemanarllauna.BackColor = QBColor(12)
      Llaunes.AddItem "# Formula Nova"
  End If
  Set rst = Nothing
End Function
Private Sub Form_Load()
   
   If comprovarsiexisteixformula Then carregar_llistallaunes
   
End Sub
Sub carregar_llistallaunes()
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, historiallauna.formula, Llaunes.activa FROM Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna  WHERE (((historiallauna.formula)='" + eformula + "') AND ((Llaunes.activa)=True));")
  If rst.EOF Then Llaunes.AddItem "# No hi ha Llaunes"
  While Not rst.EOF
    Llaunes.AddItem UCase(rst!numllauna)
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
