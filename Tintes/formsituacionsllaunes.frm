VERSION 5.00
Begin VB.Form formsituacio 
   Caption         =   "Situació de la llauna"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4530
   Icon            =   "formsituacionsllaunes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Canvi de magatzem"
      Height          =   1200
      Left            =   45
      TabIndex        =   11
      Top             =   2265
      Width           =   2160
      Begin VB.TextBox cnommagatzem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   795
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Magatzem"
         Height          =   540
         Left            =   1080
         TabIndex        =   13
         Top             =   225
         Width           =   1020
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Impresores"
         Height          =   540
         Left            =   45
         TabIndex        =   12
         Top             =   225
         Width           =   1020
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   660
      Left            =   3015
      Picture         =   "formsituacionsllaunes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Acceptar canvis"
      Top             =   3015
      Width           =   675
   End
   Begin VB.CommandButton sortirs 
      Height          =   660
      Left            =   3750
      Picture         =   "formsituacionsllaunes.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Sortir sense canvis"
      Top             =   3000
      Width           =   660
   End
   Begin VB.Frame Frame1 
      Caption         =   "Situació"
      Height          =   2175
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   2115
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
         Height          =   330
         Left            =   270
         TabIndex        =   3
         Top             =   1620
         Width           =   1650
      End
      Begin VB.CommandButton situacio 
         Caption         =   "Exterior"
         Height          =   540
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Tag             =   "EXT"
         Top             =   885
         Width           =   1665
      End
      Begin VB.CommandButton situacio 
         Caption         =   "Impresores"
         Height          =   540
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Tag             =   "IMP"
         Top             =   270
         Width           =   1665
      End
      Begin VB.Label etmaxim 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   105
         TabIndex        =   15
         Top             =   1950
         Width           =   1875
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Llaunes"
      Height          =   2880
      Left            =   2250
      TabIndex        =   4
      Top             =   60
      Width           =   2130
      Begin VB.CommandButton Command2 
         Height          =   435
         Left            =   1665
         Picture         =   "formsituacionsllaunes.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Acceptar canvis"
         Top             =   210
         Width           =   405
      End
      Begin VB.TextBox numllauna 
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
         Left            =   120
         TabIndex        =   6
         Top             =   210
         Width           =   1545
      End
      Begin VB.ListBox llistadellaunes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         ItemData        =   "formsituacionsllaunes.frx":1628
         Left            =   135
         List            =   "formsituacionsllaunes.frx":162A
         TabIndex        =   5
         Top             =   705
         Width           =   1830
      End
   End
   Begin VB.Label missatge 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   105
      TabIndex        =   10
      Top             =   3450
      Width           =   2820
   End
End
Attribute VB_Name = "formsituacio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  
End Sub

Function existeixlallauna(numllauna As String, vactiva As Boolean) As Boolean
   Dim rst As Recordset
   existeixlallauna = True
   Set rst = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(numllauna) + "'")
   If rst.EOF Then
      existeixlallauna = False
       Else: vactiva = rst!activa
   End If
   Set rst = Nothing
End Function

Private Sub combosituacio_Click()
   comprovar_maximpersituacio
   numllauna.SetFocus
   cnommagatzem.Text = ""
End Sub

Private Sub combosituacio_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
  comprovar_maximpersituacio
  numllauna.SetFocus
End Sub

Private Sub combosituacio_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
  comprovar_maximpersituacio
  numllauna.SetFocus
End Sub

Function comprovar_silesllaunesesnecessitenafabrica() As String

End Function

Private Sub Command1_Click()
   Dim i As Byte
   Dim vnumllauna As String
   Dim vllaunesfabrica As String
   If cnommagatzem = "" And combosituacio.Text = "" Then MsgBox "Falta escullir el destí de les llaunes", vbCritical, "Error": Exit Sub
'   If cnommagatzem.Text = "Magatzem" Then
'     vllaunesfabrica = comprovar_silesllaunesesnecessitenafabrica
'     If vllaunesfabrica <> "" Then If MsgBox("Les llaunes " + vllaunesfabrica + " es necessitaran a fàbrica d'aqui a poc temps." + Chr(10) + "Segur que vols pujar-les a magatzem?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
'   End If
If combosituacio.BackColor = QBColor(12) Then If UCase(InputBox("Aquesta UBICACIÓ ja està plena," + vbNewLine + "VOLS POSSAR IGUALMENT AQUEST BIDÓ EN AQUESTA UBICACIÓ? " + vbNewLine + "ESCRIU [SI] per acceptar-ho igualment", "CANVI UBICACIÓ")) <> "SI" Then Exit Sub
   If combosituacio.Text <> "" Then
        numllauna.SetFocus
        If llistadellaunes.ListCount = 0 Then MsgBox "No hi ha cap llauna entrada.", , "Atenció": Exit Sub
        If MsgBox("Segur que vols passar aquestes llaunes a l'estat de " + combosituacio + "?", vbInformation + vbYesNo, "Atenció") = vbNo Then Exit Sub
        For i = 0 To llistadellaunes.ListCount
           If atrim(UCase(llistadellaunes.List(i))) <> "" Then
                 dbtintes.Execute "update  llaunes set situacio='" + atrim(UCase(combosituacio)) + "' where numllauna='" + atrim(UCase(llistadellaunes.List(i))) + "'"
                 dbtintes.Execute "insert into historialsituacions (data,situacio,numllauna) values (now,'" + atrim(UCase(combosituacio)) + "','" + atrim(UCase(llistadellaunes.List(i))) + "')"
                 If combosituacio <> "IMP" Then imprimir_etiqueta atrim(UCase(llistadellaunes.List(i)))
           End If
        Next i
   End If
   If cnommagatzem <> "" Then
     numllauna.SetFocus
        If llistadellaunes.ListCount = 0 Then MsgBox "No hi ha cap llauna entrada.", , "Atenció": Exit Sub
        If MsgBox("Segur que vols passar aquestes llaunes a " + cnommagatzem + "?", vbInformation + vbYesNo, "Atenció") = vbNo Then Exit Sub
        For i = 0 To llistadellaunes.ListCount
           If atrim(UCase(llistadellaunes.List(i))) <> "" Then
                 dbtintes.Execute "update  llaunes set aimpresores=" + IIf(cnommagatzem = "Impresores", "True", "False") + " where numllauna='" + atrim(UCase(llistadellaunes.List(i))) + "'"
           End If
        Next i
   End If
   Unload Me
End Sub

Private Sub Command2_Click()
   afegeixllauna
End Sub

Private Sub Command3_Click()
  combosituacio.Text = ""
  cnommagatzem.Text = "Impresores"
  numllauna.SetFocus
End Sub

Private Sub Command4_Click()
  combosituacio.Text = ""
  cnommagatzem.Text = "Magatzem"
  numllauna.SetFocus
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    numllauna.SetFocus
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

Private Sub llistadellaunes_DblClick()
   If llistadellaunes.ListIndex < 0 Then Exit Sub
   llistadellaunes.RemoveItem llistadellaunes.ListIndex
End Sub

Private Sub numllauna_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
      afegeixllauna
  End If
  numllauna.BackColor = QBColor(15)
End Sub
Function encaranohies(numllauna) As Boolean
   Dim i As Byte
   encaranohies = True
   For i = 0 To llistadellaunes.ListCount
      If llistadellaunes.List(i) = numllauna Then encaranohies = False
   Next i
   
End Function
Sub afegeixllauna()
   Dim vactiva As Boolean
   missatge = ""
   If existeixlallauna(numllauna, vactiva) Then
    If Not vactiva Then
        If MsgBox("Aquesta llauna no està activa." + vbNewLine + "Vols utilitzar-la igualment?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then vactiva = True
    End If
    If vactiva Then
     If encaranohies(numllauna) Then
       llistadellaunes.AddItem UCase(numllauna)
       numllauna.Text = ""
         Else: missatge = "Aquesta llauna ja està a la llista"
     End If
       Else
          numllauna.BackColor = QBColor(12)
          missatge = "Aquesta llauna no està ACTIVA"
    End If
        Else: MsgBox "Aquesta llauna no existeix", vbCritical, "Error": Exit Sub
    End If
    
End Sub

Private Sub situacio_Click(Index As Integer)
   Dim i As Byte
   For i = 0 To situacio.Count - 1
      situacio(i).BackColor = formsituacio.BackColor
   Next i
   situacio(Index).BackColor = QBColor(12)
   combosituacio.Text = situacio(Index).tag
   cnommagatzem.Text = ""
   numllauna.SetFocus
End Sub

Private Sub sortirs_Click()
    Unload Me
End Sub
