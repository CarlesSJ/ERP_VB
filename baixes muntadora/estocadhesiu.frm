VERSION 5.00
Begin VB.Form estocadhesiu 
   Caption         =   "Estoc Cinta adhesiva"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   4455
   Icon            =   "estocadhesiu.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0FF&
      Height          =   390
      Left            =   2400
      Picture         =   "estocadhesiu.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Guardar dades"
      Top             =   5625
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0080FF80&
      Height          =   390
      Left            =   60
      Picture         =   "estocadhesiu.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Guardar dades"
      Top             =   5625
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      Height          =   1000
      Left            =   75
      TabIndex        =   26
      Top             =   15
      Width           =   4245
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   3300
         Top             =   480
      End
      Begin VB.TextBox o1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         MaxLength       =   100
         TabIndex        =   37
         ToolTipText     =   "Observacions"
         Top             =   645
         Width           =   2460
      End
      Begin VB.TextBox adhesiu1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   75
         TabIndex        =   30
         Top             =   150
         Width           =   2985
      End
      Begin VB.TextBox q1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3090
         TabIndex        =   29
         Top             =   165
         Width           =   705
      End
      Begin VB.CommandButton Command1 
         Height          =   285
         Left            =   3870
         Picture         =   "estocadhesiu.frx":0F9E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Left            =   3885
         Picture         =   "estocadhesiu.frx":1528
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   315
      End
      Begin VB.Label etcomanda1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1770
         MouseIcon       =   "estocadhesiu.frx":1AB2
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   645
         Width           =   2040
      End
   End
   Begin VB.CommandButton bcomandafeta 
      BackColor       =   &H008080FF&
      Caption         =   "Comanda Feta"
      Height          =   435
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5595
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      Height          =   1000
      Left            =   75
      TabIndex        =   0
      Top             =   930
      Width           =   4245
      Begin VB.TextBox o2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         MaxLength       =   100
         TabIndex        =   38
         ToolTipText     =   "Observacions"
         Top             =   690
         Width           =   2460
      End
      Begin VB.TextBox adhesiu2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   75
         TabIndex        =   4
         Top             =   180
         Width           =   2985
      End
      Begin VB.TextBox q2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3090
         TabIndex        =   3
         Top             =   195
         Width           =   705
      End
      Begin VB.CommandButton Command4 
         Height          =   285
         Left            =   3870
         Picture         =   "estocadhesiu.frx":203C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         Height          =   285
         Left            =   3885
         Picture         =   "estocadhesiu.frx":25C6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   315
      End
      Begin VB.Label etcomanda2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1725
         MouseIcon       =   "estocadhesiu.frx":2B50
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   630
         Width           =   2040
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1000
      Left            =   75
      TabIndex        =   5
      Top             =   1845
      Width           =   4245
      Begin VB.TextBox o3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         MaxLength       =   100
         TabIndex        =   39
         ToolTipText     =   "Observacions"
         Top             =   675
         Width           =   2460
      End
      Begin VB.TextBox q3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3090
         TabIndex        =   8
         Top             =   195
         Width           =   705
      End
      Begin VB.CommandButton Command6 
         Height          =   285
         Left            =   3870
         Picture         =   "estocadhesiu.frx":30DA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton Command5 
         Height          =   285
         Left            =   3885
         Picture         =   "estocadhesiu.frx":3664
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   315
      End
      Begin VB.TextBox adhesiu3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   75
         TabIndex        =   9
         Top             =   180
         Width           =   2985
      End
      Begin VB.Label etcomanda3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1725
         MouseIcon       =   "estocadhesiu.frx":3BEE
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   630
         Width           =   2040
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1000
      Left            =   75
      TabIndex        =   20
      Top             =   2760
      Width           =   4245
      Begin VB.TextBox o4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         MaxLength       =   100
         TabIndex        =   40
         ToolTipText     =   "Observacions"
         Top             =   675
         Width           =   2460
      End
      Begin VB.TextBox adhesiu4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   75
         TabIndex        =   24
         Top             =   180
         Width           =   2985
      End
      Begin VB.TextBox q4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3090
         TabIndex        =   23
         Top             =   180
         Width           =   705
      End
      Begin VB.CommandButton Command12 
         Height          =   285
         Left            =   3870
         Picture         =   "estocadhesiu.frx":4178
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton Command11 
         Height          =   285
         Left            =   3885
         Picture         =   "estocadhesiu.frx":4702
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   315
      End
      Begin VB.Label etcomanda4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1725
         MouseIcon       =   "estocadhesiu.frx":4C8C
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   645
         Width           =   2040
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1000
      Left            =   75
      TabIndex        =   15
      Top             =   3675
      Width           =   4245
      Begin VB.TextBox o5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         MaxLength       =   100
         TabIndex        =   41
         ToolTipText     =   "Observacions"
         Top             =   690
         Width           =   2460
      End
      Begin VB.CommandButton Command10 
         Height          =   285
         Left            =   3885
         Picture         =   "estocadhesiu.frx":5216
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   315
      End
      Begin VB.CommandButton Command9 
         Height          =   285
         Left            =   3870
         Picture         =   "estocadhesiu.frx":57A0
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   405
         Width           =   330
      End
      Begin VB.TextBox q5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3090
         TabIndex        =   17
         Top             =   195
         Width           =   705
      End
      Begin VB.TextBox adhesiu5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   75
         TabIndex        =   16
         Top             =   180
         Width           =   2985
      End
      Begin VB.Label etcomanda5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1755
         MouseIcon       =   "estocadhesiu.frx":5D2A
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   630
         Width           =   2040
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1000
      Left            =   75
      TabIndex        =   10
      Top             =   4590
      Width           =   4245
      Begin VB.TextBox o6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         MaxLength       =   100
         TabIndex        =   42
         ToolTipText     =   "Observacions"
         Top             =   675
         Width           =   2460
      End
      Begin VB.CommandButton Command8 
         Height          =   285
         Left            =   3885
         Picture         =   "estocadhesiu.frx":62B4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   315
      End
      Begin VB.CommandButton Command7 
         Height          =   285
         Left            =   3870
         Picture         =   "estocadhesiu.frx":683E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   405
         Width           =   330
      End
      Begin VB.TextBox q6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3090
         TabIndex        =   12
         Top             =   195
         Width           =   705
      End
      Begin VB.TextBox adhesiu6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   75
         TabIndex        =   11
         Top             =   180
         Width           =   2985
      End
      Begin VB.Label etcomanda6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1755
         MouseIcon       =   "estocadhesiu.frx":6DC8
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   630
         Width           =   2040
      End
   End
   Begin VB.Menu mestocminim 
      Caption         =   "Estoc mínim"
      Begin VB.Menu mminim1 
         Caption         =   "Estoc mínim 1"
      End
      Begin VB.Menu minim2 
         Caption         =   "Estoc mínim 2"
      End
      Begin VB.Menu minim3 
         Caption         =   "Estoc mínim 3"
      End
      Begin VB.Menu minim4 
         Caption         =   "Estoc mínim 4"
      End
      Begin VB.Menu minim5 
         Caption         =   "Estoc mínim 5"
      End
      Begin VB.Menu minim6 
         Caption         =   "Estoc mínim 6"
      End
   End
   Begin VB.Menu m1 
      Caption         =   "                                "
      Enabled         =   0   'False
      Index           =   1
      Visible         =   0   'False
   End
   Begin VB.Menu mcomandafeta 
      Caption         =   "Comanda Feta"
      Index           =   1
      Visible         =   0   'False
   End
End
Attribute VB_Name = "estocadhesiu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bcomandafeta_Click()
   If llegir_ini("Baixes", "programaamaquina", "comandes.ini") = 1 Then MsgBox "Aquesta opció nomes es pot canviar desde oficines", vbInformation, "Atenció": Exit Sub
   If MsgBox("Estas segur que la comanda està feta al proveïdor?", vbCritical + vbYesNo + vbDefaultButton2, "Comfirmació comanda") = vbNo Then Exit Sub
    escriure_ini "Valors", "enviat", "si", fitxerini
    MsgBox "Pensa a marcar les cintes que has demanat.", vbInformation, "Atenció"
    comprovarestatenvio
End Sub

Private Sub breclamarcomanda_Click()
    
End Sub

Private Sub ccomanda_Click()

End Sub

Private Sub comanda1_Click()

End Sub

Private Sub comanda2_Click()

End Sub

Private Sub Command1_Click()
  q1 = cadbl(q1) - 1
End Sub

Private Sub Command10_Click()
  q5 = cadbl(q5) + 1
End Sub

Private Sub Command11_Click()
q4 = cadbl(q4) + 1
End Sub

Private Sub Command12_Click()
  q4 = cadbl(q4) - 1
End Sub

Private Sub Command13_Click()
guardarvalors
End Sub

Private Sub Command14_Click()
   If MsgBox("Vols guardar canvis?", vbDefaultButton2 + vbYesNo, "Guardar canvis") = vbYes Then guardarvalors
  ' enviar_mail_acompres
   estocadhesiu.Hide
End Sub

Private Sub Command2_Click()
   q1 = cadbl(q1) + 1
End Sub

Private Sub Command3_Click()
  q2 = cadbl(q2) + 1
End Sub

Private Sub Command4_Click()
q2 = cadbl(q2) - 1
End Sub

Private Sub Command5_Click()
 q3 = cadbl(q3) + 1
End Sub

Private Sub Command6_Click()
  q3 = cadbl(q3) - 1
End Sub

Private Sub Command7_Click()
  q6 = cadbl(q6) - 1
End Sub

Private Sub Command8_Click()
    q6 = cadbl(q6) + 1
End Sub

Private Sub Command9_Click()
    q5 = cadbl(q5) - 1
End Sub

Private Sub etcomanda1_Click()
  assignarcomanda 1
End Sub

Private Sub etcomanda2_Click()
assignarcomanda 2
End Sub

Private Sub etcomanda3_Click()
assignarcomanda 3
End Sub

Private Sub etcomanda4_Click()
assignarcomanda 4
End Sub

Private Sub etcomanda5_Click()
assignarcomanda 5
End Sub

Private Sub etcomanda6_Click()
assignarcomanda 6
End Sub

Function contrasenyacorrecte() As Boolean
  If LCase(InputBox("Entra la contrasenya de modificació", "Atenció")) <> "inplacsa" Then
     contrasenyacorrecte = False
      Else: contrasenyacorrecte = True
  End If
End Function
Sub assignarcomanda(v As Byte)
  Dim vcomanda As String
  If Not contrasenyacorrecte Then Exit Sub
  vcomanda = InputBox("Entra el numero de comanda que vols assignar a aquest adhesiu." + vbNewLine + "Escriu 0 o sense valor per borrar la que hi ha ara.", "Comanda")
  If StrPtr(vcomanda) = 0 Then Exit Sub
  If cadbl(vcomanda) = 0 Then If MsgBox("Vols desassignar el valor de la Comanda associada?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  dbbaixes.Execute "update  estoccintaadhesiva set comanda" + atrim(v) + "=" + atrim(cadbl(vcomanda))
  carregarvalors
  
End Sub
Private Sub Form_Click()
 'enviar_mail_acompres
' guardarvalors
End Sub

Private Sub Form_Load()
 ' mcomandafeta(1).visible = False
  estocadhesiu.tag = "carregant"
  escriure_ini "Muntadora", "tancarfinestraadhesius", "no", rutadelfitxer(cami) + "valorsprograma.ini"
  
  'comprovarestatenvio
  
  'If llegir_ini("Valors", "enviat", fitxerini) = "esperantcomanda" Then
  '      enviar_mail_acompres False  'nomes enviarà un per dia
  'End If
  
End Sub
Sub comprovarestocsminims()
  If estocadhesiu.tag <> "carregant" Then
      guardarvalors
      carregarvalors
  End If
End Sub
Sub framecomanda(vvisible As Boolean)
 '   fcomanda.visible = vvisible
'    lcomanda.visible = vvisible
  '  If vvisible Then
  '       estocadhesiu.width = 5250
  '        Else
  '          estocadhesiu.width = 4700
  '  End If
 '   If llegir_ini("Baixes", "programaamaquina", "comandes.ini") = 1 Then
'       fcomanda.Enabled = False
  '    Else: fcomanda.Enabled = True
  '  End If
End Sub
Sub comprovarestatenvio()
   
   If llegir_ini("Valors", "enviat", fitxerini) = "esperantcomanda" Then
        bcomandafeta.visible = True
        bcomandafeta.caption = "Esperant fer comanda"
        bcomandafeta.BackColor = &H8080FF
        framecomanda True
   End If
   If llegir_ini("Valors", "enviat", fitxerini) = "si" Then
        bcomandafeta.visible = True
        bcomandafeta.caption = "Comanda Feta"
        bcomandafeta.BackColor = &HFF8080
        framecomanda True
   End If
   If llegir_ini("Valors", "enviat", fitxerini) = "no" Then
        bcomandafeta.visible = False
        framecomanda False
   End If
       
End Sub
Sub enviar_mail_acompres(Optional confirmacio As Boolean)
  Dim i As Byte
  Dim cosmissatge As String
  Dim rst As Recordset
  Dim vcalenviar As Boolean
  Dim vassumpte As String
  
  If confirmacio Then
    If MsgBox("S'enviarà un mail a compres per fer la comanda corresponent." + Chr(10) + "Si es correcte i vols enviar-la fer Acceptar", vbInformation + vbYesNo, "Enviar per fer comanda") = vbNo Then Exit Sub
  End If
  cosmissatge = Chr(10) + Chr(10)
  guardarvalors
  Set rst = dbbaixes.OpenRecordset("select * from estoccintaadhesiva")
  For i = 1 To 6
     If UCase(atrim(rst.Fields("nomadhesiu" + atrim(i)))) <> "" Then
        cosmissatge = cosmissatge + atrim(i) + "-" + UCase(atrim(rst.Fields("nomadhesiu" + atrim(i)))) + "-> A: " + atrim(rst.Fields("estoc" + atrim(i))) + "  M: " + atrim(rst.Fields("minim" + atrim(i))) + IIf(cadbl(rst.Fields("comanda" + atrim(i))) > 0, " NC:" + atrim(rst.Fields("comanda" + atrim(i))), "") + Chr(10)
     End If
     If cadbl(rst.Fields("comanda" + atrim(i))) = 0 And (cadbl(rst.Fields("estoc" + atrim(i)))) < cadbl(rst.Fields("minim" + atrim(i))) Then vcalenviar = True
  Next i
  Set rst = Nothing
  'MsgBox cosmissatge
  If vcalenviar Then
    vassumpte = Format(Now, "dd/mm/yy") + " -> S´ha de fer comanda aviseu l´Encarregat "
    If ultimassumpteesdiferent(vassumpte) Then
      enviaremailgeneric "compres@inplacsa.com", vassumpte, cosmissatge
      dbbaixes.Execute "update estoccintaadhesiva set ultimassumpteemail='" + vassumpte + "'"
    End If
    'passaravis 0, 0, Format(Now, "dd/mm/yy") + " -> S´ha de fer comanda aviseu l´Encarregat ", "Adhesius", cosmissatge, 0
  End If
End Sub
Function ultimassumpteesdiferent(vassumpte As String) As Boolean
  Dim rst As Recordset
  Set rst = dbbaixes.OpenRecordset("select ultimassumpteemail from estoccintaadhesiva")
  If rst.EOF Then Exit Function
  If atrim(rst!ultimassumpteemail) <> atrim(vassumpte) Then ultimassumpteesdiferent = True
  Set rst = Nothing
End Function
Sub guardarvalors()
   Dim r As String
   Dim rst As Recordset
   Dim i As Byte
   If llegir_ini("Muntadora", "tancarfinestraadhesius", rutadelfitxer(cami) + "valorsprograma.ini") = "si" Then Exit Sub
   If vtancarestoc Then Exit Sub
   If atrim(estocadhesiu.Controls("adhesiu" + atrim(1))) = "" Then Exit Sub
   dbbaixes.Execute "update estoccintaadhesiva set horaultimaentrada=now"
   Set rst = dbbaixes.OpenRecordset("select * from estoccintaadhesiva")
   If rst.EOF Then
      rst.AddNew
       Else: rst.Edit
   End If
   For i = 1 To 6
      rst.Fields("nomadhesiu" + atrim(i)) = atrim(estocadhesiu.Controls("adhesiu" + atrim(i)))
      rst.Fields("estoc" + atrim(i)) = cadbl(estocadhesiu.Controls("q" + atrim(i)))
      rst.Fields("observacio" + atrim(i)) = atrim(estocadhesiu.Controls("o" + atrim(i)))
   Next i
   rst.Update
   Set rst = Nothing
End Sub
Function tancarfinestraremota(vh As Date) As Boolean
   Dim c As Byte
   If DateDiff("s", vh, Now) < 10 Then
       MsgBox "Hi ha algú editant l'estoc en un altra ordinador espera uns segons i torna-ho a provar", vbCritical, "Error"
       Me.tag = "noguardar"
       tancarfinestraremota = True
       vtancarestoc = True
       Unload Me
       Exit Function
   End If
   MsgBox "Hi ha la finestra de control cinta adhesiva oberta en un altra ordinador, la tancaré d'allà per poder editar els canvis des d'aquest ordinador.", vbCritical, "Atenció"
   c = 0
   escriure_ini "Muntadora", "tancarfinestraadhesius", "si", rutadelfitxer(cami) + "valorsprograma.ini"
   While llegir_ini("Muntadora", "tancarfinestraadhesius", rutadelfitxer(cami) + "valorsprograma.ini") = "si" And c < 5
      wait 1
      c = c + 1
   Wend
   If c > 4 Then MsgBox "No he pogut tancar la finestra remota.", vbCritical, "Error"
End Function
Sub carregarvalors()
   Dim r As String
   Dim i As Byte
   Dim rst As Recordset
   Set rst = dbbaixes.OpenRecordset("select * from estoccintaadhesiva")
   If rst.EOF Then Exit Sub
   If Not IsNull(rst!horaultimaentrada) And estocadhesiu.tag = "carregant" Then If tancarfinestraremota(rst!horaultimaentrada) Then Exit Sub
   If Me.tag = "noguardar" Then Exit Sub
   dbbaixes.Execute "update estoccintaadhesiva set horaultimaentrada=now"
   For i = 1 To 6
      estocadhesiu.Controls("adhesiu" + atrim(i)) = atrim(rst.Fields("nomadhesiu" + atrim(i)))
      estocadhesiu.Controls("q" + atrim(i)) = atrim(rst.Fields("estoc" + atrim(i)))
      estocadhesiu.Controls("q" + atrim(i)).ToolTipText = "Estoc mínim " + atrim(cadbl(rst.Fields("minim" + atrim(i))))
      estocadhesiu.Controls("etcomanda" + atrim(i)).visible = True
      estocadhesiu.Controls("o" + atrim(i)) = atrim(rst.Fields("observacio" + atrim(i)))
      If cadbl(rst.Fields("comanda" + atrim(i))) > 0 Then estocadhesiu.Controls("etcomanda" + atrim(i)) = "Comanda: " + IIf(cadbl(rst.Fields("comanda" + atrim(i))) = 0, "?", atrim(cadbl(rst.Fields("comanda" + atrim(i)))))
      If cadbl(rst.Fields("estoc" + atrim(i))) < cadbl(rst.Fields("minim" + atrim(i))) Then
         estocadhesiu.Controls("q" + atrim(i)).BackColor = QBColor(12)
         estocadhesiu.Controls("etcomanda" + atrim(i)) = "Comanda: " + IIf(cadbl(rst.Fields("comanda" + atrim(i))) = 0, "?", atrim(cadbl(rst.Fields("comanda" + atrim(i)))))
           Else:
             estocadhesiu.Controls("q" + atrim(i)).BackColor = QBColor(15)
             If cadbl(rst.Fields("comanda" + atrim(i))) > 0 Then
               'If estocadhesiu.tag <> "carregant" Then
               ' If MsgBox(UCase(atrim(rst.Fields("nomadhesiu" + atrim(i)))) + " té una comanda associada per reposar estoc." + Chr(10) + "VOLS DESVINCULAR-LA?", vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
               '    If contrasenyacorrecte Then
               '      estocadhesiu.Controls("etcomanda" + atrim(i)) = ""
               '      estocadhesiu.Controls("etcomanda" + atrim(i)).visible = False
               '      rst.Edit: rst.Fields("comanda" + atrim(i)) = 0: rst.Update
               '    End If
               ' End If
               'End If
                Else:
                   estocadhesiu.Controls("etcomanda" + atrim(i)).visible = True
                   estocadhesiu.Controls("etcomanda" + atrim(i)) = "Comanda: " + atrim(cadbl(rst.Fields("comanda" + atrim(i))))
             End If
             
      End If
   Next i
   Set rst = Nothing
End Sub

Private Sub lcomanda_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Me.tag <> "noguardar" And Not vtancarestoc Then
    guardarvalors
    enviar_mail_acompres
    dbbaixes.Execute "update estoccintaadhesiva set horaultimaentrada=null"
  End If
  escriure_ini "Muntadora", "tancarfinestraadhesius", "no", rutadelfitxer(cami) + "valorsprograma.ini"
End Sub

Private Sub minim2_Click()
possarestocminim 2
End Sub

Private Sub minim3_Click()
possarestocminim 3
End Sub

Private Sub minim4_Click()
possarestocminim 4
End Sub

Private Sub minim5_Click()
possarestocminim 5
End Sub

Private Sub minim6_Click()
possarestocminim 6
End Sub

Private Sub mminim1_Click()
   possarestocminim 1
End Sub
Sub possarestocminim(v As Byte)
   Dim r As String
   Dim valor As Double
   Dim rst As Recordset
   Set rst = dbbaixes.OpenRecordset("select * from estoccintaadhesiva")
   If rst.EOF Then rst.AddNew
   valor = cadbl(InputBox("Entra el valor de l'estoc minim per [" + atrim(rst.Fields("nomadhesiu" + atrim(v))) + "]", "Estoc minim", atrim(rst.Fields("minim" + atrim(v)))))
   If valor = 0 Then GoTo fi
   If rst.EOF Then
     rst.AddNew
      Else: rst.Edit
   End If
   rst.Fields("minim" + atrim(v)) = valor
   rst.Update
   comprovarestocsminims
fi:
   Set rst = Nothing
End Sub

Private Sub q1_Change()
  comprovarestocsminims
End Sub

Private Sub q2_Change()
comprovarestocsminims
End Sub

Private Sub q3_Change()
comprovarestocsminims
End Sub

Private Sub q4_Change()
comprovarestocsminims
End Sub

Private Sub q5_Change()
comprovarestocsminims
End Sub

Private Sub q6_Change()
comprovarestocsminims
End Sub
Sub passaravis(p As Double, b As Double, avis, Optional comanda As String, Optional explicacio As String, Optional mtrsajust As Double)
   Dim rutamdb As String
   Dim dbavisos As Database
   Dim rsta As Recordset
   rutamdb = rutadelfitxer(cami) + "avisosincidencies.mdb"
   'MsgBox explicacio
   
   Set dbavisos = DBEngine.OpenDatabase(rutamdb)
'   MsgBox "insert into avisos_baixes  (data,seccio,nomoperari,numoperari,palet,bobina,avis,comanda) values (now,'" + atrim(lletraseccio) + "','" + atrim(Form1.nomoperari) + "','" + atrim(numop) + "','" + atrim(palet) + "','" + atrim(bobina) + "','" + treure_apostruf(avis) + "','" + atrim(comanda) + "')"
   Set rsta = dbavisos.OpenRecordset("select * from avisos_baixes where seccio='M' and comanda='" + atrim(comanda) + "' and avis='" + treure_apostruf(avis) + "'")
   If rsta.EOF Then
    dbavisos.Execute ("insert into avisos_baixes  (data,seccio,nomoperari,numoperari,palet,bobina,avis,comanda,mtrsassignats,mtrsrestants,mtrsgastats,observacio) values (now,'M','Encarregat','0','" + atrim(p) + "','" + atrim(b) + "','" + treure_apostruf(avis) + "','" + atrim(comanda) + "',0," + atrim(cadbl(mtrsrestants)) + "," + atrim(cadbl(mtrsgastats) + cadbl(mtrsajust)) + ",'" + treure_apostruf(explicacio) + "')")
   End If
   Set rsta = Nothing
   dbavisos.Close
   Set dbavisos = Nothing
End Sub

Private Sub Timer1_Timer()
  If estocadhesiu.tag = "carregant" Then
    carregarvalors
    If Me.tag <> "noguardar" And Not vtancarestoc Then
      comprovarestocsminims
      estocadhesiu.tag = ""
    End If
  End If
  If llegir_ini("Muntadora", "tancarfinestraadhesius", rutadelfitxer(cami) + "valorsprograma.ini") = "si" Then
   If estocadhesiu.tag <> "carregant" Then
    Me.tag = "noguardar"
    Unload Me
   End If
  End If
End Sub
