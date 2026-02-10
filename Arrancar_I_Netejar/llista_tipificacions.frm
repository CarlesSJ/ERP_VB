VERSION 5.00
Begin VB.Form Formtipificacions 
   Caption         =   "Estat del clixé"
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5385
   Icon            =   "llista_tipificacions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bok 
      BackColor       =   &H00C0FFC0&
      Height          =   465
      Left            =   510
      Picture         =   "llista_tipificacions.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Acceptar tipificació"
      Top             =   3975
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   5025
      Picture         =   "llista_tipificacions.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Eliminar tipificació"
      Top             =   3960
      Width           =   315
   End
   Begin VB.CommandButton Command1 
      Height          =   285
      Left            =   4740
      Picture         =   "llista_tipificacions.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Afegir Tipificació"
      Top             =   3975
      Width           =   285
   End
   Begin VB.ListBox llistatipificacions 
      BackColor       =   &H00EEE4D7&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      ItemData        =   "llista_tipificacions.frx":1628
      Left            =   45
      List            =   "llista_tipificacions.frx":1632
      TabIndex        =   0
      Top             =   45
      Width           =   5265
   End
End
Attribute VB_Name = "Formtipificacions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bok_Click()
   If llistatipificacions.ListIndex = -1 Then MsgBox "Escull una tipificació primer.", vbCritical, "Error": Exit Sub
   Formtipificacions.Tag = llistatipificacions.List(llistatipificacions.ListIndex)
   Formtipificacions.Hide
End Sub

Private Sub Command1_Click()
  
   Dim resp As String
   If UCase(InputBox("Entra la contrasenya per poder afegir tipificació.", "Contrasenya")) <> "INPLACSA" Then MsgBox "Contrasenya incorrecte", vbCritical, "Error": Exit Sub
   
   resp = UCase(InputBox("Entra la descripció de la tipificació que vols crear.", "Tipificació nova"))
   If resp <> "" Then
       dbbaixes.Execute "insert into neteja_tipificacions (descripcio) values ('" + treure_apostruf(resp) + "')"
   End If
   carregar_tipificacions
End Sub

Private Sub Command2_Click()
 Dim resp As String
   Dim vtipificacio As String
   'If llistatipificacions.ListIndex = -1 Then MsgBox "Primer escull quina tipificació vols eliminar.", vbCritical, "Escull primer": Exit Sub
   vtipificacio = llistatipificacions.List(llistatipificacions.ListIndex)
   If vtipificacio = "- BORRAR -" Or vtipificacio = "- TOT CORRECTE -" Then MsgBox "Aquesta tipificació no es pot eliminar.", vbCritical, "Atenció": Exit Sub
   If UCase(InputBox("Entra la contrasenya per poder afegir tipificació.", "Contrasenya")) <> "INPLACSA" Then MsgBox "Contrasenya incorrecte", vbCritical, "Error": Exit Sub
   
   If MsgBox("Estàs segur que vols eliminar la tipificació? " + Chr(10) + "<" + vtipificacio + ">", vbInformation + vbDefaultButton2 + vbYesNo, "Eliminar") = vbYes Then
      dbbaixes.Execute "delete * from neteja_tipificacions where descripcio='" + vtipificacio + "'"
   End If
   carregar_tipificacions

End Sub

Private Sub Form_Load()
  carregar_tipificacions
End Sub
Sub carregar_tipificacions()
  Dim rst As Recordset
  Set rst = dbbaixes.OpenRecordset("select * from neteja_tipificacions order by descripcio")
  llistatipificacions.Clear
  llistatipificacions.AddItem "- TOT CORRECTE -"
  While Not rst.EOF
    llistatipificacions.AddItem UCase(rst!descripcio)
    rst.MoveNext
  Wend
  llistatipificacions.AddItem "- BORRAR -"
  Set rst = Nothing
End Sub

