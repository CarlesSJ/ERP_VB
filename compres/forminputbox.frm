VERSION 5.00
Begin VB.Form forminputbox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Escriu una descripció "
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8505
   Icon            =   "forminputbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Height          =   285
      Left            =   75
      Picture         =   "forminputbox.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Afegir referencia"
      Top             =   930
      Width           =   285
   End
   Begin VB.CommandButton borrarliniesdescripcio 
      Height          =   285
      Left            =   75
      Picture         =   "forminputbox.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Borrar referencia"
      Top             =   1215
      Width           =   285
   End
   Begin VB.CommandButton bcupo 
      Height          =   435
      Index           =   5
      Left            =   5865
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1005
      Width           =   1095
   End
   Begin VB.CommandButton bcupo 
      Height          =   435
      Index           =   4
      Left            =   4770
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1005
      Width           =   1095
   End
   Begin VB.CommandButton bcupo 
      Height          =   435
      Index           =   3
      Left            =   3675
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1005
      Width           =   1095
   End
   Begin VB.CommandButton bcupo 
      Height          =   435
      Index           =   2
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1005
      Width           =   1095
   End
   Begin VB.CommandButton bcupo 
      Height          =   435
      Index           =   1
      Left            =   1485
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1005
      Width           =   1095
   End
   Begin VB.CommandButton bcupo 
      Height          =   435
      Index           =   0
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1005
      Width           =   1095
   End
   Begin VB.CommandButton bacceptar 
      Height          =   360
      Left            =   7215
      Picture         =   "forminputbox.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Acceptar canvis"
      Top             =   1020
      Width           =   930
   End
   Begin VB.TextBox cresposta 
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   525
      Width           =   8115
   End
   Begin VB.Label etmissatge 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   450
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   7515
   End
End
Attribute VB_Name = "forminputbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bacceptar_Click()
   bacceptar.tag = "1"
  Me.Hide
End Sub

Private Sub bcupo_Click(index As Integer)
  If bcupo(index).BackColor = QBColor(12) Then
     comandescompra.capcalera.Database.Execute "delete * from referenciescupusproveidors where descripcio='" + bcupo(index).caption + "' and codiproveidor=" + atrim(comandescompra.capcalera.Recordset!codiproveidor)
     possar_nom_codiscupus
     netejar_botons
     Exit Sub
  End If
   cresposta = "COD. " + bcupo(index).caption
End Sub
Sub netejar_botons()
  Dim i As Byte
  For i = 0 To bcupo.Count - 1
    bcupo(i).BackColor = &H8000000F
    If bcupo(i).caption = "" Then
      bcupo(i).visible = False
        Else: bcupo(i).visible = True
    End If
  Next i
End Sub

Private Sub borrarliniesdescripcio_Click()
  Dim i As Byte
  If bcupo(index).BackColor = QBColor(12) Then netejar_botons: Exit Sub
  For i = 0 To bcupo.Count - 1
    bcupo(i).BackColor = QBColor(12)
  Next i
  MsgBox "Prem sobre la referencia que vols eliminar", vbExclamation, "Atenció"
End Sub

Private Sub Command1_Click()
   
End Sub

Private Sub Command2_Click()
   Dim resp As String
   Dim rst As Recordset
    Dim dbcompres As Database
   resp = UCase(InputBox("Entra la nova referència que vols utilitzar.", "Nova refrencia"))
   If resp = "" Then Exit Sub
   If Len(resp) > 10 Then MsgBox "Aquesta referencia es massa llarga no pot passar de 10 caracters", vbCritical, "Error": Exit Sub
   Set dbcompres = comandescompra.capcalera.Database
   Set rst = dbcompres.OpenRecordset("Select * from referenciescupusproveidors where codiproveidor=" + atrim(comandescompra.capcalera.Recordset!codiproveidor))
   rst.FindFirst "descripcio='" + atrim(resp) + "'"
   If Not rst.NoMatch Then MsgBox "Aquesta referencia ja existeix", vbCritical, "Error": Exit Sub
   dbcompres.Execute "insert into referenciescupusproveidors (codiproveidor,descripcio) values (" + atrim(comandescompra.capcalera.Recordset!codiproveidor) + ",'" + treure_apostruf(resp) + "')"
   possar_nom_codiscupus
   netejar_botons
   Set dbcompres = Nothing
   Set rst = Nothing
End Sub

Private Sub Form_Load()
  possar_nom_codiscupus
  netejar_botons
  
End Sub
Sub possar_nom_codiscupus()
 Dim rst As Recordset
 Dim dbcompres As Database
 Dim i As Byte
 Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
 Set rst = dbcompres.OpenRecordset("Select * from referenciescupusproveidors where codiproveidor=" + atrim(comandescompra.capcalera.Recordset!codiproveidor))
 i = 0
 While i < 6
   If Not rst.EOF Then
    bcupo(i).caption = UCase(rst!descripcio)
    rst.MoveNext
     Else: bcupo(i).caption = ""
   End If
   i = i + 1
   
 Wend
 Set rst = Nothing
 Set dbcompres = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Screen.ActiveForm.Name = "forminputbox" Then
   Cancel = 1
   cresposta.Text = ""
   Me.Hide
  End If
End Sub
