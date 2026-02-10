VERSION 5.00
Begin VB.Form formmeslots 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mes Lots"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   " Auto Scan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Acceptar canvis"
      Top             =   435
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      Height          =   435
      Left            =   2370
      Picture         =   "formdemanameslots.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Acceptar canvis"
      Top             =   0
      Width           =   1275
   End
   Begin VB.TextBox vnumlot 
      Alignment       =   2  'Center
      BackColor       =   &H00EAD9CE&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   75
      TabIndex        =   1
      Top             =   435
      Width           =   2295
   End
   Begin VB.ListBox llistallaunes 
      BackColor       =   &H00F3B378&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   75
      TabIndex        =   0
      Top             =   975
      Width           =   3540
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Guió davant per canvi de LOT"
      ForeColor       =   &H00ED823A&
      Height          =   240
      Left            =   -225
      TabIndex        =   6
      Top             =   0
      Width           =   2565
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "* per LOT manual."
      ForeColor       =   &H00ED823A&
      Height          =   240
      Left            =   165
      TabIndex        =   4
      Top             =   195
      Width           =   2145
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2400
      Picture         =   "formdemanameslots.frx":058A
      Stretch         =   -1  'True
      Top             =   450
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "Nº de LOT"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   195
      Width           =   1740
   End
End
Attribute VB_Name = "formmeslots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim vmsg As String
  If formmeslots.tag <> "noautocarregar" Then
    If Not hihatotselslotsdelscomponents(vmsg) Then
         If UCase(InputBox("Falten lots de: " + Chr(10) + vmsg + Chr(10) + Chr(10) + "Escriu la contrasenya per continuar igualment.", "Error")) <> "INPLACSA" Then Exit Sub
    End If
  End If
  formmeslots.Hide
End Sub
Function hihatotselslotsdelscomponents(vmsg As String) As Boolean
   Dim vformula As String
   Dim rst As Recordset
   Dim rstlots As Recordset
   Dim rsttinta As Recordset
   Dim i As Long
   Dim vcoditinta As Double
   Dim vcoditrobat As Boolean
   hihatotselslotsdelscomponents = True
   vformula = formtintes.possarformulapredeterminada(formtintes.tintes.Recordset!idtinta)
   If vformula = "" Then GoTo fi
   Set rst = dbtintes.OpenRecordset("select idformula from formules where codiformula='" + atrim(vformula) + "'")
   If Not rst.EOF Then
      Set rst = dbtintes.OpenRecordset("SELECT  DetallFormules.[%decomponent],DetallFormules.IDFormula, DetallFormules.IdComponente, Componentsbase.nomcomponent FROM DetallFormules INNER JOIN Componentsbase ON DetallFormules.IdComponente = Componentsbase.idcomponent Where idformula = " + atrim(rst!idformula))
      While Not rst.EOF
        Set rstlots = dbtintes.OpenRecordset("select numerodelot from detallnumeroslotsbase where idcomponent=" + atrim(rst!idcomponente) + " order by data desc")
        If Not rstlots.EOF Then
            Set rsttinta = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, tintes.codi FROM Llaunes INNER JOIN tintes ON Llaunes.idtinta = tintes.idtinta where numllauna='" + atrim(rstlots!numerodelot) + "'")
            vcoditrobat = False
            If Not rsttinta.EOF Then
                vcoditinta = rsttinta!codi
                For i = 0 To llistallaunes.ListCount - 1
                    Set rsttinta = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, tintes.codi FROM Llaunes INNER JOIN tintes ON Llaunes.idtinta = tintes.idtinta where numllauna='" + atrim(llistallaunes.List(i)) + "'")
                    If Not rsttinta.EOF Then If vcoditinta = cadbl(rsttinta!codi) Then vcoditrobat = True
                Next i
                If Not vcoditrobat Then vmsg = vmsg + Chr(10) + "Falta -> " + atrim(rst![%decomponent]) + "% de " + rst!nomcomponent
            End If
        End If
        rst.MoveNext
      Wend
   End If
fi:
   If vmsg <> "" Then hihatotselslotsdelscomponents = False
   Set rst = Nothing
   Set rsttinta = Nothing
   Set rstlots = Nothing
End Function


Private Sub Command2_Click()
   autoescan
End Sub
Sub autoescan()
   Dim vformula As String
   Dim rst As Recordset
   Dim rstlots As Recordset
   
   vformula = formtintes.possarformulapredeterminada(formtintes.tintes.Recordset!idtinta)
   If vformula = "" Then MsgBox "No hi ha cap formula relacionada.", vbCritical, "Error": Exit Sub
   Set rst = dbtintes.OpenRecordset("select idformula from formules where codiformula='" + atrim(vformula) + "'")
   If Not rst.EOF Then
      Set rst = dbtintes.OpenRecordset("select * from detallformules where idformula=" + atrim(rst!idformula))
      While Not rst.EOF
        Set rstlots = dbtintes.OpenRecordset("select numerodelot from detallnumeroslotsbase where idcomponent=" + atrim(rst!idcomponente) + " order by data desc")
        If Not rstlots.EOF Then If nohihaellot(UCase(rstlots!numerodelot)) Then llistallaunes.AddItem UCase(rstlots!numerodelot)
        rst.MoveNext
      Wend
   End If
   Set rst = Nothing
End Sub
Function nohihaellot(vnumlot As String) As Boolean
   Dim i As Long
   nohihaellot = True
   For i = 0 To llistallaunes.ListCount - 1
     If UCase(llistallaunes.List(i)) = UCase(vnumlot) Then nohihaellot = False
   Next i
End Function
Private Sub Form_Activate()
  vnumlot.SetFocus
  If formmeslots.tag <> "noautocarregar" Then autoescan
End Sub

Private Sub Image1_Click()
  vnumlot.SetFocus
End Sub

Private Sub llistallaunes_DblClick()
   If llistallaunes.ListIndex < 0 Then Exit Sub
   llistallaunes.RemoveItem llistallaunes.ListIndex
End Sub

Private Sub llistallaunes_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Then
    If llistallaunes.ListIndex < 0 Then Exit Sub
    llistallaunes.RemoveItem llistallaunes.ListIndex
  End If
End Sub

Private Sub vnumlot_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then KeyAscii = 0: afegir_lot_alallista (vnumlot): vnumlot = ""
End Sub
Function agafarellotdelcomponent(vdosificador As String, Optional vnomlot As String, Optional vcoditinta As String) As String
  Dim rst As Recordset
  vdosificador = UCase(vdosificador)
  If InStr(1, vdosificador, "A") Then agafarellotdelcomponent = UCase(vdosificador): GoTo fi
  vdosificador = substituir("  " + vdosificador, "I", "")
  If cadbl(vdosificador) = 0 Then GoTo fi
  sql = "SELECT Componentsbase.nomcomponent AS nomlot,componentsbase.coditintarelacionada, detallnumeroslotsbase.numerodelot AS codilot FROM detallnumeroslotsbase INNER JOIN Componentsbase ON detallnumeroslotsbase.idcomponent = Componentsbase.idcomponent "
  sql = sql + " WHERE Componentsbase.numdosificador=" + atrim(vdosificador) + " order by data DESC"
  Set rst = dbtintes.OpenRecordset(sql)
  If Not rst.EOF Then agafarellotdelcomponent = atrim(rst!codilot)
fi:
  Set rst = Nothing
End Function
Sub substituir_lot(vnumloterroni As String)

End Sub
Public Function CnvDec(ByVal S As String) As Double
    Dim P As Integer
    Dim N As String
    N = S
    P = InStr(N, ",")
    Do While P > 0
        Mid(N, P, 1) = "."
        P = InStr(N, ",")
    Loop
    CnvDec = Val(N)
End Function
Sub afegir_lot_alallista(vnum As String)
 Dim i As Integer
 Dim vkg As Double
 Dim v As String
 i = 0
 If Mid(vnum + "  ", 1, 1) = "-" Then vnumlot.tag = vnum: formmeslots.Hide
 If Mid(vnum + "  ", 1, 1) <> "*" Then
    vnum = agafarellotdelcomponent(vnum)
 End If

 'afegeig el lot si no hi era a la llista
 If atrim(vnum) <> "" Then
    If nohihaellot(treure_apostruf(vnum)) Then
          While vkg = 0
            v = InputBox("Entra els kg de tinta que afegeixes, no val zero." + vbNewLine + "(Es un valor estadistic)", "Kg de tinta afegits")
            vkg = CnvDec(v)
          Wend
          llistallaunes.AddItem treure_apostruf(vnum)
          llistallaunes.ItemData(llistallaunes.NewIndex) = Redondejar(vkg, 2) * 100
    End If
 End If
fi:
End Sub
