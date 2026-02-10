VERSION 5.00
Begin VB.Form formescanejarllaunes 
   Caption         =   "Escanejar llaunes"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9255
   ControlBox      =   0   'False
   Icon            =   "escanejarllaunestinta.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bsortir 
      Height          =   465
      Left            =   8505
      Picture         =   "escanejarllaunestinta.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Tancar"
      Top             =   195
      Width           =   585
   End
   Begin VB.ListBox llistadellaunes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      ItemData        =   "escanejarllaunestinta.frx":0B14
      Left            =   75
      List            =   "escanejarllaunestinta.frx":0B16
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   9030
   End
   Begin VB.CommandButton Command1 
      Height          =   465
      Left            =   4575
      Picture         =   "escanejarllaunestinta.frx":0B18
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Acceptar llauna"
      Top             =   150
      Width           =   585
   End
   Begin VB.TextBox cnumllauna 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2625
      TabIndex        =   0
      Top             =   165
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Nº de Llauna:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   270
      TabIndex        =   1
      Top             =   180
      Width           =   2580
   End
End
Attribute VB_Name = "formescanejarllaunes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bsortir_Click()
  If llistadellaunes.ListCount > 0 Then
    If MsgBox("Hi ha dades entrades, vols sortir sense guardar-les?", vbCritical + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then cnumllauna.SetFocus: Exit Sub
  End If
  Unload Me
End Sub

Private Sub cnumllauna_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      afegir_llauna cnumllauna
      KeyCode = 0
      cnumllauna = ""
   End If
End Sub
Sub afegir_llauna(vnumllaunacontrol As String)
  Dim rst As Recordset
  Dim rstp As Recordset
  Dim vmsg As String
  Dim vmsg2 As String
  Dim i As Integer
  Dim j As Integer
  Dim vnumllauna As String
  Dim vnomlot As String
  Dim vnumtinter As Integer
  Dim vllistalots(50) As String
  Dim vesbona As Boolean
  Dim vnumtintervisual As String
  Dim vprimeralletra As String
  Dim vtintarelacionada As String
  Dim vnomproveidor As String
  
  vprimeralletra = Mid(atrim(UCase(vnumllaunacontrol)) + " ", 1, 1)
  If vprimeralletra <> "A" And vprimeralletra <> "I" Then
    cnumllauna = "": cnumllauna.SetFocus
    vmsg = "#Error: S'ha escanejat una fórmula.  " + vnumllaunacontrol
    GoTo fi
  End If
  vnumllauna = form1.agafarellotdelcomponent("#" + vnumllaunacontrol, vnomlot, vtintarelacionada)
  valtreslots = form1.buscarlots(vnomlot, vnumllaunacontrol, vllistalots)
  If vtintarelacionada = "" Then
    If valtreslots <> "" Then vnumllauna = vnumllauna + "+" + valtreslots
      Else: vnumllauna = UCase(vnumllaunacontrol)
  End If
  i = 0
  While (vllistalots(i) <> "" And vmsg = "") Or vtintarelacionada <> ""
        If vtintarelacionada = "" Then
         vnumllauna = vllistalots(i)
         Set rst = dbtintes.OpenRecordset("select * from dadesllaunestotes where numllauna='" + atrim(vnumllauna) + "'")
          Else:
             Set rst = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(vtintarelacionada) + "'")
             vtintarelacionada = ""
             i = -1
        End If
        If rst.EOF Then GoTo proxima 'vmsg = "#Error No existeix la llauna": GoTo fi
        vnumtinter = 0
        For j = 0 To formaniloxos.tintacomanda.Count - 1
           vesbona = False
           If cadbl(formaniloxos.tintacomanda.Item(j).tag) = cadbl(rst!codi) Then
                vesbona = True
             Else:
                If latintaesiguala(cadbl(formaniloxos.tintacomanda.Item(j).tag), cadbl(rst!codi)) Then
                      'comprovo que si hi ha un @ a la tinta el proveidor de la tinta ha de ser el que diu despres de @
                   If InStr(1, atrim(formaniloxos.tintacomanda.Item(j)), "@") > 0 Then
                     vnomproveidor = Mid(atrim(formaniloxos.tintacomanda.Item(j)), InStr(1, atrim(formaniloxos.tintacomanda.Item(j)), "@") + 1)
                     vnomproveidor = Trim(Mid(vnomproveidor, 1, InStr(1, atrim(rst!descripcio) + " ", " ") - 1))
                     Set rstp = dbtintes.OpenRecordset("select nomproveidor from dadesllaunestotes where numllauna='" + atrim(vnumllauna) + "'")
                     If Not rstp.EOF Then
                         If InStr(1, rstp!nomproveidor, vnomproveidor) = 0 Then
                              GoTo alarmatintaerroneaaldosificador
                         End If
                     End If
                   End If
                   vesbona = True  'es bona si arriba aqui si no haurà saltat l'alarma
                End If
                If hihatintaalternativa(cadbl(rst!codi), cadbl(formaniloxos.ordre(j).tag)) Then vesbona = True
           End If
           If vesbona Then
                If vnumtinter = 0 Then
                     vnumtinter = j + 1
                     vnumtintervisual = atrim(vnumtinter)
                    Else
                       vnumtinter = cadbl(atrim(vnumtinter) + atrim(j + 1))
                       vnumtintervisual = vnumtintervisual + "," + atrim(j + 1)
                End If
                vmsg = "#" + atrim(vnumtintervisual) + " " + atrim(rst!descripcio)
           End If
        Next j
proxima:
        i = i + 1
  Wend
  If vmsg = "" Then
       vnumllauna = vnumllaunacontrol
       vmsg = "#Error- " + nomdelatinta(vnumllauna)
    Else: vmsg2 = generarliniadecodis(vnumllauna, vllistalots)
  End If
fi:
  llistadellaunes.AddItem vnumllauna + " - " + vmsg
  llistadellaunes.ItemData(llistadellaunes.NewIndex) = vnumtinter
  If vmsg2 <> "" Then
     llistadellaunes.AddItem "          - " + vmsg2
     llistadellaunes.ItemData(llistadellaunes.NewIndex) = vnumtinter
  End If
  If InStr(1, vmsg, "#Error") > 0 Then
      llistadellaunes.BackColor = QBColor(12)
      sonar_sirena "continuu"
       Else: sonar_sirena "intermitent"
  End If
  Set rst = Nothing
  Set rstp = Nothing
  Exit Sub
alarmatintaerroneaaldosificador:
  While MsgBox("Error de proveidor al dosificador amb la tinta " + Chr(10) + atrim(rst!descripcio) + " --> " + atrim(rstp!nomproveidor) + Chr(10) + "HAURIA DE SER " + vnomproveidor, vbSystemModal + vbCritical + vbOKCancel + vbDefaultButton2, "ERROR DE PROVEIDOR") = vbCancel
    DoEvents
  Wend
  While UCase(r) <> "PARAR LA COMANDA"
    r = InputBox("S'HA DE PARAR LA COMANDA FINS QUE ES CANVI EL PROVEIDOR DE LA TINTA " + atrim(rst!descripcio) + Chr(10) + "DEL DOSIFICADOR " + vnumllaunacontrol + Chr(10) + Chr(10) + " ESCRIU [PARAR LA COMANDA] PER CONTINUAR", "ERROR NO ES POT CONTINUAR")
  Wend
  Set rstp = Nothing
End Sub
Function hihatintaalternativa(vcoditinta As Double, vid_tinter As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from tintes_alternatives where id_tinter=" + atrim(vid_tinter))
   If Not rst.EOF Then
       'If Not rst.EOF Then
       '   rst.FindFirst "coditinta='" + atrim(vcoditinta) + "'"
       '   If Not rst.NoMatch Then hihatintaalternativa = True
       'End If
       While Not rst.EOF
          If latintaesiguala(vcoditinta, rst!coditinta) Then hihatintaalternativa = True
          rst.MoveNext
       Wend
   End If
   If hihatintaalternativa = False Then
       Set rst = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(vid_tinter))
       If Not rst.EOF Then
            If Not rst.EOF Then
                rst.FindFirst "coditinta='" + atrim(vcoditinta) + "'"
                If Not rst.NoMatch Then hihatintaalternativa = True
            End If
       End If
   End If
   Set rst = Nothing
End Function
Function nomdelatinta(vnumllauna As String) As String
  Dim rst As Recordset
  If UCase(Mid(vnumllauna + "  ", 1, 1)) = "I" Then
   Set rst = dbtintes.OpenRecordset("select nomcomponent from componentsbase where numdosificador=" + Mid(vnumllauna, 2))
   If Not rst.EOF Then nomdelatinta = UCase(atrim(rst!nomcomponent))
   GoTo fi
  End If
  Set rst = dbtintes.OpenRecordset("select * from dadesllaunes where numllauna='" + atrim(vnumllauna) + "'")
  If Not rst.EOF Then
       nomdelatinta = atrim(rst!descripcio)
   End If
fi:
   Set rst = Nothing
End Function
Function latintaesiguala(vcodi1 As Double, vcodi2 As Double) As Boolean
  Dim rst1 As Recordset
  Dim rst2 As Recordset
  Dim vsql As String
  Set rst1 = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(vcodi1) + "'")
  Set rst2 = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(vcodi2) + "'")
  If rst1.EOF Or rst2.EOF Then GoTo fi
  If InStr(1, rst1!referenciacolor, "P-") Then GoTo fi
  'If rst1!idserie = rst2!idserie And
  If rst1!idfamilia = rst2!idfamilia And rst1!idsubfamilia = rst2!idsubfamilia Then
      If rst1!idfamcolor = rst2!idfamcolor And rst1!idsubfamcolor = rst2!idsubfamcolor Then
         latintaesiguala = True
      End If
  End If
fi:
  Set rst1 = Nothing
  Set rst2 = Nothing
End Function
Function generarliniadecodis(vnumllauna As String, vllistalots As Variant) As String
  Dim i As Integer
  While vllistalots(i) <> ""
     If vllistalots(i) <> vnumllauna Then generarliniadecodis = generarliniadecodis + " " + vllistalots(i)
     i = i + 1
  Wend
End Function

Private Sub cnumllauna_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Command1_Click()
   If cnumllauna = "" Then
     If llistadellaunes.BackColor = QBColor(12) Then
       If MsgBox("Hi ha un ERROR amb les llaunes estàs d'acord en acceptar les llaunes que son correctes?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then cnumllauna.SetFocus: Exit Sub
     End If
     acceptar_llaunes
        Else: afegir_llauna cnumllauna: cnumllauna = ""
   End If
   
End Sub
Sub acceptar_llaunes()
  Dim j As Integer
  For j = 0 To llistadellaunes.ListCount - 1
   If Mid(llistadellaunes.List(j), 1, 12) <> "          - " Then
    guardar_llaunesdelafila j
   End If
  Next j
  
  Unload Me
End Sub
Sub possartinters(v As String, vtinters As Variant)
  Dim j As Byte
  v = atrim(Mid(v, 1, InStr(1, v, " ")))
  v = substituir(v, ",", " ")
  If cadbl(v) > 0 Then vtinters(j) = v: j = j + 1
  While InStr(1, v, " ")
    
    vtinters(j) = cadbl(Mid(v, 1, InStr(1, v, " ")))
    v = atrim(Mid(v, InStr(1, v, " ")))
    If Len(v) > 0 Then v = v + " "
    j = j + 1
  Wend
End Sub
Sub guardar_llaunesdelafila(vfila As Integer)
  Dim vllista(100) As String
  Dim vlinia As String
  Dim vtinters(10) As Byte
  Dim vnumc As Double
  Dim vidtinter As Double
  Dim j As Integer
  vlinia = llistadellaunes.List(vfila) + "  "
  If InStr(1, v, "#Error") > 0 Then Exit Sub
  ' cadbl(Mid(vlinia, InStr(1, vlinia, "#"), 2))
  possartinters Mid(vlinia, InStr(1, vlinia, "#") + 1), vtinters
  vnumc = cadbl(form1.comanda)
  j = 1
  vllista(0) = atrim(Mid(vlinia, 1, InStr(2, vlinia, " - ")))
  If Mid(vllista(0), 1, 1) = "I" Then vllista(0) = " "
  If vfila + 1 < llistadellaunes.ListCount Then
      vlinia = llistadellaunes.List(vfila + 1) + "  "
      If Mid(vlinia, 1, 12) = "          - " Then
       vlinia = substituir(vlinia, "          - ", "")
       While atrim(vlinia) <> ""
         vllista(j) = atrim(Mid(vlinia, 1, InStr(1, LTrim(vlinia), " ")))
             vlinia = atrim(Mid(vlinia, InStr(1, vlinia, " ") + 1)) + "  "
         j = j + 1
       Wend
      End If
  End If
  j = 0
  While cadbl(vtinters(j)) > 0 And j < 10
    i = 0
    ' If vllista(0) <> "" Then dbtmpb.Execute "delete * from impresores_llaunesgastades where comanda=" + atrim(vnumc) + " and tinter=" + atrim(vtinters(j)) + " and tipus='I'"
     While vllista(i) <> ""
       dbtmpb.Execute "delete * from impresores_llaunesgastades where comanda=" + atrim(vnumc) + " and tinter=" + atrim(vtinters(j)) + " and tipus='I' and numllauna='" + atrim(vllista(i)) + "'"
       vidtinter = cadbl(formaniloxos.ordre(cadbl(vtinters(j) - 1)).tag)
       dbtmpb.Execute "insert into impresores_llaunesgastades (numllauna,tinter,comanda,tipus,id_tinter) values ('" + vllista(i) + "'," + atrim(vtinters(j)) + "," + atrim(vnumc) + ",'I'," + atrim(vidtinter) + ")"
       If i = 0 Then canviarTINTA_llistadetintersFORMANILOXOS vllista(i), vtinters(j)
       i = i + 1
     Wend
    
    j = j + 1
  Wend
End Sub
Sub canviarTINTA_llistadetintersFORMANILOXOS(vnumllauna As String, vnumtinter As Byte)
   Dim rst As Recordset
   
   Set rst = dbtintes.OpenRecordset("select * from dadesllaunestotes where numllauna='" + atrim(vnumllauna) + "'")
   If Not rst.EOF Then
        formaniloxos.tintacomanda(vnumtinter - 1) = atrim(rst!descripcio)
        formaniloxos.tintacomanda(vnumtinter - 1).tag = atrim(rst!codi)
   End If
   Set rst = Nothing
End Sub
Private Sub Form_Activate()
  cnumllauna.SetFocus
End Sub

Private Sub llistadellaunes_GotFocus()
  cnumllauna.SetFocus
End Sub
