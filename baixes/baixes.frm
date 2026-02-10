VERSION 5.00
Begin VB.Form entradabaixes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Baixes de Seccions"
   ClientHeight    =   1515
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   7380
   Icon            =   "baixes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Proves"
      Height          =   270
      Left            =   45
      TabIndex        =   16
      Top             =   15
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CheckBox borrar 
      Caption         =   "Borrar Seccio"
      Height          =   210
      Left            =   5505
      TabIndex        =   12
      Top             =   0
      Width           =   2130
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5130
      Top             =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   210
      Width           =   7275
      Begin VB.CommandButton estat3 
         Caption         =   "V"
         Height          =   285
         Left            =   3570
         TabIndex        =   15
         Top             =   375
         Width           =   255
      End
      Begin VB.CommandButton estat2 
         Caption         =   "V"
         Height          =   285
         Left            =   6930
         TabIndex        =   14
         Top             =   465
         Width           =   255
      End
      Begin VB.CommandButton estat1 
         Caption         =   "V"
         Height          =   285
         Left            =   6930
         TabIndex        =   13
         Top             =   135
         Width           =   255
      End
      Begin VB.CommandButton linkcomanda2 
         BackColor       =   &H0000FF00&
         Height          =   300
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   450
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton linkcomanda1 
         BackColor       =   &H0000FF00&
         Height          =   300
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Acceptar"
         Height          =   420
         Left            =   3855
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   270
         Width           =   1230
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Entrega"
         Height          =   285
         Left            =   6060
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "V"
         Top             =   762
         Width           =   1185
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Soldadores"
         Height          =   285
         Left            =   4860
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "S"
         Top             =   762
         Width           =   1185
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rebobinadora"
         Height          =   285
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "R"
         Top             =   762
         Width           =   1185
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Laminadores"
         Height          =   285
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "L"
         Top             =   762
         Width           =   1185
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Impressores"
         Height          =   285
         Left            =   1305
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "I"
         Top             =   762
         Width           =   1185
      End
      Begin VB.TextBox comanda 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1875
         TabIndex        =   2
         Text            =   "117704"
         Top             =   360
         Width           =   1650
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Extrussores"
         Height          =   285
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "E"
         Top             =   762
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Comanda:"
         Height          =   270
         Left            =   825
         TabIndex        =   3
         Top             =   405
         Width           =   975
      End
   End
   Begin VB.Menu pasaranoacavada 
      Caption         =   "Pasar seccio a NoAcavada"
   End
   Begin VB.Menu mpassarbobinesaacabades 
      Caption         =   "Passar bobines a acabades"
   End
   Begin VB.Menu mllistats 
      Caption         =   "Llistats"
      Begin VB.Menu mllistatentregues 
         Caption         =   "Entregues per dates"
      End
      Begin VB.Menu entregadatesavançat 
         Caption         =   "Entregues per dates (Avançat)"
      End
   End
End
Attribute VB_Name = "entradabaixes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbvendes As Database
Private Sub borrar_Click()
 If borrar = 1 Then MsgBox "Ara clica sobre la secció que vols eliminar", vbCritical + vbOKOnly, "Atenció"

End Sub

Private Sub comanda_Change()
 If Len(comanda) = 6 Then
   comprovar_comanda
 End If
End Sub
Sub comprovar_comanda()
   Dim rst As Recordset
   Dim rst2 As Recordset
   Set rst = dbtmp.OpenRecordset("select producte,proximaseccio,seccioactual,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(cadbl(comanda)))
   possar_ruta ""
   If Not rst.EOF Then
      Set rst2 = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
      If Not rst2.EOF Then possar_ruta rst2!ruta, rst!proximaseccio: Label1.Tag = rst2!ruta: pr = atrim(rst!proximaseccio): sa = atrim(rst!seccioactual)
      linkcomanda1.Caption = atrim(IIf(cadbl(rst!linkcomanda1) > 0, rst!linkcomanda1, ""))
      linkcomanda2.Caption = atrim(IIf(cadbl(rst!linkcomanda2) > 0, rst!linkcomanda2, ""))
      linkcomanda1.Visible = IIf(linkcomanda1.Caption = "", False, True)
      linkcomanda2.Visible = IIf(linkcomanda2.Caption = "", False, True)
      Set rst2 = dbtmp.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(cadbl(rst!linkcomanda1)))
      estat1.Caption = "": estat2.Caption = ""
      If cadbl(rst!linkcomanda1) > 0 Then estat1.Caption = atrim(rst2!proximaseccio)
      estat1.Visible = IIf(estat1.Caption = "", False, True)
      Set rst2 = dbtmp.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(cadbl(rst!linkcomanda2)))
      
      If cadbl(rst!linkcomanda2) > 0 Then estat2.Caption = atrim(rst2!proximaseccio)
      estat2.Visible = IIf(estat2.Caption = "", False, True)
      
      Set rst2 = dbtmp.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(cadbl(comanda)))
      estat3.Caption = atrim(rst2!proximaseccio)
      
      'If InStr(1, rst!producte, "PC") Then Command6.BackColor = QBColor(4): Command6.Enabled = False
      
   End If
   Set rst = Nothing
End Sub
Sub possar_ruta(ruta As String, Optional seccio As String)
   Dim objecte As Object
   Dim color As Byte
   ruta = ruta + "VPT"
   If InStr(1, "VPT", seccio) Then
      Command6.Tag = seccio
       Else: Command6.Tag = "V"
   End If
   If ruta = "" Then
     For Each objecte In entradabaixes
       If objecte.Tag <> "" Then objecte.BackColor = QBColor(8) ': objecte.Enabled = False
     Next
     Exit Sub
   End If
       For Each objecte In entradabaixes
        If TypeOf objecte Is CommandButton Then
         If objecte.Tag <> "" Then
          If InStr(1, ruta, objecte.Tag) Then
             objecte.Enabled = True
             color = 14
             If InStr(1, ruta, objecte.Tag) > InStr(1, ruta, seccio) Then color = 4
             If InStr(1, ruta, objecte.Tag) = InStr(1, ruta, seccio) Then color = 12
             objecte.BackColor = QBColor(color)
             'objecte.Enabled = IIf(color <> 4, True, False)
            Else: If objecte.Tag <> "" Then objecte.BackColor = QBColor(8): objecte.Enabled = False
          End If
         End If
        End If
       Next
       If InStr(1, "T", seccio) Then Command6.BackColor = QBColor(14)
       Command6.Enabled = True
End Sub
Private Sub comanda_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 112 Then KeyCode = 0: Command7_Click ' acceptar_boto
End Sub

Private Sub comanda_LostFocus()
  'escriure_ini "Baixes", "ultimacomanda", comanda.Text, fitxerini
End Sub

Private Sub Command1_Click()
 On Error Resume Next
 Unload Extrussores
 On Error GoTo 0
 If borrar = 1 Then borrar_seccio "Ext": Exit Sub
 If Not comprovar_seccio Then Exit Sub

 Extrussores.Show

  
End Sub
Sub borrar_seccio(seccio As String)
Dim dbtmpb As Database
 borrar = 0
 If sa <> Mid(seccio, 1, 1) And sa <> "V" Then
    MsgBox "Aquesta secció no es la última i no es pot borrar"
    Exit Sub
 End If
 Set dbtmpb = OpenDatabase(cami)
 If seccio = "Ext" Then
    Set rsttmp = dbtmpb.OpenRecordset("select id from extrussores where comanda=" + atrim(cadbl(comanda.Text)))
    While Not rsttmp.EOF
     dbtmpb.Execute "delete * from bobinesext where controlid=" + atrim(rsttmp!ID)
     rsttmp.MoveNext
    Wend
    dbtmpb.Execute "delete * from extrussores where comanda=" + atrim(cadbl(comanda.Text))
    dbtmp.Execute "update comandes set proximaseccio='E' where comanda=" + atrim(cadbl(comanda.Text))
 End If
 If seccio = "Imp" Then
    Set rsttmp = dbtmpb.OpenRecordset("select id from impressores where comanda=" + atrim(cadbl(comanda.Text)))
    While Not rsttmp.EOF
     dbtmpb.Execute "delete * from bobinesimp where controlid=" + atrim(rsttmp!ID)
     rsttmp.MoveNext
    Wend
    dbtmpb.Execute "delete * from impressores where comanda=" + atrim(cadbl(comanda.Text))
    dbtmp.Execute "update comandes set proximaseccio='I' where comanda=" + atrim(cadbl(comanda.Text))
 End If
 If seccio = "Lam" Then
    Set rsttmp = dbtmpb.OpenRecordset("select id from laminadores where comanda=" + atrim(cadbl(comanda.Text)))
    While Not rsttmp.EOF
     dbtmpb.Execute "delete * from bobineslam where controlid=" + atrim(rsttmp!ID)
     rsttmp.MoveNext
    Wend
    dbtmpb.Execute "delete * from laminadores where comanda=" + atrim(cadbl(comanda.Text))
    dbtmp.Execute "update comandes set proximaseccio='L' where comanda=" + atrim(cadbl(comanda.Text))
 End If
 If seccio = "Reb" Then
    Set rsttmp = dbtmpb.OpenRecordset("select id from rebobinadores where comanda=" + atrim(cadbl(comanda.Text)))
    While Not rsttmp.EOF
     dbtmpb.Execute "delete * from bobinesreb where controlid=" + atrim(rsttmp!ID)
     rsttmp.MoveNext
    Wend
    dbtmpb.Execute "delete * from rebobinadores where comanda=" + atrim(cadbl(comanda.Text))
    dbtmp.Execute "update comandes set proximaseccio='R' where comanda=" + atrim(cadbl(comanda.Text))
 End If
 If seccio = "Sol" Then
    Set rsttmp = dbtmpb.OpenRecordset("select id from soldadores where comanda=" + atrim(cadbl(comanda.Text)))
    While Not rsttmp.EOF
     dbtmpb.Execute "delete * from bobinessol where controlid=" + atrim(rsttmp!ID)
     rsttmp.MoveNext
    Wend
    dbtmpb.Execute "delete * from soldadores where comanda=" + atrim(cadbl(comanda.Text))
    dbtmp.Execute "update comandes set proximaseccio='S' where comanda=" + atrim(cadbl(comanda.Text))
 End If
 Set dbtmpb = Nothing
 comprovar_comanda
End Sub
Function comprovar_seccio() As Boolean
  Dim continua As Boolean
  Dim seccioactual As String
  
  continua = False
  If Not canvissortirseccio Then
     seccioactual = r
    Else: seccioactual = Screen.ActiveControl.Tag
  End If
  Set rsttmp = dbtmp.OpenRecordset("select producte from comandes where comanda=" + atrim(cadbl(comanda)))
  If Not rsttmp.EOF Then
     Set rsttmp = dbtmp.OpenRecordset("select ruta from productes where codi='" + atrim((rsttmp!producte)) + "'")
     If Not rsttmp.EOF Then If InStr(1, rsttmp!ruta, seccioactual) <> 0 Then continua = True
  End If
  If continua = False Then
     MsgBox "Aquesta seccio no està dins de la ruta d'aquesta comanda..."
     If Not canvissortirseccio Then End
  End If
fi:
  comprovar_seccio = continua
  If continua Then Me.Visible = False
End Function
Sub acceptar_boto()
 On Error Resume Next
 'Unload Extrussores
 Unload Impressores
 On Error GoTo 0
 'Extrussores.Show
 Impressores.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
  Unload Impressores
On Error GoTo 0
If Not hihaalgualanterior("I") Then MsgBox "No hi ha res a la secció anterior.", vbCritical, "Atenció": Exit Sub
 If borrar = 1 Then borrar_seccio "Imp": Exit Sub
If Not comprovar_seccio Then Exit Sub
'acceptar_boto
Impressores.Show
End Sub

Private Sub Command3_Click()
On Error Resume Next
 Unload Laminadores
 On Error GoTo 0
If Not hihaalgualanterior("L") Then MsgBox "No hi ha res a la secció anterior.", vbCritical, "Atenció": Exit Sub
 If borrar = 1 Then borrar_seccio "Lam": Exit Sub
If Not comprovar_seccio Then Exit Sub

 Set rst = dbtmp.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(cadbl(linkcomanda1.Caption)))
 If Not rst.EOF Then
    If rst!proximaseccio <> "V" And rst!proximaseccio <> "T" And rst!proximaseccio <> "P" Then
       If MsgBox("Vols passar l'estat de la comanda [" + linkcomanda1.Caption + "]  a  V ?", vbInformation + vbYesNo, "Atenció") = vbYes Then
          dbtmp.Execute "update comandes set proximaseccio='V' where comanda=" + atrim(cadbl(linkcomanda1.Caption))
       End If
    End If
 End If
 Laminadores.Show

End Sub

Private Sub Command4_Click()
On Error Resume Next
 Unload Rebobinadores
 On Error GoTo 0
If Not hihaalgualanterior("R") Then MsgBox "No hi ha res a la secció anterior.", vbCritical, "Atenció": Exit Sub
 If borrar = 1 Then borrar_seccio "Reb": Exit Sub
If Not comprovar_seccio Then Exit Sub

 Rebobinadores.Show

End Sub

Private Sub Command5_Click()
 On Error Resume Next
  Unload Soldadores
 On Error GoTo 0
If Not hihaalgualanterior("S") Then MsgBox "No hi ha res a la secció anterior.", vbCritical, "Atenció": Exit Sub

 If borrar = 1 Then borrar_seccio "Sol": Exit Sub
If Not comprovar_seccio Then Exit Sub
 

 Soldadores.Show
End Sub

Private Sub Command6_Click()
 On Error Resume Next
 Unload Entrega
 On Error GoTo 0
If Not hihaalgualanterior("V") Then MsgBox "No hi ha res a la secció anterior.", vbCritical, "Atenció": Exit Sub
Me.Visible = False

 Entrega.Show
End Sub
Function hihaalgualanterior(sec As String) As Boolean
   Dim v As String
   If InStr(1, Label1.Tag, "V") = 0 Then Label1.Tag = Label1.Tag + "V"
   v = Mid(Label1.Tag, InStr(1, Label1.Tag, sec) - 1, 1)
   v = nomseccio(v)
   Set dbtmpb = OpenDatabase(cami)
   Set rsttmp = dbtmpb.OpenRecordset("select id from " + v + " where tipus='F' and comanda=" + atrim(cadbl(comanda.Text)))
   If rsttmp.EOF Then
      hihaalgualanterior = False
     Else: hihaalgualanterior = True
   End If
   Set dbtmpb = Nothing
   Set rsttmp = Nothing
End Function
Function nomseccio(v As String) As String
   Select Case v
     Case "E"
       nomseccio = "extrussores"
     Case "I"
       nomseccio = "impressores"
     Case "R"
       nomseccio = "rebobinadores"
     Case "L"
       nomseccio = "laminadores"
     Case "S"
       nomseccio = "soldadores"
   End Select
End Function
Private Sub Command7_Click()
  Set rsttmp = dbtmp.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(cadbl(comanda)))
  If Not rsttmp.EOF Then
     Select Case atrim(rsttmp!proximaseccio)
       Case "E", "": Command1_Click
       Case "I": Command2_Click
       Case "L": Command3_Click
       Case "R": Command4_Click
       Case "S": Command5_Click
       Case "V", "P": Command6_Click
       Case "T": MsgBox "Aquesta comanda ja està ENTREGADA..."
     End Select
    Else: MsgBox "Aquesta comanda no existeix": comanda.SetFocus
  End If
  
End Sub

Private Sub Command8_Click()
cami = llegir_ini("General", "camibaixesprova", fitxerini)
camicomandes = llegir_ini("General", "camiprova", fitxerini)
If Me.BackColor = QBColor(13) Then
    cami = "": Me.BackColor = QBColor(15)
    camicomandes = ""
  Else: Me.BackColor = QBColor(13)
End If
Form_Load
'MsgBox cami, , "baixes"
'MsgBox dbtmp.Name, , "comandes"
End Sub

Private Sub entregadatesavançat_Click()
   v = InputBoxEx("Escriu la contrasenya per treure aquest llistat.", "Llistat protegit amb contrasenya", , , , , , SPassword)
   If UCase(v) <> "INPLACSA123" Then MsgBox "La contrasenya no es correcte.": Exit Sub
   llistatentreguesperdates "A"
End Sub

Private Sub Form_Activate()

comanda_Change
'comanda = ""
comanda.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then End

End Sub
Function ObtenerLíneaComando(Optional MaxArgs)
    'Declara las variables.
    Dim c, LíneaComando, LonLínComando, ArgIn, i, NúmArgs
    'Ver si MaxArgs está.
    If IsMissing(MaxArgs) Then MaxArgs = 10
    'Crea una matriz del tamaño correcto.
    ReDim ArgArray(MaxArgs)
    NúmArgs = 0: ArgIn = False
    'Obtiene los argumentos de la línea de comandos.
    LíneaComando = Command()
    LonLínComando = Len(LíneaComando)
    'Recorre la línea de comando carácter a carácter
    'a la vez.

For i = 1 To LonLínComando
        c = Mid(LíneaComando, i, 1)
        'Comprueba espacio o tabulación.
        If (c <> " " And c <> vbTab) Then
            'Ningún espacio o tabulación.
            'Comprueba si está en el argumento.
            If Not ArgIn Then
            'Empieza el nuevo argumento.
            'Comprueba para más argumentos.
                If NúmArgs = MaxArgs Then Exit For
                    NúmArgs = NúmArgs + 1
                    ArgIn = True
                End If
            'Agrega el carácter al argumento actual.

ArgArray(NúmArgs) = ArgArray(NúmArgs) + c
        Else
            'Encontró un espacio o tabulador.
            'Establece ArgIn a False.
            ArgIn = False
        End If
    Next i
    'Redimensiona la matriz lo suficiente para contener los argumentos.
    'ReDim Preserve ArgArray(NúmArgs)
    'Devuelve la matriz en nombre de la función.
    ObtenerLíneaComando = ArgArray()
End Function
Private Sub Form_Load()
  Dim anarsec As String
  Dim anarcomanda As String
  Dim dbtmpb As Database
  Dim arguments As Variant
arguments = ObtenerLíneaComando
fitxerini = "comandes.ini"
If atrim(arguments(1)) <> "" Then fitxerini = atrim(arguments(1))
  canvissortirseccio = True
  If camicomandes = "" Then camicomandes = llegir_ini("General", "cami", fitxerini)
  If cami = "" Then cami = llegir_ini("General", "camibaixes", fitxerini)
  If cami = "{[}]" Then
    escriure_ini "General", "camibaixes", InputBox("Entra la ruta de baixes", "Atenció", "y:\comandes\baixes.mdb"), fitxerini
  End If
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  
  '"c:\misdoc~1\commandes\comandes.mdb"
  centerscreen Me
  Set dbtmp = OpenDatabase(camicomandes)
  Set dbtmpb = OpenDatabase(cami)
  
  anarsec = llegir_ini("General", "anarbaixasec", fitxerini)
  anarcomanda = llegir_ini("General", "anarbaixacom", fitxerini)
  escriure_ini "General", "anarbaixasec", "", fitxerini
  escriure_ini "General", "anarbaixacom", "", fitxerini
  comanda = cadbl(llegir_ini("Baixes", "ultimacomanda", fitxerini))
  If cadbl(anarcomanda) > 0 Then
     comanda = anarcomanda
     canvissortirseccio = False
     r = anarsec
     Select Case anarsec
       Case "E", "": Command1_Click
       Case "I": Command2_Click
       Case "L": Command3_Click
       Case "R": Command4_Click
       Case "S": Command5_Click
       Case "V", "P": Command6_Click
     End Select
     
  End If
  If anarsec = "" Then dbtmpb.Execute ("delete * from impressores where datainici=null")
  'Set dbtmpb = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
If canvissortirseccio Then escriure_ini "Baixes", "ultimacomanda", comanda.Text, fitxerini
End
End Sub

Private Sub linkcomanda1_Click()
  comanda = linkcomanda1.Caption
End Sub

Private Sub linkcomanda2_Click()
comanda = linkcomanda2.Caption
End Sub

Private Sub mllistatentregues_Click()
    llistatentreguesperdates "B"
    
End Sub
Sub llistatentreguesperdates(vTipus As String)
Dim vdata As String
    Dim vdatafi As String
    
    vdata = InputBox("Entra la data d'inici de la consulta" + Chr(10) + "Ex: 01/01/2016", "Atenció", Date)
    If Not IsDate(vdata) Then MsgBox "Error amb la data d'inici", vbCritical, "Error": Exit Sub
    vdatafi = InputBox("Entra la data de fi de la consulta" + Chr(10) + "Ex: 01/01/2016", "Atenció", Date)
    If Not IsDate(vdatafi) Then MsgBox "Error amb la data de fi.", vbCritical, "Error": Exit Sub
    
    ferllistatentregues CVDate(vdata), CVDate(vdatafi), vTipus
    
End Sub
Function shaborratlataulatemporal() As Boolean
   On Error GoTo fi
   shaborratlataulatemporal = True
   dbtmpb.Execute "delete * from tmp_llistatentreguesdiaries"
   Exit Function
fi:
   shaborratlataulatemporal = False
End Function
Function espesordelmaterial(numc As Double, numc2 As Double, numc3 As Double) As Double
    Dim rst As Recordset
    If numc2 = 0 Then numc2 = -9999
    If numc3 = 0 Then numc3 = -9999
    Set rst = dbstocks.OpenRecordset("SELECT DISTINCT Parcials.comanda, Palets.micres as espesor FROM Palets INNER JOIN Parcials ON Palets.Idpalet = Parcials.idpalet WHERE (((Parcials.comanda)='" + atrim(numc) + "' Or (Parcials.comanda)='" + atrim(numc2) + "' Or (Parcials.comanda)='" + atrim(numc3) + "'));")
    While Not rst.EOF
       espesordelmaterial = espesordelmaterial + cadbl(rst!espesor)
       rst.MoveNext
    Wend
    Set rst = Nothing
End Function
Sub imprimirelllistatentregues(vdata As Date, vdatafi As Date, vTipus As String)

' Dim rst As Recordset
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatentreguesdiaries.rpt", 1)
  oreport.Database.Tables.Item(1).Location = cami
  'oreport.RecordSelectionFormula = "{diferenciescomandaitreball.comanda}=" + atrim(numc)
  oreport.FormulaFields.GetItemByName("tipusllistat").Text = """" + atrim(vTipus) + """"
  If vdata = vdatafi Then
     oreport.FormulaFields.GetItemByName("nomdelllistat").Text = """" + Format(vdata, "dd-mmm-yyyy") + """"
      Else: oreport.FormulaFields.GetItemByName("nomdelllistat").Text = """" + "Entre " + Format(vdata, "dd-mmm-yyyy") + " i " + Format(vdatafi, "dd-mmm-yyyy") + """"
  End If
  oreport.DiscardSavedData
   
  'If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
   ' Else
   '   oreport.PrintOut False, 1
 ' End If
 End Sub
 Function calcularpesxrpeçall(rst As Recordset, pesgrmcm2 As Double) As Double
    calcularpesxrpeçall = pesgrmcm2 * ((cadbl(rst!amplesol) + cadbl(rst!solapasol)) * cadbl(rst!longitudsol))
    calcularpesxrpeçall = calcularpesxrpeçall * IIf(rst!migelaboratsol = "L", 1, 2)
End Function
Sub ferllistatentregues(vdata As Date, vdatafi As Date, vTipus As String)
    Dim rst As Recordset
    Dim rstb As Recordset
    Dim rstc As Recordset
    Dim rstcli As Recordset
    Dim rstextra As Recordset
    Dim rstmesures As Recordset
    Dim vsqldata As String
    Dim vpreucomanda As Boolean
    Dim vquantalbara As Double
    Dim vhihaabono As Boolean
    Dim vclixesaSAP As Boolean
    Dim vpreuimpostenvasos As Double
    
    Set dbtmp = OpenDatabase(camicomandes)
    Set dbtmpb = OpenDatabase(cami)
    Dim pespeça As Double
    Set dbstocks = OpenDatabase(rutadelfitxer(camicomandes) + "palets.mdb")
    Set dbclixes = OpenDatabase(rutadelfitxer(camicomandes) + "clixesnous.mdb")
    Set dbvendes = OpenDatabase(rutadelfitxer(cami) + "vendes.mdb")
    Set rstmesures = dbtmp.OpenRecordset("select * from mesures")
    wait 2
    vsqldata = "(((bobinesent.data)>=#" + Format(vdata, "mm/dd/yyyy") + "#) and ((bobinesent.data)<=#" + Format(vdatafi, "mm/dd/yyyy") + "#))"
    Set rst = dbtmpb.OpenRecordset("SELECT first(bobinesent.numalbara) as primeralbara,bobinesent.data, bobinesent.comanda,first(bobinesent.seccio) as fseccio, Sum(bobinesent.metresisacs) AS SumaDemetresisacs, Sum(bobinesent.kilosiunitats) AS SumaDekilosiunitats From bobinesent GROUP BY bobinesent.data, bobinesent.comanda HAVING " + vsqldata + ";")
    vsqldata = "(((dataalbara)>=#" + Format(vdata, "mm/dd/yyyy") + "#) and ((dataalbara)<=#" + Format(vdatafi, "mm/dd/yyyy") + "#))"
    If rst.EOF Then MsgBox "No hi ha cap entrega aquesta data", vbInformation, "No entregues": Exit Sub
    If Not shaborratlataulatemporal Then MsgBox "Error fent el llistat, algu altra també està fent el llistat.", vbCritical, "Error": Exit Sub
    Set rstb = dbtmpb.OpenRecordset("select * from tmp_llistatentreguesdiaries")
    While Not rst.EOF
      Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rst!comanda))
      If Not rstc.EOF Then
        vhihaabono = False
        vclixesaSAP = False
        vpreuimpostenvasos = 0
        If rstc!producte = "PCI3" Then GoTo proxim
        Set rstcli = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(rstc!client))
        If rstcli.EOF Then GoTo proxim
        rstb.AddNew
        rstb!comanda = rst!comanda
        rstb!nomclient = rstcli!nom
        rstb!producte = rstc!producte
        rstb!ample = IIf(rst!fseccio = "S", rstc!amplesol, IIf(rst!fseccio = "R", rstc!amplereb, rstc!ampleesq))
        rstb!espesor = espesordelmaterial(rstc!comanda, cadbl(rstc!linkcomanda1), cadbl(rstc!linkcomanda2))
        rstb!comandapartida = IIf(cadbl(rstc!refilate) > 0, "*", "")
        If rst!fseccio <> "S" Then
             rstb!metres = rst!SumaDemetresisacs
             rstb!kilos = rst!SumaDekilosiunitats
               Else:
                    Set rstextra = dbtmp.OpenRecordset("select solpesgrmcm2 from comandes_extres where comanda=" + atrim(rstc!comanda))
                    If Not rstextra.EOF Then
                        pespeça = calcularpesxrpeçall(rstc, rstextra!solpesgrmcm2)
                        rstb!unitats = rst!SumaDemetresisacs
                         rstb!kilos = rst!SumaDemetresisacs * pespeça
                    End If
        End If
        rstb!clixesfacturats = buscarsishanfacturatclixes(cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), rstc!comanda, vclixesaSAP)
        rstb!clixesaSAP = vclixesaSAP
        rstb!dataalbara = rst!data
        rstb!parcialototal = atrim(rstc!proximaseccio) + buscarsientregatenvaries(rstc!comanda)
        vquantalbara = 0: vpreucomanda = False
        rstb!eurocomanda = IIf(cadbl(rstc!pvpdolar) > 0, cadbl(rstc!pvpdolar), cadbl(rstc!pvp))
        rstb!totalalbara = buscarpreualbara(cadbl(rst!primeralbara), rstc!comanda, vsqldata, rstb!eurocomanda, vpreucomanda, vquantalbara, vpreuimpostenvasos)
        rstb!preuImpostEnvasos = vpreuimpostenvasos
        rstb!quantitatalbara = vquantalbara
        
        rstmesures.FindFirst "codi=" + atrim(cadbl(rstc!mesurapvp))
        If Not rstmesures.NoMatch Then rstb!unitatpvp = IIf(cadbl(rstc!pvpdolar) > 0, substituir(atrim(rstmesures!unitatinterna), "€", "$"), rstmesures!unitatinterna)
        If rstb!totalalbara = 0 Or vpreucomanda Then
          If rstb!totalalbara = 0 Then rstb!totalalbara = rstb!kilos * cadbl(rstb!eurocomanda)
           rstb!preudecomanda = True
        End If
        rstb!preuSAP = buscarpreuSAP(rstc!comanda, cadbl(rst!primeralbara), vhihaabono)
        rstb!hihaabono = vhihaabono
        rstb!numalbara = atrim(cadbl(rst!primeralbara))
        rstb.Update
      End If
proxim:
      rst.MoveNext
    Wend
    wait 2
    dbtmpb.Execute "delete * from tmp_llistatentreguesdiaries where comanda=0"
    wait 3 'esperem que els datos arrivin a la base de dades
    imprimirelllistatentregues vdata, vdatafi, vTipus
    Set rst = Nothing
    Set rstb = Nothing
    Set rstc = Nothing
    Set rstcli = Nothing
    Set rstmesures = Nothing
    Set rstextra = Nothing
    Set dbclixes = Nothing
    Set dbvendes = Nothing
End Sub
Function buscarsientregatenvaries(vnumc As Double) As String
    Dim rst As Recordset
    Set rst = dbtmpb.OpenRecordset("SELECT distinct bobinesent.data, bobinesent.comanda FROM bobinesent where comanda=" + atrim(vnumc))
    If Not rst.EOF Then
       If rst.RecordCount > 1 Then buscarsientregatenvaries = "*"
    End If
    Set rst = Nothing
End Function
Function buscarpreuSAP(vnumc As Double, vnumalb As Double, vhihaabono As Boolean) As Double
    Dim rst As Recordset
    Dim vdata As Date
    vnumalb = cadbl(vnumalb)
  '  If vnumc = 216023 Then Stop
    Set rst = dbtmpb.OpenRecordset("select * from  Importada_LiniesFacturesSAP_Inplacsa where U_GSP_INFABLOTE='" + atrim(vnumc) + "' and instr(1,' '+Albara_Produccio,'" + atrim(vnumalb) + "')>0")
    'busco la data de l'albara per buscar tots els albarans fets aquest dia
    If Not rst.EOF Then vdata = rst!dataalbara
    Set rst = dbtmpb.OpenRecordset("select * from  Importada_LiniesFacturesSAP_Inplacsa where U_GSP_INFABLOTE='" + atrim(vnumc) + "' and dataalbara=#" + Format(vdata, "mm/dd/yy") + "#")
    While Not rst.EOF
       If rst!tipusdelinia = "F" Then
          buscarpreuSAP = buscarpreuSAP + (passaradecimal(rst!quantity) * passaradecimal(rst!price))
           Else: vhihaabono = True
       End If
       rst.MoveNext
    Wend
       'si no ho trobo a Inplacsa ho miro a Plasel
    If buscarpreuSAP = 0 Then
      
    Set rst = dbtmpb.OpenRecordset("select * from  Importada_LiniesFacturesSAP_Plasel where U_GSP_INFABLOTE='" + atrim(vnumc) + "' and instr(1,' '+Albara_Produccio,'" + atrim(vnumalb) + "')>0")
    'busco la data de l'albara per buscar tots els albarans fets aquest dia
    If Not rst.EOF Then vdata = rst!dataalbara
    Set rst = dbtmpb.OpenRecordset("select * from  Importada_LiniesFacturesSAP_Plasel where U_GSP_INFABLOTE='" + atrim(vnumc) + "' and dataalbara=#" + Format(vdata, "mm/dd/yy") + "#")
    While Not rst.EOF
       If rst!tipusdelinia = "F" Then
         buscarpreuSAP = buscarpreuSAP + (passaradecimal(rst!quantity) * passaradecimal(rst!price))
          Else: vhihaabono = True
       End If
       rst.MoveNext
    Wend
    End If
End Function
Function buscarpreualbara(vnumalb As Double, vnumc As Double, vsqldata As String, vpvpcomanda As Double, vpreucomanda As Boolean, vquantitatalbara As Double, ByRef vpreuimpostenvasos As Double) As Double
  Dim rst As Recordset
  Dim vdata As Date
  Dim vpreuimpost As Double
  'Set rst = dbvendes.OpenRecordset("select * from liniesalbara where numalbara=" + atrim(vnumalb) + " and lotinplacsa=" + atrim(vnumc))
  Set rst = dbvendes.OpenRecordset("SELECT capcaleraalbara.numalbara, capcaleraalbara.dataalbara, capcaleraalbara.dataenvioasap, liniesalbara.*, Clients_envios.pais FROM (liniesalbara INNER JOIN capcaleraalbara ON liniesalbara.numalbara = capcaleraalbara.numalbara) LEFT JOIN Clients_envios ON capcaleraalbara.id_direnvio = Clients_envios.id where lotinplacsa=" + atrim(vnumc) + " and " + vsqldata + " order by dataalbara")
 ' Clipboard.Clear
 ' Clipboard.SetText "SELECT capcaleraalbara.numalbara,capcaleraalbara.dataalbara, capcaleraalbara.dataenvioasap, liniesalbara.* FROM liniesalbara INNER JOIN capcaleraalbara ON liniesalbara.numalbara = capcaleraalbara.numalbara where lotinplacsa=" + atrim(vnumc) + " and " + vsqldata + " order by dataalbara"
  If Not rst.EOF Then
     rst.FindFirst "capcaleraalbara.numalbara=" + atrim(cadbl(vnumalb))
     If Not rst.NoMatch Then vdata = rst!dataalbara
     rst.MoveFirst
  End If
  'Clipboard.Clear
  'Clipboard.SetText "SELECT capcaleraalbara.dataalbara, capcaleraalbara.dataenvioasap, liniesalbara.* FROM liniesalbara INNER JOIN capcaleraalbara ON liniesalbara.numalbara = capcaleraalbara.numalbara where lotinplacsa=" + atrim(vnumc) + " and " + vsqldata
  'If Not rst.EOF Then buscarpreualbara = Redondejar(cadbl(rst!preuvenda) * cadbl(rst!quantitat), 2)
  While Not rst.EOF
    If rst!dataalbara = vdata Then
        vpreu = cadbl(rst!preuvenda)
        If vpreu = 0 Then vpreu = vpvpcomanda: vpreucomanda = True
        If rst!pais = "ES" Then vpreuimpost = cadbl(rst!kgimpostenvasos) * 0.45
        buscarpreualbara = buscarpreualbara + Redondejar((cadbl(vpreu) * cadbl(rst!quantitat) + vpreuimpost), 2)
        vquantitatalbara = vquantitatalbara + cadbl(rst!quantitat)
        vpreuimpostenvasos = vpreuimpostenvasos + vpreuimpost
    End If
    rst.MoveNext
  Wend
  Set rst = Nothing
End Function
Function buscarsishanfacturatclixes(vnumtreball As Double, vordremodificacio, vnumc As Double, vclixesaSAP As Boolean) As Double
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from pressupostos where id_treball=" + atrim(vnumtreball) + " and ordremodificacio=" + atrim(vordremodificacio))
   If Not rst.EOF Then
      If cadbl(rst!lotambelqueshafacturat) = vnumc Then buscarsishanfacturatclixes = cadbl(rst!preu)
   End If
   'sihi ha preu miro si estan a sap a inplacsa o plasel
   If buscarsishanfacturatclixes > 0 Then
        'Plasel
        Set rst = dbtmpb.OpenRecordset("select * from Importada_LiniesFacturesSAP_Plasel where u_gsp_infablote='" + atrim(vnumc) + "' and itemcode='PLATES'")
        If Not rst.EOF Then
            If passaradecimal(rst!price) = buscarsishanfacturatclixes Then vclixesaSAP = True
               Else
                'Inplacsa
                Set rst = dbtmpb.OpenRecordset("select * from Importada_LiniesFacturesSAP_Inplacsa where u_gsp_infablote='" + atrim(vnumc) + "' and itemcode='PLATES'")
                If Not rst.EOF Then
                   If passaradecimal(rst!price) = buscarsishanfacturatclixes Then vclixesaSAP = True
                End If
            End If
   End If
   Set rst = Nothing
End Function

Private Sub mpassarbobinesaacabades_Click()
   Bobinesassignades.Show 1
End Sub

Private Sub pasaranoacavada_Click()
    Dim numc As Double
    Dim seccio As String
    Set dbtmpb = OpenDatabase(cami)
    numc = cadbl(InputBox("Entra el numero de comanda a modificar.", "Atenció"))
    If numc = 0 Then Exit Sub
    seccio = UCase(InputBox("Entra la inicial de la seccio que vols modificar pasar a no acavada (I,L,R,S)", "Atenció"))
    If seccio = "" Then Exit Sub
    If InStr(1, "ILR", seccio) = 0 Then MsgBox "Aquesta seccio no existeix", vbCritical, "Error": Exit Sub
    If seccio = "I" Then seccio = "impressorestot"
    If seccio = "R" Then seccio = "rebobinadorestot"
    If seccio = "L" Then seccio = "laminadorestot"
    dbtmpb.Execute "update " + seccio + " set acavada=0 where comanda=" + atrim(numc)
    
    
End Sub

Private Sub Timer1_Timer()
  'controldeteclat
  canviarelscolorsdelscontrolsalentrar
End Sub
