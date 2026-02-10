VERSION 5.00
Begin VB.Form paletsajust 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Palets Ajust"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3930
   Begin VB.Frame framedemanabobina 
      BackColor       =   &H0000FF00&
      Caption         =   "Bobina d'Ajust"
      Height          =   2115
      Left            =   300
      TabIndex        =   18
      Top             =   1935
      Visible         =   0   'False
      Width           =   3885
      Begin VB.CommandButton bok 
         Height          =   390
         Left            =   2835
         Picture         =   "paletsajust.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1665
         Width           =   795
      End
      Begin VB.TextBox cnumbobina 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1350
         TabIndex        =   21
         Top             =   1665
         Width           =   1470
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00F1B75F&
         Caption         =   "Desbobinador 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   2250
         Picture         =   "paletsajust.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   195
         Width           =   1425
      End
      Begin VB.CommandButton bdesb1 
         BackColor       =   &H00EEE4D7&
         Caption         =   "Desbobinador 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   195
         Picture         =   "paletsajust.frx":0E47
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   195
         Width           =   1425
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Bobina:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   315
         TabIndex        =   22
         Top             =   1740
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command3 
      Height          =   450
      Left            =   555
      Picture         =   "paletsajust.frx":16FD
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Borrar totes les bobines d'ajust."
      Top             =   75
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.CommandButton Command2 
      Height          =   390
      Left            =   2865
      Picture         =   "paletsajust.frx":1C87
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Marcar bobina acabada."
      Top             =   1005
      Width           =   390
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   2850
      Picture         =   "paletsajust.frx":20C5
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Marcar bobina acabada."
      Top             =   1500
      Width           =   390
   End
   Begin VB.CommandButton impbobajust2 
      Height          =   375
      Left            =   3255
      Picture         =   "paletsajust.frx":2503
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Imprimir l'etiqueta de bobina d'entrada parcial."
      Top             =   1500
      Width           =   390
   End
   Begin VB.CommandButton impbobajust1 
      Height          =   390
      Left            =   3255
      Picture         =   "paletsajust.frx":2A8D
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprimir l'etiqueta de bobina d'entrada parcial."
      Top             =   1005
      Width           =   390
   End
   Begin VB.CommandButton alta 
      Height          =   450
      Left            =   120
      Picture         =   "paletsajust.frx":3017
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Alta  Registres"
      Top             =   75
      Width           =   405
   End
   Begin VB.CommandButton sortir 
      Height          =   510
      Left            =   3285
      Picture         =   "paletsajust.frx":35A1
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Sortida."
      Top             =   75
      Width           =   420
   End
   Begin VB.TextBox palet1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1005
      Width           =   1050
   End
   Begin VB.TextBox bob1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1005
      Width           =   495
   End
   Begin VB.TextBox mtrs1 
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
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1005
      Width           =   765
   End
   Begin VB.TextBox palet2 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1485
      Width           =   1050
   End
   Begin VB.TextBox bob2 
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
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1485
      Width           =   495
   End
   Begin VB.TextBox mtrs2 
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
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1485
      Width           =   780
   End
   Begin VB.Label xrmodificar 
      BackStyle       =   0  'Transparent
      Caption         =   "Per Modificar els metres fes dos clics a sobre del camp."
      Height          =   450
      Left            =   975
      TabIndex        =   16
      Top             =   75
      Width           =   2145
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Per bobines de material per llençar posseu: 11.111 / 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   285
      TabIndex        =   13
      Top             =   1965
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      Height          =   1005
      Left            =   270
      Top             =   915
      Width           =   3420
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   990
      Width           =   420
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "2."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   1500
      Width           =   420
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Palet   Bob  Mtrs"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   510
      TabIndex        =   6
      Top             =   555
      Width           =   2445
   End
End
Attribute VB_Name = "paletsajust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command23_Click()

End Sub


Function noexisteixbobina(p As Double, b As Double) As Boolean
  Dim rstp As Recordset
  Set rstp = dbstocks.OpenRecordset("select idpalet from bobines where idpalet=" + atrim(p) + " and idbobina=" + atrim(b))
  If Not rstp.EOF Then
      noexisteixbobina = False
     Else: noexisteixbobina = True
  End If
  If p = 11111 And b = 1 Then noexisteixbobina = False
End Function
Sub demanarpaletibobina(vpaletib As String)
   framedemanabobina.visible = True
   framedemanabobina.Top = 1
   framedemanabobina.Left = 1
   cnumbobina = ""
   cnumbobina.SetFocus
   While framedemanabobina.visible
      DoEvents
   Wend
   vpaletib = cnumbobina
End Sub
Private Sub alta_Click()
   Dim rst As Recordset
   Dim palet As Double
   Dim bobina As Double
   Dim metres As Double
   Dim vpalet As Double
   Dim vbob As Double
   Dim paletp As String
   Dim vobservacio As String
   
   If palet2 <> "" Then MsgBox "No es poden entrar mes palets d'ajust": Exit Sub
   'paletp = InputBox("Entra el numero de palet d'ajust:" + Chr(10) + "POTS ESCANEJAR-LO SI VOLS", "Ajust")
   demanarpaletibobina paletp
   vpalet = 0: vbob = 0
   convertirScanambPaletiBobina paletp, vpalet, vbob
   If vpalet > 0 And vbob > 0 Then
     palet = cadbl(vpalet)
     bobina = cadbl(vbob)
       Else: palet = cadbl(paletp): bobina = cadbl(InputBox("Entra el numero de bobina d'ajust:", "Ajust"))
   End If
   paletp = ""
   
   If palet = 111111 Then palet = 11111
   If palet > 1000000 Then palet = 11111
   If palet = 11111 Then bobina = 1
   
   If noexisteixbobina(palet, bobina) Then MsgBox "Aquest Palet/Bobina no existeix", vbCritical, "Atenció": Exit Sub
   If cadbl(palet) = cadbl(palet1) And cadbl(bobina) = cadbl(bob1) Then MsgBox "Aquesta bobina ja està entrada", 64, "Atenció": Exit Sub
   If palet1 = "" Then
     If palet <> 11111 Then
        If Not comprovar_ajust_porllençar(palet, bobina) Then Exit Sub
     End If
   End If
   metres = cadbl(InputBox("Entra els metres gastats d'aquesta bobina per l'ajust:", "Ajust"))
   
   If palet > 20000 And bobina > 0 And metres > 0 Then
      vobservacio = ""
      Set rst = dbstocks.OpenRecordset("select comanda from parcials where (comanda='2500' or observacions='#llençar') and idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina)))
      If Not rst.EOF Then vobservacio = "#llençar"
      dbstocks.Execute "insert into parcials (idpalet,idbobina,metres,comanda,data,seccio,utilitzada,orcomassignacio,operari,observacions) values (" + atrim(palet) + "," + atrim(bobina) + "," + atrim(metres) + ",'" + atrim(cadbl(form1.comanda)) + "',now,'" + lletraseccio + "',true,500," + atrim(numop) + ",'" + vobservacio + "')"
      If UCase(InputBox("Bobina fisicament acabada?" + Chr(10) + Chr(13) + "Escriu SI per donar per acabada.", "Bobina acabada?")) = "SI" Then
          dbstocks.Execute "delete * from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and not utilitzada and comanda='" + atrim(form1.comanda) + "'"
          dbstocks.Execute "delete * from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and not utilitzada and (cdbl(comanda)>2000 and cdbl(comanda)<3000)"
          explicacio = "DONADA DE BAIXA FENT L'AJUST "
          mantenimentbobina.passarbobinaaacavada palet, bobina
          'MsgBox "Bobina acabada.", vbInformation, "Informació"
      End If
   End If
   If Not (palet > 0 And bobina > 0 And metres > 0) Then Exit Sub
   If form1.impresores.Recordset.EditMode = 0 Then form1.impresores.Recordset.Edit
   If palet1 <> "" Then
     palet2 = palet
     bob2 = bobina
     mtrs2 = metres
     
     form1.impresores.Recordset!paletprova2 = palet
     form1.impresores.Recordset!bobinaprova2 = bobina
     form1.impresores.Recordset!metresprova2 = metres
   End If
   If palet1 = "" Then
     palet1 = palet
     bob1 = bobina
     mtrs1 = metres
     form1.impresores.Recordset!paletprova = palet
     form1.impresores.Recordset!bobinaprova = bobina
     form1.impresores.Recordset!metresprova = metres
   End If
   form1.impresores.Recordset!mtrsprova = cadbl(form1.impresores.Recordset!metresprova) + cadbl(form1.impresores.Recordset!metresprova2)
   form1.impresores.Recordset!paletbobprova = atrim(form1.impresores.Recordset!paletprova) + "-" + atrim(form1.impresores.Recordset!bobinaprova) + "/"
   paletp = atrim(form1.impresores.Recordset!paletbobprova) + atrim(form1.impresores.Recordset!paletprova2) + "-" + atrim(form1.impresores.Recordset!bobinaprova2)
   paletp = Mid(paletp, 1, 20)
   form1.impresores.Recordset!paletbobprova = paletp
   form1.impresores.Recordset.Update
   mantenimentbobina.actualitzarmetresgrupsestoc cadbl(palet), cadbl(bobina)
   bobinesdentrada.actualitzar_metres_disponibles palet, bobina
   'bobinesdentrada.imprimir_bobinaparcial palet, bobina, , 1
End Sub

Sub matarbobina(palet As String, bobina As String)
   ratoli "normal"
   explicacio = ""
   If UCase(InputBox("Bobina fisicament acabada?" + Chr(10) + Chr(13) + "Escriu SI per donar per acabada.", "Bobina acabada?")) = "SI" Then
       explicacio = InputBox("Vols entrar una explicació del perquè dones per acabada aquesta bobina?" + Chr(10) + Chr(13) + "SI HAS DONAT PER ACABADA ACCIDENTALMENT ESCRIU [ERROR] A LA CASELLA.", "Bobina acabada")
       ratoli "espera"
       If UCase(explicacio) <> "ERROR" Then mantenimentbobina.passarbobinaaacavada cadbl(palet), cadbl(bobina)
       ratoli "normal"
   End If
End Sub

Private Sub bdesb1_Click()
  cnumbobina = carregarbobinadeldesbobinador(1)
End Sub
Function carregarbobinadeldesbobinador(vnumdesb As Byte)
'   Dim rst As Recordset
'   Set rst = dbtmpb.OpenRecordset("select *  from bobinesdesbobinadors where numdesbobinador=" + atrim(cadbl(vnumdesb)) + " and maquina=" + atrim(nummaq) + " order by data desc")
'   If Not rst.EOF Then
'       carregarbobinadeldesbobinador = atrim(rst!palet) + "/" + atrim(rst!bobina)
'   End If
'   Set rst = Nothing
   carregarbobinadeldesbobinador = llegir_ini("Bobines_Desbobinadors_" + atrim(nummaq), "Bobina" + atrim(vnumdesb), rutadelfitxer(cami) + "valorsprograma.ini")
   If carregarbobinadeldesbobinador = "{[}]" Then carregarbobinadeldesbobinador = ""
End Function

Private Sub bok_Click()
  framedemanabobina.visible = False
End Sub

Private Sub Command1_Click()
  If cadbl(palet2) > 11111 Then matarbobina palet2, bob2
End Sub

Private Sub Command2_Click()
   If cadbl(palet1) > 11111 Then matarbobina palet1, bob1
End Sub

Private Sub Command3_Click()
   Dim v1 As Boolean
   Dim v2 As Boolean
   If form1.impresores.Recordset.EditMode = 0 Then form1.impresores.Recordset.Edit
   v1 = eliminarbobinaajust(cadbl(palet1), cadbl(bob1), cadbl(form1.comanda))
   bobinesdentrada.actualitzar_metres_disponibles cadbl(palet1), cadbl(bob1)
   v2 = eliminarbobinaajust(cadbl(palet2), cadbl(bob2), cadbl(form1.comanda))
   bobinesdentrada.actualitzar_metres_disponibles cadbl(palet2), cadbl(bob2)
   If palet2 <> "" And v2 Then
     palet2 = ""
     bob2 = ""
     mtrs2 = 0
     form1.impresores.Recordset!paletprova2 = 0
     form1.impresores.Recordset!bobinaprova2 = 0
     form1.impresores.Recordset!metresprova2 = 0
   End If
   If palet1 <> "" And v1 Then
     palet1 = ""
     bob1 = ""
     mtrs1 = 0
     form1.impresores.Recordset!paletprova = 0
     form1.impresores.Recordset!bobinaprova = 0
     form1.impresores.Recordset!metresprova = 0
   End If
   'Form1.impresores.Recordset!mtrsprova = 0
   'Form1.impresores.Recordset!paletbobprova = ""
   form1.impresores.Recordset!mtrsprova = cadbl(form1.impresores.Recordset!metresprova) + cadbl(form1.impresores.Recordset!metresprova2)
   form1.impresores.Recordset!paletbobprova = atrim(form1.impresores.Recordset!paletprova) + "-" + atrim(form1.impresores.Recordset!bobinaprova) + "/"
   paletp = atrim(form1.impresores.Recordset!paletbobprova) + atrim(form1.impresores.Recordset!paletprova2) + "-" + atrim(form1.impresores.Recordset!bobinaprova2)
   paletp = Mid(paletp, 1, 20)
   form1.impresores.Recordset!paletbobprova = paletp
   form1.impresores.Recordset.Update
   
End Sub
Function eliminarbobinaajust(palet As Double, bobina As Double, comanda As Double) As Boolean
   If palet = 11111 Or palet = 0 Then Exit Function
   eliminarbobinaajust = False
   If MsgBox("Segur que vols borrar la informació d'ajust del palet " + atrim(palet) + "/" + atrim(bobina) + "?", vbCritical + vbYesNo, "Atenció") = vbYes Then
     dbstocks.Execute "delete * from parcials where comanda='" + atrim(comanda) + "' and orcomassignacio='500' and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina)
     eliminarbobinaajust = True
   End If
End Function
Function famesde10minuts() As Boolean
   Dim rst As Recordset
   Set rst = form1.impresores.Recordset.Clone
   If Not rst.EOF Then 'rst.MoveLast
      If DateDiff("n", (atrim(rst!datainici) + " " + atrim(rst!horafi)), Now) > 10 Then famesde10minuts = True
   End If
   
End Function

Private Sub Command4_Click()
   cnumbobina = carregarbobinadeldesbobinador(2)
End Sub

Private Sub Form_Activate()
     palet2 = catrim(form1.impresores.Recordset!paletprova2)
     bob2 = catrim(form1.impresores.Recordset!bobinaprova2)
     mtrs2 = catrim(form1.impresores.Recordset!metresprova2)
     
     palet1 = catrim(form1.impresores.Recordset!paletprova)
     bob1 = catrim(form1.impresores.Recordset!bobinaprova)
     mtrs1 = catrim(form1.impresores.Recordset!metresprova)
     If cadbl(form1.impresores.Recordset!metresprova) = 0 Then
        mtrs1 = catrim(form1.impresores.Recordset!mtrsprova)
     End If
    
     
     If famesde10minuts Then
       xrmodificar.visible = False
       Command3.visible = False
        Else
             xrmodificar.visible = True
             Command3.visible = True
     End If
   'bloquejar els canvis a màquina// ho he tret perquè l'alicia ha demanat
     'que els operaris volen fer canvis.
    '  If llegir_ini("Baixes", "programaamaquina", fitxerini) <> "1" Then
        xrmodificar.visible = True
        Command3.visible = True
    ' End If
    
     
End Sub
Function catrim(ByVal num As Variant) As String
   
   If cadbl(num) = 0 Then
      catrim = ""
     Else: catrim = atrim(num)
   End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 110 Then KeyCode = 188
   
End Sub

Private Sub Form_Load()
 obrestocks
 If cadbl(llegir_ini("finestreajust", "top", fitxerini)) > 0 Then
   paletsajust.Top = cadbl(llegir_ini("finestreajust", "top", fitxerini))
   paletsajust.Left = cadbl(llegir_ini("finestreajust", "left", fitxerini))
     Else:
       paletsajust.Top = (Screen.Height / 2) - (paletsajust.Height / 2)
       paletsajust.Left = (Screen.width / 2) - (paletsajust.width / 2)
 End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 2 Then
     xrmodificar.visible = True
     Command3.visible = True
  End If
End Sub

Private Sub impbobajust1_Click()
  If cadbl(palet1) > 0 And cadbl(bob1) > 0 And cadbl(palet1) <> 11111 Then
    bobinesdentrada.imprimir_bobinaparcial cadbl(palet1), cadbl(bob1), , 1
  End If
End Sub

Private Sub impbobajust2_Click()
If cadbl(palet2) > 0 And cadbl(bob2) > 0 Then
    bobinesdentrada.imprimir_bobinaparcial cadbl(palet2), cadbl(bob2), , 1
  End If
End Sub

Private Sub mtrs1_DblClick()
  Dim metresanteriors As Double
  Dim metresnous As Double
  Dim resp As String
  Dim paletp As String
  If Not xrmodificar.visible Then Exit Sub
  metresanteriors = cadbl(mtrs1)
  If metresanteriors > 0 Then
   resp = InputBox("Entra els metres gastats d'ajust pel palet " + atrim(palet1) + "/" + atrim(bob1), "Modificació metres")
   If IsNumeric(resp) Then
     metresnous = cadbl(resp)
     dbstocks.Execute "update parcials set metres=" + atrim(metresnous) + " where orcomassignacio='500' and idpalet=" + atrim(cadbl(palet1)) + " and idbobina=" + atrim(cadbl(bob1)) + " and metres=" + atrim(metresanteriors) + " and comanda='" + atrim(form1.comanda) + "'"
     If Not (cadbl(palet1) > 0 And cadbl(bob1) > 0 And metresnous > 0) Then Exit Sub
     If form1.impresores.Recordset.EditMode = 0 Then form1.impresores.Recordset.Edit
     form1.impresores.Recordset!metresprova = metresnous
     form1.impresores.Recordset!mtrsprova = cadbl(form1.impresores.Recordset!metresprova) + cadbl(form1.impresores.Recordset!metresprova2)
     form1.impresores.Recordset.Update
     mtrs1 = metresnous
     bobinesdentrada.actualitzar_metres_disponibles cadbl(palet1), cadbl(bob1)
   End If
  End If
End Sub

Private Sub mtrs2_DblClick()
Dim metresanteriors As Double
  Dim metresnous As Double
  Dim resp As String
  Dim paletp As String
  If Not xrmodificar.visible Then Exit Sub
  metresanteriors = cadbl(mtrs2)
  If metresanteriors > 0 Then
   resp = InputBox("Entra els metres gastats d'ajust pel palet " + atrim(palet2) + "/" + atrim(bob2), "Modificació metres")
   If IsNumeric(resp) Then
     metresnous = cadbl(resp)
     dbstocks.Execute "update parcials set metres=" + atrim(metresnous) + " where orcomassignacio='500' and idpalet=" + atrim(cadbl(palet2)) + " and idbobina=" + atrim(cadbl(bob2)) + " and metres=" + atrim(metresanteriors) + " and comanda='" + atrim(form1.comanda) + "'"
     If Not (cadbl(palet2) > 0 And cadbl(bob2) > 0 And metresnous > 0) Then Exit Sub
     If form1.impresores.Recordset.EditMode = 0 Then form1.impresores.Recordset.Edit
     form1.impresores.Recordset!metresprova2 = metresnous
     form1.impresores.Recordset!mtrsprova = cadbl(form1.impresores.Recordset!metresprova) + cadbl(form1.impresores.Recordset!metresprova2)
     form1.impresores.Recordset.Update
     mtrs2 = metresnous
     bobinesdentrada.actualitzar_metres_disponibles cadbl(palet2), cadbl(bob2)
   End If
  End If
End Sub

'Sub convertirScanambPaletiBobina(vcodi As String, vpalet As Long, vbob As Long)
'   Dim vcont As Double
'   vcodi = atrim(vcodi)
'   While vcont < Len(vcodi)
'      If Not IsNumeric(Mid(vcodi, vcont + 1, 1)) Then
'        vpalet = cadbl(Mid(vcodi, 1, vcont))
'        If Len(vcodi) >= vcont + 2 Then vbob = cadbl(Mid(vcodi, vcont + 2))
'        GoTo sortir
'      End If
'      vcont = vcont + 1
'   Wend
'sortir:
'End Sub

Private Sub palet1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub palet1_LostFocus()
  Dim vpalet As Double
  Dim vbob As Double
  convertirScanambPaletiBobina palet1, vpalet, vbob
  If vpalet > 0 And vbob > 0 Then
     palet1 = atrim(vpalet)
     bob1 = atrim(vbob)
     Command1_Click
  End If
End Sub

Private Sub palet2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub palet2_LostFocus()
  Dim vpalet As Double
  Dim vbob As Double
  convertirScanambPaletiBobina palet2, vpalet, vbob
  If vpalet > 0 And vbob > 0 Then
     palet2 = atrim(vpalet)
     bob2 = atrim(vbob)
  End If
End Sub

Private Sub sortir_Click()
   
   escriure_ini "finestreajust", "top", paletsajust.Top, fitxerini
   escriure_ini "finestreajust", "left", paletsajust.Left, fitxerini
   comprovar_materials_diferents
   Unload paletsajust
End Sub
Sub comprovar_materials_diferents()
   Dim rst As Recordset
   Dim vsql As String
   Dim vultim As String
   vsql = "SELECT Parcials.idpalet, materials.familia, familiesmaterials.descripcio  FROM familiesmaterials RIGHT JOIN ((Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi) ON familiesmaterials.codi = materials.familia where parcials.comanda='" + atrim(cadbl(form1.comanda)) + "'"
   Set rst = dbstocks.OpenRecordset(vsql)
   If Not rst.EOF Then vultim = rst!descripcio
   While Not rst.EOF
      If vultim <> rst!descripcio Then MsgBox "Hi ha dos tipus de materials a la bobina de sortida d'ajust." + Chr(10) + "Comprova si s'ha de separar els materials", vbCritical, "Atenció": GoTo fi
      rst.MoveNext
   Wend
fi:
   Set rst = Nothing
   
End Sub
Function comprovar_ajust_porllençar(vp As Double, vb As Double) As Boolean
   Dim rsta As Recordset
   Dim rst As Recordset
   Dim vdemanarexplicacions As Boolean
   Dim vexplicacions As String
   comprovar_ajust_porllençar = True
   Set rsta = dbstocks.OpenRecordset("select mtrsajust,sistemadajust from opcionsdajust where comanda=" + atrim(cadbl(form1.comanda)))
   If rsta.EOF Then GoTo fi
   If rsta!sistemadajust = 1 Then
     Set rst = dbstocks.OpenRecordset("select comanda from parcials where comanda='2500' or observacions='#llençar' and idpalet=" + atrim(cadbl(vp)) + " and idbobina=" + atrim(cadbl(vb)))
     If rst.EOF Then vdemanarexplicacions = True: GoTo fi
   End If
   
fi:
   Set rst = Nothing
   Set rsta = Nothing
   If vdemanarexplicacions Then
       If MsgBox("Els palets que has utilitzats no son per llençar i l'ajust s'havia de fer amb matrial per llençar." + Chr(10) + "ÈS CORRECTE?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then comprovar_ajust_porllençar = False: Exit Function
       vexplicacions = InputBox("El material que has utilitzat per ajust no es dels escullits per llençar" + Chr(10) + "ESCRIU PERQUÈ NO HAS AGAFAT BOBINES PER LLENÇAR.", "Atenció")
       mantenimentbobina.passaravis cadbl(vp), cadbl(vb), "S'ha utilitzat el palet " + atrim(vp) + "/" + atrim(vb) + " a ajust que no era per llençar.", form1.comanda, vexplicacions
   End If
End Function
