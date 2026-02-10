VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formqualitat 
   Caption         =   "Control Albarans i Certificats de Qualitat"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18225
   Icon            =   "FormControlQualitat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   18225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btot 
      BackColor       =   &H00F1B75F&
      Caption         =   "Tot"
      Height          =   405
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   300
      Width           =   885
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pendent de documentació"
      Height          =   405
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   300
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Només a punt de verificar"
      Height          =   405
      Left            =   1095
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   300
      Width           =   2265
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   7380
      Left            =   165
      TabIndex        =   0
      Top             =   810
      Width           =   17820
      _ExtentX        =   31433
      _ExtentY        =   13018
      _Version        =   393216
      ForeColorSel    =   -2147483643
      AllowBigSelection=   0   'False
      FocusRect       =   2
      MergeCells      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Caducada"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8535
      TabIndex        =   9
      Top             =   30
      Width           =   1200
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sense CQ"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9780
      TabIndex        =   8
      Top             =   30
      Width           =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CQ Revisat o ALB xr Revisar"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   12825
      TabIndex        =   7
      Top             =   45
      Width           =   2400
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CQ NO Revisat"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   11055
      TabIndex        =   6
      Top             =   30
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concertat o sense Control de CQs"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   15300
      TabIndex        =   5
      Top             =   45
      Width           =   2565
   End
   Begin VB.Label etactualitzant 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualitzant les dades..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4380
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   3315
   End
End
Attribute VB_Name = "formqualitat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbqualitat As Database
Dim rstqualitat As Recordset
Dim rstescanejades As Recordset
Dim vordrereixa As String
Dim vteclaapretada As Double
Dim vfilareixa As Double

Private Sub btot_Click()
    Dim vrow As Double
   vrow = 1
   reixa.Col = 2
   reixa.Redraw = False
   While vrow < reixa.Rows
      reixa.Row = vrow
      reixa.RowHeight(vrow) = reixa.RowHeight(0)
      vrow = vrow + 1
   Wend
   reixa.Redraw = True
End Sub

Private Sub Command1_Click()
   Dim vrow As Double
   btot_Click
   vrow = 1
   reixa.Col = 2
   reixa.Redraw = False
   While vrow < reixa.Rows
      reixa.Row = vrow
      If reixa.CellBackColor <> QBColor(10) Then reixa.RowHeight(vrow) = 0
      vrow = vrow + 1
   Wend
   reixa.Redraw = True
End Sub

Private Sub Command2_Click()
Dim vrow As Double
   btot_Click
   vrow = 1
   reixa.Col = 2
   reixa.Redraw = False
   While vrow < reixa.Rows
      reixa.Row = vrow
      If reixa.CellBackColor = QBColor(10) Then reixa.RowHeight(vrow) = 0  'si el proveidor es verd clar amago la fila
      vrow = vrow + 1
   Wend
   reixa.Redraw = True
End Sub

Private Sub Form_Activate()
vordrereixa = "nomproveidor"
carregar_reixa
enviament_LOTS_pentdents
End Sub
Sub enviament_LOTS_pentdents()
  Dim vmsg As String
  Dim vrow As Double
  Dim i As Byte
  Dim vnomfitxer As String
  If llegir_ini("Qualitat", "emailreclamacioCQs", rutadelfitxer(cami) + "valorsprograma.ini") = Format(Now, "dd/mm/yy") Then Exit Sub
  vnomfitxer = "c:\temp\RelaciodeCQspendents.csv"
  If existeix(vnomfitxer) Then Kill vnomfitxer
  btot_Click
   vrow = 1
   reixa.Col = 2
   reixa.Redraw = False
   Open vnomfitxer For Output As #2
   Print #2, "Data;Albarà;Proveidor;Lot1;Lot2;Lot3;Lot4;Lot5;Lot6;Lot7;Lot8;Lot9;Lot10"
   While vrow < reixa.Rows
      reixa.Row = vrow
      reixa.Col = 2
      If reixa.CellBackColor = 0 And reixa.TextMatrix(vrow, 4) <> "" Then
           vmsg = reixa.TextMatrix(vrow, 1) + ";" + reixa.TextMatrix(vrow, 2) + ";" + reixa.TextMatrix(vrow, 3)
           For i = 4 To reixa.Cols - 1
             reixa.Col = i
             If reixa.CellBackColor = 0 And reixa.Text <> "" Then
                    vmsg = vmsg + ";[" + reixa.Text + "]"
             End If
           Next i
           Print #2, vmsg
      End If
      vrow = vrow + 1
   Wend
   Close #2
   reixa.Redraw = True
   If existeix(vnomfitxer) Then
       enviaremailgenericambadjunt "compres@inplacsa.com", "Llistat de CQ's pendents de linkar. " + Format(Now, "dd/mm/yyyy"), "Adjunto fitxer amb Lots pendents de CQ.", vnomfitxer
       escriure_ini "Qualitat", "emailreclamacioCQs", Format(Now, "dd/mm/yy"), rutadelfitxer(cami) + "valorsprograma.ini"
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   vteclaapretada = KeyCode
  
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   vteclaapretada = 0

End Sub

Private Sub Form_Load()
 Dim arguments As Variant
   arguments = ObtenerLíneaComando
   cami = llegir_ini("General", "cami", "comandes.ini")
   Set dbqualitat = OpenDatabase(rutadelfitxer(cami) + "qualitat.mdb")
   configurar_reixa
   
   
End Sub
Sub carregar_reixa()
   Dim vsql As String
   Dim rst As Recordset
   etactualitzant.Visible = True
   reixa.Redraw = False
   ratoli "espera"
   DoEvents
   configurar_reixa
   vfilareixa = 0
   Set rst = dbqualitat.OpenRecordset("select * from albaransbip order by data desc")
   vsql = "select * from registre_escanejades_expedicions where tipus='ALB' and datarev=null"
   Set rstqualitat = dbqualitat.OpenRecordset(vsql + " order by " + vordrereixa)
   If Not rstqualitat.EOF Then rstqualitat.MoveLast: rstqualitat.MoveFirst
   Set rstescanejades = dbqualitat.OpenRecordset("select * from registre_escanejades_expedicions order by data desc")
   While Not rstqualitat.EOF
'      If rstqualitat!numalbara = "11105" Then Stop
      rst.FindFirst "numalbaraprov='" + atrim(rstqualitat!numalbara) + "'"
      If rst.NoMatch Then
           rst.FindFirst "numalbaraprov='" + substituirtots(rstqualitat!numalbara, "_", "/") + "'"
      End If
      If rst.NoMatch Then MsgBox "No trobo el numero d'albarà " + atrim(rstqualitat!numalbara)
      carregar_filaalareixa rstqualitat, IIf(rst.NoMatch, "", atrim(rst!Data))
      rstqualitat.MoveNext
   Wend
   reixa.Redraw = True
   etactualitzant.Visible = False
fi:
   Set rst = Nothing
   ratoli "normal"
End Sub
Function comprovar_VerificacioCQ(vnumlot As String, vcodiproveidor As String) As Byte
   If vnumlot = "" Then Exit Function
   rstescanejades.FindFirst "numlotproveidor='" + atrim(vnumlot) + "' and codiproveidor='" + atrim(vcodiproveidor) + "'"
   If rstescanejades.NoMatch Then rstescanejades.FindFirst "numlotproveidor='" + substituirtots(atrim(vnumlot), "/", "_") + "' and codiproveidor='" + atrim(vcodiproveidor) + "'"
   If Not rstescanejades.NoMatch Then
        If Not IsNull(rstescanejades!datarev) Then
             comprovar_VerificacioCQ = 1
              Else: comprovar_VerificacioCQ = 2
        End If
   End If
End Function
Sub carregar_filaalareixa(rst As Recordset, vdataalbara As String)
   Dim rstCQ As Recordset
   Dim vCQ_verificat As Byte
   Dim i As Byte
   Dim vtotselslotsescanejats As Boolean
   Dim vdataconcertada As String
   Dim vtipuscontrolCQ As String
   Dim vcolor As Byte
   
   buscar_tipus_controlCQ vtipuscontrolCQ, vdataconcertada, rst!codiproveidor
   'If vtipuscontrolCQ = "" Then GoTo fi
   If reixa.Rows <= vfilareixa + 1 Then reixa.Rows = vfilareixa + 2
   reixa.TextMatrix(vfilareixa + 1, 0) = rst!ID
   reixa.TextMatrix(vfilareixa + 1, 1) = atrim(Format(vdataalbara, "dd/mm/yy"))
   reixa.TextMatrix(vfilareixa + 1, 2) = " " + atrim(rst!numalbara)
   reixa.TextMatrix(vfilareixa + 1, 3) = " " + atrim(rst!nomproveidor)
  ' If InStr(1, atrim(rst!nomproveidor), "UTIEL") > 0 Then Stop
   
   Set rstCQ = dbqualitat.OpenRecordset("SELECT distinct albaransbip.numlotproveidor From albaransbip WHERE (((albaransbip.numalbaraprov)='" + atrim(rst!numalbara) + "') and codiproveidorcomercial=" + atrim(cadbl(rst!codiproveidor)) + ")")
   If rstCQ.EOF Then Set rstCQ = dbqualitat.OpenRecordset("SELECT distinct albaransbip.numlotproveidor From albaransbip WHERE (((albaransbip.numalbaraprov)='" + atrim(substituirtots(rst!numalbara, "_", "/")) + "') and codiproveidorcomercial=" + atrim(cadbl(rst!codiproveidor)) + ")")
   If Not rstCQ.EOF Then rstCQ.MoveLast: rstCQ.MoveFirst
   i = 4
   'vtotselslotsescanejats = False
   
   vtotselslotsescanejats = True
   While Not rstCQ.EOF
      reixa.TextMatrix(vfilareixa + 1, i + rstCQ.AbsolutePosition) = atrim(rstCQ!numlotproveidor)
      If vtipuscontrolCQ <> "Calidad concertada" And vtipuscontrolCQ <> "" Then
             vCQ_verificat = comprovar_VerificacioCQ(atrim(rstCQ!numlotproveidor), rst!codiproveidor)
              Else: vCQ_verificat = 1
      End If
      If vCQ_verificat > 0 Then
             reixa.Row = vfilareixa + 1
             reixa.Col = i + rstCQ.AbsolutePosition
             vcolor = IIf(vCQ_verificat = 1, 10, 12) ' 10 verd   12 vermell
             If vtipuscontrolCQ = "" Then vcolor = 11
             If vtipuscontrolCQ = "Calidad concertada" Then
                   If DateDiff("d", Now, vdataconcertada) <= 0 Then
                         vcolor = 13  'Concertada però amb data caducada  ' 13 fucsia clar
                          Else: vcolor = 11  'concertada amb data correcte ' 11 agua marina clar
                   End If
             End If
             reixa.CellBackColor = QBColor(vcolor)
              Else: vtotselslotsescanejats = False
      End If
      rstCQ.MoveNext
   Wend
   If vtotselslotsescanejats Then
      reixa.Row = vfilareixa + 1
      reixa.Col = 2
      reixa.CellBackColor = QBColor(10) 'verd
      reixa.Col = 3
      reixa.CellBackColor = QBColor(10) 'verd
   End If
   vfilareixa = vfilareixa + 1
fi:
   Set rstCQ = Nothing
End Sub
Sub buscar_tipus_controlCQ(vtipus As String, vdata As String, vcodicomptable As String)
   Dim rst As Recordset
   Dim vsql As String
   vsql = "SELECT proveidors_comercial.codiproduccio, proveidors_comercial.nom, proveidors_comercial.codicomptable, proveidors_qualitat.tipuscontrolCQ, proveidors_qualitat.CQ_datacaducitat "
   vsql = vsql + " FROM proveidors_comercial LEFT JOIN proveidors_qualitat ON proveidors_comercial.codiproduccio = proveidors_qualitat.codiproveidor;"
   Set rst = dbqualitat.OpenRecordset(vsql)
   If Not rst.EOF Then
       rst.FindFirst "codicomptable='" + atrim(vcodicomptable) + "'"
       If Not rst.NoMatch Then
           vtipus = atrim(rst!tipuscontrolcq)
           vdata = atrim(rst!cq_datacaducitat)
           If vdata = "" Then vdata = "01/10/2000"
       End If
   End If
   Set rst = Nothing
End Sub
Sub guardar_nom_camps(rst As Recordset)
   Dim valb As String
   Dim vnomp As String
   Dim vcodiprov As String
   valb = Mid(rst!nomfitxer, 1, InStr(1, rst!nomfitxer, "[") - 1)
   vcodiprov = Mid(rst!nomfitxer, InStr(1, rst!nomfitxer, "[") + 1)
   vcodiprov = Mid(vcodiprov, 1, InStr(1, vcodiprov, "]-") - 1)
   vnomp = Mid(rst!nomfitxer, InStr(1, rst!nomfitxer, "]-") + 2)
   vnomp = Mid(vnomp, 1, InStr(1, LCase(vnomp), ".pdf"))
   rst.Edit: rst!numalbara = valb: rst!nomproveidor = vnomp: rst!codiproveidor = vcodiprov: rst.Update
End Sub
Sub configurar_reixa(Optional vtreureordre As Boolean)
    Dim i As Byte
    Dim vcol As Double
    If vtreureordre Then GoTo treure_ordre
    reixa.FixedCols = 0
    reixa.FixedRows = 1
    reixa.Row = 0
    reixa.Cols = 19
    reixa.ColWidth(0) = 0
    reixa.ColWidth(1) = 1000
    reixa.ColWidth(2) = 1500
    reixa.ColWidth(3) = 3000
    reixa.TextMatrix(0, 1) = "id"
    reixa.TextMatrix(0, 1) = "Data Alb."
    reixa.TextMatrix(0, 2) = "Albarà Prov."
    reixa.TextMatrix(0, 3) = "Nom Prov."
    For i = 4 To 18
       reixa.ColWidth(i) = 1200
       reixa.TextMatrix(0, i) = "LOT_" + atrim(i - 3)
    Next i
treure_ordre:
    vcol = reixa.Col
    For i = 0 To reixa.Cols - 1
       reixa.Col = i
       reixa.CellTextStyle = flexTextFlat
    Next i
    reixa.Col = vcol
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

Private Sub Form_Resize()
  On Error Resume Next
   
   reixa.Width = formqualitat.Width - 500
   reixa.Height = formqualitat.Height - reixa.Top - 900
End Sub

Private Sub reixa_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
' Compare Date Values and Indicate which comes first.
    ' Giving a Date Sort.
    With reixa
      If .TextMatrix(Row1, 1) <> "" Then
       If .CellTextStyle = flexTextFlat Then
            Cmp = IIf(CDate(.TextMatrix(Row1, 1)) > CDate(.TextMatrix(Row2, 1)), 1, -1)
          Else: Cmp = IIf(CDate(.TextMatrix(Row1, 1)) < CDate(.TextMatrix(Row2, 1)), 1, -1)
       End If
      End If
    End With
End Sub
Function totselsCQrevisats(vrow As Double) As Boolean
   Dim i As Byte
   totselsCQrevisats = True
   reixa.Row = vrow
   For i = 4 To reixa.Cols - 1
      reixa.Col = i
      If reixa.Text = "" Then Exit For
      If reixa.CellBackColor <> QBColor(10) And reixa.CellBackColor <> QBColor(11) Then totselsCQrevisats = False
   Next i
End Function
Sub filtrar_valorcolumna()
Dim vrow As Double
Dim vvalor As String
Dim vcol As Double
   vcol = reixa.Col
   vvalor = reixa.Text
   btot_Click
   vrow = 1
   reixa.Col = vcol
   reixa.Redraw = False
   While vrow < reixa.Rows
      reixa.Row = vrow
      If reixa.Text <> vvalor Then reixa.RowHeight(vrow) = 0
      vrow = vrow + 1
   Wend
   reixa.Redraw = True
End Sub
Private Sub reixa_DblClick()
   Dim vnomfitxer As String
   Dim vresp As Double
   Dim vcodiproveidor As String
   
   If reixa.Col = 3 Or reixa.Col = 1 Then filtrar_valorcolumna
   If reixa.Col = 2 Then 'albarà
      rstescanejades.FindFirst "id=" + reixa.TextMatrix(reixa.Row, 0)
      If Not rstescanejades.NoMatch Then
            vnomfitxer = rstescanejades!rutadestilocal + rstescanejades!nomfitxer
            obrir_document vnomfitxer
            If existeix(vnomfitxer) Then
               If totselsCQrevisats(reixa.Row) Then
                     wait 2
                     vresp = MsgBox("Es correcte aquest Albarà?" + vbNewLine + "VOLS MARCAR-LO COM A VERIFICAT?", vbExclamation + vbDefaultButton2 + vbYesNo, "VERIFICACIÓ")
                       Else: MsgBox "Per poder firmar l'albarà primer has de revisat tots els CQ's.", vbCritical, "Atenció"
               End If
                 Else: MsgBox "No he trobat el PDF de l'albarà: " + vbNewLine + vnomfitxer
            End If
            If vresp = 6 Then firmar_albara rstescanejades
      End If
   End If
   If reixa.Col >= 4 And reixa.Text <> "" Then 'certifiat
      rstescanejades.FindFirst "id=" + reixa.TextMatrix(reixa.Row, 0)
      If Not rstescanejades.NoMatch Then
            If vteclaapretada = 17 Then vresp = 6: GoTo firmar
            vcodiproveidor = atrim(rstescanejades!codiproveidor)
            rstescanejades.FindFirst "numlotproveidor='" + atrim(reixa.Text) + "' and codiproveidor='" + vcodiproveidor + "'"
            If rstescanejades.NoMatch Then rstescanejades.FindFirst "numlotproveidor like '" + substituirtots(atrim(reixa.Text), "/", "?") + "' and codiproveidor='" + vcodiproveidor + "'"
            If Not rstescanejades.NoMatch Then
                 vnomfitxer = rstescanejades!rutadestilocal + rstescanejades!nomfitxer
                 obrir_document vnomfitxer
                 If existeix(vnomfitxer) Then
                     wait 2
                     If Not rstescanejades!revisat Then
                            vresp = MsgBox("Es correcte aquest Certificat?" + vbNewLine + "VOLS MARCAR-LO COM A VERIFICAT?", vbExclamation + vbDefaultButton2 + vbYesNo, "VERIFICACIÓ")
                     End If
                 End If
                   Else: MsgBox "No s'ha trobat el CQ escanejat.", vbCritical, "Error"
            End If
firmar:
            If vresp = 6 Then firmar_certificat rstescanejades
      End If
   End If
End Sub
Sub firmar_certificat(rstescanejades As Recordset)
    Dim vvalues As String
    vvalues = " revisat=true,datarev=now,operari='" + nomordinador + "' "
    dbqualitat.Execute "update registre_escanejades_expedicions set " + vvalues + " where id=" + atrim(rstescanejades!ID)
    reixa.CellBackColor = QBColor(10)

End Sub
Sub firmar_albara(rstescanejades As Recordset)
    Dim vvalues As String
    vvalues = " revisat=true,datarev=now,operari='" + nomordinador + "' "
    dbqualitat.Execute "update registre_escanejades_expedicions set " + vvalues + " where id=" + atrim(rstescanejades!ID)
    
    reixa.RemoveItem reixa.Row
End Sub

Private Sub reixa_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = 67 Then
      Clipboard.Clear
      Clipboard.SetText reixa.Text
  End If
End Sub

Private Sub reixa_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   
  If y < reixa.CellHeight Then
      'If reixa.Col = 1 Then vordrereixa = "data": carregar_reixa
   '   If reixa.Col = 3 Then vordrereixa = "nomproveidor": carregar_reixa
        reixa.Row = 0
        If reixa.CellTextStyle = flexTextFlat Then
            configurar_reixa True
            reixa.CellTextStyle = flexTextRaised
            reixa.Sort = IIf(reixa.Col = 1, 9, flexSortGenericAscending)
             Else
              configurar_reixa True
              reixa.CellTextStyle = flexTextFlat
              reixa.Sort = IIf(reixa.Col = 1, 9, flexSortGenericDescending)
        End If
  End If
End Sub

Private Sub reixa_OLEDragDrop(Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  If Data.GetFormat(vbCFFiles) Then
      If existeix(Data.Files(1)) Then guarda_CQ_sicorrespon Data.Files(1)
  End If
End Sub
Sub guarda_CQ_sicorrespon(vfitxer As String)
   'guardar_CQ_sicorrespon(vcodiproveidor As String, vnomproveidor As String, vlotproveidor As String)
  Dim vnomfitxerfinal As String
  Dim vcodiproveidor As String
  Dim vlotproveidor As String
  Dim vnomproveidor As String
  Dim vnumCQ As String
  
  If reixa = "" Or reixa.Col < 4 Or reixa.CellBackColor <> 0 Then Exit Sub
  rstescanejades.FindFirst "id=" + reixa.TextMatrix(reixa.Row, 0)
  If rstescanejades.NoMatch Then Exit Sub
  vlotproveidor = atrim(reixa.Text)
  vcodiproveidor = atrim(rstescanejades!codiproveidor)
  vnomproveidor = atrim(rstescanejades!nomproveidor)
  If vfitxer = "" Or vlotproveidor = "" Or vcodiproveidor = "" Or vnomproveidor = "" Then Exit Sub
  vnomfitxerfinal = "CQ_" + vlotproveidor + " [" + atrim(vcodiproveidor) + "]-" + atrim(vnomproveidor) + ".pdf"
  vnomfitxerfinal = treuresimbolsnovalidsnomfitxer(vnomfitxerfinal)
  If MsgBox("Vols assignar aquest PDF al lot " + atrim(vlotproveidor) + vbNewLine + "del proveïdor: " + vnomproveidor, vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  FileCopy vfitxer, rutadelfitxer(cami) + "Cache_escanejarexpedicions\CQ\" + vnomfitxerfinal
  If existeix(rutadelfitxer(cami) + "Cache_escanejarexpedicions\CQ\" + vnomfitxerfinal) Then
      dbqualitat.Execute "update albaransbip set lotescanejat=true where numlotproveidor='" + atrim(vlotproveidor) + "' and codiproveidorcomercial=" + atrim(vcodiproveidor)
      reixa.CellBackColor = QBColor(5)
      mirarsihihaunaltralotigual atrim(vnomproveidor), atrim(vlotproveidor)
  End If
End Sub
Sub mirarsihihaunaltralotigual(vnomproveidor As String, vlotproveidor As String)
  Dim vcol As Double
  Dim vrow As Double
  Dim x As Long
  Dim y As Long
  vcol = reixa.Col
  vrow = reixa.Row
  For y = 1 To reixa.Rows - 1
    For x = 4 To reixa.Cols - 1
      If atrim(reixa.TextMatrix(y, 3)) = atrim(vnomproveidor) Then
           If reixa.TextMatrix(y, x) = vlotproveidor Then
                reixa.Row = y: reixa.Col = x: reixa.CellBackColor = QBColor(5)
                  Else: If reixa.TextMatrix(y, x) = "" Then Exit For
           End If
            Else: Exit For
      End If
    Next x
  Next y
  reixa.Col = vcol: reixa.Row = vrow
End Sub
Function treuresimbolsnovalidsnomfitxer(desc As String) As String
   desc = substituir(desc, "\", "_")
   desc = substituir(desc, "/", "_")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ":", ";")
   desc = substituir(desc, "?", "¿")
   desc = substituir(desc, "*", "x")
   desc = substituir(desc, """", "'")
   desc = substituir(desc, ">", "+")
   desc = substituir(desc, "<", "-")
   treuresimbolsnovalidsnomfitxer = desc
End Function


Private Sub reixa_OLEDragOver(Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  Dim i As Double
  Dim vcol As Double
  Dim vrow As Double
  Dim viant As Double
  reixa.SetFocus
 ' Me.Caption = atrim(x) + "  -  " + atrim(y) + "  -  " + atrim(reixa.RowPos(2))
  
  For i = 0 To reixa.Cols - 1
        If reixa.ColPos(i) > x Then
            vcol = i - 1: Exit For
             Else: vcol = i
        End If
  Next i
  viant = 0
  i = 0
  While i < reixa.Rows
  'For i = 0 To reixa.Rows - 1
    If reixa.RowIsVisible(i) Then
     If reixa.RowHeight(i) > 0 Then
      If reixa.RowPos(i) > y Then vrow = viant: i = reixa.Rows ': Exit for
      viant = i
     End If
    End If
    i = i + 1
  'Next i
  Wend
  If i = reixa.Rows Then vrow = viant
  If vcol >= 0 And vcol < reixa.Cols And vrow > 0 And vrow <= reixa.Rows Then
       reixa.Col = vcol: reixa.Row = vrow
  End If
End Sub
