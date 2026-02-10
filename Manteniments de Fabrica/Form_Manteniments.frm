VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniments periodics a le Seccions"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18600
   Icon            =   "Form_Manteniments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   18600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CheckFetsSenseRevisar 
      Caption         =   "Fets pero sense revisar."
      Height          =   285
      Left            =   3135
      TabIndex        =   8
      Top             =   1935
      Width           =   2610
   End
   Begin VB.CheckBox CheckVeureTots 
      Caption         =   "Veure tots els registres."
      Height          =   285
      Left            =   450
      TabIndex        =   7
      Top             =   1920
      Width           =   2610
   End
   Begin VB.ComboBox Comboequipament 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   420
      TabIndex        =   5
      Top             =   1230
      Width           =   3195
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   4950
      Left            =   165
      TabIndex        =   4
      Top             =   2250
      Width           =   18240
      _ExtentX        =   32173
      _ExtentY        =   8731
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox cfirma 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4380
      TabIndex        =   2
      Top             =   510
      Width           =   4365
   End
   Begin VB.ComboBox ComboSeccio 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   375
      TabIndex        =   0
      Top             =   375
      Width           =   3195
   End
   Begin VB.Label Label3 
      Caption         =   "Equipament"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   900
      TabIndex        =   6
      Top             =   885
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "Nom del que Firma"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5055
      TabIndex        =   3
      Top             =   195
      Width           =   2910
   End
   Begin VB.Label Label1 
      Caption         =   "Secció"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   855
      TabIndex        =   1
      Top             =   30
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cnnDirecta As New ADODB.Connection
    

Private Sub cfirma_LostFocus()
   escriure_ini "Manteniments", "nomdelquefirma", cfirma, "comandes.ini"
End Sub

Private Sub CheckFetsSenseRevisar_Click()
  If Comboequipament <> "" Then Comboequipament_Click
End Sub

Private Sub CheckVeureTots_Click()
  ' If Comboequipament <> "" Then
      emplenar_combo_Seccions
      ComboSeccio_Click ': Comboequipament_Click
  ' End If
End Sub

Private Sub Comboequipament_Click()
  Dim vsql As String
  Dim vsql2 As String
  vsql = " cdsec=" + atrim(ComboSeccio.ItemData(ComboSeccio.ListIndex)) + " and cdeq=" + atrim(Comboequipament.ItemData(Comboequipament.ListIndex))
  If rs.State = adStateOpen Then rs.Close
  If CheckVeureTots.Value = 0 Then
      vsql2 = " and isdate(FechaPrev)  and Fecha is null " + " order by FechaPrev asc"
        Else: vsql2 = " order by FechaPrev desc"
  End If
  If CheckFetsSenseRevisar.Value = 1 Then
       vsql2 = " and Fecha is not null  and  cdcgrev is null" + " order by Fecha desc"
  End If
  
  rs.Open "SELECT * FROM [Manteniments] where " + vsql + vsql2, cnnDirecta, adOpenStatic, adLockOptimistic
  v = llegir_ini("Manteniments", "amplades", "comandes.ini")
  If v = "{[}]" Then v = ""
  v = "Desscar=0,Deseq=0,Cdregman=0,cdmaq=5,Cdscar=0,Cdeq=5,Cdsec=0,Cdcgrev=0" + IIf(v <> "", "," + v, "")
  carregar_reixa rs, , v
    
End Sub

Sub emplenar_combo_Equipament(rs As ADODB.Recordset)
    Dim vid As Long
    Dim v As Long
    Comboequipament.Clear
    While Not rs.EOF
      Comboequipament.AddItem rs!Deseq
      Comboequipament.ItemData(Comboequipament.NewIndex) = rs!Cdeq
      rs.MoveNext
    Wend
  
End Sub



Public Sub carregar_reixa(ByRef rs As ADODB.Recordset, _
                          Optional ByVal campsExclosos As String = "", _
                          Optional ByVal ampladesManuals As String = "")
    
    ' campsExclosos: llista separada per comes, ex: "id,password,timestamp"
    ' ampladesManuals: llista camp=valor, ex: "Nom=3000,Data=1200"

    Dim i As Integer, j As Integer
    Dim colIndex As Integer
    Dim nomCamp As String
    Dim ampleText As Single, ampleMaxim As Single
    Dim trobat As Boolean
    reixa.Rows = 1
    reixa.Clear
    reixa.Redraw = False
    
    If rs.EOF And rs.BOF Then
        reixa.Redraw = True
        Exit Sub
    End If

    ' 1. Comptar quantes columnes reals tindrem (excloent les de la llista)
    Dim numCols As Integer
    numCols = 0
    For i = 0 To rs.Fields.Count - 1
        If InStr(1, "," & campsExclosos & ",", "," & rs.Fields(i).Name & ",", vbTextCompare) = 0 Then
            numCols = numCols + 1
        End If
    Next i

    reixa.Cols = numCols
    reixa.Rows = 1
    
    ' 2. Crear Capçaleres (només dels camps permesos)
    colIndex = 0
    For i = 0 To rs.Fields.Count - 1
        nomCamp = rs.Fields(i).Name
        ' Si el camp NO està a la llista d'exclosos...
        If InStr(1, "," & campsExclosos & ",", "," & nomCamp & ",", vbTextCompare) = 0 Then
            reixa.TextMatrix(0, colIndex) = nomCamp
            reixa.Row = 0: reixa.Col = colIndex
            reixa.CellFontBold = True
            colIndex = colIndex + 1
        End If
    Next i

    ' 3. Omplir Dades
    rs.MoveFirst
    i = 1
    Do While Not rs.EOF
        reixa.AddItem ""
        colIndex = 0
        For j = 0 To rs.Fields.Count - 1
            nomCamp = rs.Fields(j).Name
            If InStr(1, "," & campsExclosos & ",", "," & nomCamp & ",", vbTextCompare) = 0 Then
                reixa.TextMatrix(i, colIndex) = IIf(IsNull(rs.Fields(j).Value), "", rs.Fields(j).Value)
                colIndex = colIndex + 1
            End If
        Next j
        i = i + 1
        rs.MoveNext
    Loop

    ' 4. Ajustar Amplades (Manual o AutoSize)
    For j = 0 To reixa.Cols - 1
        nomCamp = reixa.TextMatrix(0, j)
        ampleMaxim = 0
        
        ' Busquem si aquest camp té una amplada manual definida (ex: "Nom=3000")
        Dim posManual As Integer
        posManual = InStr(1, ampladesManuals, nomCamp & "=", vbTextCompare)
        
        If posManual > 0 Then
            ' Extraiem el valor numèric després del "="
            Dim startPos As Integer, endPos As Integer
            startPos = posManual + Len(nomCamp) + 1
            endPos = InStr(startPos, ampladesManuals & ",", ",")
            reixa.ColWidth(j) = Val(Mid(ampladesManuals, startPos, endPos - startPos))
        Else
            ' Si no hi ha amplada manual, fem l'AutoSize d'abans
            For i = 0 To reixa.Rows - 1
                ampleText = Me.TextWidth(reixa.TextMatrix(i, j))
                If ampleText > ampleMaxim Then ampleMaxim = ampleText
            Next i
            reixa.ColWidth(j) = ampleMaxim + 200
        End If
    Next j

    reixa.FixedRows = 1
    reixa.Redraw = True
End Sub

Private Sub ComboSeccio_Click()
  Dim vsql As String
  Dim vsql2 As String
  escriure_ini "Manteniments", "NomSeccioPredeterminada", ComboSeccio, "comandes.ini"
  escriure_ini "Manteniments", "IdSeccioPredeterminada", ComboSeccio.ItemData(ComboSeccio.ListIndex), "comandes.ini"
  vsql = " cdsec=" + atrim(ComboSeccio.ItemData(ComboSeccio.ListIndex))
  vsql2 = IIf(CheckVeureTots.Value = 1, " fecha is not null and ", " fecha is null and ")
  If rs.State = adStateOpen Then rs.Close
  rs.Open "SELECT distinct cdeq,deseq FROM [Manteniments] where " + vsql2 + vsql, cnnDirecta, adOpenStatic, adLockOptimistic
  emplenar_combo_Equipament rs
  Comboequipament.Text = ""
  reixa.Clear
End Sub

Private Sub Form_Load()
    'Set cnn = New ADODB.Connection
    Dim strConn As String
    Dim strConnUpdate
    Set rs = New ADODB.Recordset
    cami = llegir_ini("General", "cami", "comandes.ini")
  '  strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + rutadelfitxer(cami) + "MantenimentMaquinesv2003.mdb;"
    strConnUpdate = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\ord_copies\Qualitat_Proveidors\Qualitat Mdb\ISO 9001-2000 - DEFINITIVO.mdb;"
    'cnn.Open strConnUpdate 'strConn
    cnnDirecta.Open strConnUpdate
    emplenar_combo_Seccions
    
    
    cfirma = atrim(llegir_ini("Manteniments", "nomdelquefirma", "comandes.ini"))
    ComboSeccio = llegir_ini("Manteniments", "NomSeccioPredeterminada", "comandes.ini")
    
  'escriure_ini "Manteniments", "IdSeccioPredeterminada", ComboSeccio.ItemData(ComboSeccio.ListIndex), "comandes.ini"
End Sub
Sub emplenar_combo_Seccions()
    Dim vid As Long
    Dim v As Long
    Dim rs As ADODB.Recordset
    Dim vsql As String
    v = -1
    vid = cadbl(llegir_ini("Manteniments", "IdSeccioPredeterminada", "comandes.ini"))
    Set rs = New ADODB.Recordset
    If rs.State = adStateOpen Then rs.Close
    vsql = " fecha is null"
    If CheckVeureTots.Value = 1 Then vsql = " fecha is not null "
    rs.Open "SELECT distinct Desscar,cdscar FROM [Manteniments] where " + vsql, cnnDirecta, adOpenStatic, adLockOptimistic
    While Not rs.EOF
      ComboSeccio.AddItem rs!Desscar
      ComboSeccio.ItemData(ComboSeccio.NewIndex) = rs!Cdscar
      If rs!Cdscar = vid Then
          v = ComboSeccio.NewIndex
      End If
      rs.MoveNext
    Wend
    If v > -1 Then ComboSeccio.ListIndex = v
End Sub
Public Function obtenir_amplades_actuals() As String
    Dim j As Integer
    Dim resultat As String
    Dim nomCamp As String
    Dim ample As Long

    resultat = ""
    
    ' Recorrem totes les columnes de la reixa
    For j = 0 To reixa.Cols - 1
        nomCamp = reixa.TextMatrix(0, j) ' Agafem el nom de la capçalera
        ample = reixa.ColWidth(j)        ' Agafem l'amplada actual en twips
        
        ' Concatenem en el format "Camp=Valor,"
        resultat = resultat & nomCamp & "=" & CStr(ample) & ","
    Next j
    
    ' Retornem la cadena (treient l'última coma si existeix)
    If Len(resultat) > 0 Then
        obtenir_amplades_actuals = Left(resultat, Len(resultat) - 1)
    Else
        obtenir_amplades_actuals = ""
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)
  Dim v As String
  v = obtenir_amplades_actuals
  escriure_ini "Manteniments", "amplades", v, "comandes.ini"
End Sub
Public Function BuscarNumColumna(ByRef reixa As MSFlexGrid, ByVal NomColumna As String) As Integer
    Dim i As Integer
    Dim indexTrobat As Integer
    
    indexTrobat = -1 ' Valor per defecte si no es troba la columna
    
    ' Recorrem totes les columnes de la reixa
    For i = 0 To reixa.Cols - 1
        ' Comparem el text de la capçalera (fila 0) amb el nom buscat
        If UCase(reixa.TextMatrix(0, i)) = UCase(NomColumna) Then
            indexTrobat = i
            Exit For
        End If
    Next i
    
    BuscarNumColumna = indexTrobat
End Function
Private Sub Timer1_Timer()

End Sub

Private Sub reixa_DblClick()
    Dim rst As ADODB.Recordset
    Dim vP As String
    Dim vQ As Long
    Dim vdataprevista As String
    Dim SQL As String
   
    
    If reixa.Row = 0 Then Exit Sub
    

    Set rst = New ADODB.Recordset
    vCdRegMan = reixa.TextMatrix(reixa.Row, BuscarNumColumna(reixa, "Cdregman"))
    If CheckFetsSenseRevisar.Value = 0 Then
            If MsgBox("Vols marcar aquesta feina com a feta a dia d'avui?", vbExclamation + vbDefaultButton2 + vbYesNo, "Feina feta?") = vbNo Then GoTo fi
    End If
    
    If rst.State = adStateOpen Then rst.Close
    rst.Open "SELECT * FROM [Registros mantenimiento] where cdregman=" + atrim(vCdRegMan), cnnDirecta, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
       vSqlFirmaRamon = "Cdcgrev = 40, FirmaRevisor = 'RAMON SIMON' "
       If CheckFetsSenseRevisar.Value = 0 Then
                 If MsgBox("Aquest manteniment també l'ha revisat l'encarregat de revisió?" + vbNewLine + "Si no ho ha revisat l'encarregat de revisió se l'hi enviarà un email.", vbExclamation + vbDefaultButton1 + vbYesNo, "Revisat?") = vbNo Then
                        vSqlFirmaRamon = "Cdcgrev = null, FirmaRevisor = '' "
                        enviaremailgeneric "miquel.inplacsa@gmail.com", "Revisió de manteniment a fàbrica pendent " + ComboSeccio.Text + "->" + Comboequipament.Text, "S'hauria de revisar algunes reparacions fetes i donar el vist-i-plau."
                 End If
                 SQL = "UPDATE [Registros mantenimiento] SET " & _
                        vSqlFirmaRamon & "," & _
                        "fecha = '" & Format(Date, "yyyy-mm-dd") & "', " & _
                        "Fecharev = '" & Format(Date, "yyyy-mm-dd") & "', " & _
                        "FirmaMant = '" & treure_apostruf(cfirma) & "', " & _
                        "Alb = '" + Format(Date, "yymmdd") + "' " & _
                        "WHERE cdregman = " & Trim$(vCdRegMan)
                
                cnnDirecta.Execute SQL, , adExecuteNoRecords   'he hagut de fer update directament a la taula principal TARDAVA MOLT
                
                
                 vQ = reixa.TextMatrix(reixa.Row, BuscarNumColumna(reixa, "periodanys"))
                 vP = reixa.TextMatrix(reixa.Row, BuscarNumColumna(reixa, "Tipoperiodo"))
                 vdataprevista = CrearDataProximaRevisio(vQ, vP)
                 vdataprevista = InputBox("Per programar una nova revisió per d'aqui " + atrim(vQ) + " " + vP + "?", "Nova Data prevista", Format(vdataprevista, "dd/mm/yy"))
                 If StrPtr(vdataprevista) = 0 Then GoTo fi
                 If Not IsDate(vdataprevista) Then MsgBox "Aquesta data no es vàlida": GoTo fi
                 
                 ClonarIColocar rst
                 rst.Fields("Cdcgrev").Value = Null
                 rst.Fields("Fecharev").Value = Null
                 rst.Fields("Fecha").Value = Null
                 rst.Fields("Alb").Value = " "
                 rst.Fields("FirmaRevisor").Value = ""
                 rst.Fields("FechaPrev").Value = vdataprevista
                 rst.Update
             Else
              If MsgBox("Aquest mantenimentl'ha revisat l'encarregat de revisió?", vbExclamation + vbDefaultButton1 + vbYesNo, "Revisat?") = vbYes Then
                 SQL = "UPDATE [Registros mantenimiento] SET " & _
                        vSqlFirmaRamon & _
                        "WHERE cdregman = " & Trim$(vCdRegMan)
                 cnnDirecta.Execute SQL, , adExecuteNoRecords   'he hagut de fer update directament a la taula principal TARDAVA MOLT
              End If
                
        End If
   End If

fi:
  If rst.State = adStateOpen Then rst.Close
  Set rst = Nothing
  Comboequipament_Click
End Sub
Function CrearDataProximaRevisio(vQ As Long, vT As String) As Date
  Dim vPeriode As String
  If vT = "Dias" Then vPeriode = "d"
  If vT = "Semanas" Then vPeriode = "ww"
  If vT = "Meses" Then vPeriode = "m"
  If vT = "Años" Then vPeriode = "yyyy"
  If vPeriode <> "" Then CrearDataProximaRevisio = DateAdd(vPeriode, vQ, Now)
 
End Function
Public Sub ClonarIColocar(ByRef rst As ADODB.Recordset)
    Dim i As Integer, noms(), valors()
    
    ' Preparem matrius ignorant el camp 0 (ID autonumèric)
    ReDim noms(rst.Fields.Count - 2)
    ReDim valors(rst.Fields.Count - 2)
    
    For i = 1 To rst.Fields.Count - 1
        noms(i - 1) = rst.Fields(i).Name
        valors(i - 1) = rst.Fields(i).Value
    Next i
    
    ' En fer AddNew amb dades i Update, el cursor es queda al nou registre
    rst.AddNew noms, valors
    rst.Update
End Sub
