VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form Formagrupartreballs 
   BackColor       =   &H00EAD9CE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agrupar treballs semblants"
   ClientHeight    =   12315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16065
   Icon            =   "Formagrupartreballs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12315
   ScaleWidth      =   16065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Veure última comanda"
      Height          =   405
      Left            =   6360
      TabIndex        =   15
      Top             =   4890
      Width           =   1860
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Vista previa PDF"
      Height          =   7005
      Left            =   240
      TabIndex        =   9
      Top             =   5265
      Width           =   15585
      Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
         Height          =   6675
         Left            =   150
         TabIndex        =   16
         Top             =   210
         Width           =   15300
         _cx             =   5080
         _cy             =   5080
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<-- Assignar número"
      Height          =   390
      Left            =   4365
      TabIndex        =   8
      Top             =   4875
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   3585
      Left            =   195
      TabIndex        =   3
      Top             =   1185
      Width           =   15720
      _ExtentX        =   27728
      _ExtentY        =   6324
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F1B75F&
      Caption         =   "Buscar"
      Height          =   1050
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   15720
      Begin VB.CommandButton Command2 
         Height          =   345
         Left            =   11160
         Picture         =   "Formagrupartreballs.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   285
         Width           =   420
      End
      Begin VB.TextBox ccodiliniav 
         Height          =   360
         Left            =   13590
         TabIndex        =   12
         Top             =   255
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox ccodilinia 
         Height          =   360
         Left            =   12870
         TabIndex        =   10
         Top             =   255
         Width           =   555
      End
      Begin VB.TextBox cpantone 
         Height          =   360
         Left            =   7530
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   3585
      End
      Begin VB.CommandButton bbuscar 
         Height          =   420
         Left            =   14355
         Picture         =   "Formagrupartreballs.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   255
         Width           =   1140
      End
      Begin VB.TextBox cmarcailinia 
         Height          =   360
         Left            =   1140
         TabIndex        =   1
         Top             =   270
         Width           =   5250
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13455
         TabIndex        =   13
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi de linia:"
         Height          =   285
         Left            =   11910
         TabIndex        =   11
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pantone 1:"
         Height          =   285
         Left            =   6705
         TabIndex        =   6
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Marca i linia:"
         Height          =   285
         Left            =   135
         TabIndex        =   2
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.Label etcodidelinia 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00ED823A&
      Height          =   435
      Left            =   390
      TabIndex        =   7
      Top             =   4875
      Width           =   3990
   End
End
Attribute VB_Name = "Formagrupartreballs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vordrereixa As Byte
Dim vensenya_ultima_comanda As Boolean
Private Sub Command1_Click()
  Dim v As String
  Dim vnum As Double
  Dim vversio As Double
  Dim rst As Recordset
  v = InputBox("Escriu el número de linia que vols assignar-hi." + vbNewLine + "NO ESCRIGUIS RES PER ASSIGNAR AUTOMÀTICAMENT UN DE NOU." + vbNewLine + vbNewLine + "ESCRIU -1 PER ELIMINAR L'ASSIGNACIÓ.", "Escullir número de linia.")
  If StrPtr(v) = 0 Then Exit Sub
  If cadbl(v) = 0 Then
      If MsgBox("Segur que vols assignar un número nou?", vbDefaultButton2 + vbYesNo + vbExclamation, "Atenció") = vbNo Then GoTo fi
      vnum = 0: vversio = 0
      assignar_numerodelinia cadbl(reixa.TextMatrix(reixa.Row, 0)), cadbl(etcodidelinia.tag), vnum, vversio
      GoTo multiseleccio
  End If
  If cadbl(v) = -1 Then vnum = -1: vversio = 0: assignar_numerodelinia cadbl(reixa.TextMatrix(reixa.Row, 0)), cadbl(etcodidelinia.tag), vnum, vversio: GoTo multiseleccio
  vnum = cadbl(v)
  Set rst = dbclixes.OpenRecordset("select * from modificacions where codidelinia=" + atrim(vnum))
  If rst.EOF Then MsgBox "Aquest CODI DE LINIA no existeix encara.", vbCritical, "Error": GoTo fi
  v = InputBox("Escriu el número de la VERSIÓ DE LINIA que vols assignar-hi." + vbNewLine + "NO ESCRIGUIS RES PER ASSIGNAR AUTOMÀTICAMENT UN DE NOU.", "Escullir número de VERSIÓ DE LINIA.")
  If StrPtr(v) = 0 Then Exit Sub
  If cadbl(v) = 0 Then
     If MsgBox("Segur que vols assignar un número nou?", vbDefaultButton2 + vbYesNo + vbExclamation, "Atenció") = vbNo Then GoTo fi
     vversio = 0
     assignar_numerodelinia cadbl(reixa.TextMatrix(reixa.Row, 0)), cadbl(etcodidelinia.tag), vnum, vversio
       Else: vversio = cadbl(v): assignar_numerodelinia cadbl(reixa.TextMatrix(reixa.Row, 0)), cadbl(etcodidelinia.tag), vnum, vversio
  End If
multiseleccio:
  reixa.TextMatrix(reixa.Row, 2) = format(vnum, "000") + "#" + atrim(vversio)
  possar_numero_multiseleccio vnum, vversio
fi:
  carregar_pdf cadbl(reixa.TextMatrix(reixa.Row, 0))
  
End Sub
Sub possar_numero_multiseleccio(vnum As Double, vversio As Double)
   Dim start_row As Double
   Dim stop_row As Double
   If reixa.Row > reixa.RowSel Then
        start_row = reixa.RowSel
        stop_row = reixa.Row
    Else
        start_row = reixa.Row
        stop_row = reixa.RowSel
   End If

   For i = start_row To stop_row
      'assignar_numerodelinia cadbl(reixa.TextMatrix(i, 0)), cadbl(reixa.TextMatrix(i, 1)), vnum, vversio
      dbclixes.Execute "update modificacions set codidelinia=" + atrim(vnum) + ",codideliniav=" + atrim(vversio) + " where id_treball=" + atrim(reixa.TextMatrix(i, 0)) + " and ordre=" + atrim(reixa.TextMatrix(i, 1))
      reixa.TextMatrix(i, 2) = format(vnum, "000") + "#" + atrim(vversio)
   Next i
End Sub
Sub assignar_numerodelinia(vtreball As Double, vversio As Double, vnumlinia As Double, vnumliniav As Double)
   Dim rst As Recordset
   'si vnumlinia =0 crea la proxima+1 si es -1 posa a zero la linia i versió
   If vnumlinia = 0 Then
       Set rst = dbclixes.OpenRecordset("select max(codidelinia) as maxcodidelinia from modificacions")
       If rst.EOF Then vnumlinia = 0 Else: vnumlinia = cadbl(rst!maxcodidelinia)
       vnumlinia = vnumlinia + 1
       vnumliniav = 1
   End If
   If vnumlinia > 0 And vnumliniav = 0 Then
       Set rst = dbclixes.OpenRecordset("select max(codideliniav) as maxcodideliniav from modificacions where codidelinia=" + atrim(vnumlinia))
       If rst.EOF Then vnumliniav = 0 Else: vnumliniav = cadbl(rst!maxcodideliniav)
       vnumliniav = vnumliniav + 1
   End If
   If vnumlinia = -1 Then vnumlinia = 0: vnumliniav = 0
   dbclixes.Execute "update modificacions set codidelinia=" + atrim(vnumlinia) + ",codideliniav=" + atrim(vnumliniav) + " where id_treball=" + atrim(vtreball) + " and ordre=" + atrim(vversio)
End Sub

Private Sub bbuscar_Click()
  buscar_treballs
End Sub

Sub buscar_treballs()
   Dim rst As Recordset
   Dim i As Long
   Dim vsql As String
   Dim vSubsqlpantones As String
   Dim vcodidelinia As String
   Dim vversio As Double
   Dim vsql_unio As String
   Static vjaheentrat As Boolean
    
   'If cmarcailinia = "" And cpantone <> "" Then If MsgBox("Aquesta consulta només amb PANTONE pot trigar molta estona." + vbNewLine + "Vols continuar igualment?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   If ccodilinia <> "" Then vsql = " (clixes.id_treball in (select id_treball from modificacions where codidelinia<>null and codidelinia=" + atrim(cadbl(ccodilinia)) + "))"
   If cmarcailinia <> "" Then vsql = vsql + IIf(vsql <> "", " and ", "") + " (marca like '*" + treure_apostruf(cmarcailinia) + "*' or linia like '*" + treure_apostruf(cmarcailinia) + "*') "
   If cpantone.tag <> "" And cpantone.tag <> "0" Then
     If Not vjaheentrat Then
      dbclixes.Execute "drop table tmp_tintes_maxversio"
      dbclixes.Execute "SELECT Max(Tintes.ordremodificacio) AS idmodificacio, Tintes.id_treball, Max(Tintes.ordremodificacio) AS ordremodificacio INTO tmp_tintes_maxversio From tintes GROUP BY Tintes.id_treball HAVING (((Max(Tintes.ordremodificacio))<>False));"
      dbclixes.Execute "UPDATE tmp_tintes_maxversio LEFT JOIN Modificacions ON (tmp_tintes_maxversio.ordremodificacio = Modificacions.ordre) AND (tmp_tintes_maxversio.id_treball = Modificacions.id_treball) SET tmp_tintes_maxversio.idmodificacio = [modificacions].[id_modificacio];"
      vjaheentrat = True
     End If
      vSubsqlpantones = "SELECT Tintes.id_treball FROM Tintes LEFT JOIN Tintes AS Tintes_1 ON Tintes.tinterlinkambid_treball = Tintes_1.id_tinter WHERE (IIf([tintes].[tinterlinkambid_treball]>0,[tintes_1].[coditinta],[tintes].[coditinta]))='" + cpantone.tag + "'"
   End If
   If vSubsqlpantones <> "" Then vsql = vsql + IIf(vsql <> "", " and ", "") + " clixes.id_treball in(" + vSubsqlpantones + ")"
   If vsql = "" Then Exit Sub
   vsqlunio = "SELECT Clixes.*, Modificacions.* FROM Clixes RIGHT JOIN (tmp_tintes_maxversio LEFT JOIN Modificacions ON tmp_tintes_maxversio.idmodificacio = Modificacions.id_modificacio) ON Clixes.id_treball = Modificacions.id_treball "
   Set rst = dbclixes.OpenRecordset(vsqlunio + " where databaixaclixe=null And " + vsql)
   ratoli "espera"
   carregar_PDFdeltreball 0, 0
   etcodidelinia = ""
   reixa.Clear
   config_reixa
   i = 1
   While Not rst.EOF
    'If Not hihaelpantoneescullit(rst!id_treball, cpantone, vversio, vcodidelinia) Then GoTo proxim
    'If Not hihaelpantoneescullit(rst!id_treball, cpantone2, vversio, vcodidelinia) Then GoTo proxim
    carregarmesdades rst![clixes.id_treball], vversio, vcodidelinia
    reixa.Rows = i + 1
    reixa.TextMatrix(i, 0) = atrim(rst![clixes.id_treball])
    reixa.TextMatrix(i, 1) = atrim(vversio)
    reixa.TextMatrix(i, 2) = atrim(vcodidelinia)
    reixa.TextMatrix(i, 3) = atrim(rst!marca) + " - " + atrim(rst!linia)
    reixa.TextMatrix(i, 4) = atrim(rst!nomclienttemporal)
    i = i + 1
proxim:
    rst.MoveNext
   Wend
   Set rst = Nothing
   ratoli "normal"
End Sub
Sub carregarmesdades(vtreball As Double, ByRef vversio As Double, ByRef vcodidelinia As String)
  Dim rst As Recordset
  Set rst = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(vtreball) + " order by ordre DESC")
  If rst.EOF Then Exit Sub
  vversio = rst!ordre
  vcodidelinia = atrim(format(cadbl(rst!codidelinia), "000")) + "#" + atrim(cadbl(rst!codideliniav))
  If cadbl(rst!codidelinia) = 0 Then vcodidelinia = ""
  Set rst = Nothing
End Sub


Sub config_reixa()
   reixa.Cols = 5
   reixa.Rows = 1
   reixa.TextMatrix(0, 0) = "Treball"
   reixa.TextMatrix(0, 1) = "Max_Versió"
   reixa.TextMatrix(0, 2) = "CdL"
   reixa.TextMatrix(0, 3) = "Marca i Linia"
   reixa.TextMatrix(0, 4) = "Client"
   reixa.ColWidth(0) = 1000
   reixa.ColWidth(1) = 500
   reixa.ColWidth(2) = 800
   reixa.ColWidth(3) = 7500
   reixa.ColWidth(4) = 4500
End Sub

Private Sub Command2_Click()
  Dim vcodi As Double
  Dim vdescripcio As String
  cpantone = "": cpantone.tag = ""
  triar_coditinta vcodi, vdescripcio
  cpantone = vdescripcio
  cpantone.tag = atrim(vcodi)
End Sub
Sub triar_coditinta(vcodi As Double, vdescripcio As String)
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  
  sql = "SELECT  idtinta,codi,descripcio,referenciacolor from tintes_tot "
  were = " order by descripcio"
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
  formseleccio.Data1.RecordSource = sql + were
  formseleccio.width = 13000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.colocar_botofiltre 2  'coloca els prismatics sobre la columna 2
  formseleccio.caption = "Escull la tinta que vols buscar"
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 1200
  formseleccio.Show 1
  If seleccioret = 1 Then
    If Not formseleccio.Data1.Recordset.EOF Then
        vcodi = atrim(formseleccio.Data1.Recordset!codi)
        vdescripcio = atrim(formseleccio.Data1.Recordset!descripcio)
    End If
  End If
  
  Unload formseleccio
End Sub


Private Sub Command3_Click()
   Set dbtmpb = dbbaixes
   Set dbtmp = dbcomandes
   If vensenya_ultima_comanda Then Unload formannex: vensenya_ultima_comanda = False: Exit Sub
   Load formannex
   'formannex.carregarcomanda 212365
   formannex.Show
   vensenya_ultima_comanda = True
   carregar_comandasical cadbl(reixa.TextMatrix(reixa.Row, 0)), cadbl(reixa.TextMatrix(reixa.Row, 1))
End Sub

Private Sub Form_Activate()
    If UCase(Environ("computername")) <> "ORD_TINTES" Then Command1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If vensenya_ultima_comanda Then Unload formannex
End Sub

Private Sub reixa_Click()
   
   carregar_pdf cadbl(reixa.TextMatrix(reixa.Row, 0))
   carregar_comandasical cadbl(reixa.TextMatrix(reixa.Row, 0)), cadbl(reixa.TextMatrix(reixa.Row, 1))
End Sub
Sub carregar_comandasical(vtreball As Double, vversio As Double)
   Dim rst As Recordset
   If Not vensenya_ultima_comanda Then Exit Sub
   Set rst = dbcomandes.OpenRecordset("select comanda from comandes where numtreball=" + atrim(vtreball) + " and numordremodificacio=" + atrim(vversio) + " order by comanda desc")
   If rst.EOF Then
      Unload formannex
      formannex.Show
      Exit Sub
   End If
   If cadbl(rst!comanda) = 0 Then
      Unload formannex
      formannex.Show
   End If
   formannex.carregarcomanda rst!comanda
   Set rst = Nothing
End Sub
Sub carregar_pdf(vtreball As Double)
   Dim vversio As Double
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(vtreball) + " order by ordre DESC")
   If Not rst.EOF Then
      vversio = rst!ordre
       Else: Exit Sub
   End If
   etcodidelinia = "Codi de linia: " + format(cadbl(rst!codidelinia), "000") + "#" + atrim(format(cadbl(rst!codideliniav), "0"))
   etcodidelinia.tag = atrim(vversio)
   carregar_PDFdeltreball vtreball, vversio
   
End Sub
Sub carregar_PDFdeltreball(vtreball As Double, vversio As Double)
  Dim rst As Recordset
  Dim vrutapdf As String
  If vtreball = 0 Then AcroPDF1.LoadFile "res":: Exit Sub
  vrutapdf = ruta_documentacio_clixes + "\" + format(vtreball, "00000") + "\PDF" + format(vtreball, "00000") + "-" + format(vversio, "000") + ".pdf"
  AcroPDF1.LoadFile "res"
  AcroPDF1.src = ""
  If existeix(vrutapdf) Then
      ratoli "espera"
      etrutapdf = vrutapdf
      'AcroPDF1.LoadFile vrutapdf
      AcroPDF1.src = vrutapdf
      AcroPDF1.setLayoutMode "SinglePage"
      AcroPDF1.setShowToolbar False
      AcroPDF1.setShowScrollbars False
      AcroPDF1.setView ("Fit")
      ratoli "normal"
  End If
  
End Sub

Private Sub reixa_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i As Byte
    If y < reixa.RowHeight(0) Then
        For i = 0 To reixa.Cols - 1
         If x > reixa.ColPos(i) And x < (reixa.ColPos(i) + reixa.ColWidth(i)) Then
              vordrereixa = i
         End If
        Next i
    End If
End Sub
