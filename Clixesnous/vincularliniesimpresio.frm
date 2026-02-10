VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form vincularliniesimpresio 
   Caption         =   "Vincular linies d'impresió"
   ClientHeight    =   10575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13800
   Icon            =   "vincularliniesimpresio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   13800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Desassignar linia"
      Height          =   645
      Left            =   4905
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CheckBox Checktots 
      Caption         =   "Tots, assignats i no assignats"
      Height          =   195
      Left            =   2610
      TabIndex        =   10
      Top             =   15
      Width           =   2580
   End
   Begin VB.CheckBox Checklimitdata 
      Caption         =   "Limit dos anys enrrera"
      Height          =   195
      Left            =   300
      TabIndex        =   9
      Top             =   15
      Value           =   1  'Checked
      Width           =   2580
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Nova linia"
      Height          =   615
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4950
      Width           =   945
   End
   Begin VB.CommandButton assignarpdf 
      BackColor       =   &H0080FF80&
      Caption         =   "Assignar el PDF  a la linia seleccionada"
      Enabled         =   0   'False
      Height          =   630
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4935
      Width           =   2430
   End
   Begin MSDBGrid.DBGrid reixaliniatreball 
      Bindings        =   "vincularliniesimpresio.frx":048A
      Height          =   2310
      Left            =   4590
      OleObjectBlob   =   "vincularliniesimpresio.frx":04A6
      TabIndex        =   3
      Top             =   240
      Width           =   1155
   End
   Begin VB.Data dataliniestreball 
      Caption         =   "dataliniestreball"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\Dades\progcomandes\dades\clixesnous.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4590
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT DISTINCT Modificacions.numerodelinia from modificacions"
      Top             =   45
      Visible         =   0   'False
      Width           =   3420
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      DragIcon        =   "vincularliniesimpresio.frx":0CF7
      DragMode        =   1  'Automatic
      Height          =   4755
      Left            =   1530
      TabIndex        =   2
      Top             =   5610
      Width           =   11940
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.Data Datalinies 
      Caption         =   "datalinies"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\Dades\progcomandes\dades\clixesnous.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   330
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"vincularliniesimpresio.frx":1281
      Top             =   2310
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Data datamarques 
      Caption         =   "datamarques"
      Connect         =   "Access"
      DatabaseName    =   "\\serverprodu\Dades\progcomandes\dades\clixesnous.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select marca from clixes"
      Top             =   375
      Visible         =   0   'False
      Width           =   3420
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "vincularliniesimpresio.frx":1347
      Height          =   2445
      Left            =   225
      OleObjectBlob   =   "vincularliniesimpresio.frx":135D
      TabIndex        =   0
      Top             =   210
      Width           =   3990
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "vincularliniesimpresio.frx":1B94
      Height          =   2205
      Left            =   195
      OleObjectBlob   =   "vincularliniesimpresio.frx":1BA9
      TabIndex        =   1
      Top             =   2670
      Width           =   5730
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF2 
      Height          =   5100
      Left            =   6015
      TabIndex        =   4
      Top             =   105
      Width           =   7500
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.Label etrutapdf 
      Height          =   225
      Left            =   1575
      TabIndex        =   8
      Top             =   10365
      Width           =   11865
   End
   Begin VB.Label etlinkar 
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
      Height          =   300
      Left            =   6075
      TabIndex        =   5
      Top             =   5235
      Width           =   6435
   End
End
Attribute VB_Name = "vincularliniesimpresio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CRViewer91_CloseButtonClicked(UseDefault As Boolean)

End Sub

Private Sub assignarpdf_Click()
   assignarliniaaltreball Datalinies.Recordset!id_treball, Datalinies.Recordset!ordre, cadbl(reixaliniatreball.Text)
End Sub

Private Sub Checklimitdata_Click()
   carregarlesmarques
End Sub

Private Sub Checktots_Click()
  carregarlesmarques
End Sub

Private Sub Command1_Click()
   assignarliniaaltreball Datalinies.Recordset!id_treball, Datalinies.Recordset!ordre, 0
     
End Sub
Sub assignarliniaaltreball(vidtreball As Double, vordre As Double, vlinia As Double)
   If vlinia = 0 Then
      If Not dataliniestreball.Recordset.EOF Then
         dataliniestreball.Recordset.MoveLast
         vlinia = cadbl(dataliniestreball.Recordset!numerodelinia) + 1
        Else: vlinia = 1
      End If
   End If
   If vlinia = -1 Then dbclixes.Execute "update modificacions set numerodelinia=null where id_treball=" + atrim(vidtreball) + " and ordre=" + atrim(vordre): GoTo cont
   If MsgBox("Vols relacionar " + UCase(atrim(Datalinies.Recordset!linia)) + Chr(10) + " amb el Nº de linia: " + atrim(vlinia) + " ?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo fi
   dbclixes.Execute "update modificacions set numerodelinia=" + atrim(vlinia) + " where id_treball=" + atrim(vidtreball) + " and ordre=" + atrim(vordre)
cont:
   Datalinies.Refresh
   dataliniestreball.Refresh

fi:
End Sub
Private Sub AcroPDF2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  etlinkar = "Relacionar treball amb la linia Nº: " + atrim(cadbl(reixaliniatreball.Text))
End Sub

Private Sub Command2_Click()
   If MsgBox("Estàs segur que vols desassignar aquesta linia a aquest treball?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
      assignarliniaaltreball Datalinies.Recordset!id_treball, Datalinies.Recordset!ordre, -1
   End If
End Sub

Private Sub dataliniestreball_Reposition()
  If dataliniestreball.Recordset.EOF Then
    AcroPDF2.LoadFile "res"
    AcroPDF2.src = ""
  End If
End Sub

Private Sub datamarques_Reposition()
   Dim vmarca As String
   'If datamarques.Recordset.EOF Then Exit Sub
   vmarca = atrim(datamarques.Recordset!marca)
   If datamarques.Recordset.EOF Then
     dataliniestreball.RecordSource = "select distinct numerodelinia from modificacions where id_treball in (SELECT id_treball FROM Clixes where marca='{[}]')"
     Datalinies.RecordSource = "SELECT Modificacions.id_treball, Modificacions.ordre, Clixes.marca, Clixes.linia, Modificacions.numerodelinia FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball WHERE  modificacions.id_treball in (SELECT clixes.id_treball FROM Clixes where marca='{[}]')"
     GoTo fi
   End If
   dataliniestreball.RecordSource = "select distinct numerodelinia from modificacions where id_treball in (SELECT id_treball FROM Clixes where marca='" + atrim(vmarca) + "') and (numerodelinia<>null and numerodelinia>0)"
   Datalinies.RecordSource = "SELECT Modificacions.id_treball, Modificacions.ordre, Clixes.marca, Clixes.linia, Modificacions.numerodelinia FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball WHERE  modificacions.id_treball in (SELECT clixes.id_treball FROM Clixes where marca='" + atrim(vmarca) + "') " + IIf(datamarques.tag <> "", " and " + datamarques.tag, "")
fi:
   Datalinies.Refresh
   dataliniestreball.Refresh
End Sub

 Function rutapdftreball(id_treball As Double, ordre As Double)
  '  MkDir ruta_documentacio_clixes + "\" + Format(id_treball, "00000")
  '  MkDir ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF"
    rutapdftreball = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordre, "000") + ".pdf"
End Function



Private Sub DBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   Dim vrutapdf As String
   AcroPDF1.LoadFile "res"
   AcroPDF1.src = ""
   etrutapdf = ""
   If Datalinies.Recordset.EOF Then Exit Sub
   vrutapdf = rutapdftreball(Datalinies.Recordset!id_treball, Datalinies.Recordset!ordre)
    
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
Sub carregarlesmarques()
   Dim vlimitdata As String
   Dim vtots As String
   Dim vmarcaproducte As String
   Dim vwhere As String
   ratoli "espera"
   If formclixes.Command18.tag <> "" Then Checklimitdata.Value = 0: DBGrid1.tag = formclixes.Command18.tag
   If DBGrid1.tag <> "" Then vmarcaproducte = " clixes.marca='" + DBGrid1.tag + "' "
   If Checktots.Value = 1 Then
      vtots = ""
       Else: vtots = "(((Modificacions.numerodelinia) Is Null Or (Modificacions.numerodelinia)=0))"
   End If
   If Checklimitdata.Value = 1 Then
      vlimitdata = " clixes.id_treball in(select numtreball from comandes where proximaseccio='T' and year(datacomanda)>" + atrim(Year(Now) - 2) + ")"
      Else: vlimitdata = ""
   End If
   vwhere = vtots + IIf(vtots <> "" And vlimitdata <> "", " and ", "") + vlimitdata + IIf((vtots <> "" Or vlimitdata <> "") And vmarcaproducte <> "", " and", "") + vmarcaproducte
   datamarques.tag = vwhere
  ' vwhere = vtots + IIf(vtots <> "" And vmarcaproducte <> "", " and", "") + vmarcaproducte
   datamarques.RecordSource = "SELECT distinct Clixes.marca FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball " + IIf(vwhere <> "", " where ", "") + vwhere
   datamarques.Refresh
   'Clipboard.Clear
   'Clipboard.SetText datamarques.RecordSource
   If DBGrid1.tag <> "" Then Datalinies.Recordset.FindFirst "id_treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio)
   ratoli "normal"
End Sub
Private Sub Form_Load()
   carregarlesmarques
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'AcroPDF1.Left = X
   'AcroPDF1.Top = Y
End Sub

Sub carregar_imatge_linia()
   Dim rst As Recordset
  Dim vrutapdf As String
  Dim vmarca As String
  AcroPDF2.LoadFile ""
  AcroPDF2.src = ""
  vmarca = atrim(datamarques.Recordset!marca)
  assignarpdf.Enabled = True
  Set rst = dbclixes.OpenRecordset("SELECT clixes.marca,modificacions.* FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball where numerodelinia=" + atrim(cadbl(reixaliniatreball.Text)) + " and clixes.marca='" + vmarca + "'")
  If rst.EOF Then Exit Sub
  vrutapdf = rutapdftreball(rst!id_treball, rst!ordre)
  If existeix(vrutapdf) Then
      ratoli "espera"
      If existeix("c:\temp\copiapdf.pdf") Then Kill "c:\temp\copiapdf.pdf"
      FileCopy vrutapdf, "c:\temp\copiapdf.pdf"
      AcroPDF2.LoadFile "c:\temp\copiapdf.pdf" 'vrutapdf
      AcroPDF2.setPageMode "none"
      AcroPDF2.setLayoutMode "SinglePage"
      
      AcroPDF2.setShowToolbar False
     ' AcroPDF1.setShowScrollbars False
       AcroPDF2.setView ("FitBH")
      ratoli "normal"
      assignarpdf.Enabled = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
formclixes.Command18.tag = "Tancat"
End Sub

Private Sub reixaliniatreball_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   carregar_imatge_linia
End Sub
