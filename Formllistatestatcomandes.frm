VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Formllistatestatcomandes 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Llistat d'estat de les comandes de Crop's"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15270
   Icon            =   "Formllistatestatcomandes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox cultimaref 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Només ultima Referència"
      Height          =   240
      Left            =   630
      TabIndex        =   12
      Top             =   75
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.CommandButton bposicioordre 
      Height          =   315
      Left            =   1695
      Picture         =   "Formllistatestatcomandes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar totes les linies"
      Top             =   1020
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton bordre 
      Height          =   315
      Left            =   0
      Picture         =   "Formllistatestatcomandes.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ordenar per..."
      Top             =   15
      Width           =   300
   End
   Begin VB.Frame Fbotons 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   13290
      TabIndex        =   4
      Top             =   -15
      Width           =   1980
      Begin VB.CommandButton sortir 
         Height          =   480
         Left            =   1260
         Picture         =   "Formllistatestatcomandes.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Sortir"
         Top             =   120
         Width           =   645
      End
      Begin VB.CommandButton exportaraxls 
         BackColor       =   &H00F0F0F0&
         Height          =   480
         Left            =   660
         Picture         =   "Formllistatestatcomandes.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exportar a Excel la sel.lecció"
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Height          =   480
         Left            =   60
         Picture         =   "Formllistatestatcomandes.frx":1F46
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Pujar a la Web"
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.ComboBox filtre 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Index           =   0
      Left            =   75
      TabIndex        =   2
      Top             =   615
      Width           =   555
   End
   Begin VB.CommandButton treurefiltre 
      Height          =   285
      Left            =   0
      Picture         =   "Formllistatestatcomandes.frx":2578
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Eliminar totes les linies"
      Top             =   330
      Width           =   300
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   6465
      Left            =   15
      TabIndex        =   0
      Top             =   930
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   11404
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorSel    =   16756318
      ForeColorSel    =   16711680
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   3
   End
   Begin VB.Label etcarregantdades 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualitzant les dades...   Un moment siusplau."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   555
      TabIndex        =   13
      Top             =   30
      Visible         =   0   'False
      Width           =   8610
   End
   Begin VB.Label etregistres 
      Height          =   270
      Left            =   75
      TabIndex        =   10
      Top             =   7650
      Width           =   15150
   End
   Begin VB.Label etordre 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   345
      TabIndex        =   9
      Top             =   60
      Width           =   3885
   End
   Begin VB.Label etmsgajuda 
      BackColor       =   &H0000FFFF&
      Height          =   270
      Left            =   300
      TabIndex        =   3
      Top             =   330
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "Formllistatestatcomandes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long

 Dim vnopoblar As Boolean
Dim fitxertmpestats As String
Dim dbplanificacio As Database
Dim whereultimfiltre As String
Dim oshell As Object
Dim objFSO As Object





Sub configreixa(Optional nocarregaramples As Boolean)
  Dim rst As Recordset
  Dim col As Integer
  Dim enes As Byte
  'reixa.LeftCol = 0
  If reixa.Rows > 1 Then reixa.TopRow = 1
  Set rst = dbconsulta.OpenRecordset("select * from consultaestats")
  col = 0
  enes = 0
  reixa.Cols = rst.Fields.Count
  For i = 0 To rst.Fields.Count - 1
       reixa.ColAlignment(col) = 2
       reixa.TextMatrix(0, col) = campsestat(i + 1, 3)
       If Not nocarregaramples Then colocarfiltre col, i + 1
       
       col = col + 1
  Next i
   
     
  reixa.Cols = reixa.Cols - (enes + 1)
  carregar_amples_reixa
  'reixa.Row = 0
  'For i = 0 To reixa.Cols - 1
  '  reixa.col = i
  '  reixa.ColSel = i
  '  reixa.CellBackColor = QBColor(8)
  'Next i
End Sub
Sub carregar_amples_reixa()
 Dim ample As String
 Dim x As Long
 Dim j As Integer
 If iniconfigreixa <> "" Then ' existeix("c:\windows\" + iniconfigreixa) Then
 
  x = reixa.Left + 35
  For j = 0 To reixa.Cols - 1
   ample = llegir_ini("AmplesReixa", UCase(reixa.TextMatrix(0, j)), iniconfigreixa)
   If ample <> "{[}]" Then
    reixa.ColWidth(j) = cadbl(ample)
    If x < reixa.Width Then
     filtre(j).Left = x
     filtre(j).Width = cadbl(ample)
     filtre(j).Visible = True
     filtre(j).ForeColor = &H808080
      Else: If filtre.Count < j - 1 Then filtre(j).Visible = False
    End If
    x = x + cadbl(ample)
   End If
 Next j
End If

End Sub

Function ordredelataula() As String
  If bordre.Tag = "" Then
     ordredelataula = " order by cvdate(datacomanda) desc"
    Else: ordredelataula = " order by " + bordre.Tag
  End If
End Function
Sub poblarlareixa(Optional were As String)
  Dim i As Byte
  Dim fila As Integer
  Dim col As Byte
  Dim rst As Recordset
  Dim apuntxrimprimir As Double
  Dim tenimmaterial As Boolean
  Dim tenimclixes As Boolean
  Dim textetaula As String
  Dim vordre As String
  Dim vultimacodi As String
  ratoli "espera"
  etregistres = ""
  reixa.Visible = False
  reixa.Clear
  reixa.BackColor = QBColor(15)
  configreixa IIf(were <> "", True, False)
  reixa.Rows = 1
  If cultimaref.Value = 1 Then
       vordre = "order by vref2,datacomanda"
        Else: vordre = ordredelataula
  End If
  Set rst = dbconsulta.OpenRecordset("select * from consultaestats " + IIf(were <> "", " where " + were, "") + vordre)
  If rst.EOF Then GoTo fi
  fila = 0
  reixa.Tag = "poblant"
  While Not rst.EOF
   If cultimaref.Value = 1 Then If rst!vref2 = vultimcodi Then GoTo proxim
   fila = fila + 1
   reixa.Rows = fila + 1
   For i = 0 To rst.Fields.Count - 1
     If campsestat(i + 1, 1) <> "" Then
      reixa.TextMatrix(fila, i) = IIf(IsNull(rst.Fields(campsestat(i + 1, 1))), "", rst.Fields(campsestat(i + 1, 1)))
      If reixa.TextMatrix(fila, i) = "0:00:00" Then reixa.TextMatrix(fila, i) = ""
      posarelcolordelcamp fila, i, campsestat(i + 1, 1)
     End If
   Next i
proxim:
   'vultimcodi = rst!vref2
   rst.MoveNext
  Wend
  etregistres.Caption = "Registres: " + atrim(rst.RecordCount) + IIf(cultimaref.Value = 1, " -> Filtrats: " + atrim(reixa.Rows - 1), "")
fi:
  Set rst = Nothing
  reixa.Tag = ""
  reixa.Visible = True
  ratoli "normal"
End Sub
Sub posarelcolordelcamp(fila As Integer, columna As Byte, vcamp As String)
  reixa.col = columna
  reixa.row = fila
  If vcamp = "observacions" Then
     reixa.CellBackColor = &HEAD9CE
      Else: reixa.CellBackColor = &H80000005
  End If
End Sub
Sub colocarfiltre(col As Integer, i As Long)
  If filtre.Count <= col Then Load filtre(col)
  filtre(col).Text = campsestat(i, 3)
  filtre(col).Tag = i
'  Load filtre(col + 1)
End Sub


Private Sub bordre_Click()
 etmsgajuda = "Prem sobre la columna que vols ordenar."
 etmsgajuda.Width = 3000
 etmsgajuda.Left = treurefiltre.Left + treurefiltre.Width + 100
 etmsgajuda.Visible = True
 bordre.BackColor = &HFFFF&
 reixa.BackColorFixed = &HFFFF&
End Sub

Private Sub Command3_Click()
   ratoli "espera"
   generar_xls True
   wait 2
   pujarfitxer "c:\temp\dadescrops.csv"
   wait 5
   ratoli "normal"
   obrir_document "www.inplacsa.com/onlinedata/importardadescsv.php"
   wait 8
   obrir_document "www.inplacsa.com/onlinedata/phpgrid/codi/basic_phpgrid.php"
End Sub
Sub pujarfitxer(vnomfitxer As String)
   Set oshell = CreateObject("Shell.Application")
   Set objFSO = CreateObject("Scripting.FileSystemObject")

   FTPUpload (vnomfitxer)
End Sub
Sub FTPUpload(path As String)

On Error Resume Next
Const FONTS = &H14&


Const FOF_SILENT = &H4&
Const FOF_RENAMEONCOLLISION = &H8&
Const FOF_NOCONFIRMATION = &H10&
Const FOF_ALLOWUNDO = &H40&
Const FOF_FILESONLY = &H80&
Const FOF_SIMPLEPROGRESS = &H100&
Const FOF_NOCONFIRMMKDIR = &H200&
Const FOF_NOERRORUI = &H400&
Const FOF_NOCOPYSECURITYATTRIBS = &H800&
Const FOF_NORECURSION = &H1000&
Const FOF_NO_CONNECTED_ELEMENTS = &H2000&


cFlags = FOF_SILENT + FOF_NOCONFIRMATION + FOF_NOERRORUI

'FTP Wait Time in ms
waitTime = 80000

FTPUser = "inplacsa"
FTPPass = "9dV1aO33"
FTPHost = "www.inplacsa.com"
FTPDir = "/www/onlinedata"

strFTP = "ftp://" & FTPUser & ":" & FTPPass & "@" & FTPHost & FTPDir
Set objFTP = oshell.NameSpace(strFTP)


'Upload single file
If objFSO.FileExists(path) Then
        Set objFile = objFSO.getFile(path)
        strParent = objFile.ParentFolder
        Set objFolder = oshell.NameSpace(strParent)
        Set objItem = objFolder.ParseName(objFile.Name)
        Wscript.Echo "Uploading file " & objItem.Name & " to " & strFTP
        objFSO.DeleteFile FTPDir + objItem, True
        objFTP.CopyHere objItem, cFlags
End If


'Upload all files in folder
If objFSO.FolderExists(path) Then

'Entire folder
Set objFolder = oshell.NameSpace(path)

Wscript.Echo "Uploading folder " & path & " to " & strFTP
objFTP.CopyHere objFolder.Items, copyType

End If


If err.Number <> 0 Then
Wscript.Echo "Error: " & err.Description
End If

'Wait for upload
Wscript.Sleep waitTime

End Sub
Sub imprimirseleccio(vexportar As Boolean)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.report
  Set oapp = New CRAXDDRT.Application
  
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistatconsultaestatclixes" + IIf(vexportar, "exportació", "") + ".rpt", 1)
  oreport.Database.Tables.Item(1).SetDataSource dbconsulta.OpenRecordset("select * from consultaestats " + IIf(whereultimfiltre <> "", " where " + whereultimfiltre, "") + ordredelataula), 3
  '"c:\temp\consultaestatstmp.mdb"
  oreport.DiscardSavedData
  If vexportar Then
   oreport.ExportOptions.DiskFileName = "c:\temp\consultaestattreballs.xls"
   oreport.ExportOptions.PDFExportAllPages = True
   oreport.ExportOptions.FormatType = crEFTExcel97 ' crEFTExcel80Tabular
   oreport.ExportOptions.DestinationType = crEDTDiskFile
   oreport.Export False
   obrir_document "c:\temp\consultaestattreballs.xls"
   GoTo fi
  End If
  oreport.PageEngine.ValueFormatOptions = crIncludeFieldValues
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
    Else
      oreport.PrintOut False, 1
  End If
fi:

End Sub

Private Sub cultimaref_Click()
   borrarelfiltre
   poblarlareixa
End Sub

Private Sub exportaraxls_Click()
    'imprimirseleccio True
    generar_xls
    wait 2
    obrir_document "c:\temp\dadescrops.csv"
End Sub
Function seleccionats(iniciofi As String) As Double
    If reixa.row > reixa.RowSel Then
      If iniciofi = "inici" Then seleccionats = reixa.RowSel
      If iniciofi = "fi" Then seleccionats = reixa.row
        Else
          If iniciofi = "inici" Then seleccionats = reixa.row
          If iniciofi = "fi" Then seleccionats = reixa.RowSel
    End If
End Function
Function esalareixa(vnumc As Double) As Boolean
   Dim vtrobat As Boolean
   Dim vcont As Double
   While Not vtrobat And vcont < reixa.Rows
      If cadbl(reixa.TextMatrix(vcont, 0)) = vnumc Then vtrobat = True
      vcont = vcont + 1
   Wend
   esalareixa = vtrobat
End Function
Sub generar_xls(Optional vpelnuvol As Boolean)
   Dim i As Byte
   Dim rst As Recordset
   Dim linia As String
   
   Set rst = dbconsulta.OpenRecordset("select * from consultaestats " + IIf(whereultimfiltre <> "", " where " + whereultimfiltre, "") + ordredelataula)
   If rst.EOF Then MsgBox "No hi ha dades per exportar", vbCritical, "Error": Exit Sub
   Open "c:\temp\dadescrops.csv" For Output As #1
   'posso els titols del csv
   If Not rst.EOF Then
    For i = 1 To rst.Fields.Count - 1
      If vpelnuvol And rst.Fields(i).Name = "estat" Then GoTo titolproximcamp
      If atrim(campsestat(i, 3)) <> "Confirmed_Deliv_date" Then
           linia = linia + IIf(linia = "", "", ";") + atrim(campsestat(i, 3))
            Else: linia = linia + ";Confirmed_Deliv_date1;Confirmed_Deliv_date2"
      End If
titolproximcamp:
    Next i
    Print #1, linia
   End If
   '---------------
   While Not rst.EOF
    linia = ""
    If (seleccionats("fi") - seleccionats("inici")) > 0 Then
       If Not tocaexportar(rst.Fields("Treball"), seleccionats("inici"), seleccionats("fi")) Then GoTo proxim
    End If
    If Not esalareixa(cadbl(rst!comanda)) Then GoTo proxim
    For i = 1 To rst.Fields.Count - 1
      If vpelnuvol And rst.Fields(i).Name = "estat" Then GoTo proximcamp
      vvalor = treure_apostruf(atrim(rst.Fields(i)))
      If IsDate(vvalor) Then vvalor = Format(vvalor, "yyyy/mm/dd")
      linia = linia + IIf(linia = "", "", ";") + """" + IIf(rst.Fields(i).Name = "refclient", " ", "") + vvalor + """"
      If rst.Fields(i).Name = "dataentregareal" Then
           linia = linia + ";""" + sumardieslaborables(atrim(rst.Fields(i)), 1) + """"
      End If
proximcamp:
    Next i
    Print #1, linia
proxim:
    rst.MoveNext
   Wend
   Close #1
      
End Sub
Function tocaexportar(treball As String, inici As Double, fi As Double) As Boolean
   Dim i As Long
   tocaexportar = False
   For i = inici To fi
      If reixa.TextMatrix(i, 0) = treball Then tocaexportar = True: GoTo fi
   Next i
fi:
End Function
Private Sub filtre_DropDown(Index As Integer)
    carregar_combo_filtre Index
   
End Sub

Private Sub filtre_GotFocus(Index As Integer)
bxrcontrolagafafocus Index
  ultimfiltre = Index
  If filtre(Index).Width < 500 Then filtre(Index).HelpContextID = filtre(Index).Width: filtre(Index).Width = 1000
End Sub
Sub bxrcontrolagafafocus(i As Integer)
  Dim cntrl As Control
  Set cntrl = Screen.ActiveControl
  If cntrl.Text <> "" Then
     If cntrl.Text = campsestat(cadbl(filtre(i).Tag), 3) Then cntrl.Text = ""
     cntrl.ForeColor = QBColor(0)
     
   Else:
       
       cntrl.Text = campsestat(cadbl(filtre(i).Tag), 3)
       cntrl.ForeColor = &H808080
  End If
End Sub


Private Sub filtre_LostFocus(Index As Integer)
  crear_i_aplicar_noufiltre Index
End Sub
Sub crear_i_aplicar_noufiltre(Index As Integer)
Dim noufiltre As String
  If Index = 998 Then whereultimfiltre = "": Exit Sub
  noufiltre = crearfiltre
  If filtre(ultimfiltre).Text = "" Then
    filtre(ultimfiltre).Text = campsestat(cadbl(filtre(ultimfiltre).Tag), 3)
    filtre(ultimfiltre).ForeColor = &H808080
    If filtre(ultimfiltre).HelpContextID <> 0 Then filtre(ultimfiltre).Width = filtre(ultimfiltre).HelpContextID
  End If
  If noufiltre <> whereultimfiltre Or Index = 999 Then
     If noufiltre <> "" Then poblarlareixa noufiltre
  End If
  If Index = 999 And noufiltre = "" Then
     poblarlareixa
  End If
  ratoli "normal"
  reixa.Visible = True
  whereultimfiltre = noufiltre
  'Me.caption = whereultimfiltre
  possaretiquetaajuda
  'Command3.tag = noufiltre ' el guardo pel llistat
  
End Sub
Sub possaretiquetaajuda()
   Dim i As Byte
   etmsgajuda.Visible = False
   For i = 0 To filtre.Count - 1
    If InStr(1, filtre(i), ",") > 0 Then
      etmsgajuda.Caption = "Una coma busca dos valors"
      etmsgajuda.Width = filtre(i).Width
      If etmsgajuda.Width < 2000 Then etmsgajuda.Width = 2000
      etmsgajuda.Left = filtre(i).Left
      etmsgajuda.Visible = True
      GoTo fi
    End If
   Next i
fi:
End Sub
Function possarweres(ByVal camp As String, condicio As String, ByVal filtre As String) As String
  Dim re As String
'camps(j, 1) + " LIKE '*" + treure_apostruf(filtre(i)) + "*'"
  filtre = filtre + ","
  If camp = "nomclient" And cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1)) > 0 Then camp = "codiclient"
  While InStr(1, filtre, ",") > 0 And filtre <> ""
    If camp <> "codiclient" Then
       
       If camp = "estatclixe" Then
         re = IIf(re <> "", re + " or ", "") + camp + " = '" + Mid(filtre, 1, InStr(1, filtre, ",") - 1) + "'"
           Else: re = IIf(re <> "", re + " or ", "") + camp + " like '*" + Mid(filtre, 1, InStr(1, filtre, ",") - 1) + "*'"
       End If
      Else: re = IIf(re <> "", re + " or ", "") + camp + " =" + atrim(cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1))) + ""
    End If
    filtre = Mid(filtre, InStr(1, filtre, ",") + 1)
  Wend
  If re <> "" Then re = "(" + re + ")"
  possarweres = re
End Function
Function crearwere(i As Integer) As String
   Dim w As String
   Dim j As Integer
   If filtre(i) = "" Then Exit Function
   j = cadbl(filtre(i).Tag)
   If campsestat(j, 2) = "date" Then
      If IsDate(filtre(i)) Then
         crearwere = campsestat(j, 1) + "=#" + Format(filtre(i), "mm/dd/yy") + "# "
      End If
      Exit Function
   End If
   If InStr(1, campsestat(j, 2), "string") > 0 Then
       crearwere = possarweres(campsestat(j, 1), "LIKE", treure_apostruf(filtre(i)))
       Exit Function
   End If
   crearwere = campsestat(j, 1) + "=" + passaradecimalpunt(atrim(cadbl(filtre(i))))
   

End Function
Function crearfiltre() As String
  Dim i As Integer
  Dim were As String
  Dim w As String
  For i = 0 To filtre.Count - 1
    If filtre(i).Text <> campsestat(cadbl(filtre(i).Tag), 3) Then  ' And campsestat(cadbl(filtre(i).Tag), 1) <> "comanda"
      w = crearwere(i)
      If were = "" Then
         were = w
        Else: If w <> "" Then were = were + " and " + w
      End If
    End If
  Next i
  crearfiltre = were
End Function


Sub carregar_combo_filtre(Index As Integer)
   Dim rst As Recordset
   Set rst = dbconsulta.OpenRecordset("select distinct " + atrim(campsestat(Index + 1, 1)) + " as valor from consultaestats " + IIf(whereultimfiltre <> "", " where " + whereultimfiltre, "") + " order by " + atrim(campsestat(Index + 1, 1)) + " asc")
   filtre(Index).Clear
   While Not rst.EOF
      If atrim(rst!valor) <> "" Then filtre(Index).AddItem rst!valor
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub Form_Activate()
  If etcarregantdades.Visible Then Exit Sub
   If reixa.Rows = 2 Then
    etcarregantdades.Visible = True
    SetActiveWindow Formllistatestatcomandes.hwnd
    DoEvents
    SetActiveWindow Formllistatestatcomandes.hwnd
    carregardadesfitxertemporal
    SetActiveWindow Formllistatestatcomandes.hwnd
    If Not vnopoblar Then poblarlareixa
    SetActiveWindow Formllistatestatcomandes.hwnd
   End If
   SetActiveWindow Formllistatestatcomandes.hwnd
   etcarregantdades.Visible = False
   'AppActivate Formllistatestatcomandes.Caption
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 And Screen.ActiveControl.Name = "filtre" Then crear_i_aplicar_noufiltre 999
End Sub

Private Sub Form_Load()
  
   If r = "nopoblar" Then vnopoblar = True: r = ""
   iniconfigreixa = "c:\windows\consultesestatcomandescrops.ini"
   fitxertmpestats = "c:\temp\consultaestatcomandescrops_tmp.mdb"
   carregartamanyform
   crearfitxertemp
   
End Sub
Sub borrarlataula()
   dbconsulta.Execute "delete * from consultaestats"
End Sub
Sub carregardadesfitxertemporal()
   Dim rst As Recordset
   Dim rstnou As Recordset
    Dim vcodiclient As Double
   vcodiclient = 6841 'crop's
   borrarlataula
   ratoli "espera"
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
   Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb", , True)
   Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificacio.mdb", , True)
   Set rst = dbtmp.OpenRecordset("select comanda from comandes where (producte<>'PC' and producte <>'PC2' and producte <>'PCP' and producte<>'PCI3') and proximaseccio<>'T' and client=" + atrim(vcodiclient))
   Set rstnou = dbconsulta.OpenRecordset("select * from consultaestats")
   While Not rst.EOF
     copiarregistreatemporal rst, rstnou
    rst.MoveNext
   Wend
  ' wait 2
  ratoli "normal"
   Set rstc2 = Nothing
   Set rst = Nothing
   Set rstnou = Nothing
   Set dbclixes = Nothing
   Set dbplanificacio = Nothing
End Sub
Function canviarlacoma(ByVal n As String) As String
   While InStr(n, ",")
     n = Mid(n, 1, InStr(1, n, ",") - 1) + "¸" + Mid(n, InStr(1, n, ",") + 1)
   Wend
   If n = "{[}]" Then n = ""
   canviarlacoma = n
End Function
Function treurerefisap(vrefcli As String, vcamp As String) As String
   Dim vref1 As String
   Dim vref2 As String
   Dim vref3 As String
   If InStr(1, vrefcli, "/") = 0 Then treurerefisap = vrefcli: Exit Function
   vrefcli = "  " + vrefcli + "  "
   vref1 = Mid(vrefcli, 1, InStr(1, vrefcli, "/") - 1)
   vref2 = Mid(vrefcli, InStr(1, vrefcli, "/") + 1)
   
   If Mid(atrim(vref1), 1, 1) = "0" Then
        vref3 = vref1
        vref1 = vref2
        vref2 = vref3
   End If
   treurerefisap = IIf(vcamp = "ref", atrim(vref2), atrim(vref1))
End Function
Function calculardataimpresio(vdata1 As String, vdata2 As String) As String
   If vdata1 = vdata2 And IsDate(vdata1) Then calculardataimpresio = sumardieslaborables(vdata1, 6) 'DateAdd("d", 3, vdata1): Exit Function
   If vdata1 <> vdata2 And IsDate(vdata2) Then calculardataimpresio = sumardieslaborables(vdata2, 6) ' DateAdd("d", 3, vdata2): Exit Function
End Function
Function sumardieslaborables(vdataE As String, vdies As Long) As String
    Dim i As Byte
    Dim vcont As Byte
    Dim vdata As String
    If Not IsDate(vdataE) Then GoTo fi
    vdata = vdataE
    For i = 1 To vdies
         vcont = 0
         vdata = DateAdd("d", 1, vdata)
         While WeekDay(vdata, vbMonday) > 5 And vcont < 10
            vdata = DateAdd("d", 1, vdata)
            vcont = vcont + 1
         Wend
    Next i
    sumardieslaborables = vdata
fi:
End Function
Sub copiarregistreatemporal(rst As Recordset, rstnou As Recordset)
   Dim rstc As Recordset
   Dim rstclixe As Recordset
   Dim vdataentrega2 As String
   Dim rstplanificacio As Recordset
   Dim vdataentregareal As String
   Set rstc = dbtmp.OpenRecordset("SELECT comandes.*, productes.ruta AS laruta, clients_codisSAP.nomclient as empresaonfacturar, comandes.linkcomanda1,comandes.linkcomanda2, comandes.comanda FROM ((comandes INNER JOIN productes ON comandes.producte = productes.codi) INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda) LEFT JOIN clients_codisSAP ON comandes_extres.codicomptable = clients_codisSAP.codiSAP WHERE (((comandes.comanda)=" + atrim(rst!comanda) + "));")
   If rstc.EOF Then MsgBox "La comanda " + atrim(rst!comanda) + " l'hi falta alguna dada de client o te un error de producte.", vbCritical, "Error": Exit Sub
   Set rstclixe = dbclixes.OpenRecordset("SELECT Modificacions.id_treball, Modificacions.ordre, Clixes.marca, Clixes.linia, Modificacions.desarroll, Modificacions.tinters, Fotogravadors.nomfotogravador FROM (Clixes LEFT JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) LEFT JOIN Fotogravadors ON Modificacions.fotograbador = Fotogravadors.codi where modificacions.id_treball = " + atrim(cadbl(rstc!numtreball)) + " And modificacions.ordre = " + atrim(cadbl(rstc!numordremodificacio)))
   If rstclixe.EOF Then MsgBox "La comanda " + atrim(rst!comanda) + " l'hi falta alguna dada de clixes.", vbCritical, "Error": Exit Sub
   Set rstplanificacio = dbplanificacio.OpenRecordset("select data2 from planificaciototes where comanda=" + atrim(cadbl(rst!comanda)))
   If Not rstplanificacio.EOF Then
        vdataentrega2 = atrim(rstplanificacio!Data2)
      Else: vdataentrega2 = ""
   End If
   rstnou.AddNew
   rstnou!comanda = cadbl(rst!comanda)
   rstnou!estat = atrim(rstc!proximaseccio)
   rstnou!fulla = IIf(cadbl(rstc!refilate) > 0, cadbl(rstc!refilate), Null)
   rstnou!marca = atrim(rstclixe!marca)
   rstnou!linia = atrim(rstclixe!linia)
   rstnou!refclient = treurerefisap(atrim(rstc!refclient), "ref")
   rstnou!refsap = treurerefisap(atrim(rstc!refclient), "sap")
   rstnou!empresaonfacturar = atrim(rstc!empresaonfacturar)
   rstnou!datacomanda = rstc!datacomanda
   rstnou!quantitatdemanada = cadbl(rstc!tubbaseext)
   rstnou!contract = atrim(rstc!comandaclient)
   rstnou!dataquevolelclient = rstc!datamaterial
   vdataentregareal = calculardataimpresio(atrim(rstc!dataentrega), atrim(vdataentrega2))
   rstnou!dataentregareal = IIf(IsDate(vdataentregareal), vdataentregareal, Null)
   If IsDate(rstnou!dataentregareal) Then rstnou!dataimpresio = "Week " + atrim(DatePart("ww", rstnou!dataentregareal, vbMonday, vbFirstFourDays) - 1)
   possarpecesproduidesentregadesireals rstnou, rstclixe!desarroll
   possardadesdelcalloff rstnou, rstclixe!desarroll
   possarespesorimaterial rstnou, rstc![comandes.comanda], rstc![comandes.linkcomanda1], rstc![comandes.linkcomanda2]
   rstnou!fotogravador = atrim(rstclixe!nomfotogravador)
   rstnou!observacions = buscarobservacio(rstnou!refclient)
   
cont:
   rstnou.Update
   Set rstc = Nothing
   Set rstclixe = Nothing
   Set rstplanificacio = Nothing
End Sub
Function buscarobservacio(vitem As String) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from Observacionsestatcomanda where item='" + atrim(vitem) + "'")
  If Not rst.EOF Then buscarobservacio = atrim(rst!observacions)
  Set rst = Nothing
End Function
Sub possarpecesproduidesentregadesireals(rstnou As Recordset, vdesarroll As Double)
   Dim rstent As Recordset
   Dim vdiffabricadesiteoriques As Double
   vdesarroll = vdesarroll / 1000
   Set rstent = dbbaixes.OpenRecordset("select sum(metresisacs) as tmetres ,first(seccio) as tseccio from bobinesent where entregat='S' and comanda=" + atrim(rstnou!comanda) + " group by comanda")
   If Not rstent.EOF Then rstnou!pecesentregades = Redondejar(cadbl(rstent!tmetres) / IIf(atrim(rstent!tseccio) = "R", vdesarroll, 1), 0)
   Set rstent = dbbaixes.OpenRecordset("select sum(metresisacs) as tmetres ,first(seccio) as tseccio from bobinesent where comanda=" + atrim(rstnou!comanda) + " group by comanda")
   If Not rstent.EOF Then rstnou!pecesproduides = Redondejar(cadbl(rstent!tmetres) / IIf(atrim(rstent!tseccio) = "R", vdesarroll, 1), 0)
   rstnou!estocreal = cadbl(rstnou!pecesproduides) - cadbl(rstnou!pecesentregades)
   Set rstent = dbbaixes.OpenRecordset("select sum(metresisacs) as tmetres ,first(seccio) as tseccio from bobinesent where (entregat='N' or entregat='' or entregat=null)and comanda=" + atrim(rstnou!comanda) + " group by numpalet")
   If Not rstent.EOF Then rstnou!descripciopalets = generardescripciopalets(rstent, vdesarroll)
   'posso peces en proces
   If InStr(1, "EITP", rstnou!estat) = 0 Then rstnou!pecesproces = cadbl(rstnou!quantitatdemanada)
   If rstnou!estat = "I" And estaamuntadora(rstnou!comanda) Then rstnou!pecesproces = cadbl(rstnou!quantitatdemanada)
   If InStr(1, "TP", rstnou!estat) > 0 Then
      vdiffabricadesiteoriques = cadbl(rstnou!quantitatdemanada) - cadbl(rstnou!pecesproduides)
      If Not (cadbl(rstnou!quantitatdemanada) < (cadbl(rstnou!pecesproduides) + (cadbl(rstnou!pecesproduides) * 0.1)) And cadbl(rstnou!quantitatdemanada) > (cadbl(rstnou!pecesproduides) - (cadbl(rstnou!pecesproduides) * 0.1))) Then
        rstnou!pecesproces = vdiffabricadesiteoriques
      End If
   End If
   
   '--------------
End Sub
Function estaamuntadora(numc As Double) As Boolean
   Dim rst As Recordset
   estamuntada = False
   Set rst = dbbaixes.OpenRecordset("SELECT muntadoratot.comanda  FROM comandes INNER JOIN muntadoratot ON comandes.comanda = muntadoratot.comanda WHERE (((muntadoratot.acabada)=True) AND ((comandes.proximaseccio)='I') and muntadoratot.comanda=" + atrim(cadbl(numc)) + ");")
   If Not rst.EOF Then estamuntada = True
   Set rst = dbbaixes.OpenRecordset("select comanda from muntadora_ordremuntatge where comanda=" + atrim(numc))
   If Not rst.EOF Then estamuntada = True
End Function

Function generardescripciopalets(rstent As Recordset, vdesarroll As Double)
   Dim vpeces As Double
   While Not rstent.EOF
      vpeces = Redondejar(cadbl(rstent!tmetres) / IIf(atrim(rstent!tseccio) = "R", vdesarroll, 1), 0)
      If vpeces > 0 Then
        generardescripciopalets = IIf(generardescripciopalets <> "", generardescripciopalets + " + ", "") + " 1 x " + atrim(vpeces) + " Pcs  "
      End If
      rstent.MoveNext
   Wend
      
End Function

Sub possardadesdelcalloff(rstnou As Recordset, vdesarroll As Double)
   Dim rstent As Recordset
   Dim rst As Recordset
   Dim rstcalloff As Recordset
   Dim vnumc As Double
   Dim vnumcalloff As String
   Dim vgeneraloconcret As Boolean
   If vdesarroll = 0 Then Exit Sub
   vnumcalloff = ""
   vnumc = rstnou!comanda
   Set rstent = dbtmp.OpenRecordset("select numcalloff as fnumcalloff,entregat as fentregat from bobinesent where (numcalloff<>'' and numcalloff<>null) and comanda=" + atrim(vnumc) + " order by entregat")
   If rstent.EOF Then
      Set rst = dbtmp.OpenRecordset("select * from calloffs_detall where comanda=" + atrim(vnumc))
      If Not rst.EOF Then
        If atrim(rstent!fentregat) <> "S" Then vnumcalloff = atrim(rst!numcalloff)
      End If
      vgeneraloconcret = True
       Else
            If atrim(rstent!fentregat) <> "S" Then vnumcalloff = atrim(rstent!fnumcalloff)
           vgeneraloconcret = False
   End If
   If vnumcalloff <> "" Then
            Set rstcalloff = dbtmp.OpenRecordset("select * from calloffs where item='" + atrim(rstnou!refsap) + "' and numcalloff='" + atrim(vnumcalloff) + "'")
            If Not rstcalloff.EOF Then
                rstnou!datacalloff = atrim(rstcalloff!Data)
                If vgeneraloconcret Then
                  rstnou!quantitatcalloff = atrim(rstcalloff!demanats) 'si es general
                   Else: rstnou!quantitatcalloff = Redondejar(totalassignatdelcalloff(vnumcalloff, vnumc) / (vdesarroll / 1000), 0)
                End If
                rstnou!numerocalloff = atrim(rstcalloff!numcalloff)
            End If
   End If
   Set rst = Nothing
   Set rstent = Nothing
   Set rstcalloff = Nothing
End Sub
Function totalassignatdelcalloff(vnumcalloff As String, vnumc As Double) As Double
   Dim rstent As Recordset
   Set rstent = dbtmp.OpenRecordset("select sum(metresisacs) as tmetres from bobinesent where (entregat='N' or entregat=null or entregat='') and numcalloff='" + atrim(vnumcalloff) + "' and comanda=" + atrim(vnumc) + " group by comanda")
   If Not rstent.EOF Then
          totalassignatdelcalloff = rstent!tmetres
   End If
   Set rstent = Nothing
End Function
Function traduir(valor As String, Idioma As String) As String
   Dim rst As Recordset
   traduir = atrim(valor)
   Set rst = dbclixes.OpenRecordset("select * from diccionari where idioma='" + atrim(Idioma) + "' and trim(pertraduir)='" + atrim(valor) + "'")
   If Not rst.EOF Then
      traduir = atrim(rst!traduit)
   End If
End Function
Function descripciomaterialconcatenat(rstmat As Recordset) As String
   Dim c As String
   Dim vnomcolor As String
   c = Mid(atrim(rstmat![familiesmaterials.descripcio]) + " ", 1, InStr(1, atrim(rstmat![familiesmaterials.descripcio]) + " ", " "))
   vnomcolor = atrim(Mid(atrim(rstmat![familiescolorants.descripcio]) + " ", 1, InStr(1, atrim(rstmat![familiescolorants.descripcio]) + " ", " ")))
   If vnomcolor = "TRANSPARENT" Then vnomcolor = ""
   vnomcolor = traduir(vnomcolor, "EN")
   c = c + " " + vnomcolor
   descripciomaterialconcatenat = c
End Function
Sub possarespesorimaterial(rstnou As Recordset, numc1 As Double, numc2 As Double, numc3 As Double)
    Dim rstmat1 As Recordset
  Dim rstmat2 As Recordset
  Dim rstmat3 As Recordset
  Dim espesormat1 As Double
  Dim espesormat2 As Double
  Dim espesormat3 As Double
  Dim descripciomat1 As String
  Dim descripciomat2 As String
  Dim descripciomat3 As String
  Dim tipusfilm As String
  Dim codimat As String
  Dim rstcomandes As Recordset
  Set rstcomandes = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc1) + " or comanda=" + atrim(numc2) + " or comanda=" + atrim(numc3))
  If rstcomandes.EOF Then Exit Sub
  rstcomandes.FindFirst "comanda=" + atrim(numc1)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat1 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));")
  rstcomandes.FindFirst "comanda=" + atrim(numc2)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat2 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));")
  rstcomandes.FindFirst "comanda=" + atrim(numc3)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat3 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));")
  If Not rstmat1.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc1)
     If Not rstcomandes.NoMatch Then
        descripciomat1 = descripciomaterialconcatenat(rstmat1)  'atrim(rstmat1![familiesmaterials.descripcio]), atrim(rstmat1![familiescolorants.descripcio]))rstmat1![subfamiliesmaterials.descripcio]
        espesormat1 = formcomandes.micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))
     End If
  End If
  If Not rstmat2.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc2)
     If Not rstcomandes.NoMatch Then
        descripciomat2 = descripciomaterialconcatenat(rstmat2)
        espesormat2 = formcomandes.micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))
     End If
  End If
  If Not rstmat3.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc3)
     If Not rstcomandes.NoMatch Then
        descripciomat3 = descripciomaterialconcatenat(rstmat3)
        espesormat3 = formcomandes.micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))
     End If
  End If
  rstnou!material = atrim(espesormat1) + " " + descripciomat1 + IIf(cadbl(espesormat2) <> 0, "+" + atrim(espesormat2) + " " + descripciomat2, "") + IIf(cadbl(espesormat3) <> 0, "+" + atrim(espesormat3) + " " + descripciomat3, "")
  Set rstmat1 = Nothing
  Set rstmat2 = Nothing
  Set rstmat3 = Nothing
  Set rstcomandes = Nothing
End Sub

Sub passardadesdeltreball(rstnou As Recordset, numtreball As Double, ordre As Double)
   Dim rstclixes As Recordset
   If numtreball < 1 Then Exit Sub
   If ordre = 0 Then ordre = 1
   Set rstclixes = dbclixes.OpenRecordset("SELECT marca,linia, descripcioquantitatlinia, tinters, desarroll FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball where clixes.id_treball = " + atrim(numtreball) + " And ordre = " + atrim(ordre))
   If rstclixes.EOF Then Exit Sub
   rstnou!texteimpresio = atrim(rstclixes!marca) + " - " + atrim(rstclixes!linia) + " #" + atrim(rstclixes!descripcioquantitatlinia)
   rstnou!tintes = cadbl(rstclixes!tinters)
   rstnou!desarrollimp = cadbl(rstclixes!desarroll)
   Set rstclixes = Nothing
End Sub
Function carregaobservacio(treball As String) As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select observacions from consultaestats where treball='" + atrim(treball) + "'")
   If Not rst.EOF Then
       carregaobservacio = atrim(rst!observacions)
   End If
End Function
Function buscardataentrega(numc As String) As Date
   Dim rst As Recordset
   Set rst = dbplanificacio.OpenRecordset("select data1 from planificaciototes where comanda=" + atrim(cadbl(numc)))
   If Not rst.EOF Then If Not IsNull(rst!Data1) Then buscardataentrega = Format(rst!Data1, "dd/mm/yy")
   Set rst = Nothing
   If buscardataentrega = "0:00:00" Then buscardataentrega = Empty
End Function
Function buscarcomandes(id_treball As Double, ordremodificacio As Double) As String
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select comanda from comandes where proximaseccio='E' and numtreball=" + atrim(id_treball) + " and numordremodificacio=" + atrim(ordremodificacio))
   While Not rst.EOF
      buscarcomandes = buscarcomandes + IIf(buscarcomandes = "", "", ", ") + atrim(rst!comanda)
      rst.MoveNext
   Wend
   If buscarcomandes = "" Then buscarcomandes = "-"
   Set rst = Nothing
End Function
Function nomfotogravador(codi As Long) As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select nomfotogravador from fotogravadors where codi=" + atrim(codi))
   If Not rst.EOF Then nomfotogravador = rst!nomfotogravador
End Function
Sub crearfitxertemp()
     
    If Not existeix(fitxertmpestats) Then
       crearfitxertemporal
    End If
   Set dbconsulta = DBEngine.OpenDatabase(fitxertmpestats)
   carregarllistadecampstemporals
   creartaula
   
End Sub
Sub crearfitxertemporal()
    borrartemps
    If Not existeix(fitxertmpestats) Then
       DBEngine.CreateDatabase fitxertmpestats, dbLangGeneral, DatabaseTypeEnum.dbVersion30
    End If
    
    
End Sub
Sub borrartemps()
   On Error Resume Next
    Kill fitxertmpestats
End Sub
Sub carregarllistadecampstemporals()
  Dim i As Byte
  i = 1
  campsestat(i, 1) = "comanda": campsestat(i, 2) = "double": campsestat(i, 3) = "Lot_Inplacsa": i = i + 1
  campsestat(i, 1) = "fulla": campsestat(i, 2) = "double": campsestat(i, 3) = "Fulla": i = i + 1
  campsestat(i, 1) = "estat": campsestat(i, 2) = "string": campsestat(i, 3) = "Estat": i = i + 1
  campsestat(i, 1) = "marca": campsestat(i, 2) = "string": campsestat(i, 3) = "Line/Label": i = i + 1
  campsestat(i, 1) = "linia": campsestat(i, 2) = "string": campsestat(i, 3) = "Description": i = i + 1
  campsestat(i, 1) = "refclient": campsestat(i, 2) = "string": campsestat(i, 3) = "Item": i = i + 1
  campsestat(i, 1) = "refsap": campsestat(i, 2) = "string": campsestat(i, 3) = "Sap": i = i + 1
  campsestat(i, 1) = "empresaonfacturar": campsestat(i, 2) = "string": campsestat(i, 3) = "Company": i = i + 1
  campsestat(i, 1) = "material": campsestat(i, 2) = "string": campsestat(i, 3) = "Material": i = i + 1
  campsestat(i, 1) = "datacomanda": campsestat(i, 2) = "date": campsestat(i, 3) = "Date_Order": i = i + 1
  campsestat(i, 1) = "quantitatdemanada": campsestat(i, 2) = "double": campsestat(i, 3) = "Quantity_Asked": i = i + 1
  campsestat(i, 1) = "contract": campsestat(i, 2) = "string": campsestat(i, 3) = "Contract": i = i + 1
  campsestat(i, 1) = "dataquevolelclient": campsestat(i, 2) = "date": campsestat(i, 3) = "Desired_Deliv_Date": i = i + 1
  campsestat(i, 1) = "dataimpresio": campsestat(i, 2) = "string": campsestat(i, 3) = "Printingdate": i = i + 1
  campsestat(i, 1) = "dataentregareal": campsestat(i, 2) = "date": campsestat(i, 3) = "Confirmed_Deliv_date": i = i + 1
  campsestat(i, 1) = "pecesproces": campsestat(i, 2) = "double": campsestat(i, 3) = "Quantity_In_process": i = i + 1
  campsestat(i, 1) = "pecesproduides": campsestat(i, 2) = "double": campsestat(i, 3) = "Quantity_Produced": i = i + 1
  campsestat(i, 1) = "pecesentregades": campsestat(i, 2) = "double": campsestat(i, 3) = "Quantity_Delivered": i = i + 1
  campsestat(i, 1) = "datacalloff": campsestat(i, 2) = "date": campsestat(i, 3) = "Date_Call-Off": i = i + 1
  campsestat(i, 1) = "quantitatcalloff": campsestat(i, 2) = "double": campsestat(i, 3) = "Quantity_Call-Off": i = i + 1
  campsestat(i, 1) = "numerocalloff": campsestat(i, 2) = "string": campsestat(i, 3) = "Call-Off_Number": i = i + 1
  campsestat(i, 1) = "estocreal": campsestat(i, 2) = "double": campsestat(i, 3) = "Real_Stock": i = i + 1
  campsestat(i, 1) = "descripciopalets": campsestat(i, 2) = "string": campsestat(i, 3) = "Real_stock_Pal_Splits": i = i + 1
  campsestat(i, 1) = "fotogravador": campsestat(i, 2) = "string": campsestat(i, 3) = "Plates": i = i + 1
  campsestat(i, 1) = "Observacions": campsestat(i, 2) = "string": campsestat(i, 3) = "Remarks": i = i + 1
  campsestat(i, 1) = "": campsestat(i, 2) = "": campsestat(i, 3) = "": i = i + 1
  
End Sub
Sub creartaula()
  Dim i As Integer
 On Error GoTo jaexisteix
  dbconsulta.Execute ("create table consultaestats (id counter)")
  On Error GoTo 0
  dbconsulta.Execute "CREATE INDEX principal ON consultaestats ([id]) witH PRIMARY;"


  For i = 1 To 100
    If campsestat(i, 1) <> "" Then
       dbconsulta.Execute ("alter table consultaestats add column " + campsestat(i, 1) + " " + campsestat(i, 2))
        Else: i = 1000
    End If
  Next i
  
  SetAllowZeroLength dbconsulta
  Exit Sub
jaexisteix:
  dbconsulta.Execute "drop table consultaestats"
  Resume
End Sub

Function SetAllowZeroLength(db As Database)
    Dim i As Integer, j As Integer
    Dim td As TableDef, fld As Field

    
    'The following line prevents the code from stopping if you do not
    'have permissions to modify particular tables, such as system
    'tables.
    On Error Resume Next
    For i = 0 To db.TableDefs.Count - 1
       Set td = db(i)
       For j = 0 To td.Fields.Count - 1
          Set fld = td(j)
          If (fld.Type = 10) And Not _
            fld.AllowZeroLength Then
             fld.AllowZeroLength = True
          End If
       Next j
    Next i
    
End Function

Sub guardar_amples_reixa()
Dim j As Integer
If iniconfigreixa <> "" Then
  For j = 0 To reixa.Cols - 1
   escriure_ini "AmplesReixa", UCase(reixa.TextMatrix(0, j)), atrim(Redondejar(reixa.ColWidth(j), 0)), iniconfigreixa
 Next j
End If
End Sub

Private Sub Form_Resize()
   If Formllistatestatcomandes.Height - reixa.Top - 800 < 1 Then Exit Sub
   reixa.Width = Formllistatestatcomandes.Width - 300
   reixa.Height = Formllistatestatcomandes.Height - reixa.Top - 800
   Fbotons.Left = Formllistatestatcomandes.Width - Fbotons.Width - 300
   etregistres.Top = reixa.Height + reixa.Top
   If Formllistatestatcomandes.Tag <> "canvianttamany" Then
    escriure_ini "TamanyForm", "ample", atrim(Formllistatestatcomandes.Width), iniconfigreixa
    escriure_ini "TamanyForm", "alt", atrim(Formllistatestatcomandes.Height), iniconfigreixa
   End If
End Sub
Sub carregartamanyform()
  If cadbl(llegir_ini("TamanyForm", "ample", iniconfigreixa)) > 0 Then
   Formllistatestatcomandes.Tag = "canvianttamany"
   Formllistatestatcomandes.Width = llegir_ini("TamanyForm", "ample", iniconfigreixa)
   Formllistatestatcomandes.Height = llegir_ini("TamanyForm", "alt", iniconfigreixa)
   Formllistatestatcomandes.Tag = ""
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  guardar_amples_reixa

End Sub

Private Sub MouseWheel1_WheelMove(bDown As Boolean)
  Dim v As Byte
  v = 3
  If reixa.Rows < 2 Then Exit Sub
  If bDown Then
     If reixa.TopRow + v < reixa.Rows Then
        reixa.TopRow = reixa.TopRow + v
       Else: reixa.TopRow = reixa.Rows - 1
     End If
    Else:
        If reixa.TopRow - v > 1 Then
           reixa.TopRow = reixa.TopRow - v
          Else: reixa.TopRow = 1
        End If
  End If
  
End Sub

Private Sub reixa_Click()
  If reixa.BackColorFixed <> treurefiltre.BackColor Then
      vordre = campsestat(reixa.col + 1, 1)
     ' If vordre = "dataobertura" Then vordre = "cvdate(dataobertura)"
      If InStr(1, bordre.Tag, vordre) > 0 Then
          If InStr(1, bordre.Tag, "ASC") > 0 Then
                bordre.Tag = " DESC"
              Else: bordre.Tag = " ASC"
          End If
           Else
              bordre.Tag = " ASC"
      End If
      etordre = campsestat(reixa.col + 1, 3) + " " + bordre.Tag
      bordre.Tag = vordre + bordre.Tag
      etmsgajuda.Visible = False
      bordre.BackColor = treurefiltre.BackColor
      reixa.BackColorFixed = treurefiltre.BackColor
      poblarlareixa whereultimfiltre
  End If
End Sub

Private Sub reixa_DblClick()
   Dim vitem As String
   vitem = atrim(reixa.TextMatrix(reixa.row, 3))
   If reixa.TextMatrix(0, reixa.col) = "Remarks" Then
    obs = InputBox("Entra la observació", "Entrada", reixa.TextMatrix(reixa.row, reixa.col))
    If obs <> "" Then
      dbtmp.Execute "insert into observacionsestatcomanda (item,observacions) values ('" + atrim(vitem) + "',' ') "
      dbtmp.Execute "update observacionsestatcomanda set observacions='" + treure_apostruf(obs) + "' where item='" + atrim(vitem) + "'"
      reixa.Text = atrim(obs)
    End If
  End If
End Sub
Sub guardarobservacio(treball As String, valornou As String)
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from consultaestats where treball='" + atrim(treball) + "'")
   If Not rst.EOF Then
       If valornou = "" Then
          rst.Delete
          GoTo fi
           Else: rst.Edit
       End If
         Else
           rst.AddNew
           rst!treball = treball
   End If
   rst!observacions = valornou
   rst.Update
fi:
   Set rst = Nothing
End Sub

Private Sub reixa_LostFocus()
    guardar_amples_reixa
End Sub
Sub borrarelfiltre()
   configreixa
   poblarlareixa
   filtre_LostFocus 998
End Sub

Private Sub sortir_Click()
  Formllistatestatcomandes.Hide
End Sub

Private Sub treurefiltre_Click()
 borrarelfiltre
End Sub
