VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{8D1418DD-FB6E-4C6F-A1DC-13E914E39989}#1.0#0"; "TBarCode11.ocx"
Begin VB.Form Formprestatgesnous 
   Caption         =   "Prestatges NOUS"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   12960
   Icon            =   "Formprestatgesnous.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Checkeditar 
      Caption         =   "Editar"
      Height          =   195
      Left            =   690
      TabIndex        =   12
      Top             =   885
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir Barres"
      Height          =   645
      Left            =   6105
      Picture         =   "Formprestatgesnous.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   330
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      ItemData        =   "Formprestatgesnous.frx":0B14
      Left            =   3435
      List            =   "Formprestatgesnous.frx":0B63
      TabIndex        =   2
      Top             =   210
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   645
      Left            =   4560
      Picture         =   "Formprestatgesnous.frx":0BB2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   330
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   6750
      Left            =   450
      TabIndex        =   0
      Top             =   1080
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   11906
      _Version        =   393216
      Rows            =   7
      Cols            =   99
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TBarCode11LibCtl.TBarCode11 tbarcode 
      Height          =   360
      Left            =   8805
      TabIndex        =   13
      Top             =   345
      Visible         =   0   'False
      Width           =   1110
      _cx             =   1958
      _cy             =   635
      BackColor       =   15790320
      BackStyle       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   "Adjust Properties"
      TextAlignment   =   0
      BarCode         =   20
      CDMethod        =   1
      CountCheckDigits=   0
      EscapeSequences =   0   'False
      Format          =   ""
      BearerBarWidth  =   -1
      BearerBarType   =   0
      ModuleWidth     =   "339"
      Orientation     =   0
      PrintDataText   =   0   'False
      PrintTextAbove  =   0   'False
      Ratio           =   ""
      RatioHint       =   "1B:2B:3B:4B:1S:2S:3S:4S"
      RatioDefault    =   "1:2:3:4:1:2:3:4"
      TextColor       =   0
      LastError       =   ""
      LastErrorNo     =   0
      MustFit         =   0   'False
      TextDistance    =   0
      NotchHeight     =   -1
      CountModules    =   222
      DrawStatus      =   0
      SuppressErrorMsg=   0   'False
      CountRows       =   1
      EncodingMode    =   0
      OptResolution   =   0   'False
      DisplayText     =   ""
      BarWidthReduction=   0
      BarWidthReductionUnit=   0
      Quality         =   514
      CompositeComponent=   0
      RSS_SegmPerRow  =   -1
      TrimSpaces      =   0
      DefaultSet      =   0
      QuietZoneUnit   =   0
      QuietZoneLeft   =   0
      QuietZoneRight  =   0
      QuietZoneTop    =   0
      QuietZoneBottom =   0
      DefaultColorForQuietZoneLeft=   -1  'True
      DefaultColorForQuietZoneRight=   -1  'True
      DefaultColorForQuietZoneTop=   -1  'True
      DefaultColorForQuietZoneBottom=   -1  'True
      QuietZoneColorLeft=   16777215
      QuietZoneColorRight=   16777215
      QuietZoneColorTop=   16777215
      QuietZoneColorBottom=   16777215
      Compression     =   0
      SizeMode        =   0
      Dpi             =   300
      Decoder         =   1
      DrawMode        =   0
      CodePage        =   1
      CodePageCustom  =   0
      PropertyInternal=   ""
      MaximumTextIndex=   5
      ActiveTextIndex =   0
      TextPositionLeft=   0
      TextPositionTop =   0
      TextBlockWidth  =   0
      TextBlockHeight =   0
      TextClipping    =   -1  'True
      WordWrappingEnabled=   -1  'True
      TextRotation    =   0
      BarShape        =   0
      BarShapeImageFile=   ""
      Options         =   ""
      CBF_Rows        =   -1
      CBF_Columns     =   -1
      CBF_RowHeight   =   -1
      CBF_RowSeparatorHeight=   -1
      CBF_Format      =   0
      DM_Size         =   0
      DM_Rectangular  =   0   'False
      DM_Format       =   0
      DM_EnforceBinary=   0   'False
      DM_AppendIndex  =   -1
      DM_AppendCount  =   -1
      DM_AppendFileID =   -1
      Aztec_Size      =   0
      Aztec_EnforceBinary=   0   'False
      Aztec_ErrorCorrection=   -1
      Aztec_Runes     =   0   'False
      Aztec_Format    =   0
      Aztec_FormatSpecifier=   ""
      Aztec_AppendActive=   0   'False
      Aztec_AppendIndex=   65
      Aztec_AppendTotal=   65
      Aztec_AppendMessageID=   ""
      DotCode_SizeMode=   -1
      DotCode_Size    =   ""
      DotCode_PrintDirection=   0
      DotCode_Format  =   0
      DotCode_FormatSpecifier=   ""
      DotCode_EnforceBinary=   0   'False
      DotCode_Mask    =   -1
      DotCode_AppendActive=   0   'False
      DotCode_AppendIndex=   1
      DotCode_AppendTotal=   1
      HanXin_Size     =   0
      HanXin_EnforceBinary=   0   'False
      HanXin_ECLevel  =   0
      HanXin_Mask     =   -1
      MAXI_Mode       =   4
      MAXI_AppendIndex=   -1
      MAXI_AppendCount=   -1
      MAXI_Undercut   =   -1
      MAXI_Preamble   =   0   'False
      MAXI_PostalCode =   ""
      MAXI_CountryCode=   ""
      MAXI_ServiceClass=   ""
      MAXI_Date       =   "96"
      PDF417_Rows     =   -1
      PDF417_Columns  =   -1
      PDF417_ECLevel  =   -1
      PDF417_EncodationMode=   0
      PDF417_RowHeight=   -1
      PDF417_FileName =   ""
      PDF417_SegmentCount=   -1
      PDF417_TimeStamp=   -1
      PDF417_Sender   =   ""
      PDF417_Addressee=   ""
      PDF417_FileSize =   -1
      PDF417_CheckSum =   -1
      PDF417_RatioRowCol=   ""
      PDF417_SegmentIndex=   -1
      PDF417_FileID   =   ""
      PDF417_LastSegment=   0   'False
      MicroPDF_Mode   =   0
      MicroPDF_Version=   0
      QR_Version      =   0
      MQR_Version     =   0
      QR_Format       =   0
      QR_FmtAppIndicator=   ""
      QR_ECLevel      =   1
      QR_Mask         =   -1
      MQR_Mask        =   -1
      QR_AppendIndex  =   -1
      QR_AppendCount  =   -1
      QR_AppendParity =   -1
      QR_KanjiChineseCompaction=   -1
   End
   Begin VB.Label F1 
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   8
      Left            =   45
      TabIndex        =   10
      Top             =   6870
      Width           =   405
   End
   Begin VB.Label F1 
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   7
      Left            =   45
      TabIndex        =   9
      Top             =   5935
      Width           =   405
   End
   Begin VB.Label F1 
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   6
      Left            =   45
      TabIndex        =   8
      Top             =   5003
      Width           =   405
   End
   Begin VB.Label F1 
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   5
      Left            =   45
      TabIndex        =   7
      Top             =   4071
      Width           =   405
   End
   Begin VB.Label F1 
      Caption         =   "F5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   4
      Left            =   45
      TabIndex        =   6
      Top             =   3139
      Width           =   405
   End
   Begin VB.Label F1 
      Caption         =   "F6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   45
      TabIndex        =   5
      Top             =   2207
      Width           =   405
   End
   Begin VB.Label F1 
      Caption         =   "F7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   45
      TabIndex        =   4
      Top             =   1275
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "Estanteria / Magatzem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   270
      TabIndex        =   3
      Top             =   510
      Width           =   3660
   End
   Begin VB.Menu mo 
      Caption         =   "Opcions"
      Begin VB.Menu mllicenciar 
         Caption         =   "Llicenciar TBarCode11"
      End
   End
End
Attribute VB_Name = "Formprestatgesnous"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
   carregar_estanteria Combo1
End Sub
Sub carregar_estanteria(vestanteria As String)
  Dim rst As Recordset
  Dim c As Integer
  Dim f As Integer
  
  configurar_reixa
  Set rst = dbtmp.OpenRecordset("select * from PrestatgesNous where estanteria='" + atrim(vestanteria) + "'")
  While Not rst.EOF
    f = reixa.Rows - rst!fila
    c = rst!columna - 1
    reixa.TextMatrix(f, c) = rst!estanteria + format(rst!columna, "00") + atrim(rst!fila)
    reixa.row = f: reixa.col = c
    reixa.CellBackColor = QBColor(10)
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub

Private Sub comboimpresores_Change()

End Sub

Private Sub Command1_Click()
   Dim r As Integer
   Dim c As Integer
   Dim fila As Byte
   Dim columna As Byte
   Dim vestanteria As String
   Dim vcolor As Double
   
   If MsgBox("Si guardes els canvis es substituiran totes les estanteries per aquestes de la Estanteria/Magatzem " + atrim(Combo1), vbCritical + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Sub
   vestanteria = Combo1
   dbtmp.Execute "delete * from prestatgesnous where estanteria='" + vestanteria + "'"
   For c = 0 To reixa.Cols - 1
    For r = reixa.Rows - 1 To 0 Step -1
      fila = reixa.Rows - r
      columna = c + 1
      If reixa.TextMatrix(r, c) <> "" Then
         reixa.TextMatrix(r, c) = UCase(vestanteria) + format(atrim(columna), "00") + atrim(fila)
         dbtmp.Execute "insert into prestatgesnous (estanteria,columna,fila) values ('" + atrim(vestanteria) + "'," + atrim(columna) + "," + atrim(fila) + ")"
      End If
      If reixa.TextMatrix(r, c) = "" Then
           vcolor = QBColor(15)
            Else: vcolor = QBColor(10)
      End If
      reixa.row = r: reixa.col = c
      reixa.CellBackColor = vcolor
    Next r
   Next c
End Sub


Private Sub Command2_Click()
 Dim llistat As CrystalReport
 Set llistat = Form1.llistat
 If Combo1 = "" Then MsgBox "Primer escull una Estanteria.", vbCritical, "Atenció": Exit Sub
 generarcodisdebarrestemporals Combo1
 wait 2
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "etiquetaestanteria.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = rutadelfitxer(cami) + "palets.mdb"
 llistat.SortFields(0) = "+{tmp_estanteriescodidebarres.estanteria}"
 llistat.SortFields(1) = "+{tmp_estanteriescodidebarres.columna}"
 llistat.SelectionFormula = "{tmp_estanteriescodidebarres.estanteria}='" + Combo1 + "'"
 'llistat.GroupCondition(0) = "{prestatgesnous.columna}"
 llistat.DiscardSavedData = True
 llistat.Formulas(1) = ""
 llistat.Formulas(0) = ""
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 llistat.Formulas(4) = ""
 llistat.Formulas(5) = ""
 llistat.Formulas(6) = ""
 llistat.Formulas(7) = ""
 llistat.Formulas(8) = ""
 llistat.Formulas(9) = ""
 llistat.Formulas(10) = ""
 llistat.Formulas(11) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
' If Form1.mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 'llistat.PrinterDriver = X.DeviceName
' llistat.PrinterName = X.DriverName
' llistat.PrinterPort = X.Port
 MsgBox "Has d'escullir a quina impresora vols imprimir." + Chr(10) + "PENSA QUE HA D'ESSER EN PAPER HORITZONTAL.", vbInformation, "ATENCIÓ"
 llistat.PrinterSelect
 
 llistat.Action = 1
 
 'destrueixo la taula temporal perquè si no pesa molt
 On Error Resume Next
' dbtmp.Execute "drop table tmp_estanteriescodidebarres"
 On Error GoTo 0
End Sub
Sub generarcodisdebarrestemporals(vestanteria As String)
  Dim rst As Recordset
  Dim rst2 As Recordset
  
  
  On Error Resume Next
  dbtmp.Execute "drop table tmp_estanteriescodidebarres"
  On Error GoTo 0
  dbtmp.Execute "create table tmp_estanteriescodidebarres ( estanteria text,columna integer, codidebarres OLEOBJECT)"
  dbtmp.Execute "CREATE INDEX 1 ON tmp_estanteriescodidebarres (estanteria,columna);"
  Set rst2 = dbtmp.OpenRecordset("tmp_estanteriescodidebarres")
  Set rst = dbtmp.OpenRecordset("select distinct estanteria,columna from prestatgesnous where estanteria='" + atrim(vestanteria) + "'")
  
  While Not rst.EOF
    If existeix("c:\temp\codidebarrespalet.bmp") Then Kill "c:\temp\codidebarrespalet.bmp"
    tbarcode.Text = UCase(rst!estanteria) + format(rst!columna, "00")
    tbarcode.SaveImage "c:\temp\codidebarrespalet", eIMBmp, 3000, 500, 600, 600
    rst2.AddNew
    rst2!estanteria = rst!estanteria
    rst2!columna = rst!columna
    'rst2!fila = rst!fila
    Form1.copiafoto "c:\temp\codidebarrespalet.bmp", rst2!codidebarres
    rst2.Update
    rst.MoveNext
  Wend
  
  Set rst2 = Nothing
  Set rst = Nothing
  
End Sub

Private Sub Command3_Click()
   
End Sub

Private Sub Form_Load()
 configurar_reixa
 tbarcode.Enabled = True
 
End Sub
Sub configurar_reixa()
 Dim i As Byte
  reixa.Clear
  For i = 0 To reixa.Rows - 1
    reixa.RowHeight(i) = 900
  Next i
  For i = 0 To reixa.Cols - 1
    reixa.ColWidth(i) = 1200
  Next i
End Sub

Private Sub mllicenciar_Click()
   tbarcode.Licensing
End Sub

Private Sub reixa_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim vr As Byte
  Dim vrs As Byte
  Dim vc As Byte
  Dim vcs As Byte
  Dim v As Byte
  If Combo1 = "" Or Checkeditar.Value <> 1 Then Exit Sub
  vr = reixa.row
  vrs = reixa.RowSel
  vc = reixa.col
  vcs = reixa.ColSel
  If vr > vrs Then v = vr: vr = vrs: vrs = v
  If vc > vcs Then v = vc: vc = vcs: vcs = v
  If MsgBox("Vols assignar forats de bobina a aquesta seleccio?", vbInformation + vbDefaultButton2 + vbYesNo, "Assignar") = vbYes Then
      For i = vr To vrs
          For j = vc To vcs
             If reixa.TextMatrix(i, j) = "" Then
                ' reixa.TextMatrix(i, j) = "?"
                 reixa.TextMatrix(i, j) = UCase(Combo1) + format(atrim(j + 1), "00") + atrim(reixa.Rows - i)
                   Else: reixa.TextMatrix(i, j) = ""
             End If
          Next j
      Next i
  End If
End Sub
