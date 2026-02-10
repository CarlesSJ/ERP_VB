VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{8D1418DD-FB6E-4C6F-A1DC-13E914E39989}#1.0#0"; "TBarCode11.ocx"
Begin VB.Form Form1 
   Caption         =   "Imprimir etiquetes Inkmaker"
   ClientHeight    =   1725
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   5745
   Icon            =   "Etiquetes Tintes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.TextBox cnumlotinplacsa 
      Height          =   345
      Left            =   2835
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   810
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir Llauna"
      Height          =   480
      Left            =   4215
      TabIndex        =   7
      Top             =   735
      Width           =   1440
   End
   Begin VB.CommandButton botoretorn 
      Caption         =   "Retorn Llauna"
      Height          =   480
      Left            =   4215
      TabIndex        =   5
      Top             =   120
      Width           =   1440
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1215
      Top             =   450
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3840
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Left            =   3735
      TabIndex        =   2
      Top             =   510
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   570
      Top             =   420
   End
   Begin VB.Label etcanvilot 
      BackStyle       =   0  'Transparent
      Caption         =   "Doble clic per canviar-lo."
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2370
      TabIndex        =   11
      Top             =   615
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NºLot Inplacsa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2070
      TabIndex        =   10
      Top             =   795
      Width           =   765
   End
   Begin VB.Label etcomprovant 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   195
      TabIndex        =   8
      Top             =   1410
      Width           =   5355
   End
   Begin TBarCode11LibCtl.TBarCode11 codidebarres 
      Height          =   825
      Left            =   1785
      TabIndex        =   6
      Top             =   555
      Visible         =   0   'False
      Width           =   1950
      _cx             =   3440
      _cy             =   1455
      BackColor       =   15724527
      BackStyle       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   "x"
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
      PrintDataText   =   -1  'True
      PrintTextAbove  =   0   'False
      Ratio           =   ""
      RatioHint       =   "1B:2B:3B:4B:1S:2S:3S:4S"
      RatioDefault    =   "1:2:3:4:1:2:3:4"
      TextColor       =   0
      LastError       =   "La operación se completó correctamente. "
      LastErrorNo     =   0
      MustFit         =   0   'False
      TextDistance    =   0
      NotchHeight     =   -1
      CountModules    =   46
      DrawStatus      =   0
      SuppressErrorMsg=   0   'False
      CountRows       =   1
      EncodingMode    =   0
      OptResolution   =   0   'False
      DisplayText     =   ""
      BarWidthReduction=   0
      BarWidthReductionUnit=   0
      Quality         =   98
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
      Dpi             =   600
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1410
      TabIndex        =   4
      Top             =   825
      Width           =   705
   End
   Begin VB.Label etpesbascula 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   150
      TabIndex        =   3
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label etrutatemp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   60
      TabIndex        =   1
      Top             =   405
      Width           =   4305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Impresió d'etiquetes de l'inkmaker activat a la ruta:"
      Height          =   210
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   3930
   End
   Begin VB.Menu mopcions 
      Caption         =   "Opcions"
      Begin VB.Menu mtipusbascula 
         Caption         =   "Tipus Bascula (1,2,3,4...)"
      End
      Begin VB.Menu mstringconnexio 
         Caption         =   "Comunicació amb la bascula COM4"
      End
      Begin VB.Menu mtancar 
         Caption         =   "Tancar programa"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vmodelbascula As Byte
Dim VnomfitxerconsumLam As String
Dim esbase As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' -- Api SetForegroundWindow Para traer la ventana al frente
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
' -- Api para desplegar el cuadro de diálogo Acerca de ...
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

' -- Estructura NOTIFYICONDATA
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type


Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Dim systray As NOTIFYICONDATA
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


Private Function EstaCorriendo(ByVal NombreDelProceso As String) As Boolean
    Const MAX_PATH As Long = 260
    Dim con As Byte
    Dim lProcesses() As Long, lModules() As Long, n As Long, lRet As Long, hProcess As Long
    Dim sName As String
    NombreDelProceso = UCase$(NombreDelProceso)
    ReDim lProcesses(1023) As Long
 con = 0
    If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
        For n = 0 To (lRet \ 4) - 1
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(n))
            If hProcess Then
                ReDim lModules(1023)
                If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
                    sName = String$(MAX_PATH, vbNullChar)
                    GetModuleBaseName hProcess, lModules(0), sName, MAX_PATH
                    sName = Left$(sName, InStr(sName, vbNullChar) - 1)
 
                    If Len(sName) = Len(NombreDelProceso) Then
                        If NombreDelProceso = UCase$(sName) Then con = con + 1
                        If con = 2 Then EstaCorriendo = True: Exit Function
                    End If
                End If
            End If
            CloseHandle hProcess
        Next n
    End If
End Function

Private Sub botoretorn_Click()
   ' ensenyar_missatge_imprimint_etiqueta "a1234"
   'ensenyar_missatge_imprimint_etiqueta "Noetiqueta"
    Form1.tag = "1"
    Timer1.Enabled = False
    formretorntintes.Show
    Timer1.Enabled = True
End Sub

Private Sub cnumlotinplacsa_DblClick()
  canviarnumlotinplacsa
End Sub
Sub canviarnumlotinplacsa()
   Dim nlot As String
   nlot = InputBox("Entra el nou numero de Lot d'Inplacsa", "Canvi de lot manual", cnumlotinplacsa)
   If atrim(nlot) = "" Then Exit Sub
   cnumlotinplacsa = atrim(nlot)
   escriure_ini "General", "lotinplacsatintes", cnumlotinplacsa, fitxerini
End Sub

Private Sub cnumlotinplacsa_GotFocus()
    etcanvilot.visible = True
End Sub

Private Sub cnumlotinplacsa_LostFocus()
  etcanvilot.visible = False
End Sub

Private Sub Command1_Click()
   Dim rstmoviment As Recordset
   Dim rstlotsbase As Recordset
   Dim clonlotsbase As Recordset
   Dim clonmoviment As Recordset
   Dim componentanterior As String
   Dim saltsdemovimentestoc As Byte
   Dim ultimadata As Date
   Dim ncontrol As Double
   Set rstmoviment = dbtintes.OpenRecordset("select * from movimentestoc order by data asc")
   Set rstlotsbase = dbtintes.OpenRecordset("select * from lotsbase order by id desc")
   'ncontrol = rstlotsbase!numcontrol
   While Not rstmoviment.EOF
     If Not existeixelcamp(rstmoviment!idcomponent, rstlotsbase) Then GoTo cont
     If rstmoviment!codilot <> atrim(rstlotsbase.Fields(treureespaisibarres(rstmoviment!idcomponent))) And rstmoviment!codilot <> "" Then
          'On Error GoTo 0
          Set clonmoviment = dbtintes.OpenRecordset("select * from movimentestoc where format(data,'ddmmyyhhnn')<" + Format(DateAdd("h", 2, rstmoviment!data), "ddmmyyhhnn") + " and format(data,'ddmmyyhhnn')>" + Format(DateAdd("h", -2, rstmoviment!data), "ddmmyyhhnn") + " order by idcomponent,data")
          Set clonlotsbase = dbtintes.OpenRecordset("select * from lotsbase where id=" + atrim(rstlotsbase!id))
          If Format(rstmoviment!data, "ddmmyy") = "090915" Then Stop
          ultimadata = rstmoviment!data
          rstlotsbase.AddNew
          For i = 1 To clonlotsbase.Fields.Count - 1
             rstlotsbase.Fields(i) = clonlotsbase.Fields(i)
          Next i
          
          
          rstlotsbase!data = rstmoviment!data
          'ncontrol = ncontrol - 1
          'rstlotsbase!numcontrol = ncontrol
          componentanterior = ""
          saltsdemovimentestoc = 0
          While Not clonmoviment.EOF
             If existeixelcamp(treureespaisibarres(clonmoviment!idcomponent), rstlotsbase) And componentanterior <> atrim(clonmoviment!codilot) Then
                rstlotsbase.Fields(treureespaisibarres(clonmoviment!idcomponent)) = IIf(clonmoviment!codilot = "", " ", clonmoviment!codilot)
             End If
             componentanterior = atrim(clonmoviment!codilot)
             clonmoviment.MoveNext
             saltsdemovimentestoc = saltsdemovimentestoc + 1
          Wend
          rstlotsbase.Update
          Set rstlotsbase = dbtintes.OpenRecordset("select * from lotsbase order by id desc")
          'rstlotsbase.Bookmark = rstlotsbase.LastModified
          
          'If ncontrol = 1 Then GoTo fi
     End If
cont:
     If saltsdemovimentestoc = 0 Then saltsdemovimentestoc = 1
     Do
       rstmoviment.MoveNext
       If rstmoviment.EOF Then GoTo fi
     Loop While Not rstmoviment.EOF And Format(DateAdd("h", -2, ultimadata), "ddmmyyhhnn") < Format(rstmoviment!data, "ddmmyyhhnn") And Format(DateAdd("h", 2, ultimadata), "ddmmyyhhnn") > Format(rstmoviment!data, "ddmmyyhhnn")
       
     
   Wend
   
fi:
  Set rstlotsbase = dbtintes.OpenRecordset("select * from lotsbase order by id desc")
  ncontrol = 1484
  While Not rstlotsbase.EOF
    dbtintes.Execute "update lotsbase set numcontrol=" + atrim(ncontrol) + " where id=" + atrim(rstlotsbase!id)
    ncontrol = ncontrol - 1
    rstlotsbase.MoveNext
  Wend

  MsgBox "Ja està"
   
End Sub

Function treureespaisibarres(nomcomponent As String) As String
   Dim n As String
   n = nomcomponent
   While InStr(n, " ")
     n = Mid(n, 1, InStr(1, n, " ") - 1) + "_" + Mid(n, InStr(1, n, " ") + 1)
   Wend
  ' While InStr(n, "/")
  '   n = Mid(n, 1, InStr(1, n, "/") - 1) + "_" + Mid(n, InStr(1, n, "/") + 1)
  ' Wend
   If n = "{[}]" Then n = ""
   treureespaisibarres = n
   
End Function
Function existeixelcamp(nomcomponent As String, rst As Recordset) As Boolean
   Dim i As String
   On Error GoTo fi
   nomcomponent = treureespaisibarres(nomcomponent)
   i = atrim(rst.Fields(nomcomponent))
   existeixelcamp = True
   Exit Function
fi:
   existeixelcamp = False
End Function

Sub retornllaunes()
  Dim numllaunaretorn As String
  Dim desctintes As String
  Dim rstll As Recordset
  Dim rsthistoria As Recordset
  Dim resp As String
  
  
  numllaunaretorn = InputBox("Entra el Nº de llauna que retornes.", "Retorn Llaunes")
  If atrim(numllaunaretorn) = "" Then Exit Sub
  Set rstll = dbtintes.OpenRecordset("SELECT Llaunes.*, tintes.descripcio FROM tintes LEFT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta where numllauna='" + atrim(numllaunaretorn) + "'")
  If rstll.EOF Then MsgBox "Aquesta llauna no existeix.", vbCritical, "Atenció": Exit Sub
  If Not rstll!activa Then MsgBox "Aquesta llauna no està activa no pots retornar-la", vbCritical, "Atenció": Exit Sub
  desctintes = Chr(10) + "Origen: " + UCase(rstll!numllauna) + " --> " + UCase(rstll!descripcio) + Chr(10)
  If MsgBox("Ès correcte la llauna que retornes?" + Chr(10) + desctintes, vbInformation + vbYesNo + vbDefaultButton2, "Retorn") = vbNo Then Exit Sub
    If cadbl(etpesbascula) - tarallauna < 1 Then
       resp = UCase(InputBox("El pes de la tinta es de menys d'1Kg." + Chr(10) + "Escriu [retorn] o [buidar]", "Poc pes de la llauna"))
    End If
    If resp = "RETORN" Then
        Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna")
        rsthistoria.AddNew
        rsthistoria!idnumllauna = rstll!id
        rsthistoria!data = Now
        rsthistoria!tipusmoviment = "R"
        rsthistoria!formula = ""
        rsthistoria!kg = cadbl(etpesbascula)
        rsthistoria.Update
    End If
    If resp = "BUIDAR" Then
        Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna")
        rsthistoria.AddNew
        rsthistoria!idnumllauna = rstll!id
        rsthistoria!data = Now
        rsthistoria!tipusmoviment = "V"
        rsthistoria!formula = ""
        rsthistoria!kg = 0
        rsthistoria.Update
    End If
  calcularkgdisponiblesllauna rstll!numllauna
  Set rsthistoria = Nothing
  Set rstll = Nothing
End Sub

Private Sub Command2_Click()
   Dim numllauna As String
   Dim rst As Recordset
   numllauna = InputBox("Entra la llauna que vols imprimir", "Imprimir Llauna")
   If atrim(numllauna) <> "" Then
        Set rst = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(numllauna) + "'")
        If Not rst.EOF Then
           imprimir_etiqueta numllauna
             Else: MsgBox "No he trobat aquest numero de llauna    " + numllauna, vbCritical, "Error"
        End If
   End If
   Set rst = Nothing
End Sub

Private Sub etpesbascula_DblClick()
   If MSComm1.tag = "error" Then etpesbascula = cadbl(InputBox("Entra el pes que vols utilitzar", "Pes manual"))
End Sub

Private Sub Form_Click()
  'inkmaker_llegir_horaultima_etiqueta
' actualitzar_estocdecomponents
'   Dim rst As Recordset
'   Dim rsthistoria As Recordset
'   Dim rsttemporal As Recordset
'   Set rst = dbtintes.OpenRecordset("select * from llaunes")
 '  While Not rst.EOF
 '    Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna where idnumllauna=" + atrim(rst!id))
 '    If Not rsthistoria.EOF Then
        'Set rsttemporal = dbtintes.OpenRecordset("select * from temporalestoc where numlata='" + atrim(rst!numllauna) + "'")
        'If Not rsttemporal.EOF Then
       '    dbtintes.Execute "insert into historiallaunalots (idhistoria,idcomponent,numlotbase,tanx100tinta,kgtinta) values (" + atrim(rsthistoria!id) + ",0,'" + atrim(rsttemporal!lofab) + "',0,0)"
       ' End If
     'End If
     
    ' rst.MoveNext
  ' Wend
  
  
  
  'Dim rstllauna As Recordset
  'Dim rstlots As Recordset
  'Set rstllauna = dbtintes.OpenRecordset("select * from llaunes where numllauna='A10905'")
  'sqllots = "SELECT dbo.tblLogBook.WorkOrder, dbo.tblLogBookDetail.IDComponente, dbo.tblLogBookDetail.CodComponente, dbo.tblLogBookDetail.BatchCode, dbo.tblLogBookDetail.DispensedQuantity FROM dbo.tblLogBook INNER JOIN dbo.tblLogBookDetail ON dbo.tblLogBook.IDLogBook = dbo.tblLogBookDetail.IDLogBook "
  ' sqllots = sqllots + " WHERE (((dbo.tblLogBook.Barcode)='18052402000'));"
  ' Clipboard.SetText sqllots
  'Set rstlots = conODBC.OpenRecordset(sqllots)
  'MsgBox elslotssondiferents(rstllauna, rstlots)
  
  
End Sub

Sub obrirportseriebascula()
  On Error GoTo errordeport
    If Not MSComm1.PortOpen Then
      MSComm1.CommPort = 4
     ' 9600 baudios, sin paridad, 7 bits de datos y 1 bit de parada.
      vmodelbascula = cadbl(llegir_ini("Bascula", "modelbascula", "comandes.ini"))
      If vmodelbascula = 0 Then
         vmodelbascula = 1
         escriure_ini "Bascula", "modelbascula", "1", "comandes.ini"
      End If
      vstringport = llegir_ini("Bascula", "connexio", "comandes.ini")
      If vstringport = "{[}]" Or vstringport = "" Then
         escriure_ini "Bascula", "connexio", "9600,n,8,1", "comandes.ini"
         vstringport = "9600,n,8,1"
      End If
      MSComm1.Settings = vstringport
     ' If nummaq = 1 Then MSComm1.Settings = "2400,n,8,1"
     ' Indicar al control que lea todo el búfer al usar Input.
      MSComm1.InputLen = 0
     
      MSComm1.RTSEnable = True 'Por si necesitas habilitar el RTS
     
     'Abrir Puertos
     
      MSComm1.PortOpen = True
    End If
    Exit Sub
errordeport:
    Me.caption = "No s'ha pogut connectar amb la bàscula"
    MSComm1.tag = "error"
    
End Sub
Function pesbascula() As Double
Static buffer As String
Static nobascula As Boolean
Dim t As String
 On Error GoTo nopossarpes
 i = 0
 buffer = buffer & MSComm1.Input
 If Len(buffer) > 60 Then buffer = Mid(buffer, Len(buffer) - 60)
 If Len(buffer) > 30 Then
   If InStr(1, buffer, "-") Then buffer = "0"
   If vmodelbascula = 1 Then If InStr(1, buffer, Chr$(13)) > 0 Then buffer = Mid(buffer, InStr(1, buffer, "+") + 1, InStr(1, buffer, Chr$(13)))
   If vmodelbascula = 2 Then If InStr(1, buffer, "ST,GS,") > 0 And InStr(1, buffer, Chr$(13)) > 0 Then buffer = Mid(buffer, InStr(1, buffer, "ST,GS,") + 6, 8)
   'pesbascula = cadbl(substituir(buffer, ".", ","))
   pesbascula = cadbl(substituir(buffer, ",", "."))
   buffer = ""
   'escriure_ini "Tintes", "pesbascula", atrim(pesbascula), fitxerini
 End If
 Exit Function
nopossarpes:
   pesbascula = 0
End Function

Function comprovarformula(rstf As Recordset, rstink As Recordset, conODBC As DAO.Connection) As Boolean
   Dim lagran As Long
   Dim rst As Recordset
   Dim hiharf As Boolean
   comprovarformula = False
   lagran = rstf!idformula
   If rstink.EOF Then Exit Function
   If rstf!descripcioformula <> rstink!Description Then
        Exit Function
   End If
   If rstf!series <> rstink!series Then
        Exit Function
   End If
   If atrim(rstf!datacreacio) <> Format(rstink!creationdateandtime, "dd/mm/yyyy") Then
        Exit Function
   End If
   If rstf!notes <> atrim(rstink!notes) Then
        Exit Function
   End If
   'Set rst = conODBC.OpenRecordset("SELECT Code, Description, DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente WHERE (((dbo.tblFormula.Code)=[Formula que vols buscar]));")
   Set rst = conODBC.OpenRecordset("SELECT Code, dbo.tblformula.Description,IdComponente,DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente where dbo.tblFormulaDetail.formulation=0 and dbo.tblformula.code='" + atrim(rstink!code) + "'")
  ' dbtintes.Execute "delete * from detallformules where idformula=" + atrim(lagran)
   Set rstdetall = dbtintes.OpenRecordset("select * from detallformules where idformula=" + atrim(lagran))
   While Not rst.EOF
      rstdetall.FindFirst "idcomponente=" + atrim(cadbl(rst!idcomponente))
      If rstdetall.NoMatch Then
           Exit Function
          Else
             'comparar percentatge
             If Redondejar(rstdetall![%decomponent], 0) <> Redondejar(cadbl(rst![%decomponent]), 0) Then
                Exit Function
             End If
      End If
     
      rst.MoveNext
   Wend
   
   comprovarformula = True
End Function
Private Sub Form_Load()
  Dim arguments As Variant
  Dim vhora As String
  fitxerini = "comandes.ini"
  'vnumidprograma = cadbl(Format(Now, "hhnnss"))
  'If llegir_ini("Tintes", "imprimir_etiqueta", fitxerini) <> "" And llegir_ini("Tintes", "imprimir_etiqueta", fitxerini) <> "{[}]" Then vnomesobrirperimprimir = True
  'vhora = llegir_ini("Tintes", "controlprogramaobert", fitxerini)
  'If IsDate(vhora) Then If DateDiff("s", CVDate(vhora), Now) < 5 And Not vnomesobrirperimprimir Then End
  arguments = ObtenerLíneaComando
  If arguments(1) <> "imprimir" Then If EstaCorriendo("Etiquetes tintes.exe") Then End
  
  cami = llegir_ini("General", "cami", fitxerini)
  ruta_relativa_docs = llegir_ini("General", "rutatmpetiquetestintes", fitxerini)
  
  If ruta_relativa_docs = "{[}]" Then
     ruta_relativa_docs = "c:\temp\"
     escriure_ini "General", "rutatmpetiquetestintes", "c:\temp\", fitxerini
  End If
  VnomfitxerconsumLam = rutadelfitxer(cami) + "Lam_consum_cola.txt"
  etrutatemp = ruta_relativa_docs
  '"c:\misdoc~1\commandes\comandes.mdb"
  If existeix("c:\ordprog.ini") Then
     cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
       Else: assignarpuntcomadecimal
  End If
  inicidragover = 0
  hora = Now
  centerscreen Me
  Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
  Set dbtintes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  Set dbcomandes = OpenDatabase(cami)
  If arguments(1) = "imprimir" Then
    imprimir_etiqueta atrim(arguments(2))
    'wait 5
    End
  End If
  obrirportseriebascula
  connexiosql
  cnumlotinplacsa = llegir_ini("General", "lotinplacsatintes", fitxerini)
  Me.visible = False
 If Not vnomesobrirperimprimir Then Call PonerSystray
'actualitzar_estocdecomponents

End Sub
Sub connexiosql()
   Set wsODBC = CreateWorkspace("", "tintes", "", dbUseODBC)
   Set conODBC = wsODBC.OpenConnection("connexiosql", dbDriverNoPrompt, , "ODBC;DATABASE=InkmakerDB;UID=sa;PWD=Mak2008;DSN=tintes")
   Set rstfink = conODBC.OpenRecordset("select * from dbo.tblFormula ", dbOpenSnapshot)
 
End Sub
Sub guardarcodi(rsttintes As Recordset)
   Dim proximalinia As String
   Dim valor As String
   If EOF(1) Then Exit Sub
   Line Input #1, proximalinia
   valor = Mid(proximalinia, InStr(1, proximalinia, "~") + 1)
   If atrim(valor) = "Codric:" Or atrim(valor) = "Desric:" Then valor = "BASE"
   rsttintes!codiformula = atrim(valor)
End Sub
Sub guardardescripcio(rsttintes As Recordset)
   Dim proximalinia As String
   Dim valor As String
   If EOF(1) Then Exit Sub
   Line Input #1, proximalinia
   valor = Mid(proximalinia, InStr(1, proximalinia, "~") + 1)
   rsttintes!descripcio = atrim(valor)
End Sub
Sub guardarformulacio(rsttintes As Recordset)
   Dim proximalinia As String
   Dim valor As String
   If EOF(1) Then Exit Sub
   Line Input #1, proximalinia
   valor = Mid(proximalinia, InStr(1, proximalinia, "~") + 1)
   'If atrim(valor) = "Codcli:" Then valor = "BASE"
   If atrim(valor) = "FType:" Then rsttintes!codiformula = "BASE"
   
End Sub
Sub guardarpedido(rsttintes As Recordset)
   Dim proximalinia As String
   Dim valor As String
   If EOF(1) Then Exit Sub
   Line Input #1, proximalinia
   valor = Mid(proximalinia, InStr(1, proximalinia, "~") + 1)
   'If atrim(valor) = "Codcli:" Then valor = "BASE"
   If atrim(valor) = "Nlav:" Then valor = "BASE"
   If rsttintes!codiformula = "BASE" Then valor = "BASE"
   rsttintes!Lotfabricacio = atrim(valor)
   If atrim(valor) = "BASE" Then esbase = True
End Sub
Sub guardardata(rsttintes As Recordset)
   Dim proximalinia As String
   Dim data As String
   Dim hora As String
   If EOF(1) Then Exit Sub
   Line Input #1, proximalinia
   data = Mid(proximalinia, InStr(1, proximalinia, "~") + 1)
   If EOF(1) Then Exit Sub
   Line Input #1, proximalinia
   If EOF(1) Then Exit Sub
   Line Input #1, proximalinia
   hora = Mid(proximalinia, InStr(1, proximalinia, "~") + 1)
   rsttintes!data = atrim(CVDate(atrim(data) + " " + atrim(hora)))
   If rsttintes!codiformula <> "BASE" Then actualitzar_nomformula rsttintes 'hi ha un error de tamany del camp de nomformula i l'actualitzo per sql per arreglarho
   
End Sub
Sub actualitzar_nomformula(rsttintes As Recordset)
   Dim rstfink As Recordset
   Dim vdata As Date
   vdata = rsttintes!data
   Set rstfink = conODBC.OpenRecordset("select * from dbo.tblLogBook WHERE year([StartEventDateandTime])=" + atrim(Year(vdata)) + " and month([StartEventDateandTime])=" + atrim(Month(vdata)) + " and day([StartEventDateandTime])=" + atrim(Day(vdata)) + " and datepart(hour,[StartEventDateandTime])=" + atrim(Hour(vdata)) + " and datepart(minute,[StartEventDateandTime])=" + atrim(Minute(vdata)))
   If Not rstfink.EOF Then
     If atrim(rstfink!formulacode) <> "" Then rsttintes!codiformula = rstfink!formulacode
   End If
   Set rstfink = Nothing
End Sub
Sub guardarkilos(rsttintes As Recordset)
   Dim proximalinia As String
   Dim valor As String
   If EOF(1) Then Exit Sub
   Line Input #1, proximalinia
   valor = Mid(proximalinia, InStr(1, proximalinia, "~") + 1)
   If separadordecimal = "," Then valor = passaradecimal(valor)
   rsttintes!kilosfabricats = cadbl(valor)
End Sub
Function separadordecimal() As String
   separadordecimal = "."
   If InStr(1, Trim(1 / 3), ",") Then separadordecimal = ","
End Function
Sub guardarcomp(rsttintes As Recordset, resp As String)
   Dim proximalinia As String
   Dim valor As String
   Dim camp As String
   If InStr(1, resp, ":") = 0 Then Exit Sub
   If EOF(1) Then Exit Sub
   camp = Mid(resp, InStr(1, resp, "~") + 1, (InStr(1, resp, ":")) - (InStr(1, resp, "~") + 1))
   Line Input #1, proximalinia
   If InStr(1, proximalinia, "~Comp") Or InStr(1, proximalinia, "~Flexo") Or InStr(1, resp, "Company") Then Exit Sub
   valor = Mid(proximalinia, InStr(1, proximalinia, "~") + 1)
   
   rsttintes.Fields(camp) = valor
End Sub
Sub imprimirinformetinta(codiformula As String, idtinta As Long)
' Dim rst As Recordset
  
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim camp As TextObject
  Dim f  As OLEObject
  Dim rstf As Recordset
  Dim rstt As Recordset
  Set rstt = dbtintes.OpenRecordset("SELECT  idtinta,codi,descripcio,referenciacolor from tintes where idtinta=" + atrim(idtinta))
  If rstt.EOF Then Exit Sub
  Set rstf = conODBC.OpenRecordset("SELECT Code, dbo.tblformula.Description,IdComponente,DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente where dbo.tblFormulaDetail.formulation=0 and dbo.tblformula.code='" + atrim(codiformula) + "'")
  If rstf.EOF Then Exit Sub
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "verificaciorelaciotintainkmaker.rpt", 1)

  oreport.FormulaFields.GetItemByName("descripcioformula").Text = "'" + atrim(rstf!code) + " ---->   " + atrim(rstf!Description) + "'"
  oreport.FormulaFields.GetItemByName("descripcio tinta").Text = "'" + atrim(rstt!descripcio) + "'"
  oreport.DiscardSavedData
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
    Else
      oreport.DisplayProgressDialog = False
      oreport.PrintOut False, 1
  End If
  Set rstt = Nothing
  Set rstf = Nothing
End Sub
Function comprovarsirelaciotintaiformula(codi As String, descripcio As String) As Double
   Dim rst As Recordset
   Dim idtinta As Long
   comprovarsirelaciotintaiformula = 0
   Set rst = dbtintes.OpenRecordset("select * from tintesformules where numformula='" + atrim(codi) + "'")
   If rst.EOF Then
      'no hi ha relacio formula i tinta s'ha de demanar la relació
      While idtinta = 0
        If idtinta = 0 Then MsgBox "Has d'escullir una tinta per la fabricació " + atrim(descripcio) + " o no podras continuar.", vbCritical, "Atenció"
        idtinta = triartinta(descripcio)
        'If idtinta <> 0 Then
        '   imprimirinformetinta codi, idtinta
        '   If UCase(InputBox("S'ha imprès un full de confirmació de relacio." + Chr(10) + "Has de portar-ho a una segona persona per d'onar l'OK de la relació i escriure [correcte]", "Confirmació relacio formula -> tinta")) <> "CORRECTE" Then
        '       idtinta = 0
        '   End If
        'End If
      Wend
      
      rst.AddNew
      rst!numformula = codi
      rst!idtinta = idtinta
      rst.Update
      Set rst = dbtintes.OpenRecordset("select * from tintesformules where idtinta=" + atrim(idtinta) + " order by predeterminada")
      If Not rst.EOF Then
         If Not rst!predeterminada Then
           rst.Edit
           rst!predeterminada = True
           rst.Update
         End If
      End If
        Else: idtinta = cadbl(rst!idtinta)
   End If
   comprovarsirelaciotintaiformula = idtinta
   Set rst = Nothing
End Function
Function triartinta(Optional ByVal descripcio As String) As Long
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
  formseleccio.cmissatge = "        Escullir tinta per: " + descripcio
  formseleccio.cmissatge.tag = "2" 'per escullir el cap primer de busqueda
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 1000
  formseleccio.Show 1
  If seleccioret = 1 Then
    triartinta = atrim(formseleccio.Data1.Recordset!idtinta)
  End If
  If seleccioret = 9 Then
    triartinta = 0
  End If
  Unload formseleccio
End Function
Sub comprovarentradanorepetida(rsttintes As Recordset)
  Dim rst As Recordset
  Dim vdate As Date
  vdata = rsttintes!data
  If Not IsDate(vdata) Then Exit Sub
  Set rst = dbtintes.OpenRecordset("select * from etiquetesgenerades where data=#" + Format(vdata, "mm/dd/yy hh:nn:ss") + "#")
  If Not rst.EOF Then
     rsttintes.CancelUpdate
     Set rsttintes = dbtintes.OpenRecordset("SELECT * FROM etiquetesgenerades ORDER BY ID asc")
     rsttintes.FindFirst "data=#" + Format(vdata, "mm/dd/yy hh:nn:ss") + "#"
     If Not rst.NoMatch Then rsttintes.Edit
  End If
End Sub
Sub llegir_etiqueta_tintes(fitxer As String)
  Dim resp As String
  Dim vnovallauna As String
  Dim rsttintes As Recordset
  Dim novallauna As String
  Dim vidtinta As Double
  Dim datafabricacio As String
  Dim idregistre As Long
  Open fitxer For Input As #1
  If EOF(1) Then Exit Sub
  Line Input #1, resp
  esbase = False
  If resp <> "" Then
     Set rsttintes = dbtintes.OpenRecordset("SELECT * FROM etiquetesgenerades ORDER BY ID asc")
     rsttintes.AddNew
  End If
  While Not EOF(1)
    If InStr(1, resp, "~Codric:") > 0 Then guardarcodi rsttintes
    If InStr(1, resp, "~Desric:") > 0 Then guardardescripcio rsttintes
    If InStr(1, resp, "~Formulation:") > 0 Then guardarformulacio rsttintes
    'If InStr(1, resp, "~Codlav:") > 0 Then guardarpedido rsttintes
    If InStr(1, resp, "~Bcode:") > 0 Then guardarpedido rsttintes
    If InStr(1, resp, "~Data:") > 0 Then guardardata rsttintes
    If InStr(1, resp, "~RealQuant:") > 0 Then guardarkilos rsttintes
    If InStr(1, resp, "~Comp") > 0 Then guardarcomp rsttintes, resp
    Line Input #1, resp
  Wend
  Close #1
  'wait 2
  comprovarentradanorepetida rsttintes
  idregistre = rsttintes!id
  
  If rsttintes.EditMode > 0 Then
     'If atrim(rsttintes!Lotfabricacio) <> "BASE" Then vidtinta = comprovarsirelaciotintaiformula(atrim(rsttintes!codiformula), atrim(rsttintes!descripcio) + "(" + atrim(rsttintes!codiformula) + ")")
     If atrim(rsttintes!codiformula) <> "BASE" Then vidtinta = comprovarsirelaciotintaiformula(atrim(rsttintes!codiformula), atrim(rsttintes!descripcio) + "(" + atrim(rsttintes!codiformula) + ")")
     rsttintes.Update
     wait 3
     rsttintes.FindFirst "id=" + atrim(idregistre)
     If Not rsttintes.EOF Then
       If rsttintes!Lotfabricacio = "BASE" Then
          novallauna = escullirllaunarecarrega(0)
          guardarhistoriallaunailotsambbase rsttintes, novallauna
         Else:
            
            novallauna = comprovarsihiharecarregarllaunes(vidtinta)
            comprovar_si_actualitzada rsttintes
            vnovallauna = guardarhistoriallaunailots(rsttintes, novallauna)
            If atrim(vnovallauna) <> atrim(novallauna) Then
              'barrejar les llaunes i imprimir etiqueta
               If atrim(novallauna) <> "" And atrim(vnovallauna) <> "" Then
                  barrejardosllaunes novallauna, vnovallauna
               End If
               ensenyar_missatge_imprimint_etiqueta vnovallauna
               imprimir_etiqueta vnovallauna
               dbtintes.Execute "update llaunes set aimpresores=true where numllauna='" + atrim(vnovallauna) + "'"
               While IsFormLoaded(formimprimintetiqueta)
                 DoEvents
               Wend
                  Else
                    'no imprimir l'etiqueta perquè es conserva numero llauna
                    ensenyar_missatge_imprimint_etiqueta "Noetiqueta"
            End If
       End If
     End If
  End If
  
  Set rsttintes = Nothing
End Sub
Sub ensenyar_missatge_imprimint_etiqueta(vnovallauna As String)
  Load formimprimintetiqueta
  If vnovallauna = "Noetiqueta" Then
     formimprimintetiqueta.etnumetiqueta = ""
     formimprimintetiqueta.etmissatge = "Aquesta llauna conserva l'etiqueta anterior."
     formimprimintetiqueta.BackColor = QBColor(12)
    Else: formimprimintetiqueta.etnumetiqueta = vnovallauna
  End If
  formimprimintetiqueta.Show
End Sub
Sub guardarhistoriallaunailotsambbase(rsttintes As Recordset, numllauna As String)
  Dim rstlots As Recordset
  Dim rstcomp As Recordset
  Dim rstllauna As Recordset
  Dim rsthistoria As Recordset
  Dim numnovallauna As Long
  Dim idtinta As Long
  Dim idllauna As Long
  Dim idhistoria As Long
  Dim sqllots As String
  Dim sqlcomponents As String
   If numllauna = "" Then Exit Sub
  'sqlcomponents = "SELECT dbo.tblFormula.Code, dbo.tblFormula.Description, dbo.tblComponenti.IdComponente, dbo.tblComponenti.DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente "
  'sqlcomponents = sqlcomponents + " WHERE (((dbo.tblFormula.Code)='" + rsttintes!codiformula + "'));"
  sqlcomponents = "SELECT dbo.tblComponenti.* FROM dbo.tblComponenti"
  Set rstcomp = conODBC.OpenRecordset(sqlcomponents)
  
  sqllots = "SELECT dbo.tblLogBook.StartEventDateandTime,dbo.tblLogBook.WorkOrder,dbo.tblLogBook.idlogbook, dbo.tblLogBookDetail.IDComponente, dbo.tblLogBookDetail.CodComponente, dbo.tblLogBookDetail.BatchCode, dbo.tblLogBookDetail.DispensedQuantity FROM dbo.tblLogBook INNER JOIN dbo.tblLogBookDetail ON dbo.tblLogBook.IDLogBook = dbo.tblLogBookDetail.IDLogBook "
  'sqllots = sqllots + " WHERE (((Format([StartEventDateandTime],""mm/dd/yyyy  hh:nn:ss""))=#5/31/2016 10:15:20#))"
  sqllots = sqllots + " WHERE year([StartEventDateandTime])=2016 and month([StartEventDateandTime])=5 and day([StartEventDateandTime])=31 and datepart(hour,[StartEventDateandTime])=15 and datepart(minute,[StartEventDateandTime])=10"
 ' sqllots = "SELECT format([StartEventDateandTime]) as prova FROM dbo.tblLogBook"

  
  
  Set rstlots = conODBC.OpenRecordset(sqllots)
  Set rstllauna = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(numllauna) + "'")
  If rstllauna.EOF Then MsgBox "Hi ha hagut algun error al vincular la carga amb la llauna, no es guardarà bé la seva historia"
  idllauna = rstllauna!id
  Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna")
  rsthistoria.AddNew
  rsthistoria!idnumllauna = idllauna
  rsthistoria!data = Now
  rsthistoria!tipusmoviment = "C"
  rsthistoria!formula = "BASE"
  rsthistoria!kg = Redondejar(rsttintes!kilosfabricats, 1)
  idhistoria = rsthistoria!id
  rsthistoria.Update
  rsthistoria.FindFirst "id=" + atrim(idhistoria)
  
  Set rsthistoria = dbtintes.OpenRecordset("select * from historiallaunalots")
  
  While Not rstlots.EOF
     If localitzarcomponent(rstcomp, "idcomponente", rstlots!idcomponente) Then
       rsthistoria.FindFirst "idhistoria=" + atrim(idhistoria) + " and idcomponent=" + atrim(rstcomp!idcomponente)
       If rsthistoria.NoMatch Then
            rsthistoria.AddNew
            rsthistoria!idhistoria = idhistoria
            rsthistoria!idcomponent = rstlots!idcomponente
            rsthistoria!numlotbase = numerolotdelcomponent(rstlots!idcomponente)
            rsthistoria!tanx100tinta = 0
            rsthistoria!kgtinta = Redondejar(cadbl(rstlots!DispensedQuantity) / 1000, 1)
            If rsthistoria!numlotbase = "0" Then MsgBox "No hi ha numero de lot possat al programa de inkmaker pel component " + UCase(atrim(rstcomp!DescComponente)) + Chr(10) + "S'ha d'arreglar el mes ràpid possible."
            rsthistoria.Update
       End If
     End If
     rstlots.MoveNext
  Wend
  calcularkgdisponiblesllauna numllauna
End Sub
Function escullirllaunarecarrega(idtintarecarrega As Integer) As String
  If idtintarecarrega = 0 Then Exit Function
  Unload formseleccio
  escullirllaunarecarrega = ""
  Load formseleccio
  formseleccio.caption = "Llaunes trobades pendents de recarregar"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
  formseleccio.Data1.RecordSource = "select numllauna,data from recarregarllaunes " + IIf(cadbl(idtintarecarrega) > 0, "where idtinta=" + atrim(idtintarecarrega) + " order by data ", "")
  formseleccio.refrescar
  If formseleccio.Data1.Recordset.EOF Then GoTo fi
  formseleccio.Data1.Recordset.MoveLast
  formseleccio.Data1.Recordset.MoveFirst
  If formseleccio.Data1.Recordset.RecordCount = 1 Then GoTo nomesuna
  'formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(0).width = 1200
  formseleccio.DBGrid2.Columns(1).width = 2000
  formseleccio.Show 1
  If seleccioret = 1 Then
nomesuna:
   escullirllaunarecarrega = atrim(formseleccio.Data1.Recordset!numllauna)
  End If
fi:
  Unload formseleccio
End Function
Function idtintadelallauna(nllauna As String) As Long
  Dim rst As Recordset
  idtintadelallauna = 0
  Set rst = dbtintes.OpenRecordset("select idtinta from llaunes where numllauna='" + atrim(nllauna) + "'")
  If Not rst.EOF Then idtintadelallauna = rst!idtinta
  Set rst = Nothing
End Function
Function comprovarsihiharecarregarllaunes(vidtinta As Double) As String
   Dim rst As Recordset
   Dim nllaunarecarrega As String
  ' vidtinta = idtintadelallauna(novallauna)
   Set rst = dbtintes.OpenRecordset("select * from recarregarllaunes where idtinta=" + atrim(vidtinta))
 
   If Not rst.EOF Then
      'si nomes hi ha una oferir amb un box d'escullir aquella
      'si hi ha mes d'una fer el formseleccio de les que hi ha
demanarllauna:
      nllaunarecarrega = escullirllaunarecarrega(CLng(vidtinta))
      If nllaunarecarrega <> "" Then
'         barrejardosllaunes nllaunarecarrega, novallauna
         dbtintes.Execute "delete * from recarregarllaunes where numllauna='" + atrim(nllaunarecarrega) + "'"
           Else
              If MsgBox("No has escullit cap llauna de recarrega, es correcte?", vbCritical + vbDefaultButton2 + vbYesNo, "Error") = vbNo Then GoTo demanarllauna
      End If
       
   End If
   comprovarsihiharecarregarllaunes = nllaunarecarrega
   Set rst = Nothing
End Function
Function proveidorperdefecte(idtinta As Long) As Double
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("select * from tintesreferencies where idtinta=" + atrim(idtinta) + " order by predeterminada")
  If Not rst.EOF Then proveidorperdefecte = rst!id
  Set rst = Nothing
End Function
Function guardarhistoriallaunailots(rsttintes As Recordset, novallauna As String) As String
  Dim rstlots As Recordset
  Dim rstcomp As Recordset
  Dim rstllauna As Recordset
  Dim rsthistoria As Recordset
  Dim numnovallauna As Long
  Dim idtinta As Long
  Dim idllauna As Double
  Dim idhistoria As Long
  Dim sqllots As String
  Dim sqlcomponents As String
  Dim vnumrecarrega As Double
  Dim vlotinplacsa As String
  
  sqlcomponents = "SELECT dbo.tblFormula.Code, dbo.tblFormula.Description, dbo.tblComponenti.IdComponente, dbo.tblComponenti.DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente "
  sqlcomponents = sqlcomponents + " WHERE (((dbo.tblFormula.Code)='" + rsttintes!codiformula + "'));"
  
  Set rstcomp = conODBC.OpenRecordset(sqlcomponents)
  
  sqllots = "SELECT dbo.tblLogBook.WorkOrder, dbo.tblLogBookDetail.IDComponente, dbo.tblLogBookDetail.CodComponente, dbo.tblLogBookDetail.BatchCode, dbo.tblLogBookDetail.DispensedQuantity FROM dbo.tblLogBook INNER JOIN dbo.tblLogBookDetail ON dbo.tblLogBook.IDLogBook = dbo.tblLogBookDetail.IDLogBook "
 ' sqllots = sqllots + " WHERE (((dbo.tblLogBook.WorkOrder)='" + rsttintes!Lotfabricacio + "'));"
 sqllots = sqllots + " WHERE (((dbo.tblLogBook.Barcode)='" + rsttintes!Lotfabricacio + "'));"
  Set rstlots = conODBC.OpenRecordset(sqllots)
  
  Set rstllauna = dbtintes.OpenRecordset("select idtinta from tintesformules where numformula='" + atrim(rsttintes!codiformula) + "'")
  If rstllauna.EOF Then Exit Function
  idtinta = cadbl(rstllauna!idtinta)
  If idtinta = 0 Then Exit Function
  If novallauna = "" Then
nova:
        Set rstllauna = dbtintes.OpenRecordset("select numllauna from contadors")
        numnovallauna = rstllauna!numllauna + 1
        dbtintes.Execute "update contadors set numllauna=[numllauna]+1"
        Set rstllauna = dbtintes.OpenRecordset("select * from llaunes")
        rstllauna.AddNew
        rstllauna!numllauna = "A" + atrim(numnovallauna)
        guardarhistoriallaunailots = rstllauna!numllauna
        rstllauna!idtinta = idtinta
        rstllauna!id_refproveidor = proveidorperdefecte(idtinta)
        rstllauna!situacio = "IMP"  'L'ESAÚ DIU QUE TREGUI AIXÓ PERQUÈ HA RECOLOCAT LES LLAUNES A MAGATZEM I NO HO VOL AIXI 02/12/24   buscarlaultimasituacio(idtinta)
        rstllauna!activa = True
        rstllauna!preuxrkilo = saber_preu_kg_tinta_llauna("", atrim(rsttintes!codiformula))
        rstllauna.Update
        vnumrecarrega = 1
        rstllauna.FindFirst "numllauna='A" + atrim(numnovallauna) + "'"
        If rstllauna.NoMatch Then MsgBox "Hi ha hagut algun error al crear la llauna no es guardarà bé la seva historia"
        idllauna = rstllauna!id
          Else
            Set rstllauna = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(novallauna) + "'")
            If rstllauna.EOF Then MsgBox ("No he trobat la llauna seleccionada " + atrim(novallauna) + " farè llauna nova"): GoTo nova
            If elslotssondiferents(rstllauna, rstlots) Then
              If cadbl(llegir_ini("General", "lotinplacsatintes", fitxerini)) > 0 Then
                 If buscarlotinplacsadelallauna(novallauna) <> atrim(llegir_ini("General", "lotinplacsatintes", fitxerini)) Then GoTo nova
                   Else: GoTo nova
              End If
            End If
            idllauna = rstllauna!id
            guardarhistoriallaunailots = rstllauna!numllauna
            vnumrecarrega = recarregamesgran(idllauna) + 1
  End If
  Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna")
  rsthistoria.AddNew
  rsthistoria!idnumllauna = idllauna
  rsthistoria!numrecarrega = vnumrecarrega
  rsthistoria!data = Now
  rsthistoria!tipusmoviment = "C"
  rsthistoria!formula = rsttintes!codiformula
  rsthistoria!kg = Redondejar(rsttintes!kilosfabricats, 1)
  rsthistoria.Update
  rsthistoria.MoveLast
  idhistoria = rsthistoria!id
  Set rsthistoria = dbtintes.OpenRecordset("select * from historiallaunalots")
  vlotinplacsa = atrim(llegir_ini("General", "lotinplacsatintes", fitxerini))
  If vlotinplacsa <> "" And vlotinplacsa <> "{[}]" Then
            rsthistoria.AddNew
            rsthistoria!idhistoria = idhistoria
            rsthistoria!idcomponent = 0
            rsthistoria!numlotbase = vlotinplacsa
            rsthistoria!tanx100tinta = 0
            rsthistoria!kgtinta = Redondejar(rsttintes!kilosfabricats, 1)
            rsthistoria.Update
  End If
  Set rstlots = conODBC.OpenRecordset(sqllots)
  Set rstcomp = conODBC.OpenRecordset(sqlcomponents)
  While Not rstlots.EOF
     If localitzarcomponent(rstcomp, "idcomponente", rstlots!idcomponente) Then
       rsthistoria.FindFirst "idhistoria=" + atrim(idhistoria) + " and idcomponent=" + atrim(rstcomp!idcomponente)
       If rsthistoria.NoMatch Then
            rsthistoria.AddNew
            rsthistoria!idhistoria = idhistoria
            rsthistoria!idcomponent = rstcomp!idcomponente
            rsthistoria!numlotbase = numerolotdelcomponent(rstlots!idcomponente)
            rsthistoria!tanx100tinta = cadbl(rstcomp![%decomponent])
            rsthistoria!kgtinta = Redondejar(cadbl(rstlots!DispensedQuantity) / 1000, 1)
            If rsthistoria!numlotbase = "0" Then MsgBox "No hi ha numero de lot possat al programa de inkmaker pel component " + UCase(atrim(rstcomp!DescComponente)) + Chr(10) + "S'ha d'arreglar el mes ràpid possible."
            rsthistoria.Update
       End If
     End If
     rstlots.MoveNext
  Wend
  calcularkgdisponiblesllauna guardarhistoriallaunailots
End Function
Function buscarlaultimasituacio(vidtinta As Long) As String
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from llaunes where situacio<>'SALA' and situacio<>'IMP' and idtinta=" + atrim(vidtinta))
   If Not rst.EOF Then buscarlaultimasituacio = atrim(rst!situacio)
   If buscarlaultimasituacio = "" Then buscarlaultimasituacio = "IMP"
End Function
Function elslotssondiferents(rstllauna As Recordset, rstlots As Recordset) As Boolean
  Dim rstlotsllauna As Recordset
  Dim vrecarga As Byte
  If rstlots.EOF Then elslotssondiferents = True: Exit Function
  Set rstlotsllauna = dbtintes.OpenRecordset("SELECT historiallauna.id, historiallauna.numrecarrega, Llaunes.numllauna, historiallaunalots.numlotbase AS lotb, historiallaunalots.idcomponent, historiallaunalots.numlotbase FROM (historiallauna INNER JOIN Llaunes ON historiallauna.idnumllauna = Llaunes.id) INNER JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria Where (((Llaunes.numllauna) = '" + atrim(rstllauna!numllauna) + "') AND ((historiallaunalots.numlotbase)<>'0' And (historiallaunalots.numlotbase)<>'')) ORDER BY historiallauna.numrecarrega DESC;")
  'Clipboard.SetText "SELECT historiallauna.id, historiallauna.numrecarrega, Llaunes.numllauna, historiallaunalots.numlotbase, historiallaunalots.idcomponent, historiallaunalots.numlotbase FROM (historiallauna RIGHT JOIN Llaunes ON historiallauna.idnumllauna = Llaunes.id) LEFT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria Where (((Llaunes.numllauna) = '" + atrim(rstllauna!numllauna) + "')) ORDER BY historiallauna.numrecarrega DESC;"
  If rstlotsllauna.EOF Then elslotssondiferents = True: GoTo fi
  vrecarga = cadbl(rstlotsllauna!numrecarrega)
  vrecarga = 1
  While Not rstlots.EOF
       If vrecarga <> cadbl(rstlotsllauna!numrecarrega) Then GoTo fi
       rstlotsllauna.FindFirst "idcomponent=" + atrim(rstlots!idcomponente) + " and numrecarrega=" + atrim(vrecarga)
       If Not rstlotsllauna.NoMatch Then
           ' rsthistoria!idcomponent = rstlots!idcomponente
            If atrim(rstlotsllauna!numlotbase) <> numerolotdelcomponent(rstlots!idcomponente) Then
               elslotssondiferents = True
            End If
             Else: elslotssondiferents = True
       End If
     rstlots.MoveNext
  Wend
fi:
  Set rstlotsllauna = Nothing
End Function
Function numerolotdelcomponent(idcomponent As Long) As String
    Dim rst As Recordset
    numerolotdelcomponent = "0"
    Set rst = conODBC.OpenRecordset("select batchcodea from dbo.tblcomponenti where idcomponente=" + atrim(idcomponent))
    If Not rst.EOF Then numerolotdelcomponent = rst!BatchCodeA
    Set rst = Nothing
End Function
Function localitzarcomponent(rst As Recordset, camp As String, component As Double) As Boolean
   localitzarcomponent = False
   If rst.EOF And rst.BOF Then Exit Function
  rst.Requery
  While Not rst.EOF
     If rst!idcomponente = component Then localitzarcomponent = True: GoTo fi
     rst.MoveNext
  Wend
fi:
 ' Set rst = Nothing
End Function
Sub comprovar_si_actualitzada(rsttintes As Recordset)
   Dim rstf As Recordset
   Dim correcte As Boolean
   Dim rstfink As Recordset
   
   
   Set rstf = dbtintes.OpenRecordset("select idtinta from tintesformules where numformula='" + atrim(rsttintes!codiformula) + "'")
   If rstf.EOF Then MsgBox "La formula: " + atrim(rsttintes!codiformula) + "  no existeix a la Base de Dades de Producció." + Chr(10) + " o bé no està relacionada amb alguna tinta." + Chr(10) + " Per poder treballar amb aquesta fòrmula s'ha de donar d'alta.", vbCritical, "Error": Exit Sub
   Set rstf = dbtintes.OpenRecordset("select * from formules where codiformula='" + atrim(rsttintes!codiformula) + "'")
   If rstf.EOF Then MsgBox "La formula: " + atrim(rsttintes!codiformula) + "  no existeix a la Base de Dades de Producció." + Chr(10) + " o bé no està relacionada amb alguna tinta." + Chr(10) + " Per poder treballar amb aquesta fòrmula s'ha de donar d'alta.", vbCritical, "Error": Exit Sub
   Set rstfink = conODBC.OpenRecordset("select * from dbo.tblFormula where code='" + atrim(rsttintes!codiformula) + "'", dbOpenSnapshot)
   correcte = comprovarformula(rstf, rstfink, conODBC)
   If Not correcte Then
      MsgBox "La formula: " + atrim(rsttintes!codiformula) + "  no està actualitzada a la Base de Dades de Producció." + Chr(10) + " Per poder treballar amb aquesta fòrmula s'ha d'actualitzar.", vbCritical, "Error": Exit Sub
      Exit Sub
       
   End If
End Sub

Function buscarubicaciollauna(numllauna As String) As String
   
End Function
Function demanarllauna(vformula As String) As String
   Dim a As Long
   Load formdemanarllauna
   formdemanarllauna.eformula = vformula
   formdemanarllauna.Show
   
   While formdemanarllauna.visible
     a = SetWindowPos(formdemanarllauna.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
     DoEvents
   Wend
   
End Function

Private Sub Form_MouseMove( _
    Button As Integer, _
    Shift As Integer, _
    x As Single, y As Single)

Dim msg As Long

    If (Me.ScaleMode = vbPixels) Then
        msg = x
    Else
        msg = x / Screen.TwipsPerPixelX
    End If
On Error Resume Next
    Select Case msg
        Case WM_LBUTTONDBLCLK Or WM_LBUTTONUP
            ' -- Si hacemos doble click con el botón izquierdo restauramos el form
           ' Me.WindowState = vbNormal
           ' Call SetForegroundWindow(Me.hwnd)
           ' If Not IsFormLoaded(formretorntintes) Then Me.Show
            formretorntintes.Show
            Call SetForegroundWindow(formretorntintes.hwnd)
            Form1.tag = "1"
        Case WM_RBUTTONUP
            Call SetForegroundWindow(Me.hwnd)
            ' -- Si hacemos Click con el boton derecho mostramos el popup Menu
            Me.WindowState = vbNormal
            Call SetForegroundWindow(Me.hwnd)
            If Not IsFormLoaded(formretorntintes) Then Me.Show
        'Case WM_LBUTTONUP
    End Select
End Sub

Private Sub RemoverSystray()
    Shell_NotifyIcon NIM_DELETE, systray
End Sub

Private Sub Form_Resize()
    If (Me.WindowState = vbMinimized) Then
        Me.Hide
        Call PonerSystray
    End If
'    Else
'        Call RemoverSystray
'    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   'Form1.tag = "sortint"
   Cancel = 1
   Exit Sub
   If MsgBox("Si tanques el programa deixerà d'imprimir etiquetes de l'INKMAKER. Si vols amagar-lo i continuar MINIMITZAL" + Chr(10) + "Segur que vols tancar?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Cancel = 1: Form1.tag = ""
   If Cancel = 0 Then
    Set conODBC = Nothing
    Call RemoverSystray
    End
   End If
End Sub

Private Sub mstringconnexio_Click()
   Dim v As String
   v = llegir_ini("Bascula", "connexio", "comandes.ini")
   v = InputBox("Escriu l'String de connexió al COM4. Ex:9600,N,8,1", "Connexió COM4", v)
   escriure_ini "Bascula", "connexio", v, "comandes.ini"
End Sub

Private Sub mtancar_Click()
    Set conODBC = Nothing
    Call RemoverSystray
    End
End Sub
Sub possar_horaultima_llauna(vhoraultima As Date)
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("select data from etiquetesgenerades order by id desc")
  While Not rst.EOF
     If IsDate(rst!data) Then vhoraultima = CVDate(rst!data): GoTo cont
     rst.MoveNext
  Wend
cont:
  Set rst = Nothing
End Sub
Function inkmaker_llegir_horaultima_etiqueta_nogravada(vhoraultima As Date) As Date
   Dim rstink As Recordset
   Dim inst_sql As String
   Dim vdata As String
   If vhoraultima = 0 Then Exit Function
   inst_sql = "SELECT dbo.tblLogBook.StartEventDateAndTime From dbo.tblLogBook order by StartEventDateAndTime desc " 'WHERE (((dbo.tblLogBook.StartEventDateAndTime)>now()));"
   Set rstink = conODBC.OpenRecordset(inst_sql, dbOpenSnapshot)
   If rstink.EOF Then MsgBox "No s'ha trobat cap component al INKMAKER.", vbCritical, "Atenció"
  ' rstink.MoveLast
   While DateDiff("s", rstink!StartEventDateAndTime, vhoraultima) < 0
      rstink.MoveNext
      If rstink.EOF Then GoTo cont
   Wend
cont:
   rstink.MovePrevious
   If rstink.BOF Then rstink.MoveNext
   If Not rstink.EOF Then inkmaker_llegir_horaultima_etiqueta_nogravada = rstink!StartEventDateAndTime
End Function
Sub guardar_etiquetagenerada(vhoraultimainkmaker As Date)
   escriure_ini "Hores", atrim(vhoraultimainkmaker), "", "c:\temp\~etiquetesgenerades.txt"
End Sub
Sub restarkgllaunaLam(v As String)
   Dim vcomanda As Double
   Dim vllauna As String
   Dim vkg As Double
   'separo les variables que estan amb ;
   If InStr(1, v, ";") > 0 Then
      vcomanda = cadbl(Mid(v, 1, InStr(1, v, ";") - 1))
      v = Mid(v, InStr(1, v, ";") + 1)
   End If
   If InStr(1, v, ";") > 0 Then
      vllauna = atrim(Mid(v, 1, InStr(1, v, ";") - 1))
      v = Mid(v, InStr(1, v, ";") + 1)
      vkg = cadbl(v)
   End If
   If vcomanda > 0 And vkg > 0 And vllauna <> "" Then
      
   End If
   
End Sub
Sub actualitzarKGalcontenidordecoladeLaminadora()
   Dim resp As String
   If Not existeix(VnomfitxerconsumLam) Then Exit Sub
   Open VnomfitxerconsumLam For Input As #1
   If EOF(1) Then Exit Sub
   Line Input #1, resp
   While resp <> ""
    restarkgllaunaLam resp
    resp = ""
    If Not EOF(1) Then Line Input #1, resp
  Wend
  Close #1
  On Error Resume Next
  If existeix(VnomfitxerconsumLam) Then Kill VnomfitxerconsumLam
    
End Sub

Private Sub mtipusbascula_Click()
   Dim v As String
   v = llegir_ini("Bascula", "modelbascula", "comandes.ini")
   v = InputBox("Escriu el model de la bascula model 1,2,3,...", "Tipus de bascula", v)
   escriure_ini "Bascula", "modelbascula", v, "comandes.ini"
End Sub

Private Sub Timer1_Timer()
   Dim resp As String
   Dim haentrat As Boolean
   Dim vhoraultima As Date
   Dim vhoraultimainkmaker As Date
   Static vcont As Byte
   
   'comprovar que no hi hagi un altra programa està obert
   If vcont > 9 Then
      If EstaCorriendo("Etiquetes tintes.exe") Then End
      vcont = 0
       Else: vcont = vcont + 1
   End If
   'Comprovar si hi ha gasto de cola a laminadora per actualitzar la llauna
   actualitzarKGalcontenidordecoladeLaminadora
   
   haentrat = False
   If vhoraultima = 0 Then possar_horaultima_llauna vhoraultima
   vhoraultimainkmaker = inkmaker_llegir_horaultima_etiqueta_nogravada(vhoraultima)
   'If CVDate(Format(vhoraultimainkmaker, "dd/mm/yy hh:nn")) > vhoraultima Then guardar_etiquetagenerada vhoraultimainkmaker
   If DateDiff("n", vhoraultima, vhoraultimainkmaker) > 0 Then guardar_etiquetagenerada vhoraultimainkmaker
'   ruta_relativa_docs = "\\serverprodu\dades"
   resp = Dir(ruta_relativa_docs + "~*.tmp")
   If resp <> "" Then actualitzarcarguescomponents
   If MSComm1.tag <> "error" Then
      etcomprovant = etcomprovant + "."
       Else: If Not vnomesobrirperimprimir And Not existeix("c:\ordprog.ini") Then RemoverSystray: End
   End If
   DoEvents
   While resp <> "" And MSComm1.tag <> "error"
      etcomprovant = "Comprovant Inkmaker.." + resp
      'elimino totes les recargues pendents de fa mes de mitja hora
      'dbtintes.Execute "delete * from recarregarllaunes where data<dateadd('n',-30,now)"
      'l'esaú m'ha dit que ho tragues per mirar si ens dona problema amb les recarregues
      wait 2
      llegir_etiqueta_tintes ruta_relativa_docs + resp
      resp = Dir
      haentrat = True
   Wend
   If haentrat Then
     borrar_temporals
     etcomprovant = ""
     actualitzar_estocdecomponents
   End If
   If Len(etcomprovant) > 10 Then etcomprovant = ""
End Sub
Sub comprovar_si_imprimir_etiqueta()
   Dim vnllauna As String
   vnllauna = llegir_ini("Tintes", "imprimir_etiqueta", fitxerini)
   If vnllauna <> "" And vnllauna <> "{[}]" Then
     escriure_ini "Tintes", "imprimir_etiqueta", "", fitxerini
     formretorntintes.cridar_imprimir_etiqueta vnllauna
     If vnomesobrirperimprimir Then End
   End If
End Sub


Sub crear_carpeta_oldtemporals()
 On Error Resume Next
 MkDir ruta_relativa_docs + "etiquetesanteriors"
End Sub

Sub borrar_temporals()
   crear_carpeta_oldtemporals
   On Error GoTo errorborrant
   Copiar_Fitxer ruta_relativa_docs + "~*.tmp", ruta_relativa_docs + "etiquetesanteriors", 8
   Kill ruta_relativa_docs + "~*.tmp" ' "~*.tmp"
   'borrar fitxer guardars mes vells de 30 dies
  '' Shell "FORFILES /P """ + ruta_relativa_docs + "etiquetesanteriors" + """ /S /D -30 /c ""CMD /c DEL /Q @PATH""", vbHide
   Exit Sub
errorborrant:
   If err.Number <> 53 Then
     
    MsgBox "Hi ha hagut un error eliminant els fitxers temporals de inkmaker" + Chr(10) + ruta_relativa_docs + "~*.tmp", vbCritical, "Error"
   End If
End Sub
Private Sub PonerSystray()
    
    With systray
        ' -- Tamaño de la estructura systray
        .cbSize = Len(systray)
        ' -- Establecemos el Hwnd, en este caso del formulario
        .hwnd = Me.hwnd

        .uId = vbNull
        ' -- Flags
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        ' -- Establecemos el mensaje callback
        .ucallbackMessage = WM_MOUSEMOVE
        ' -- establecemos el icono, en este caso el que tiene el form, puede ser otro
        .hIcon = Me.Icon
        ' -- Establecemos el tooltiptext
        .szTip = Me.caption & vbNullChar
        ' -- Ponemos el icono en el systray
        Shell_NotifyIcon NIM_ADD, systray
    End With

End Sub
 
Sub possarpesbascula()
    If MSComm1.tag = "" Then etpesbascula = Redondejar(pesbascula, 1)
    'If MSComm1.tag = "error" Then etpesbascula = "0"
   ' If cadbl(llegir_ini("Tintes", "pesbascula", fitxerini)) > 0 Then etpesbascula = llegir_ini("Tintes", "pesbascula", fitxerini)
End Sub
Private Sub Timer2_Timer()
   Static jaheentrat As Boolean
   escriure_ini "Tintes", "controlprogramaobert", Now, fitxerini
   possarpesbascula
   comprovar_si_imprimir_etiqueta
   'A LES 12 i quart DE LA NIT FA ACTUALITZACIÓ DE LES FORMULES INKMAKER
   If Hour(Now) = 0 And Minute(Now) = 15 Then
      If cadbl(llegir_ini("Tintes", "dia_actualitzarformulesinkmaker", "comandes.ini")) <> Day(Now) Then
        Shell llegir_ini("General", "rutallistats", "comandes.ini") + "manteniment tintes.exe actualitzarformules"
        escriure_ini "Tintes", "dia_actualitzarformulesinkmaker", atrim(Day(Now)), "comandes.ini"
      End If
   End If
   If InStr(1, etcomprovant, "~") > 0 Then Exit Sub
   If Not jaheentrat And cadbl(etpesbascula) > 1 Then
      'If Form1.visible Then Exit Sub
      Form1.Hide
      If Form1.tag = "sortint" Then Exit Sub
      jaheentrat = True
      Timer1.Enabled = False
      botoretorn.tag = ""
      If Not IsFormVisible(formretorntintes) Then
         Form1.visible = False
         Form1.WindowState = 0
         formretorntintes.Show
         While IsFormVisible(formretorntintes)
           wait 1
         Wend
         If formretorntintes.visible = False Then
              Unload formretorntintes
              GoTo cont
         End If
cont:
         Form1.visible = False
      End If
         
      
      Timer1.Enabled = True
   End If
   If cadbl(etpesbascula) < 1 Then jaheentrat = False
End Sub

Private Sub Timer3_Timer()

End Sub

