VERSION 5.00
Begin VB.Form formcapturaetiqueta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Finestra captura etiqueta"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8760
   ControlBox      =   0   'False
   Icon            =   "formcapturaetiqueta.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Height          =   630
      Left            =   8085
      Picture         =   "formcapturaetiqueta.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Configuració"
      Top             =   7080
      Width           =   660
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   7395
      Top             =   105
   End
   Begin VB.CommandButton sortir 
      Height          =   510
      Left            =   7725
      Picture         =   "formcapturaetiqueta.frx":0ADD
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Sortir"
      Top             =   7740
      Width           =   1020
   End
   Begin VB.Frame Frame2 
      Height          =   1380
      Left            =   7845
      TabIndex        =   8
      Top             =   2370
      Width           =   795
      Begin VB.CommandButton Command4 
         Height          =   615
         Left            =   90
         Picture         =   "formcapturaetiqueta.frx":1067
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   570
      End
      Begin VB.TextBox vrotate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   90
         TabIndex        =   9
         Text            =   "0"
         Top             =   735
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1950
      Left            =   7830
      TabIndex        =   4
      Top             =   4965
      Width           =   795
      Begin VB.TextBox ctanxcent 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   90
         TabIndex        =   7
         Text            =   "80"
         Top             =   735
         Width           =   585
      End
      Begin VB.CommandButton Command2 
         Height          =   615
         Left            =   90
         Picture         =   "formcapturaetiqueta.frx":189D
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   570
      End
      Begin VB.CommandButton Command3 
         Height          =   615
         Left            =   90
         Picture         =   "formcapturaetiqueta.frx":1F76
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1260
         Width           =   570
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0025EFAD&
      Caption         =   "Capturar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   7695
      Picture         =   "formcapturaetiqueta.frx":2642
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   750
      Width           =   1005
   End
   Begin VB.PictureBox p 
      Height          =   180
      Left            =   6165
      ScaleHeight     =   120
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   270
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   7845
      Left            =   105
      ScaleHeight     =   519
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   498
      TabIndex        =   1
      Top             =   450
      Width           =   7530
   End
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   6750
      Top             =   45
   End
   Begin VB.PictureBox Picture1 
      Height          =   630
      Left            =   7755
      ScaleHeight     =   570
      ScaleWidth      =   780
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "formcapturaetiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle _
    As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess _
    As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Sub ShellAndWait(ByVal program_name As String, _
                         Optional ByVal window_style As VbAppWinStyle = vbNormalFocus, _
                         Optional ByVal max_wait_seconds As Long = 0)
Dim lngProcessId As Long
Dim lngProcessHandle As Long
Dim datStartTime As Date
Const WAIT_TIMEOUT = &H102
Const SYNCHRONIZE As Long = &H100000
Const INFINITE As Long = &HFFFFFFFF

    ' Start the program.
    On Error GoTo ShellError
    lngProcessId = Shell(program_name, window_style)
    On Error GoTo 0
    
    DoEvents

    ' Wait for the program to finish.
    ' Get the process handle.
    lngProcessHandle = OpenProcess(SYNCHRONIZE, 0, lngProcessId)
    If lngProcessHandle <> 0 Then
        datStartTime = Now
        Do
          If WaitForSingleObject(lngProcessHandle, 250) <> WAIT_TIMEOUT Then
            Exit Do
          End If
          DoEvents
          If max_wait_seconds > 0 Then
            If DateDiff("s", datStartTime, Now) > max_wait_seconds Then Exit Do
          End If
        Loop
        CloseHandle lngProcessHandle
    End If
    Exit Sub
    
ShellError:
End Sub

Private Sub comboorigen_Change()

End Sub

Private Sub Command1_Click()
  If Not existeix("c:\temp\capturaetiqueta.jpg") Then Exit Sub
  Timer1.Enabled = False
  If MsgBox("Veus correctament l'etiqueta de la bobina?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
     ConvertirFormats "c:\temp\capturaetiqueta.jpg", "c:\temp\capturaetiqueta.jpg", 72
     FileCopy "c:\temp\capturaetiqueta.jpg", "c:\temp\capturaetiqueta_OK.jpg"
       Else: Timer1.Enabled = True: Exit Sub
  End If
  capCaptureStop lwndC
  esperamigsegon
  escriure_ini "Baixes", "camaratanxcent", atrim(ctanxcent), "comandes.ini"
  Unload formcapturaetiqueta
End Sub

Private Sub Command2_Click()
 ctanxcent = ctanxcent + 10
  If ctanxcent > 99 Then ctanxcent = 90
End Sub

Private Sub Command3_Click()
   ctanxcent = ctanxcent - 10
  If ctanxcent < 1 Then ctanxcent = 10
  
End Sub

Private Sub Command4_Click()
  vrotate = vrotate + 90
  If vrotate > 300 Then vrotate = 0
End Sub


Private Sub Command5_Click()
   capDlgVideoSource lwndC
End Sub

Private Sub Form_Activate()
  Timer1.Enabled = True
End Sub

Private Sub Form_Load()
  
    
  
   On Error Resume Next
   Kill "c:\temp\capturaetiqueta_OK.jpg"
   On Error GoTo 0
   ctanxcent = cadbl(llegir_ini("Baixes", "camaratanxcent", "comandes.ini"))
   comboorigen = llegir_ini("Baixes", "cameraultimdriver", "comandes.ini")
   If ctanxcent = "0" Then ctanxcent = "60"

   
   obrir_camera
   Timer1_Timer

   ' seleccionar_cameraescullida
End Sub
Sub obrir_camera()
    Dim info As BITMAPINFO
    Dim lpszName As String * 100
    Dim lpszVer As String * 100
    Dim Caps As CAPDRIVERCAPS
    
    '//Create Capture Window

    capGetDriverDescriptionA 0, lpszName, 100, lpszVer, 100  '// Retrieves driver info
    lwndC = capCreateCaptureWindowA(lpszName, WS_CHILD Or WS_VISIBLE, 0, 0, 160, 120, Picture1.hwnd, 0)

    '// Connect the capture window t-o the driver
    capDriverConnect lwndC, 0
    '// Get the capabilities of the capture driver
    capDriverGetCaps lwndC, VarPtr(Caps), Len(Caps)
    If Caps.fCaptureInitialized = 0 Then
        formcapturaetiqueta.Caption = "No s'ha pogut inicialitzar l'escaner."
        GoTo fi
    End If
    capGetVideoFormat lwndC, VarPtr(info), Len(info)
    'info.bmiHeader.biWidth = 640
    'info.bmiHeader.biHeight = 480
    info.bmiHeader.biCompression = 32595559       ' YUY2 compression
    'info.bmiHeader.biSizeImage =  640 * 480 * 2
    capSetVideoFormat lwndC, VarPtr(info), Len(info)
    
    '// Set the video stream callback function
    capSetCallbackOnVideoStream lwndC, AddressOf MyVideoStreamCallback
    capSetCallbackOnFrame lwndC, AddressOf MyFrameCallback
    
    '// Set the preview rate in milliseconds
    capPreviewRate lwndC, 500
fi:
   

End Sub
Private Sub sortir_Click()
   Timer1.Enabled = False
   capCaptureStop lwndC
   esperamigsegon
   Unload formcapturaetiqueta
End Sub
Sub capturar_imatge_alfitxer(vnomfitxer As String)
  capGrabFrame lwndC
  capFileSaveDIB lwndC, vnomfitxer
  capCaptureStop lwndC

End Sub
Sub esperamigsegon()
    IniTime = GetTickCount()
    While GetTickCount() < (IniTime + 500)
        DoEvents
    Wend
End Sub
Private Sub Timer1_Timer()
'  Clipboard.SetData Picture1.Image
Dim vnomfitxer As String
Dim vnomfitxercrop As String

Dim x As Long
Dim y As Long
Dim d As Long
 
 'SI JA SOC DINS NO TORNO A ENTRAR
If Timer1.Tag <> "" Then Exit Sub

Timer1.Tag = "esticdins"

vnomfitxer = "c:\temp\capturaetiqueta.bmp"
vnomfitxercrop = "c:\temp\capturaetiqueta_crop.jpg"

'ELIMINO ELS DOS FITXERS QUE TREBALLARÉ PER CAPTURAR I RETALLAR LA IMATGE
 On Error Resume Next
  Kill vnomfitxer
  Kill vnomfitxercrop
  Kill "c:\temp\capturaetiqueta.jpg"
  On Error GoTo 0
    'HO HE FET AMB EL PROGRAMA AQUEST PER CAPTURAR IMATGE PERO AMB EL AVICAP TAMBÉ HAURIA D'ANAR BÉ
      'SI CLAGUES ES PODRIA UTILITZAR JA HO HAVIA FET I FUNCIONAVA
 'ShellAndWait "\\serverprodu\Dades\progcomandes\aplicacio\CapturarWebCam\WebCamImageSave.exe  /capture /Filename """ + vnomfitxer + """", vbHide, 10
 capturar_imatge_alfitxer vnomfitxer
 
  'RETALLO LA IMATGE CAPTURADA A LA MIDA DE LO QUE SEMBLI QUE ES POT AUTORETALLAR
   retallarimatgeFitxer vnomfitxer, vnomfitxercrop, cadbl(ctanxcent)
   If existeix(vnomfitxercrop) Then
     If FileLen(vnomfitxercrop) < 1000 Then FileCopy vnomfitxer, vnomfitxercrop
     RotarImatge vnomfitxercrop, vnomfitxercrop, cadbl(vrotate)
   End If

 'ESCALO LA IMATGE PER VEURE-LA CORRECTAMENT AL PICTURE
  p.AutoSize = True
  If Timer1.Tag = "sortint" Then GoTo fi
  Picture2.Picture = LoadPicture("")
  If Not existeix(vnomfitxercrop) Then GoTo fi
  Set p = LoadPicture(vnomfitxercrop)
  If p.ScaleHeight > p.ScaleWidth Then
      d = p.ScaleHeight - p.ScaleWidth
       Else: d = p.ScaleWidth - p.ScaleHeight
  End If
  x = IIf(p.ScaleHeight > p.ScaleWidth, ((p.ScaleWidth) / 100) * (50000 / p.ScaleHeight), 500)
  y = IIf(p.ScaleHeight > p.ScaleWidth, 500, ((p.ScaleHeight) / 100) * (50000 / p.ScaleWidth))
  If y = 0 Then y = 1
  If x = 0 Then x = 1
  Picture2.PaintPicture p, 0, 0, x, y
  Picture2.AutoRedraw = True
'COPIO LA IMATGE TALLADA COM A IMATGE CORRECTE PER TAL DE DEIXAR-LA APUNT SI L'USUARI L'ACCEPTA
  FileCopy vnomfitxercrop, "c:\temp\capturaetiqueta.jpg"
fi:
  Timer1.Tag = ""
  
End Sub



Private Sub Timer2_Timer()
 'If Formdesbobinadors.Tag = "escanejant" Then
 '    Picture2.Picture = LoadPicture("")
 '    DoEvents
 '    Timer1.Enabled = True: Formdesbobinadors.Tag = ""
 ' End If
End Sub
