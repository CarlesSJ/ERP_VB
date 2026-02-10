VERSION 5.00
Begin VB.Form formControladorServidor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servidor Controlador Producció"
   ClientHeight    =   1680
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   4365
   ControlBox      =   0   'False
   Icon            =   "formControladorServidor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H006BEBB1&
      Caption         =   "Compactar ara TOT"
      Height          =   390
      Index           =   1
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   795
      Width           =   1725
   End
   Begin VB.TextBox linia 
      Height          =   285
      Left            =   3705
      TabIndex        =   5
      Top             =   45
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   3765
      Top             =   750
   End
   Begin VB.CommandButton Command3 
      Height          =   315
      Left            =   3975
      Picture         =   "formControladorServidor.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Tancar programa"
      Top             =   30
      Width           =   360
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Servei Producció"
      Height          =   390
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1215
      Width           =   1725
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Obrir programa serveis producció al iniciar Windows."
      Height          =   270
      Left            =   255
      TabIndex        =   2
      Top             =   390
      Width           =   4065
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Compactar Diumenges"
      Height          =   270
      Left            =   270
      TabIndex        =   1
      Top             =   90
      Width           =   2760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Minimitzar__"
      Height          =   390
      Left            =   3105
      TabIndex        =   0
      Top             =   1230
      Width           =   1140
   End
End
Attribute VB_Name = "formControladorServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, _
                                ByVal bInvert As Long) As Long



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

' -- Constantes para las acciones
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

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Dim systray As NOTIFYICONDATA
Private Function EstaCorriendo(ByVal NombreDelProceso As String) As Boolean
    Const MAX_PATH As Long = 260
    Dim con As Byte
    Dim lProcesses() As Long, lModules() As Long, N As Long, lRet As Long, hProcess As Long
    Dim sName As String
    NombreDelProceso = UCase$(NombreDelProceso)
    ReDim lProcesses(1023) As Long
 con = 0
    If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
        For N = 0 To (lRet \ 4) - 1
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(N))
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
        Next N
    End If
End Function




Private Sub Check1_Click()
    escriure_ini "Controlador_Servidor", "compactardiumenges", atrim(Check1.Value), "enviarservidor.ini"
End Sub

Private Sub Check2_Click()
    escriure_ini "Controlador_Servidor", "obrirserveialiniciarwindows", atrim(Check1.Value), "enviarservidor.ini"
End Sub

Private Sub Command1_Click()
  Me.Hide
 Call PonerSystray

End Sub

Private Sub ensenya_totes_Click()

End Sub



Private Sub Command2_Click(Index As Integer)
  Dim r As String
  If Index = 0 Then obrir_programa_controlproduccio
  If Index = 1 Then
    If MsgBox("Estas segur que vols compactar ara?" + Chr(10) + "TOTS ELS PROGRAMES HAN D'ESTAR TANCATS.", vbCritical + vbDefaultButton2 + vbYesNo, "COMPACTAR") = vbNo Then Exit Sub
    formControladorServidor.Tag = "compactant"
    KillProcess "EnviarIncidenciesServidor.exe"
    r = Shell(llegir_ini("General", "camillistats", "enviarservidor.ini") + "Compactarbasesdedades.exe c:\dades\progcomandes\dades", vbNormalFocus)
  End If
End Sub

Private Sub Command3_Click()
  If MsgBox("Segur que vols tancar definitivament el programa?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then End
End Sub

Sub compactarbasededades()
   Dim r As String
   r = llegir_ini("Controlador_Servidor", "ultimcompactardiumenges", "enviarservidor.ini")
   If IsDate(r) Then
      If Day(r) = Day(Now) Then Exit Sub
   End If
   formControladorServidor.Tag = "compactant"
   KillProcess "EnviarIncidenciesServidor.exe"
   Sleep 1000
   r = Shell(llegir_ini("General", "camillistats", "enviarservidor.ini") + "Compactarbasesdedades.exe c:\dades\progcomandes\dades", vbNormalFocus)
   'r = Shell(llegir_ini("General", "camillistats", "enviarservidor.ini") + "Compactarbasesdedades.exe C:\temp\compact")
   escriure_ini "Controlador_Servidor", "ultimcompactardiumenges", Trim(Now), "enviarservidor.ini"
End Sub


Private Sub Form_Initialize()
 Me.Hide
Call PonerSystray
ensenyar_errorservidors
End Sub
Function atrim(valo As Variant) As String
On Error Resume Next
  If IsNull(valo) Then valo = ""
  atrim = Trim(valo)
End Function
Function cadbl(valo As Variant) As Double
  If Not IsNumeric(valo) Then valo = 0
  cadbl = CDbl(valo)
End Function


Private Sub Form_MouseMove( _
    Button As Integer, _
    Shift As Integer, _
    X As Single, Y As Single)

Dim msg As Long

    If (Me.ScaleMode = vbPixels) Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If

    Select Case msg
        Case WM_LBUTTONDBLCLK
            ' -- Si hacemos doble click con el botón izquierdo restauramos el form
            Me.WindowState = vbNormal
            Call SetForegroundWindow(Me.hwnd)
            Me.Show

        Case WM_RBUTTONUP
            Call SetForegroundWindow(Me.hwnd)
            ' -- Si hacemos Click con el boton derecho mostramos el popup Menu
            Me.WindowState = vbNormal
            Call SetForegroundWindow(Me.hwnd)
            Me.Show
        Case WM_LBUTTONUP
    End Select
End Sub

Private Sub RemoverSystray()
    Shell_NotifyIcon NIM_DELETE, systray
End Sub

Private Sub Form_Resize()
    If (Me.WindowState = vbMinimized) Then
        Me.Hide
        Call PonerSystray
    Else
        Call RemoverSystray
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' -- cuando descargamos el form removemos el Icono del systray
    RemoverSystray
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
        .szTip = Me.Caption & vbNullChar
        ' -- Ponemos el icono en el systray
        Shell_NotifyIcon NIM_ADD, systray
    End With

End Sub




Private Sub Form_Load()
 Dim arguments As Variant
 If App.PrevInstance Then End
 If EstaCorriendo("comprovaretreb.exe") Then End
 arguments = ObtenerLíneaComando
 If Trim(arguments(1)) <> "" Then cami = Trim(arguments(1))
 cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
   DoEvents
   Me.Visible = False
 Check1.Value = cadbl(llegir_ini("Controlador_Servidor", "compactardiumenges", "enviarservidor.ini"))
 Check2.Value = cadbl(llegir_ini("Controlador_Servidor", "obrirserveialiniciarwindows", "enviarservidor.ini"))
  If Check1.Value = 1 And WeekDay(Now, vbMonday) = 7 Then
       compactarbasededades
  End If
  Timer1_Timer
End Sub
Function ObtenerLíneaComando(Optional MaxArgs)
    'Declara las variables.
    Dim C, LíneaComando, LonLínComando, ArgIn, i, NúmArgs
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
        C = Mid(LíneaComando, i, 1)
        'Comprueba espacio o tabulación.
        If (C <> " " And C <> vbTab) Then
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

ArgArray(NúmArgs) = ArgArray(NúmArgs) + C
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

Private Sub m_historialsoks_Click()
  
End Sub

Private Sub m_Pararcomprovacioitancar_Click()

End Sub

Private Sub norecordar_Click()

End Sub

Sub comprovar_siestaobertservidorproduccio()
  Dim v As Double
  Dim r As String
  If formControladorServidor.Tag <> "" Then
     If Not estaobertelprograma("Compactarbasesdedades.exe") Then
       formControladorServidor.Tag = ""
       If existeix("c:\temp\logcompactar.txt") Then
         enviar_avis "Compactar bases de dades de Producció.", " Log compactar base de dades.", "c:\temp\logcompactar.txt"
       End If
       Exit Sub
     End If
  End If
  If Not estaobertelprograma("enviarincidenciesservidor.exe") Then
     If Not estaobertelprograma("Compactarbasesdedades.exe") Then obrir_programa_controlproduccio
  End If
End Sub
Sub obrir_programa_controlproduccio()
     Dim r As String
     escriure_ini "General", "dataprogramafuncionant", atrim(Now), "enviarservidor.ini"
     r = Shell(llegir_ini("General", "camillistats", "enviarservidor.ini") + "EnviarIncidenciesServidor.exe", vbHide)
End Sub
Sub capturarpantalla(vfitxer As String)
   Dim r As String
   r = Shell(llegir_ini("General", "camillistats", "enviarservidor.ini") + "Capturarpantalla\capturarpantalla.bat " + vfitxer, vbHide)
End Sub
Private Sub Timer1_Timer()
  Dim vdata As String
  
  'timer cada 30 segons
  If errorservidors <> "" Then
     ensenyar_errorservidors
      Else: Unload formsplash
  End If

  If Check2.Value = 1 Then
      comprovar_siestaobertservidorproduccio
  End If
 ' vdata = llegir_ini("General", "dataprogramafuncionant", "enviarservidor.ini")
 ' If Not IsDate(vdata) Then vdata = Now
 ' If DateDiff("n", vdata, Now) > 10 And DateDiff("h", vavisenviat, Now) > 0 And formControladorServidor.Tag = "" Then
 '     SendKeys "{PGDN}" 'envio una tecla perquè es desperti lordinador
 '     Sleep 3000
 '     capturarpantalla "c:\temp\capturapantalla.jpg"
 '     Sleep 1000  'espera un segon
 '     KillProcess "EnviarIncidenciesServidor.exe"
 '     enviar_avis "Incidencia programa Producció.", " Servei producció.", "c:\temp\capturapantalla.jpg"
 '     vavisenviat = Now
 ' End If
  
End Sub
Function errorservidors() As String
   If Not existeix("\\servidorsap\seidor_COMUNICADOR") Then errorservidors = errorservidors + "Servidorsap" + vbNewLine
     If Not existeix("\\serverprodu\dades") Then errorservidors = errorservidors + "SERVERPRODU" + vbNewLine
     If Not existeix("\\ord_copies\documentacioclixes") Then errorservidors = errorservidors + "ORD_COPIES" + vbNewLine
End Function
Sub ensenyar_errorservidors()
     If errorservidors = "" Then Exit Sub
     Load formsplash
     formsplash.eterror = "Error Connexió:" + vbNewLine + errorservidors
     formsplash.Show
End Sub
Function enviar_avis(vassumpte As String, vcos As String, vfitxeradjunt As String)
    Dim emails As String
    Dim destinatari As String
    Dim vusuari As String
    Dim vpassword As String
    Dim vresp As String
    vusuari = llegir_ini("dadesservidor", "usrsmtp", "enviarservidor.ini")
    vpassword = llegir_ini("dadesservidor", "passsmtp", "enviarservidor.ini")
    destinatari = "miquel.inplacsa@gmail.com"
    If InStr(1, destinatari, "@") > 0 Then
      vresp = enviaremailswitchmail("incidenciesinplacsa@gmail.com", destinatari, vassumpte, vcos, vfitxeradjunt, , vusuari, vpassword)
    End If
End Function
Function enviaremailswitchmail(vremitent As String, sSendTo As String, sSubject As String, sText As String, Optional adjunt As String, Optional vidavis As Long, Optional vusuari As String, Optional vcontrasenya As String) As Boolean
  Dim usuarim As String
  Dim contrasenyam As String
  Dim destinatari As String
  Dim instream  As Object
  Dim vcont As Double
  Dim r As String
  
   enviaremailswitchmail = False
   usuarim = vusuari
   contrasenyam = vcontrasenya
   If llegir_ini("General", "camillistats", "enviarservidor.ini") = "{[}]" Then escriure_ini "General", "camillistats", "\\serverprodu\dades\progcomandes\aplicacio\", "enviarservidor.ini"
   Open llegir_ini("General", "camillistats", "enviarservidor.ini") + "\enviomailswithmail.xml" For Input As #1
   linia.Text = Input(LOF(1), #1)
   Close #1
   
   destinatari = llegir_ini("destinataris", sSendTo, "enviarservidor.ini")
   If destinatari = "{[}]" Then destinatari = sSendTo
   sSendTo = destinatari
   linia = Mid(linia, 4)
   substituirtextbox linia, "#remitent#", vremitent
   substituirtextbox linia, "#destinatari#", sSendTo
   substituirtextbox linia, "#asumpte#", sSubject
   'substituir "#cosdelmisatge#", "CreateObject(""Scripting.FileSystemObject"").OpenTextFile(""C:\temp\cosmissatge.txt"", 1).ReadAll"
   substituirtextbox linia, "#fitxeradjunt#", adjunt
   If InStr(1, LCase(sText), "\cosmissatge.txt") > 0 Then
       substituirtextbox linia, "#fitxertxtcosmissatge#", sText
        Else: substituirtextbox linia, "#fitxertxtcosmissatge#", "c:\temp\cosmissatge.txt"
   End If
   'If adjunt = "" Then substituir "objMessage.AddAttachment """"", ""
   substituirtextbox linia, "#usuarigmail#", usuarim
   substituirtextbox linia, "#contrasenyagmail#", contrasenyam
   If InStr(1, LCase(sText), "\cosmissatge.txt") = 0 Then
        Open "c:\temp\cosmissatge.txt" For Output As #2
        Print #2, sText
   '     passarliniesdavisosalfitxertxt vidavis
        Close #2
   End If
    On Error Resume Next
    Kill "c:\temp\enviomail_Ctrl.xml"
    Kill "c:\temp\registreemail_Ctrl.txt"
    On Error GoTo 0
    Set instream = CreateObject("ADODB.Stream")
    With instream
        .Open
        .Type = 2
        .Charset = "utf-8"
        .LineSeparator = 10 'Or whatever you need.
        .WriteText linia.Text, 0
        .SaveToFile "c:\temp\enviomail_Ctrl.xml"
        .Close
    End With
   'Open "c:\temp\enviomail.xml" For Output As #2
   'Print #2, linia.Text
   'Close #2
   
   r = Shell(llegir_ini("General", "camillistats", "enviarservidor.ini") + "\SwithMail.exe /s /l 'c:\temp\registreemail_Ctrl.txt' /x 'c:\temp\enviomail_Ctrl.xml'", vbHide)
   
End Function

Function substituirtextbox(cadena As TextBox, buscar As String, canviar As String) As String
   Dim comença As Integer
   Dim acaba As Integer
   If buscar = canviar Then GoTo fi
   While InStr(1, cadena, buscar) > 0
    comença = InStr(1, cadena, buscar) - 1
    If comença < 1 Then substituirtextbox = cadena: Exit Function
    acaba = comença + Len(buscar) + 1
    cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   Wend
fi:
   substituirtextbox = cadena
  ' MsgBox cadena
End Function
