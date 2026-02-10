VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviament d'Incidències"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "enviomails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4080
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Caption         =   "Processar Escanejats"
      Height          =   450
      Left            =   2775
      TabIndex        =   34
      Top             =   885
      Width           =   1800
   End
   Begin VB.Data bobines 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4590
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1470
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data bobinesent 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4545
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   930
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton botoensenyarpacking 
      Caption         =   "Command11"
      Height          =   195
      Left            =   4545
      TabIndex        =   33
      Top             =   1230
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.FileListBox File3 
      Height          =   285
      Left            =   345
      TabIndex        =   32
      Top             =   1245
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   315
      TabIndex        =   31
      Top             =   1455
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.FileListBox File2 
      Height          =   285
      Left            =   135
      TabIndex        =   30
      Top             =   1350
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   15
      TabIndex        =   29
      Top             =   1440
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Log"
      Height          =   270
      Left            =   3990
      TabIndex        =   27
      Top             =   300
      Width           =   630
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PDF_secc"
      Height          =   345
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   375
      Width           =   930
   End
   Begin VB.CheckBox checkfertasquesdemitjanit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Fer ara tasques mitjanit."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2550
      TabIndex        =   20
      Top             =   1425
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4005
      TabIndex        =   23
      ToolTipText     =   "Tancar aplicació"
      Top             =   0
      Width           =   600
   End
   Begin VB.CheckBox sidrivenolocal 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Passar comandes a Drive peró no copiar a Local"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   285
      TabIndex        =   22
      Top             =   3375
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.CommandButton Command7 
      Height          =   600
      Left            =   3930
      Picture         =   "enviomails.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Spooler de Mails"
      Top             =   3045
      Width           =   675
   End
   Begin VB.CheckBox exportarpdfapng 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Exportar PDF a PNG"
      Height          =   225
      Left            =   1155
      TabIndex        =   16
      Top             =   435
      Width           =   2265
   End
   Begin VB.Timer Timercadaminut 
      Interval        =   60000
      Left            =   2085
      Top             =   1230
   End
   Begin VB.CheckBox checkplanificaciotmp 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Generar temporal planificació (de 6h a 18h)"
      Height          =   225
      Left            =   1170
      TabIndex        =   12
      Top             =   645
      Width           =   3375
   End
   Begin VB.Timer TimerSAP 
      Interval        =   5000
      Left            =   1815
      Top             =   750
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Sincronització SAP"
      Height          =   345
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1650
      Width           =   1980
   End
   Begin VB.CommandButton Command5 
      Caption         =   "no serveix(borrar)"
      Height          =   345
      Left            =   2955
      TabIndex        =   10
      Top             =   795
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Min _"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3375
      TabIndex        =   2
      ToolTipText     =   "Minimitzar a la barra de tasques."
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CheckBox exportarauto 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Exportar comandes auto.."
      Height          =   225
      Left            =   1155
      TabIndex        =   9
      Top             =   210
      Width           =   2535
   End
   Begin VB.CheckBox noenviar 
      BackColor       =   &H00C0C0FF&
      Caption         =   "No enviar incidències"
      Height          =   225
      Left            =   1155
      TabIndex        =   8
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton bexportarcomandes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exportar comandes"
      Height          =   510
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   735
      Width           =   930
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Test d'enviament"
      Height          =   270
      Left            =   30
      TabIndex        =   5
      Top             =   1680
      Width           =   2070
   End
   Begin VB.TextBox linia 
      Height          =   285
      Left            =   165
      TabIndex        =   4
      Top             =   1710
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comprovar"
      Height          =   330
      Left            =   2925
      TabIndex        =   1
      Top             =   2325
      Width           =   1770
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   495
      Top             =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Opcions"
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   885
   End
   Begin VB.Label vstatus 
      Caption         =   "Label1"
      Height          =   345
      Left            =   165
      TabIndex        =   28
      Top             =   3750
      Width           =   4260
   End
   Begin VB.Label eterrorservidorsap 
      BackStyle       =   0  'Transparent
      Caption         =   "Error Servidor SAP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   765
      Left            =   30
      TabIndex        =   26
      Top             =   2550
      Visible         =   0   'False
      Width           =   4410
   End
   Begin VB.Label etrutaetiquetesbobina 
      BackColor       =   &H00DADAFE&
      Height          =   195
      Left            =   0
      TabIndex        =   24
      Top             =   2805
      Width           =   4230
   End
   Begin VB.Label etpujantadrive 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   60
      TabIndex        =   19
      Top             =   2985
      Width           =   4305
   End
   Begin VB.Label etrutacomandes 
      BackColor       =   &H00DADAFE&
      Height          =   195
      Left            =   0
      TabIndex        =   18
      Top             =   2610
      Width           =   4230
   End
   Begin VB.Label etrutaescanercomandes 
      BackColor       =   &H00DADAFE&
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   17
      Top             =   2415
      Width           =   4230
   End
   Begin VB.Label etnointernet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   60
      TabIndex        =   15
      Top             =   3135
      Width           =   4185
   End
   Begin VB.Label etrutaescaneralbarans 
      BackColor       =   &H009196FB&
      Height          =   195
      Left            =   0
      TabIndex        =   14
      Top             =   2025
      Width           =   4230
   End
   Begin VB.Label etrutaalbarans 
      BackColor       =   &H009196FB&
      Height          =   195
      Left            =   0
      TabIndex        =   13
      Top             =   2220
      Width           =   4230
   End
   Begin VB.Label etexportar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   1335
      Width           =   3075
   End
   Begin VB.Label estat 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1125
      TabIndex        =   3
      Top             =   870
      Width           =   3405
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 Private Const LOCALE_IDIGITS = &H11
 Private Const LOCALE_USER_DEFAULT = &H400
Private Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long

Private Declare Function inet_addr Lib "wsock32" _
   (ByVal s As String) As Long

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
   
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
    
Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type
    
Private Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long 'formerly integer
   'Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

    
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
Private Const LOCALE_SDECIMAL = &HE

Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long

Public Function GetLocaleDecimalSep() As String
Dim strBuffer As String

strBuffer = String(255, " ")

GetLocaleInfo GetUserDefaultLCID, LOCALE_SDECIMAL, strBuffer, 255

GetLocaleDecimalSep = Trim(substituirtot(strBuffer, Chr(0), ""))

End Function

Public Sub PonerConfgRegional(lngTipo As Long, strNuevoValor As String)
    Dim intRetorno As Integer
    

    intRetorno = SetLocaleInfo(LOCALE_USER_DEFAULT, lngTipo, strNuevoValor)

    Exit Sub
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



Private Sub bexportarcomandes_Click()
   If bexportarcomandes.Tag <> "exportant" Then
      If Not impresorapdfpredeterminada Then MsgBox "No hi ha impresora pdfcreator per predeterminar." + Chr(10) + "Ha d'estar instal.lada i configurada per guardar automaticament a la c:\temp\exportar ", vbCritical, "Error": Exit Sub
      bexportarcomandes.BackColor = QBColor(12)
      DoEvents
      exportarlescomandes
      
        Else: bexportarcomandes.Tag = "": bexportarcomandes.BackColor = QBColor(15)
   End If
End Sub

Sub exportarlescomandes()
    Dim dbcomandes As Database
    Dim rst As Recordset
    bexportarcomandes.Tag = "exportant"
    impresorapdfpredeterminada
    Set dbtmp = OpenDatabase(cami)
    If existeix("c:\ordprog.ini") Then FileCopy "c:\ordprog.ini", "c:\ordprog2.ini": eliminar_fitxer "c:\ordprog.ini"
    Set dbcomandes = OpenDatabase(cami)
    Set rst = dbcomandes.OpenRecordset("select * from comandes where proximaseccio='T' and producte<>'PC' and producte<>'PC2' and producte<>'PCP'  order by comanda asc")
    rst.MoveLast
    rst.MoveFirst
    While Not rst.EOF And bexportarcomandes.Tag = "exportant"
      etexportar = "Exportant... Comanda: " + atrim(rst!comanda)
      DoEvents
      If Not existeix(llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini")) Then bexportarcomandes.Tag = "": etexportar = "No es pot accedir a la ruta desti.": Exit Sub
      If Not buscarlacomanda(rst!comanda) Then
          exportarlacomanda rst!comanda
          
      End If
      rst.MoveNext
    Wend
    bexportarcomandes.Tag = ""
    etexportar.Caption = "Procès acabat"
End Sub
Sub exportarlacomanda(numc As String)
   Dim carpetadesti As String
   Dim horacomençament As Date
   Dim carpetaprincipal As String
   If existeix("c:\ordprog.ini") Then FileCopy "c:\ordprog.ini", "c:\ordprog2.ini": eliminar_fitxer "c:\ordprog.ini"
   horacomençament = Now
   escriure_ini "baixes", "imprimircomanda", numc, "comandes.ini"
   escriure_ini "General", "exportant", "1", fitxerini
   ShellAndWait llegir_ini("General", "rutallistats", "comandes.ini") + "comandes.exe comandes.ini exportar", vbNormalFocus
   'wait (1)
   
   While llegir_ini("General", "exportant", "comandes.ini") = "1" And DateDiff("n", horacomençament, Now) < 3
     DoEvents
   Wend
   carpetadesti = llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini")
   If Not existeix(carpetadesti + "\cache_Fabricacio") Then MkDir carpetadesti + "\cache_Fabricacio"
   carpetadesti = carpetadesti + "\cache_Fabricacio"
   carpetaprincipal = "Les_" + atrim(atrim(Int(cadbl(numc) / 1000)) + "000")
   If Not existeix(carpetadesti + "\" + carpetaprincipal) Then MkDir carpetadesti + "\" + carpetaprincipal
   If Not existeix(carpetadesti + "\" + carpetaprincipal + "\" + atrim(numc)) Then MkDir carpetadesti + "\" + carpetaprincipal + "\" + atrim(numc)
   wait 5
   If hihaalgualtemporal Then
     copiarfitxersaldesti numc, carpetadesti + "\" + carpetaprincipal + "\" + atrim(numc)
     eliminar_fitxer "c:\temp\exportar\*.*"
     dbtmp.Execute "insert into comandesexportades (comanda,data) values (" + atrim(numc) + ",now)"
   End If
End Sub
Sub copiarfitxersaldesti(numc As String, carpetadesti As String)
   Dim d As String
   Dim contador As Byte
   d = Dir("c:\temp\exportar\*.pdf")
   contador = 1
   While d <> ""
   'llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\" + numc +
     FileCopy "c:\temp\exportar\" + d, carpetadesti + "\" + numc + "-" + atrim(contador) + ".pdf"
     contador = contador + 1
     d = Dir
   Wend
End Sub
Function hihaalgualtemporal() As Boolean
   Dim d As String
   d = Dir("c:\temp\exportar\*.pdf")
   If d <> "" Then hihaalgualtemporal = True
End Function

Function buscarlacomanda(numc As String) As Boolean
  'Dim d As String
'  d = Dir(llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\*.*", vbDirectory)
'  While d <> "" And d <> numc
'     d = Dir
'  Wend
 ' If d = numc Then buscarlacomanda = True
 Dim rst As Recordset
 buscarlacomanda = False
 Set rst = dbtmp.OpenRecordset("select * from comandesexportades where comanda=" + atrim(numc))
 If Not rst.EOF Then buscarlacomanda = True
End Function

Private Sub checkplanificaciotmp_Click()
    escriure_ini "general", "planificaciotmp", atrim(checkplanificaciotmp.Value), "enviarservidor.ini"
End Sub

Private Sub checkplanificaciotmp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then Shell llegir_ini("General", "rutallistats", "comandes.ini") + "planificacio.exe GENERARFITXERTEMPORAL", vbNormalFocus
End Sub

Private Sub Command1_Click()
  comprovar_incidencies
End Sub

Private Sub Command10_Click()
   obrir_document "c:\temp\Log_EnviarMails_servidor.txt"
End Sub

Private Sub Command11_Click()
 obrir_tancar_taules True
 'actualitzar_proveidors
 'assignarCQs
 organitzar_fitxers_Escanejats_Expedicions
End Sub

Private Sub Command3_Click()
 Form_Initialize
End Sub

Private Sub Command4_Click()
Dim emails As String
Dim destinatari As String
Dim vresp As Boolean
'Dim i As Byte
'  For i = 1 To 10
    Dim vusuari As String
    Dim vpassword As String
    vusuari = llegir_ini("dadesservidor", "usrsmtp", "enviarservidor.ini")
    vpassword = llegir_ini("dadesservidor", "passsmtp", "enviarservidor.ini")
    destinatari = InputBox("Escriu a quin correu vols enviar la prova.", "Enviament de prova", "miquel.inplacsa@gmail.com")
    If InStr(1, destinatari, "@") > 0 Then
      'enviaremail destinatari, "Envio de prova d'Incidències", "Cos del missatge. Una prova d 'enviament d'incidencies."
      vresp = enviaremailswitchmail("incidenciesinplacsa@gmail.com", destinatari, "Envio de prova d'Incidències", "Cos del missatge. Una prova d 'enviament d'incidencies.", "c:\temp\cosmissatge.txt", , vusuari, vpassword)
      If vresp Then
         MsgBox "Missatges enviats a: " + destinatari
         Else: MsgBox "Error enviant el email"
      End If
    End If
End Sub
Function possar_caracters_html(ByVal v As String) As String
    'v = substituirtot(v, "&", "&amp;")
    v = substituirtot(v, "À", "&Agrave;")
    v = substituirtot(v, "Á", "&Aacute;")
    v = substituirtot(v, "Â", "&Acirc;")
    v = substituirtot(v, "Ã", "&Atilde;")
    v = substituirtot(v, "Ä", "&Auml;")
    v = substituirtot(v, "È", "&Egrave;")
    v = substituirtot(v, "É", "&Eacute;")
    v = substituirtot(v, "Ê", "&Ecirc;")
    v = substituirtot(v, "Ë", "&Euml;")
    v = substituirtot(v, "Ì", "&Igrave;")
    v = substituirtot(v, "Í", "&Iacute;")
    v = substituirtot(v, "Î", "&Icirc;")
    v = substituirtot(v, "Ï", "&Iuml;")
    v = substituirtot(v, "Ò", "&Ograve;")
    v = substituirtot(v, "Ó", "&Oacute;")
    v = substituirtot(v, "Ô", "&Ocirc;")
    v = substituirtot(v, "Õ", "&Otilde;")
    v = substituirtot(v, "Ö", "&Ouml;")
    v = substituirtot(v, "Ù", "&Ugrave;")
    v = substituirtot(v, "Ú", "&Uacute;")
    v = substituirtot(v, "Û", "&Ucirc;")
    v = substituirtot(v, "Ü", "&Uuml;")
    v = substituirtot(v, "Š", "&Scaron;")
    v = substituirtot(v, "Ý", "&Yacute;")
    v = substituirtot(v, "Ÿ", "&Yuml;")
    v = substituirtot(v, "à", "&agrave;")
    v = substituirtot(v, "á", "&aacute;")
    v = substituirtot(v, "â", "&acirc;")
    v = substituirtot(v, "ã", "&atilde;")
    v = substituirtot(v, "ä", "&auml;")
    v = substituirtot(v, "è", "&egrave;")
    v = substituirtot(v, "é", "&eacute;")
    v = substituirtot(v, "ê", "&ecirc;")
    v = substituirtot(v, "ë", "&euml;")
    v = substituirtot(v, "ì", "&igrave;")
    v = substituirtot(v, "í", "&iacute;")
    v = substituirtot(v, "î", "&icirc;")
    v = substituirtot(v, "ï", "&iuml;")
    v = substituirtot(v, "ò", "&ograve;")
    v = substituirtot(v, "ó", "&oacute;")
    v = substituirtot(v, "ô", "&ocirc;")
    v = substituirtot(v, "õ", "&otilde;")
    v = substituirtot(v, "ö", "&ouml;")
    v = substituirtot(v, "ù", "&ugrave;")
    v = substituirtot(v, "ú", "&uacute;")
    v = substituirtot(v, "û", "&ucirc;")
    v = substituirtot(v, "ü", "&uuml;")
    v = substituirtot(v, "š", "&scaron;")
    v = substituirtot(v, "ý", "&yacute;")
    v = substituirtot(v, "ÿ", "&yuml;")
    v = substituirtot(v, "ç", "&#231;")
    v = substituirtot(v, "Ç", "&#199;")
    v = substituirtot(v, "ñ", "&#241;")
    v = substituirtot(v, "Ñ", "&#209,")

    
    
    possar_caracters_html = v
End Function

Function enviaremailswitchmail(vremitent As String, sSendTo As String, sSubject As String, sText As String, Optional adjunt As String, Optional vidavis As Long, Optional vusuari As String, Optional vcontrasenya As String, Optional adjunt2 As String, Optional adjunt3 As String) As Boolean
  Dim usuarim As String
  Dim contrasenyam As String
  Dim destinatari As String
  Dim instream  As Object
  Dim vcont As Double
  
   enviaremailswitchmail = False
   usuarim = vusuari
   contrasenyam = vcontrasenya
   If llegir_ini("General", "camillistats", "enviarservidor.ini") = "{[}]" Then escriure_ini "General", "camillistats", "\\serverprodu\dades\progcomandes\aplicacio\", "enviarservidor.ini"
   Open llegir_ini("General", "camillistats", "enviarservidor.ini") + "\enviomailswithmail.xml" For Input As #1
   linia.Text = Input(LOF(1), #1)
   Close #1
   escriure_log "Enviar_pendents -Switch 1.", "c:\temp\Log_EnviarMails_servidor.txt"
   destinatari = llegir_ini("destinataris", sSendTo, "enviarservidor.ini")
   If destinatari = "{[}]" Then destinatari = sSendTo
   sSendTo = destinatari
   linia = Mid(linia, 4)
   
   'poso els valors HTML als caràcters
   'sSubject = possar_caracters_html(sSubject)
   'adjunt = possar_caracters_html(adjunt)
   'sText = possar_caracters_html(sText)
   
   substituirtextbox linia, "#remitent#", vremitent
   substituirtextbox linia, "#destinatari#", sSendTo
   substituirtextbox linia, "#asumpte#", sSubject
   'substituir "#cosdelmisatge#", "CreateObject(""Scripting.FileSystemObject"").OpenTextFile(""C:\temp\cosmissatge.txt"", 1).ReadAll"
   'adjunt = substituirtot(adjunt, "|", "</AttachmentPath>" + vbNewLine + "<AttachmentPath>")
   escriure_log "Enviar_pendents -Switch 2.", "c:\temp\Log_EnviarMails_servidor.txt"
   If existeix(adjunt2) Then adjunt = adjunt + "</AttachmentPath>" + vbNewLine + "<AttachmentPath>" + adjunt2
   If existeix(adjunt3) Then adjunt = adjunt + "</AttachmentPath>" + vbNewLine + "<AttachmentPath>" + adjunt3
   
   substituirtextbox linia, "#fitxeradjunt#", adjunt
   If InStr(1, LCase(sText), "\cosmissatge.txt") > 0 Then
       substituirtextbox linia, "#fitxertxtcosmissatge#", sText
        Else: substituirtextbox linia, "#fitxertxtcosmissatge#", "c:\temp\cosmissatge.txt"
   End If
   'If adjunt = "" Then substituir "objMessage.AddAttachment """"", ""
   substituirtextbox linia, "#usuarigmail#", usuarim
   substituirtextbox linia, "#contrasenyagmail#", contrasenyam
   escriure_log "Enviar_pendents -Switch 3.", "c:\temp\Log_EnviarMails_servidor.txt"
   If InStr(1, LCase(sText), "\cosmissatge.txt") = 0 Then
        Open "c:\temp\cosmissatge.txt" For Output As #2
        Print #2, sText
        passarliniesdavisosalfitxertxt vidavis
        Close #2
   End If
   escriure_log "Enviar_pendents -Switch 4.", "c:\temp\Log_EnviarMails_servidor.txt"
    On Error Resume Next
    eliminar_fitxer "c:\temp\enviomail.xml"
    eliminar_fitxer "c:\temp\registreemail.txt"
    On Error GoTo 0
    escriure_log "Enviar_pendents -Switch 5.", "c:\temp\Log_EnviarMails_servidor.txt"
    Set instream = CreateObject("ADODB.Stream")
    With instream
        .Open
        .Type = 2
        .Charset = "utf-8"
        .LineSeparator = 10 'Or whatever you need.
        .WriteText linia.Text, 0
        .SaveToFile "c:\temp\enviomail.xml"
        .Close
    End With
   'Open "c:\temp\enviomail.xml" For Output As #2
   'Print #2, linia.Text
   'Close #2
   escriure_log "Enviar_pendents -Switch INICI executa el programa.", "c:\temp\Log_EnviarMails_servidor.txt"
   r = Shell(llegir_ini("General", "camillistats", "enviarservidor.ini") + "\SwithMail.exe /s /l 'c:\temp\registreemail.txt' /x 'c:\temp\enviomail.xml'", vbHide)
   escriure_log "Enviar_pendents -Switch FI executa el programa .", "c:\temp\Log_EnviarMails_servidor.txt"
   'espero que es generi el fitxer de registre i com a màxim esperar 10 segons
   vcont = 1
   While Not existeix("c:\temp\registreemail.txt") And vcont < 60
     wait 1
     vcont = vcont + 1
   Wend
   escriure_log "Enviar_pendents -Switch 6.", "c:\temp\Log_EnviarMails_servidor.txt"
   If existeix("c:\temp\registreemail.txt") Then enviaremailswitchmail = revisarsioklenviament("c:\temp\registreemail.txt")
   escriure_log "Enviar_pendents -Switch 7 FI.", "c:\temp\Log_EnviarMails_servidor.txt"
End Function

Function revisarsioklenviament(vlog As String) As Boolean
   Dim vlinia As String
   Open vlog For Input As #3
   vlinia = Input(LOF(3), #3)
   Close #3
   If InStr(1, vlinia, "- Success -") > 0 Then revisarsioklenviament = True
End Function

Private Sub Command5_Click()
   Dim r As String
   Dim carpetaprincipal As String
   Dim carpetadesti As String
   carpetadesti = llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini")
   
   r = Dir(carpetadesti + "\*.*", vbDirectory)
   While r <> ""
     If cadbl(r) > 0 Then
        carpetaprincipal = "Les_" + atrim(atrim(Int(cadbl(r) / 1000)) + "000")
        If Not existeix(carpetadesti + "\" + carpetaprincipal) Then MkDir carpetadesti + "\" + carpetaprincipal
        'If Not existeix(carpetadesti + "\" + carpetaprincipal + "\" + r) Then MkDir carpetadesti + "\" + carpetaprincipal + "\" + r
        Copiar_Fitxer carpetadesti + "\" + r, carpetadesti + "\" + carpetaprincipal + "\"
        'Kill carpetadesti + "\" + r + "\*.*"
        'Shell "c:\windows\system32\cmd.exe /c rd " + carpetadesti + "\" + r + " /q", vbMaximizedFocus
     End If
     r = Dir
     Me.Caption = r
     DoEvents
   Wend
   
End Sub

Private Sub Command6_Click()
  If obrir_tancar_taules(True) Then
    sincronitzar_taulesmestra
    revisar_clients_donatsdebaixa_ambcomandescirculant
    revisar_clientsalbaransclixes
  End If
  'obrir_tancar_taules False
  Me.Caption = "Enviament d'Incidències": DoEvents
End Sub
Function noexisteixelclientasap(vclient As Double, vempresa As String) As Boolean
   Dim rst As Recordset
   Set dbcomandes = OpenDatabase(cami)
   vempresa = UCase(vempresa)
   If vempresa = "I" Then
      vempresa = ""
        Else: vempresa = "PLASEL"
   End If
   Set rst = dbcomandes.OpenRecordset("select * from clients_codisSAP" + vempresa + " where codiSAP=" + atrim(vclient))
   If rst.EOF Then noexisteixelclientasap = True
   Set rst = Nothing
End Function
Sub revisar_clientsalbaransclixes()
   Dim rst As Recordset
   Dim rstp As Recordset
   Dim vlogerrors As String
   Dim vnomempresa As String
   Dim vcontrol As Integer
        Me.Caption = "Clients albarans clixes " + atrim(cadbl(vcontrol)): vcontrol = cadbl(vcontrol) + 1: DoEvents
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
        Me.Caption = "Clients albarans clixes " + atrim(cadbl(vcontrol)): vcontrol = cadbl(vcontrol) + 1: DoEvents
   Set rst = dbclixes.OpenRecordset("SELECT First(Modificacions.empresafacturadora) AS Pempresafacturadora, First(Modificacions.codiclientfactclixes) AS Pcodiclientfactclixes, Clixes_albarans.id_treball as numtreball, Clixes_albarans.ordremodificacio as numordre FROM Clixes_albarans INNER JOIN Modificacions ON (Clixes_albarans.id_treball = Modificacions.id_treball) AND (Clixes_albarans.ordremodificacio = Modificacions.ordre) Where (((Clixes_albarans.facturat) = False)) GROUP BY Clixes_albarans.id_treball, Clixes_albarans.ordremodificacio HAVING (((First(Modificacions.codiclientfactclixes))>'0' Or (First(Modificacions.codiclientfactclixes)) Is Not Null));")
        Me.Caption = "Clients albarans clixes " + atrim(cadbl(vcontrol)): vcontrol = cadbl(vcontrol) + 1: DoEvents
   While Not rst.EOF
      Set rstp = dbclixes.OpenRecordset("select preu from pressupostos where id_treball=" + atrim(rst!numtreball) + " and ordremodificacio=" + atrim(rst!numordre))
      Me.Caption = "Clients albarans clixes (while)": DoEvents
      If rstp.EOF Then GoTo proxim
      If cadbl(rstp!preu) < 1 Then GoTo proxim
      vnomempresa = IIf(atrim(rst!Pempresafacturadora) = "P", "PLASEL", "INPLACSA")
      If noexisteixelclientasap(cadbl(rst!Pcodiclientfactclixes), atrim(rst!Pempresafacturadora)) Then
         vlogerrors = vlogerrors + " El client " + atrim(cadbl(rst!Pcodiclientfactclixes)) + " de l'empresa " + vnomempresa + " no existeix al SAP s'hauria de donar d'alta. (" + atrim(rst!numtreball) + "/" + atrim(rst!numordre) + ")" + Chr(10)
      End If
proxim:
      rst.MoveNext
   Wend
        Me.Caption = "Clients albarans clixes " + atrim(cadbl(vcontrol)): vcontrol = cadbl(vcontrol) + 1: DoEvents
   If vlogerrors <> "" Then enviaremail "incidenciesillistatsSAPcomptabilitat", "Codis de client per actualitzar a SAP", vlogerrors
        Me.Caption = "Clients albarans clixes " + atrim(cadbl(vcontrol)): vcontrol = cadbl(vcontrol) + 1: DoEvents
  ' Set dbclixes = Nothing
End Sub

Private Sub Command7_Click()
  formspooler.Show
End Sub

Private Sub Command8_Click()
 ' -- cuando descargamos el form removemos el Icono del systray
    'If MsgBox("Segur que vols tancar, es deixarà d'enviar incidencies i altres operacions del programa de producció.", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Cancel = 1: Exit Sub
    If existeix("c:\ordprog2.ini") Then FileCopy "c:\ordprog2.ini", "c:\ordprog.ini": eliminar_fitxer "c:\ordprog2.ini"
    Unload formspooler
    RemoverSystray
    db.Close
    End
End Sub


Private Sub Command9_Click()
   comprovar_PDF_baixesseccions
End Sub

Private Sub exportarauto_Click()
   escriure_ini "general", "exportarauto", atrim(exportarauto.Value), "enviarservidor.ini"
End Sub

Private Sub exportarpdfapng_Click()
  escriure_ini "general", "exportarpdfapng", atrim(exportarpdfapng.Value), "enviarservidor.ini"
End Sub


Sub calcular_credit_tots_clients()
  'calcular_credit_delclient
  Dim vrisc As TipusVrisc
  Dim rstc As Recordset
  Dim rst As Recordset
  Dim vmsgCREDITSUPERAT As String
  Dim vcreditsuperat As Double
  Dim vcodiclient As String
  Dim rstf As Recordset
  
  obrir_tancar_taules True
  Set dbtmp = dbcomandes
  'Set rst = dbcomandes.OpenRecordset("SELECT  distinct clients_codisSAP.codiSAP, clients_codisSAP.nomclient FROM comandes RIGHT JOIN (comandes_extres RIGHT JOIN clients_codisSAP ON comandes_extres.codicomptable = clients_codisSAP.codiSAP) ON comandes.comanda = comandes_extres.comanda WHERE (((Year([datacomanda]))>Year(Now())-2) AND( ((comandes.proximaseccio)<>'T') AND ((comandes.producte)<>'PC' And (comandes.producte)<>'PC2' And (comandes.producte)<>'PCP')));")
  'Set rst = dbcomandes.OpenRecordset("SELECT  distinct clients_codisSAP.codiSAP, clients_codisSAP.nomclient FROM comandes RIGHT JOIN (comandes_extres RIGHT JOIN clients_codisSAP ON comandes_extres.codicomptable = clients_codisSAP.codiSAP) ON comandes.comanda = comandes_extres.comanda WHERE CODISAP=43000007445 ;")
  Set rst = dbcomandes.OpenRecordset("SELECT DISTINCT  clients_codisSAP.codiSAP, clients_codisSAP.nomclient FROM comandes RIGHT JOIN (comandes_extres RIGHT JOIN clients_codisSAP ON comandes_extres.codicomptable = clients_codisSAP.codiSAP) ON comandes.comanda = comandes_extres.comanda WHERE (((Year([datacomanda]))>Year(Now())-2) AND ((comandes.producte)<>'PC' And (comandes.producte)<>'PC2' And (comandes.producte)<>'PCP') AND ((DateDiff('d',[comandes_extres].[dataentrega],Now()))<30));")
  Set rstc = dbcomandes.OpenRecordset("Select * from clients_codisSAP")
  rst.MoveLast
  rst.MoveFirst
  While Not rst.EOF
    'If rst!codisap = 43000007445# Then Stop
    vcreditsuperat = 0
    calcular_credit_delclient rst!codisap, vrisc
    rstc.FindFirst "codisap=" + atrim(rst!codisap)
'    If rst!codisap = 43000006998# Then Stop
    If Not rstc.NoMatch Then
        rstc.Edit
        rstc!creditsap = cadbl(vrisc.creditsap)
        rstc!valordiferencial = Redondejar(cadbl(vrisc.valordiferencial), 0)
        rstc!creditgastatsap = cadbl(vrisc.creditgastatsap)
        rstc!valorestoc = cadbl(vrisc.valorestoc)
        rstc!valorpendent = cadbl(vrisc.valorpendent)
        rstc!valorproduccio = cadbl(vrisc.valorproduccio)
        rstc!valordelsclixes = cadbl(vrisc.valordelsclixes)
        rstc!valoralbaranspendentsSAP = cadbl(vrisc.valoralbaranspendentsSAP)
        rstc.Update
        vcreditsuperat = rstc!creditsap - rstc!creditgastatsap - rstc!valoralbaranspendentsSAP
        If vcreditsuperat < 0 Then
            Set rstf = dbcomandes.OpenRecordset("select * from capcaleraalbara where codiclient=" + atrim(rst!codisap) + " and dataalbara=#" + Format(DateAdd("d", -1, Now), "mm/dd/yy") + "#")
'            Clipboard.Clear
'            Clipboard.SetText "select * from capcaleraalbara where codiclient=" + atrim(rst!codisap) + " and dataalbara=#" + Format(DateAdd("d", -1, Now), "mm/dd/yy") + "#"
            If Not rstf.EOF Then
                    vmsgCREDITSUPERAT = vmsgCREDITSUPERAT + justificar(atrim(rst!codisap) + "-" + Mid(vrisc.nomdelclient, 1, 40), 45, "E", "_") + atrim(vcreditsuperat) + "€" + vbNewLine
            End If
        End If
    End If
    Me.Caption = "Calculant credit: " + Trim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount)
    DoEvents
    rst.MoveNext
  Wend
  Set rstf = Nothing
  Set rst = Nothing
  Set rstc = Nothing
  If vmsgCREDITSUPERAT <> "" And Hour(Now) = 8 Then enviaremail "ainoaduch@inplacsa.com", "Relació de CLIENTS AMB CREDIT SUPERAT " + Format(Now, "dd/mm/yy"), "LLista d'incidències: " + vbNewLine + vbNewLine + vmsgCREDITSUPERAT
  Me.Caption = "Enviament d'Incidències"
End Sub

Private Sub Form_Activate()
 PonerConfgRegional LOCALE_IDIGITS, "4"
 
 
 obrir_tancar_taules True
 comprovar_albarans_tintes_nous
 'revisar_clientsalbaransclixes
' informe_doblefirma_disposiciomaterials
'revisar_clients_donatsdebaixa_ambcomandescirculant
'  comprovar_firmes_pendents_PVP
 'guardar_doc_a_docx "c:\temp\docx\Imp00300.doc"
 'obrir_tancar_taules True
 'actualitzar_proveidors
 'assignarCQs
' organitzar_fitxers_Escanejats_Expedicions
 'llistat_comprespendents_1rdemes
 
' revisar_clients_donatsdebaixa_ambcomandescirculant
'EnviarPaletsSenseImpost

' Command6_Click
'Dim vrisc As TipusVrisc

'obrir_tancar_taules True
'Set dbtmp = dbcomandes
 'actualitzar_credit_clients "INPLACSA"
'End

 ' revisarPackinglistDescuadrats
  'diferencies_entre_comandes 211350, 211333
'obrir_tancar_taules True
'mirarcomandesambrefinplacsanovalidades
'calcular_credit_tots_clients
'actualitzar_credit_clients "INPLACSA"

'End
'comprovar_mesuraPVPvsCLIENT
'End
'GRUPS_revisarsihihaprous_metres_assignats
'obrir_tancar_taules True
'actualitzar_CQ_lots
'actualitzar_clients "inplacsa"
'actualitzar_clients "plasel"
'enviaremail "miquel.inplacsa@gmail.com", "tec", "prova"
'llistat_clixes_tintespendentsderevisar
'  revisarBaixesPDFdelesseccions
'  eterrorservidorsap.Visible = IIf(Not ferPing("servidorsap"), True, False)
  'If Not impresorapdfpredeterminada Then MsgBox "No hi ha impresora pdfcreator per predeterminar." + Chr(10) + "Ha d'estar instal.lada i configurada per guardar automaticament a la c:\temp\exportar ", vbCritical, "Error": Exit Sub
  'comprovar_albarans_tintes_nous
End Sub
Function proximcamp(camps As String) As String
   If camps = "#" Then Exit Function
   proximcamp = Mid(camps, 2, InStr(2, camps, "#") - 2)
   camps = Mid(camps, InStr(2, camps, "#"))
End Function
Function cvavalorcamp(rst As Recordset, nomcamp As String) As String
   Select Case nomcamp
      Case "ampleesq"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "plegatesq"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "solapa"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "espessor"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "micropex"
           cvavalorcamp = atrim((rst.Fields(nomcamp)))
           If cvavalorcamp = "" Then cvavalorcamp = "N"
      Case "oberturaex"
           cvavalorcamp = atrim((rst.Fields(nomcamp)))
           If cvavalorcamp = "" Then cvavalorcamp = "N"
      Case "ampleutil"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "simulteneitatlam"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "tipusadhesiu"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "migelaborat"
           cvavalorcamp = atrim(rst.Fields(nomcamp))
      Case "amplereb"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "simulteneitatreb"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
            Case "migelaboratsol"
           cvavalorcamp = atrim(rst.Fields(nomcamp))
      Case "amplesol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "ampleplegsol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "longitudsol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "solapasol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "fuellebasesol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "fuellebocasol"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "troquel"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "ansa"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
      Case "cinta"
           cvavalorcamp = atrim(cadbl(rst.Fields(nomcamp)))
       Case Else
           cvavalorcamp = atrim(rst.Fields(nomcamp))
   End Select
End Function

Function diferencies_entre_comandes(vcomandanova As Double, vcomandavella As Double) As String
   Dim rstn As Recordset
   Dim rstv As Recordset
   Dim rstmatv As Recordset
   Dim rstmatn As Recordset
   Dim camps As String
   Dim nomcamp As String
   Dim vdiferencia As String
   Dim vruta As String
   Dim rstp As Recordset
   Dim valordelcamp As String
   Dim link1n As Double
   Dim link2n As Double
   Dim link1v As Double
   Dim link2v As Double
   link1v = -1
Set dbcomandes = OpenDatabase(cami)

   If vcomandavella = 0 Or vcomandanova = 0 Then Exit Function
   Set rstn = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(vcomandanova))
   Set rstv = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(vcomandavella))
   If rstn.EOF Or rstv.EOF Then Exit Function
   If link1v = -1 Then
    link1n = rstn!linkcomanda1: link2n = rstn!linkcomanda2
    link1v = rstv!linkcomanda1: link2v = rstv!linkcomanda2
   End If
inici:
   Set rstn = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(vcomandanova))
   Set rstv = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(vcomandavella))
   If rstn.EOF Or rstv.EOF Then Exit Function
   Set rstp = dbcomandes.OpenRecordset("select * from productes where codi='" + atrim(rstn!producte) + "'")
   If Not rstp.EOF Then vruta = rstp!ruta
   If InStr(1, rstn!producte, "PC") > 0 Then
       camps = "#Etubolam#Eampleesq#Eplegatesq#Esolapa#Eespessor#Emicropex#Eoberturaex#Ematerialex#"
      Else
        camps = "#Etubolam#Eampleesq#Eplegatesq#Esolapa#Eespessor#Emicropex#Eoberturaex#Ematerialex#Inumtreball#Lampleutil#Lsimulteneitatlam#Rmigelaborat#Ramplereb#Rsimulteneitatreb"
        camps = camps + "#Smigelaboratsol#Samplesol#Sampleplegsol#Slongitudsol#Ssolapasol#Sfuellebasesol#Sfuellebocasol#Stroquel#Sansa#Scinta#"
   End If
   nomcamp = proximcamp(camps)
   Set rstmatn = dbcomandes.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstn!materialex)))
   While nomcamp <> ""
      If InStr(1, vruta, Mid(nomcamp, 1, 1)) > 0 Then
       nomcamp = Mid(nomcamp, 2)
       valordelcamp = cvavalorcamp(rstn, nomcamp)
       If nomcamp = "numtreball" Then GoTo cont
       If nomcamp = "materialex" Then
           Set rstmatv = dbcomandes.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstv!materialex)))
           If rstmatv.EOF Then vdiferencia = vdiferencia + "[Error de material]": GoTo cont
           If cadbl(rstmatv!familia) <> cadbl(rstmatn!familia) Or cadbl(rstmatv!subfamilia) <> cadbl(rstmatn!subfamilia) Or cadbl(rstmatv!familiacol) <> cadbl(rstmatn!familiacol) Then
             vdiferencia = vdiferencia + "[families MAT diferents]"
           End If
           GoTo cont
       End If
       If atrim(valordelcamp) <> cvavalorcamp(rstv, nomcamp) Then
         vdiferencia = vdiferencia + "[" + nomcamp + " = " + atrim(valordelcamp) + "<>" + cvavalorcamp(rstv, nomcamp) + "]"
       End If
     End If
cont:
      nomcamp = proximcamp(camps)
  Wend
  If vdiferencia <> "" And vcomandanova <> 0 Then diferencies_entre_comandes = diferencies_entre_comandes + vbTab + atrim(vcomandanova) + " " + atrim(rstn!producte) + "->" + vdiferencia + vbNewLine
  vdiferencia = ""
  If vcomandanova <> link1n And vcomandanova <> link2n Then vcomandanova = link1n: vcomandavella = link1v: GoTo inici
  If vcomandanova = link1n Then
      vcomandanova = link2n: vcomandavella = link2v
        Else: vcomandanova = 0
  End If
  If vcomandanova <> 0 Then GoTo inici
   Set rstp = Nothing
   Set rstmatv = Nothing
   Set rstmatn = Nothing
   Set rstn = Nothing
   Set rstv = Nothing
   
End Function





Sub mirarcomandesambrefinplacsanovalidades()
  Dim rst As Recordset
  Dim vmsg As String
  Dim vdif As String
  
  Set dbcomandes = OpenDatabase(cami)
  Set rst = dbcomandes.OpenRecordset("SELECT COMANDES.impressio,comandes.marcailinia,comandes_extres.comandaduplicadade,comandes_extres.comanda, clients.nom, comandes_extres.refinplacsa_validada, comandes_extres.comandaimpresa,comandes.producte FROM comandes_extres LEFT JOIN (comandes LEFT JOIN clients ON comandes.client = clients.codi) ON comandes_extres.comanda = comandes.comanda WHERE comandes_extres.refinplacsa_validada=False AND comandes_extres.comandaimpresa=True;")
  While Not rst.EOF
        If rst!comanda > 0 And InStr(1, atrim(rst!producte), "PC") = 0 Then
            If atrim(rst!impressio) = "N" Then
               vdif = "COMANDA NOVA."
                Else: vdif = diferencies_entre_comandes(cadbl(rst!comanda), cadbl(rst!comandaduplicadade))
            End If
            If vdif <> "" Then
                 vmsg = vmsg + atrim(rst!comanda) + " - " + atrim(rst!nom) + " [" + atrim(rst!marcailinia) + "]" + vbNewLine
                 vmsg = vmsg + " Diferencies: " + vbNewLine + vdif + vbNewLine + vbNewLine
            End If
        End If
        rst.MoveNext
  Wend
  If vmsg <> "" Then
     vmsg = "Comandes amb referencia inplacsa nova pendents de revisió:" + vbNewLine + vbNewLine + vmsg
     enviaremail "incidenciesdePVP", "Relacio de comandes pendents de revisar RefInplacsa. " + Format(Now, "dd/mm/yy hh:nn") + "", vmsg
  End If
  Set rst = Nothing
End Sub
Sub mirarcomandesenproducciosensepreu()
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim vlinia As String
  Set dbcomandes = OpenDatabase(cami)
  If existeix("c:\temp\llistatcomandesproducciosensepreu.txt") Then borrar_fitxer "c:\temp\llistatcomandesproducciosensepreu.txt"
  Set rst = dbcomandes.OpenRecordset("SELECT comandes.*, clients.nom as nomclient,clients.grupdeclient FROM comandes LEFT JOIN clients ON comandes.client = clients.codi Where comanda>147000 and proximaseccio<>'T' and proximaseccio<>'E' and pvp=0")
  If rst.EOF Then Exit Sub
  Open "c:\temp\llistatcomandesproducciosensepreu.txt" For Output As 1
  While Not rst.EOF
        If InStr(1, atrim(rst!producte), "PC") > 0 Then GoTo fi
        If atrim(rst!proximaseccio) = "T" Then GoTo fi
        If atrim(rst!numpressupost) = "PROVA" Then GoTo fi
        If atrim(rst!grupdeclient) = "INPLACSA" Then GoTo fi
        If InStr(1, atrim(rst!marcailinia), "FINGERPRINT") > 0 Then GoTo fi
        If atrim(rst!grupdeclient) <> "ARDO" Then
              If cadbl(rst!pvp) = 0 Then
                 Set rst2 = dbcomandes.OpenRecordset("select * from comandeS_observacioPVP where trim(observacio)<>'' and comanda=" + atrim(cadbl(rst!comanda)))
                 vlinia = "Comanda: " + atrim(rst!comanda) + vbNewLine + atrim(rst!client) + " - " + atrim(rst!nomclient) + vbNewLine + "Ref.Client: " + atrim(rst!refclient) + vbNewLine + "Texte Imp.:" + atrim(rst!marcailinia) + vbNewLine
                 If Not rst2.EOF Then vlinia = vlinia + atrim(rst2!observacio) + vbNewLine + vbNewLine
                 Print #1, vlinia
              End If
        End If
fi:
        rst.MoveNext
  Wend
  Close 1
  Set rst = Nothing
  Set rst2 = Nothing
  If existeix("c:\temp\llistatcomandesproducciosensepreu.txt") Then
        'incidenciesdePVP
      enviaremail "incidenciesdePVP", "Relacio de comandes en producció sense PVP.", "c:\temp\llistatcomandesproducciosensepreu.txt"
  End If


End Sub
Sub comprovar_PDF_baixesseccions()
  Dim rst As Recordset
  Dim vcarpetadesti As String
  Dim vmsg As String
  Dim dbbaixes As Database
  Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
  'muntadores
  vmsg = vmsg + Chr(13) + Chr(10) + "MUNTADORES" + Chr(13) + Chr(10)
  Set rst = dbbaixes.OpenRecordset("select comanda from muntadores where operari1<>0 and datafi<>null and year(datainici)>2020")
  rst.MoveLast
  rst.MoveFirst
  While Not rst.EOF
    'Me.Caption = atrim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount)
    vcarpetadesti = "\\ord_copies\comandespdf\Les_" + Mid(atrim(rst!comanda), 1, 3) + "000" + "\" + atrim(rst!comanda)
    If Not existeix(vcarpetadesti + "\" + atrim(rst!comanda) + "_BaixaMuntadora.pdf") Then
      vmsg = vmsg + " - " + atrim(rst!comanda)
    End If
    rst.MoveNext
    DoEvents
  Wend
  Set rst = Nothing
  
  
  'impresores
  vmsg = vmsg + Chr(13) + Chr(10) + "IMPRESORES" + Chr(13) + Chr(10)
  Set rst = dbbaixes.OpenRecordset("select comanda from impressorestot where operari<>0 and acavada='1' and year(dataimpressio)>2020")
  rst.MoveLast
  rst.MoveFirst
  While Not rst.EOF
    'Me.Caption = atrim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount)
    vcarpetadesti = "\\ord_copies\comandespdf\Les_" + Mid(atrim(rst!comanda), 1, 3) + "000" + "\" + atrim(rst!comanda)
    If Not existeix(vcarpetadesti + "\" + atrim(rst!comanda) + "_BaixaImpresores.pdf") Then
      vmsg = vmsg + " - " + atrim(rst!comanda)
    End If
    rst.MoveNext
    DoEvents
  Wend
  Set rst = Nothing
  
  'laminadores
  vmsg = vmsg + Chr(13) + Chr(10) + "LAMINADORES" + Chr(13) + Chr(10)
  Set rst = dbbaixes.OpenRecordset("select comanda from laminadorestot where operari<>0 and acavada='1' and year(datalaminacio)>2020")
  rst.MoveLast
  rst.MoveFirst
  While Not rst.EOF
    'Me.Caption = atrim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount)
    vcarpetadesti = "\\ord_copies\comandespdf\Les_" + Mid(atrim(rst!comanda), 1, 3) + "000" + "\" + atrim(rst!comanda)
    If Not existeix(vcarpetadesti + "\" + atrim(rst!comanda) + "_BaixaLaminadores.pdf") Then
      vmsg = vmsg + " - " + atrim(rst!comanda)
    End If
    rst.MoveNext
    DoEvents
  Wend
  Set rst = Nothing

    'rebobinadores
  vmsg = vmsg + Chr(13) + Chr(10) + "REBOBINADORES" + Chr(13) + Chr(10)
  Set rst = dbbaixes.OpenRecordset("select comanda from rebobinadorestot where operari<>0 and acavada='1' and year(datarebobinat)>2020")
  rst.MoveLast
  rst.MoveFirst
  While Not rst.EOF
    'Me.Caption = atrim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount)
    vcarpetadesti = "\\ord_copies\comandespdf\Les_" + Mid(atrim(rst!comanda), 1, 3) + "000" + "\" + atrim(rst!comanda)
    If Not existeix(vcarpetadesti + "\" + atrim(rst!comanda) + "_BaixaRebobinadores.pdf") Then
      vmsg = vmsg + " - " + atrim(rst!comanda)
    End If
    rst.MoveNext
    DoEvents
  Wend
  Set rst = Nothing
  enviaremail "jmiralles@inplacsa.com", "Baixes de secció sense PDF.", "Informe setmanal." + Chr(13) + Chr(10) + vmsg
  'enviaremail "miquel.inplacsa@gmail.com", "Baixes de secció sense PDF.", "Informe setmanal." + Chr(13) + Chr(10) + vmsg
End Sub
Function impresorapdfpredeterminada() As Boolean
   Dim X As Printer
   Dim obj_Impresora As Object
   If Printers.Count > 0 Then
    For Each X In Printers
        DoEvents
        If InStr(1, LCase(X.DeviceName), "pdfcreator") > 0 Then
           impresorapdfpredeterminada = True
           Set obj_Impresora = CreateObject("WScript.Network")
           obj_Impresora.setdefaultprinter X.DeviceName
        End If
    Next X
   End If
End Function
Sub generartotselsPDFpetits()
   Dim vdir As String
   Dim vdirs As String
   vdirs = Dir("\\ord_copies\documentacioclixes\*.*", vbDirectory)
   While vdirs <> ""
      buscardins "\\ord_copies\documentacioclixes\" + vdirs
      vdirs = Dir
   Wend
   
End Sub
Sub buscardins(vdirs As String)
   Dim vdir As String
   'If Mid(vdirs + " ", 1, 1) = "." Then Exit Sub
   vdir = Dir(vdirs + "\pdf?????-???.pdf")
   While vdir <> ""
      If FileLen(vdirs + "\" + vdir) / 1000 > 5000 Then
          MsgBox vdir + "  " + atrim(FileLen(vdirs + "\" + vdir))
      End If
      vdir = Dir
   Wend
End Sub

Private Sub Form_Click()
'organitzar_fitxers_Escanejats_Expedicions
'mirarcomandesenproducciosensepreu
 'calcular_credit_tots_clients
' comprovar_albarans_tintes_nous
' organitzar_fitxers_Escanejats_Expedicions
'calcular_credit_tots_clients
 'obrir_tancar_taules True
 'netejar_referencies_tarifes
'comprovar_PDF_baixesseccions
'llistat_comprespendents_1rdemes
'llistat_llaunes_1rdemes
'Revisarcomandaacabadaimpresoressinohihaliniesdefuncionamentabaixes
' comprovar_albarans_tintes_nous
'GRUPS_revisarsihihaprous_metres_assignats
'comprovar_albarans_tintes_nous
'revisar_clients_donatsdebaixa_ambcomandescirculant
'revisarsihihacomandesalallistadeimpresiosensepackinglist
'  calcular_credit_tots_clients
'actualitzar_clients "PLASEL"
'comprovar_albarans_tintes_nous
  'enviarinformedebobinessensenumerodepalet
  'convertirPDFeditablesaPDFpetits
 ' generartotselsPDFpetits

'revisar_metresicanutorebobinadora_ambcomandescirculant
 ' actualitzarCSVnetejalaser_a_aniloxos
  'calcular_estadisticaaniloxos
 ' MsgBox "fet"
  ' revisar_clientsalbaransclixes
'comprovar_comandesafabricaciosenseescanejar
 'MsgBox Enviar_Mail_CDO("smtp.gmail.com", "miquel.inplacsa@gmail.com", "miquel.inplacsa@gmail.com", "Una prova", " Cos del missatge", "c:\temp\prova.pdf", 465, "miquel.inplacsa@gmail.com", "ipc990900ipc", True, True)
' enviaremail "miquel.inplacsa@gmail.com", "una prova subjecte", "una prova cos del missatge", "c:\temp\prova.pdf"
  'guardarllistatestocamagatzem
'Timercadaminut_Timer
'comprovarmodificacionscomandesienviarles
'estadistica_llaunesinplacsa
'comprovar_comandesafabricaciosenseescanejar
  'mirarsieshoradexportarpdfapng True
' enviarinformedecontenidorsperrecuperar "miquel.inplacsa@gmail.com"
' comprovar_compres_datadentregapasada "M"
 'passar_informe_comandesdesactivades
 'obrir_tancar_taules (True)
  ' actualitzar_credit_clients "INPLACSA"
'comprovar_compres_datadentregapasada "V"
'comprovar_error_taules
'  actualitzar_facturesSAP
 ' comprovarestocminimdellaunes
   'enviaremail "miqueltec@gmail.com", "envio a
  ' passar_resumalbaransfotogravadors
'   obrir_tancar_taules True
'   comprovarsihihacomandesambtintessensecodi
'   obrir_tancar_taules False
   'comprovarestocdeadhesiuamuntadora
 '  comprovar_lesllaunesdelsdosificadors
   
End Sub
Sub comprovar_lesllaunesdelsdosificadors()
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim vllauna As String
   Dim dbtintes As Database
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
   Set rst = dbtintes.OpenRecordset("SELECT * fROM Componentsbase")
   While Not rst.EOF
      Set rstc = dbtintes.OpenRecordset("select * from detallnumeroslotsbase where idcomponent=" + atrim(rst!idcomponent) + " order by data desc")
      If Not rstc.EOF Then
         vllauna = atrim(rstc!numerodelot)
         rstc.MoveNext
         If Not rstc.EOF Then comprovar_llaunacoincideixambdosificador atrim(rstc!numerodelot), vllauna, rst!nomcomponent, dbtintes
      End If
      rst.MoveNext
   Wend
   Set rst = Nothing
  ' Set dbtintes = Nothing
End Sub
Sub comprovar_llaunacoincideixambdosificador(vllaunavella As String, vllaunanova As String, vdescripciodelcomponent As String, dbtintes As Database)
   Dim rstlln As Recordset
   Dim rstllv As Recordset
   If Mid(vllaunavella + " ", 1, 1) <> "A" Or Mid(vllaunanova + " ", 1, 1) <> "A" Then Exit Sub
   Set rstllv = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, tintes.idfamilia, tintes.idsubfamilia, tintes.idfamcolor, tintes.idsubfamcolor FROM Llaunes INNER JOIN tintes ON Llaunes.idtinta = tintes.idtinta where numllauna='" + vllaunavella + "'")
   Set rstlln = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, tintes.idfamilia, tintes.idsubfamilia, tintes.idfamcolor, tintes.idsubfamcolor FROM Llaunes INNER JOIN tintes ON Llaunes.idtinta = tintes.idtinta where numllauna='" + vllaunanova + "'")
   If rstlln.EOF Or rstllv.EOF Then enviaremail "controlestoctintes", "Error de llauna al dosificador de " + atrim(vdescripciodelcomponent), "La llauna nova " + atrim(vllaunanova) + " o l'anterior " + atrim(vllaunavella) + " no existeixen a la base de dades.": GoTo fi
   If rstllv!idfamilia <> rstlln!idfamilia Or rstllv!idsubfamilia <> rstlln!idsubfamilia Or rstllv!idfamcolor <> rstlln!idfamcolor Or rstllv!idsubfamcolor <> rstlln!idsubfamcolor Then
        enviaremail "controlestoctintes", "Error de llauna al dosificador de " + atrim(vdescripciodelcomponent), "La llauna " + atrim(vllaunanova) + " del dosificador " + atrim(vdescripciodelcomponent) + " no correspont amb l'anterior " + atrim(vllaunavella)
       'enviar missatge families equivocades per aquest dosificador
   End If
fi:
End Sub

Sub comprovarcalloffdecomandesjafabricades()
  Dim rst As Recordset
  Dim vlogcomandes As String
  Set dbcomandes = OpenDatabase(rutadelfitxer(cami) + "comandes.mdb")
  Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda,comandes.refclient, comandes.proximaseccio FROM calloffs_detall LEFT JOIN comandes ON calloffs_detall.comanda = comandes.comanda WHERE (((comandes.proximaseccio)='V' Or (comandes.proximaseccio)='P'));")
  While Not rst.EOF
    vlogcomandes = "La comanda " + atrim(rst!comanda) + " referencia " + atrim(rst!refclient) + " te un Call-Off total assignat i ara ja té bobines produïdes, assigna-les deseguida que puguis." + Chr(10)
    rst.MoveNext
  Wend
  If atrim(vlogcomandes) <> "" Then
      enviaremail "destinatari1", "Call-Offs que ja tenen producció. S'han d'assignar els palets.", vlogcomandes
   End If
   Set rst = Nothing
End Sub
Sub comprovarsihihacomandesambtintessensecodi()
  Dim rst As Recordset
  Dim rsttintes As Recordset
  Dim vtintes As String
  Dim vlogcomandes As String
  Dim vnomtintes As String
  Set rst = dbcomandes.OpenRecordset("select * from comandes where proximaseccio<>'T'")
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  While Not rst.EOF
    If cadbl(rst!numtreball) > 0 Then
        Set rsttintes = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(cadbl(rst!numtreball)) + " and ordremodificacio=" + atrim(cadbl(rst!numordremodificacio)) + " and (coditinta='0' or coditinta='' or coditinta=null)")
        If Not rsttintes.EOF Then
           vnomtintes = ""
           While Not rsttintes.EOF
              If atrim(rsttintes!Color) <> "" Then vnomtintes = vnomtintes + " [" + atrim(rsttintes!Color) + "]"
              rsttintes.MoveNext
           Wend
           vlogcomandes = "La comanda " + atrim(rst!comanda) + " Treball " + atrim(rst!numtreball) + "  te tintes sense CODI DE TINTA." + vnomtintes + Chr(10)
        End If
    End If
    rst.MoveNext
  Wend
  If atrim(vlogcomandes) <> "" Then
      enviaremail "destinatarirevisartintes", "Comandes amb tintes sense codi de tinta assignada al treball.", vlogcomandes
   End If
  Set rsttintes = Nothing
  Set rst = Nothing
 ' Set dbclixes = Nothing
End Sub

Private Sub Form_Initialize()
' Me.Hide
Call PonerSystray
End Sub

Private Sub Form_Load()
  
  If App.PrevInstance Then End
  escriure_log "---------------------------------------------------------", "c:\temp\Log_EnviarMails_servidor.txt"
  escriure_log "-------- INICI PROGRAMA " + atrim(Now) + "------------", "c:\temp\Log_EnviarMails_servidor.txt"
  escriure_log "---------------------------------------------------------", "c:\temp\Log_EnviarMails_servidor.txt"
  
  fitxerini = "Comandes.ini"
  bexportarcomandes.BackColor = QBColor(15)
 ' sendmail1.Move -1000, -1000
  cami = llegir_ini("General", "cami", fitxerini)
  Set db = DBEngine.OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
  noenviar.Value = cadbl(llegir_ini("general", "noenviar", "enviarservidor.ini"))
  exportarauto.Value = cadbl(llegir_ini("general", "exportarauto", "enviarservidor.ini"))
  exportarpdfapng.Value = cadbl(llegir_ini("general", "exportarpdfapng", "enviarservidor.ini"))
  checkplanificaciotmp.Value = cadbl(llegir_ini("general", "planificaciotmp", "enviarservidor.ini"))
  etrutaalbarans = llegir_ini("General", "rutaalbarans", "enviarservidor.ini")
  etrutaescaneralbarans = llegir_ini("General", "rutaescaneralbarans", "enviarservidor.ini")
  etrutacomandes = llegir_ini("General", "rutacomandes", "enviarservidor.ini")
  etrutaescanercomandes = llegir_ini("General", "rutaescanercomandes", "enviarservidor.ini")
  etrutaetiquetesbobina = llegir_ini("General", "rutaetiquetesbobinaDRIVE", "enviarservidor.ini")
  If etrutaalbarans = "{[}]" Or etrutacomandes = "{[}]" Then
     escriure_ini "General", "rutaalbarans", "\\ord_josepm\Albarans Proveidors\", "enviarservidor.ini"
     escriure_ini "General", "rutaescaneralbarans", "\\ord_copies\Proveidors Albarans Scanner\", "enviarservidor.ini"
     escriure_ini "General", "rutacomandes", "\\ord_josepm\Comandes\", "enviarservidor.ini"
     escriure_ini "General", "rutaetiquetesbobina", "\\ord_josepm\EtiquetesBobina\", "enviarservidor.ini"
     escriure_ini "General", "rutaescanercomandes", "\\ord_copies\comandespdf\", "enviarservidor.ini"
     etrutacomandes = llegir_ini("General", "rutacomandes", "enviarservidor.ini")
     etrutaescanercomandes = llegir_ini("General", "rutaescanercomandes", "enviarservidor.ini")
     etrutaalbarans = llegir_ini("General", "rutaalbarans", "enviarservidor.ini")
     etrutaescaneralbarans = llegir_ini("General", "rutaescaneralbarans", "enviarservidor.ini")
     etrutaetiquetesbobina = llegir_ini("General", "rutaetiquetesbobina", "enviarservidor.ini")
  End If
   ' sincronitzar_taulesmestra
  If llegir_ini("General", "rutaCQLotsDRIVE", "enviarservidor.ini") <> "{[}]" Then
   escriure_ini "ruta", "rutaCQLotsDRIVE", llegir_ini("General", "rutaCQLotsDRIVE", "enviarservidor.ini"), rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
   escriure_ini "ruta", "rutaAlbaransSAPDRIVE", llegir_ini("General", "rutaAlbaransSAPDRIVE", "enviarservidor.ini"), rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
   escriure_ini "ruta", "rutaAlbaransProveidorsDRIVE", llegir_ini("General", "rutaAlbaransProveidorsDRIVE", "enviarservidor.ini"), rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
  End If
  Load formspooler
End Sub




Private Sub Command2_Click()
  Shell "notepad.exe 'c:\windows\enviarservidor.ini'", vbNormalFocus
End Sub




Private Sub Form_MouseMove( _
    Button As Integer, _
    Shift As Integer, _
    X As Single, Y As Single)

    Dim Msg As Long

    If (Me.ScaleMode = vbPixels) Then
        Msg = X
    Else
        Msg = X / Screen.TwipsPerPixelX
    End If

    Select Case Msg
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

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Me.WindowState = vbNormal
            Call SetForegroundWindow(Me.hwnd)
            Me.Show
End Sub

Private Sub Form_Resize()
    If (Me.WindowState = vbMinimized) Then
        Me.Hide
        Call PonerSystray
    Else
        Call RemoverSystray
    End If
 '   etxactivar.Width = competreb.Width - 300
 '   tancar.Left = competreb.Width - 1440
End Sub

Private Sub noenviar_Click()
   escriure_ini "general", "noenviar", atrim(noenviar.Value), "enviarservidor.ini"
End Sub

Private Sub rellotgeSAP_Timer()
   
End Sub


Function ferPing(vServer)

  'This function will return TRUE or FALSE after pinging a server and
  'checking it's response.
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com
    Dim oShell, oFSO
    Dim sTemp, sTempFile
    Dim fFile
    Dim sResults
    
    On Error Resume Next
    Const OpenAsDefault = -2
    Const FailIfNotExist = 0
    Const ForReading = 1

    Set oShell = CreateObject("WScript.Shell")
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    sTemp = oShell.ExpandEnvironmentStrings("%TEMP%")
    sTempFile = sTemp & "\runresult.tmp"

    oShell.Run "%comspec% /c ping -n 2 " & vServer & ">" & sTempFile, 0, True

    Set fFile = oFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, _
    OpenAsDefault)

    sResults = fFile.ReadAll
    fFile.Close
    oFSO.DeleteFile (sTempFile)
            
    ferPing = (InStr(sResults, "TTL=") > 0)
    
    Set oShell = Nothing
    Set oFSO = Nothing

End Function

Sub sincronitzar_taulesmestra()
   Dim vnomusuari As String
   If Not ferPing("servidorsap") Then Exit Sub
   Command6.BackColor = QBColor(12)
   DoEvents
    Me.Caption = "Sincronització mestres (CLIENTS INPLACSA)": DoEvents
   actualitzar_clients "INPLACSA"
   Me.Caption = "Sincronització mestres (CLIENTS PLASEL)": DoEvents
   actualitzar_clients "PLASEL"
   Me.Caption = "Sincronització mestres (CREDIT INPLACSA)": DoEvents
   actualitzar_credit_clients "INPLACSA"
   Me.Caption = "Sincronització mestres (PROVEIDORS)": DoEvents
   actualitzar_proveidors
   Me.Caption = "Sincronització mestres ( TRANSPORTISTES)": DoEvents
   actualitzar_transportistes
   Me.Caption = "Sincronització mestres (FACTURES SAP)": DoEvents
   actualitzar_facturesSAP
   Command6.BackColor = Command3.BackColor
   vnomusuari = llegir_ini("General", "sincronitzarsapusuari", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini")
   If vnomusuari = "" Or vnomusuari = "{[}]" Then vnomusuari = " (Automàtic)"
   enviaremail "miquel.inplacsa@gmail.com", "Sincronització SAP " + vnomusuari, ""
   Me.Caption = "Enviament d'Incidències"
End Sub
Sub actualitzar_facturesSAP()

  On Error Resume Next
  dbsap.Execute "drop table Importada_Albarans_Compres_Inplacsa"
  dbsap.Execute "drop table Importada_Albarans_Compres_Plasel"
  dbsap.Execute "drop table Importada_LiniesFacturesSAP_Inplacsa"
  dbsap.Execute "drop table Importada_LiniesFacturesSAP_Plasel"
  dbsap.Execute "drop table Importada_FormesdepagamentClients_Inplacsa"
  dbsap.Execute "drop table Importada_RebutspendentsClients_Inplacsa"
  
  dbsap.Execute "SELECT Factures_albarans_Inplacsa.* INTO Importada_Albarans_Compres_Inplacsa FROM Factures_albarans_Inplacsa"
  dbsap.Execute "SELECT Factures_albarans_Plasel.* INTO Importada_Albarans_Compres_Plasel FROM Factures_albarans_Plasel"
'  dbsap.Execute "SELECT Factures_albarans_Plasel.* INTO Importada_Albarans_Compres_Inplacsa FROM Factures_albarans_Plasel"
  dbsap.Execute "SELECT * INTO Importada_LiniesFacturesSAP_Inplacsa FROM LiniesFacturesSAP_Inplacsa"
  dbsap.Execute "SELECT * INTO Importada_LiniesFacturesSAP_Plasel FROM LiniesFacturesSAP_Plasel"
  dbsap.Execute "SELECT * INTO Importada_FormesdepagamentClients_Inplacsa FROM FormapagamentClients_Inplacsa"
  dbsap.Execute "SELECT * INTO Importada_RebutspendentsClients_Inplacsa FROM RebutsPendentsFacturesVenda"
End Sub
Sub actualitzar_transportistes()
   Dim rsttransports As Recordset
   Dim rstsap As Recordset
   Dim taula_transports_sap As String
   Dim taula_transports_produccio As String
   
   taula_transports_sap = "Transportistes_Inplacsa"
   taula_transports_produccio = "transportistes"
plasel:
   Set rstsap = dbsap.OpenRecordset("select * from " + taula_transports_sap)
   Set rsttransports = dbcomandes.OpenRecordset("select * from " + taula_transports_produccio)
   While Not rstsap.EOF
     rsttransports.FindFirst "codi=" + atrim(rstsap!TrnspCode)
     If Not rsttransports.NoMatch Then
         If atrim(rsttransports!descripcio) <> atrim(rstsap!TrnspName) Then
              rsttransports.Edit
              rsttransports!descripcio = atrim(rstsap!TrnspName)
              rsttransports.Update
         End If
          Else
            rsttransports.AddNew
            rsttransports!descripcio = atrim(rstsap!TrnspName)
            rsttransports!codi = atrim(rstsap!TrnspCode)
            rsttransports.Update
     End If
     rstsap.MoveNext
   Wend
      'plasel ja no es fa
  ' If taula_transports_sap <> "Transportistes_Plasel" Then
  '    taula_transports_sap = "Transportistes_Plasel"
  '    taula_transports_produccio = "transportistes_plasel"
      'GoTo plasel
  ' End If
   Set rsttransports = Nothing
   Set rstsap = Nothing
End Sub
Sub actualitzar_clients(vempresa As String)
   Dim rstclients As Recordset
   Dim rstsap As Recordset
   Dim vlogcanvis As String
   vempresa = UCase(vempresa)
   obrir_tancar_taules True
   If vempresa = "INPLACSA" Then vempresa = ""
   Set rstsap = dbsap.OpenRecordset("select * from clients" + vempresa + " where inactiu='N'")
   Set rstclients = dbcomandes.OpenRecordset("select * from clients_codissap" + vempresa)
   dbcomandes.Execute "update clients_codissap" + vempresa + " set actualitzat=false"
   While Not rstsap.EOF
    If cadbl(rstsap!Codigo) > 0 Then
     rstclients.FindFirst "codisap=" + atrim(rstsap!Codigo)
     If Not rstclients.NoMatch Then
         rstclients.Edit
         If atrim(rstclients!nomclient) <> atrim(rstsap!nombre) Then
              vlogcanvis = vlogcanvis + "Modificació:  " + atrim(rstsap!Codigo) + " - " + atrim(rstclients!nomclient) + " -->  " + atrim(rstsap!nombre) + Chr(10)
              rstclients!nomclient = atrim(rstsap!nombre)
         End If
         If atrim(rstclients!moneda) <> IIf(atrim(rstsap!moneda) = "EUR", "Euros", "Dolars") Then
             vlogcanvis = vlogcanvis + "Modificació moneda:  " + atrim(rstsap!Codigo) + " - " + atrim(rstclients!moneda) + " -->  " + atrim(rstsap!moneda) + Chr(10)
             rstclients!moneda = IIf(atrim(rstsap!moneda) = "EUR", "Euros", "Dolars")
         End If
         If vempresa = "" Then
            If atrim(rstclients!Formadepagament) <> atrim(rstsap!Formadepagament) Then
                 vlogcanvis = vlogcanvis + "Modificació Formadepagament:  " + atrim(rstsap!Codigo) + " - " + atrim(rstclients!Formadepagament) + " -->  " + atrim(rstsap!Formadepagament) + Chr(10)
                rstclients!Formadepagament = atrim(rstsap!Formadepagament)
            End If
         End If
         If vempresa = "" Then
            If atrim(rstclients!nif) <> atrim(rstsap!LicTradNum) Then
                 vlogcanvis = vlogcanvis + "Modificació Nif:  " + atrim(rstsap!Codigo) + " - " + atrim(rstclients!nif) + " -->  " + atrim(rstsap!LicTradNum) + Chr(10)
                 rstclients!nif = atrim(rstsap!LicTradNum)
            End If
         End If
         rstclients!actualitzat = True
         rstclients.Update
          Else
           If atrim(rstsap!nombre) <> "" Then
            rstclients.AddNew
            rstclients!nomclient = atrim(rstsap!nombre)
            rstclients!codisap = atrim(rstsap!Codigo)
            rstclients!moneda = IIf(atrim(rstsap!moneda) = "EUR", "Euros", "Dolars")
            rstclients!nif = atrim(rstsap!LicTradNum)
            If vempresa = "" Then rstclients!Formadepagament = atrim(rstsap!Formadepagament)
            rstclients.Update
            vlogcanvis = vlogcanvis + "Nou: " + atrim(rstsap!Codigo) + " - " + atrim(rstsap!nombre) + Chr(10)
           End If
     End If
    End If
     rstsap.MoveNext
   Wend
   Set rstclients = dbcomandes.OpenRecordset("select * from clients_codissap" + vempresa + " where not actualitzat")
   While Not rstclients.EOF
      vlogcanvis = vlogcanvis + "Eliminat: " + atrim(rstclients!codisap) + " - " + atrim(rstclients!nomclient) + Chr(10)
      rstclients.MoveNext
   Wend
   'MsgBox vlogcanvis
  ' Clipboard.Clear
  ' Clipboard.SetText vlogcanvis
   dbcomandes.Execute "delete * from clients_codissap" + vempresa + " where not actualitzat"
   dbcomandes.Execute "delete * from clients_codissap" + vempresa + " where nomclient=''"
   Set rstclients = Nothing
   Set rstsap = Nothing
   If atrim(vlogcanvis) <> "" And vempresa = "" Then
      enviaremail "destinatariclientdesvinculatsSAP", "Modificacions taula clients SAP vs Producció", vlogcanvis
   End If
End Sub
Sub actualitzar_credit_clients(vempresa As String)
   Dim rstclients As Recordset
   Dim rstsap As Recordset
   Dim vlogcanvis As String
   vempresa = UCase(vempresa)
   If vempresa = "INPLACSA" Then vempresa = ""
   Set rstsap = dbsap.OpenRecordset("select * from Credit_Clients_Inplacsa")
   dbcomandes.Execute "delete * from credit_clients_inp"
   Set rstclients = dbcomandes.OpenRecordset("select * from Credit_Clients_Inp")
   While Not rstsap.EOF
     rstclients.AddNew
     'If rstsap!cardcode = 43000006998# Then Stop
     For i = 0 To rstsap.Fields.Count - 1
        If rstclients.Fields(rstsap.Fields(i).Name).Type = 7 Then
                rstclients.Fields(rstsap.Fields(i).Name) = cadbl_sap(rstsap.Fields(i))
             Else: rstclients.Fields(rstsap.Fields(i).Name) = atrim(rstsap.Fields(i))
        End If
        'rstclients.Fields(rstsap.Fields(i).Name) = IIf(rstclients.Fields(rstsap.Fields(i).Name).Type = 7, cadbl_sap(rstsap.Fields(i)), atrim(rstsap.Fields(i)))
     Next i
     rstclients.Update
     rstsap.MoveNext
   Wend
   Set rstsap = Nothing
   Set rstclients = Nothing
End Sub
Function cadbl_sap(v As Variant) As Double
   'If Len(v) > 5 Then If MsgBox(v, vbCritical + vbDefaultButton2 + vbYesNo, "a") = vbNo Then End
   If GetLocaleDecimalSep = "," Then
       v = substituirtot(atrim(v), ".", "")
        Else: If GetLocaleDecimalSep = "." Then v = substituirtot(atrim(v), ",", "")
   End If
 '  v = substituirtot(atrim(v), ",", "#")
 '  v = substituirtot(atrim(v), ".", ",")
 '  v = substituirtot(atrim(v), "#", ".")
    cadbl_sap = cadbl(v)
End Function
Function buscarprovincia(rst As Recordset, vcodipostal As String) As String
   'rst.FindFirst "codipostal='" + Mid(vcodipostal + "    ", 1, 3) + "'"
   'If rst.NoMatch Then
   rst.FindFirst "codipostal='" + Mid(vcodipostal + "    ", 1, 2) + "x'"
   If Not rst.NoMatch Then buscarprovincia = atrim(rst!provincia)
End Function
Sub actualitzar_proveidors()
   Dim rstprov As Recordset
   Dim rstsap As Recordset
   Dim codinouprov As Double
   Dim codinouprovcomercial As Double
   Dim rstcodisprovsap As Recordset
   Dim rstprovincies As Recordset
   Dim vcodipostal As String
   Dim vprovinciaproveidor  As String
   
   Set rstsap = dbsap.OpenRecordset("select * from proveidors order by Activo")
   Set rstprovincies = dbsap.OpenRecordset("select * from [CodisPostals-Provincies]")
   Set rstprov = dbcomandes.OpenRecordset("select * from proveidors_comercial")
   Set rstcodisprovsap = dbcomandes.OpenRecordset("select * from proveidors_codisSAP")
   dbcomandes.Execute "update proveidors_codissap set actualitzat=false"
   '''''   HI HA UNA FUNCIO ACTUALITZAR_PROVEIDORS_IALTAPRODUCCIO  QUE TAMBÉ FA LA ALTA DE CADA REGISTRES
     '''    A LA TAULA DE PROVEIDORS I DE PROVEIDORS_COMERCIALS
   While Not rstsap.EOF
     rstcodisprovsap.FindFirst "codiSAP=" + atrim(cadbl(rstsap!Codigo))
     If atrim(rstsap!Country) = "ES" Then
            If Not rstcodisprovsap.NoMatch Then vprovinciaproveidor = buscarprovincia(rstprovincies, atrim(rstsap!zipcode))
              Else: vprovinciaproveidor = atrim(rstsap!county)
     End If
     If Not rstcodisprovsap.NoMatch Then
         rstcodisprovsap.Edit
         If atrim(rstcodisprovsap!nomproveidor) <> atrim(rstsap!nombre) Then  ' Or atrim(rstprov!aliastintes) <> atrim(rstsap!aliastintes)
              rstcodisprovsap!nomproveidor = Mid(atrim(rstsap!nombre), 1, rstcodisprovsap("nomproveidor").Size)
         End If
         rstcodisprovsap!nif = atrim(rstsap!nif)
         rstcodisprovsap!actualitzat = True
         rstcodisprovsap!provincia = Mid(vprovinciaproveidor, 1, 20)
         rstcodisprovsap.Update
          Else
            'codissap de produccio nou
            rstcodisprovsap.AddNew
            rstcodisprovsap!codisap = atrim(cadbl(rstsap!Codigo))
            rstcodisprovsap!nomproveidor = Mid(atrim(rstsap!nombre), 1, rstcodisprovsap("nomproveidor").Size)
            rstcodisprovsap!nif = atrim(rstsap!nif)
            rstcodisprovsap!provincia = vprovinciaproveidor
            
           ' rstcodisprovsap!aliastintes = atrim(rstsap!aliastintes)
            rstcodisprovsap.Update
            ''''''''''
     End If
     rstprov.FindFirst "codicomptable='" + atrim(cadbl(rstsap!Codigo)) + "'"
     'If rstsap!Activo = "N" And Not rstprov.NoMatch Then Stop
     If Not rstprov.NoMatch Then dbcomandes.Execute "update proveidors set databaixa=" + IIf(rstsap!Activo = "Y", "Null", "NOW") + " where codi=" + atrim(rstprov!codiproduccio)
     rstsap.MoveNext
   Wend
   dbcomandes.Execute "delete * from proveidors_codissap where not actualitzat"
   dbcomandes.Execute "delete * from proveidors_codissap where nomproveidor=''"
   Set rstsap = Nothing
   Set rstprov = Nothing
   Set rstcodisprovsap = Nothing
   
End Sub
Sub actualitzar_proveidors_ialtaproduccio()
   Dim rstprov As Recordset
   Dim rstsap As Recordset
   Dim codinouprov As Double
   Dim codinouprovcomercial As Double
   Dim rstcodisprovsap As Recordset
   Set rstsap = dbsap.OpenRecordset("select * from proveidors")
   Set rstprov = dbcomandes.OpenRecordset("select * from proveidors_comercial")
   Set rstcodisprovsap = dbcomandes.OpenRecordset("select * from proveidors_codisSAP")
   While Not rstsap.EOF
     rstprov.FindFirst "codicomptable='" + atrim(rstsap!Codigo) + "'"
     If Not rstprov.NoMatch Then
         If atrim(rstprov!nom) <> atrim(rstsap!nombre) Then  ' Or atrim(rstprov!aliastintes) <> atrim(rstsap!aliastintes)
              rstprov.Edit
              rstprov!nom = atrim(rstsap!nombre)
              rstprov.Update
                 ' aqui també s'hauria de canvia el nom de la paura de proveidors que
                   ' es la que utilitzem a producció, però no ho faig perquè diria que
                   ' el nom de produccioó iel de adalt noha de ser igual per força
              rstcodisprovsap.FindFirst "codisap=" + atrim(rstsap!Codigo)
              If Not rstcodisprovsap.NoMatch Then
                 rstcodisprovsap.Edit
                 rstcodisprovsap!nomproveidor = atrim(rstsap!nombre)
                 rstcodisprovsap!aliastintes = atrim(rstsap!aliastintes)
                 rstcodisprovsap.Update
              End If
         End If
          Else
            'codissap de produccio nou
            rstcodisprovsap.AddNew
            rstcodisprovsap!codisap = atrim(rstsap!Codigo)
            rstcodisprovsap!nomproveidor = atrim(rstsap!nombre)
           ' rstcodisprovsap!aliastintes = atrim(rstsap!aliastintes)
            rstcodisprovsap.Update
            ''''''''''
            'codis nous de proveidor i comercial copiant totes les dades
            codinouprov = nouproveidornocomercial(rstsap!nombre, "")   ' les cometes seria rstsap!aliastintes
            codinouprovcomercial = mesgranmesundeprovcomercial
            rstprov.AddNew
            rstprov!codi = codinouprovcomercial
            rstprov!codicomptable = rstsap!Codigo
            rstprov!codiproduccio = codinouprov
            rstprov!alta_desde_sap = True
            rstprov!nom = atrim(rstsap!nombre)
            rstprov.Update
     End If
     rstsap.MoveNext
   Wend
   Set rstsap = Nothing
   Set rstprov = Nothing
   Set rstcodisprovsap = Nothing
End Sub
Function mesgranmesundeprovcomercial() As Double
Dim codi As Double
Dim rsttmp As Recordset
        'busco el mes gran i el poso a codi +1
        Set rsttmp = dbcomandes.OpenRecordset("select max(codi) as [grancodi] from proveidors_comercial")
        If Not rsttmp.EOF Then
          codi = cadbl(rsttmp!grancodi) + 1
              Else: codi = 1
        End If
        mesgranmesundeprovcomercial = codi
        Set rsttmp = Nothing
End Function
Function nouproveidornocomercial(nom As String, alias As String) As Double
Dim codi As Double
Dim rttmp As Recordset
        'busco el mes gran i el poso a codi +1
        Set rsttmp = dbcomandes.OpenRecordset("select max(codi) as [grancodi] from proveidors")
        If Not rsttmp.EOF Then
          codi = cadbl(rsttmp!grancodi) + 1
              Else: codi = 1
        End If
        Set rsttmp = dbcomandes.OpenRecordset("select * from proveidors")
        rsttmp.AddNew
        rsttmp!codi = codi
        rsttmp!nom = nom
        rsttmp!aliastintes = alias
        rsttmp!alta_desde_sap = True
        rsttmp.Update
        nouproveidornocomercial = codi
        Set rttmp = Nothing
End Function
Function obrir_tancar_taules(obrir As Boolean) As Boolean
  On Error GoTo errorobrint
  obrir_tancar_taules = True
   If obrir Then
      Set dbsap = OpenDatabase(rutadelfitxer(cami) + "connexiosap.mdb")
      Set dbcomandes = OpenDatabase(rutadelfitxer(cami) + "comandes.mdb")
      Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
   End If
'   If Not obrir Then
'      Set dbsap = Nothing
'      Set dbcomandes = Nothing
'      Set dbclixes = Nothing
'   End If
  Exit Function
errorobrint:
  obrir_tancar_taules = False
  
End Function
Sub comprovarestocdeadhesiuamuntadora()
  Dim i As Byte
  Dim valor As Double
  Dim valorq As Double
  Dim unpersota As Boolean
  Dim fitxerinimuntadora As String
  fitxerinimuntadora = rutadelfitxer(cami) + "muntadora.ini"
  For i = 1 To 6
    valor = cadbl(llegir_ini("Valors", "minim" + atrim(i), fitxerinimuntadora))
    valorq = cadbl(llegir_ini("Valors", "q" + atrim(i), fitxerinimuntadora))
    If valor > 0 And valorq > 0 Then
       If valorq <= valor Then
           unpersota = True
       End If
    End If
  Next i
  If unpersota Then
     If llegir_ini("Valors", "enviat", fitxerinimuntadora) = "no" Then
        enviaremail "destinatari1", "Estoc de Adhesiu a muntadora per sota de mínim i sense comanda, s'hauria de revisar.", "Aquest missatge es genera per assegurar que s'ha rebut el missatge demanant adhesiu, si ja l'heu rebut no en feu cas d'aquest"
     End If
  End If
End Sub
Sub comprovar_albarans_tintes_nous()
   Dim verroralbarans As Boolean
   Dim verrorcomandes As Boolean
   Dim verroretiquetespalets As Boolean
   Dim verroretiquetespressupostos As Boolean
   Dim vrutaetiquetesbobinesLOCAL As String
   Dim etrutapressupostosLOCAL As String
   Dim etrutapressupostos As String
   Dim vrutacachepressupostos As String
   Dim vrutapressupostosLOCAL As String
   Static vdins As Boolean
   
   If bexportarcomandes.Tag = "exportant" Then GoTo fi
   If vdins Then Exit Sub
   vdins = True
   etrutaalbarans = llegir_ini("General", "rutaalbarans", "enviarservidor.ini")
   etrutaescaneralbarans = llegir_ini("General", "rutaescaneralbarans", "enviarservidor.ini")
   etrutacomandes = llegir_ini("General", "rutacomandes", "enviarservidor.ini")
   etrutaescanercomandes = llegir_ini("General", "rutaescanercomandes", "enviarservidor.ini")
   etrutaetiquetesbobina = llegir_ini("General", "rutaetiquetesbobinaDRIVE", "enviarservidor.ini")
   vrutaetiquetesbobinesLOCAL = llegir_ini("General", "rutaetiquetesbobinaLOCAL", "enviarservidor.ini")
   etrutapressupostos = llegir_ini("General", "rutapressupostosDRIVE", "enviarservidor.ini")
   vrutapressupostosLOCAL = llegir_ini("General", "rutapressupostosLOCAL", "enviarservidor.ini")
   If Not existeix(etrutaalbarans) Then etrutaalbarans = "Error ruta ordinador drive d'albarans.": verroralbarans = True
   If Not existeix(etrutaescaneralbarans) Then etrutaescaneralbarans = "Error ruta escaner albarans.": verroralbarans = True
   If Not existeix(etrutaescanercomandes) Then etrutaescanercomandes = "Error ruta escaner comandes.": verrorcomandes = True
   If Not existeix(etrutacomandes) Then etrutacomandes = "Error ruta ordinador drive de comandes.": verrorcomandes = True
   If Not existeix(etrutaetiquetesbobina) Then etrutaetiquetesbobina = "Error ruta ordinador drive de Etiquetes Bobina.": verroretiquetespalets = True
   If Not existeix(etrutapressupostos) Then etrutaetiquetesbobina = "Error ruta ordinador drive de Pressupostos.": verroretiquetespressupostos = True
   If Not verroralbarans Then
    organitzar_fitxers etrutaescaneralbarans, "Tinta\"
    organitzar_fitxers etrutaescaneralbarans, "Material Film Bobines\"
    organitzar_fitxers etrutaescaneralbarans, "Varis\"
   End If
   If Not verrorcomandes Then
      escriure_log "1 (organitzar comandes)", "c:\temp\Log_EnviarMails_servidor.txt"
      organitzar_fitxers_comandes etrutaescanercomandes, etrutacomandes
       Else: escriure_ini "General", "pujantadrive", "no", etrutaescanercomandes + "\cache_originals\organitzar.ini"
   End If
   If Not verroretiquetespalets Then
      escriure_log "1 (organitzar etiquetespalets)", "c:\temp\Log_EnviarMails_servidor.txt"
      organitzar_fitxers_etiquetespalets rutadelfitxer(cami) + "cache_EtiquetesBobinesProveidor", etrutaetiquetesbobina, vrutaetiquetesbobinesLOCAL
       Else: escriure_ini "General", "pujantadrive", "no", etrutaescanercomandes + "\cache_originals\organitzar.ini"
   End If
   If Not verroretiquetespressupostos Then
      escriure_log "1 (organitzar pressupostos)", "c:\temp\Log_EnviarMails_servidor.txt"
      vrutacachepressupostos = rutadelfitxer(cami) + "Cache_escanejarpressupostos\cache"
      If UCase(llegir_ini("General", "pujantadrive", vrutacachepressupostos + "\organitzar.ini")) = "NO" Then
       escriure_ini "General", "pujantadrive", "si", vrutacachepressupostos + "\organitzar.ini"
       organitzar_fitxers_etiquetespalets vrutacachepressupostos, etrutapressupostos, vrutapressupostosLOCAL, "Pressupostos"
       escriure_ini "General", "pujantadrive", "no", vrutacachepressupostos + "\organitzar.ini"
      End If
       Else: escriure_ini "General", "pujantadrive", "no", etrutapressupostosLOCAL + "\cache\organitzar.ini"
   End If
   
   organitzar_fitxers_Escanejats_Expedicions
   
fi:
   vdins = False
End Sub
Sub organitzar_fitxers_Escanejats_Expedicions()
  Dim vrutadestiLOCAL As String
  Dim vrutadestiDRIVE As String
  Dim vrutaorigen As String
  Dim vnomfitxer As String
  Dim vany As String
  Static socdins As Boolean
  
  If socdins Then Exit Sub
  socdins = True
  escriure_log "1 (organitzar_fitxers_Escanejats_Expedicions)", "c:\temp\Log_EnviarMails_servidor.txt"
  'CONTROL QUALITAT DELS LOTS
  vrutaorigen = rutadelfitxer(cami) + "Cache_escanejarexpedicions\CQ\"
  vrutadestiLOCAL = llegir_ini("General", "rutaCQLotsLOCAL", "enviarservidor.ini")
  vrutadestiDRIVE = llegir_ini("General", "rutaCQLotsDRIVE", "enviarservidor.ini")
  
  processar_fitxers_local_drive "CQ", vrutaorigen, vrutadestiLOCAL, vrutadestiDRIVE
  
  escriure_log "2 (organitzar_fitxers_Escanejats_Expedicions)", "c:\temp\Log_EnviarMails_servidor.txt"
  'ALBARANSSAP FIRMATS
  vrutaorigen = rutadelfitxer(cami) + "Cache_escanejarexpedicions\AlbaransSAP\"
  vrutadestiLOCAL = llegir_ini("General", "rutaAlbaransSAPLOCAL", "enviarservidor.ini")
  vrutadestiDRIVE = llegir_ini("General", "rutaAlbaransSAPDRIVE", "enviarservidor.ini")
  processar_fitxers_local_drive "SAP", vrutaorigen, vrutadestiLOCAL, vrutadestiDRIVE
  
  escriure_log "3 (organitzar_fitxers_Escanejats_Expedicions)", "c:\temp\Log_EnviarMails_servidor.txt"
  'ALBARANS DELS PROVEIDORS
  vrutaorigen = rutadelfitxer(cami) + "Cache_escanejarexpedicions\AlbaransProveidor\"
  vrutadestiLOCAL = llegir_ini("General", "rutaAlbaransProveidorsLOCAL", "enviarservidor.ini")
  vrutadestiDRIVE = llegir_ini("General", "rutaAlbaransProveidorsDRIVE", "enviarservidor.ini") + "EscanejatDesdeExpedicions\" + atrim(Year(Now)) + "\"
  processar_fitxers_local_drive "ALB", vrutaorigen, vrutadestiLOCAL, vrutadestiDRIVE

  escriure_log "FI (organitzar_fitxers_Escanejats_Expedicions)", "c:\temp\Log_EnviarMails_servidor.txt"
  socdins = False
End Sub
Function errorcreantlaruta(vrutadestiLOCAL As String, vrutadestiDRIVE As String) As Boolean
   On Error GoTo fi
   If Not existeix(vrutadestiLOCAL) Then MkDir vrutadestiLOCAL
   If Not existeix(vrutadestiDRIVE) Then MkDir vrutadestiDRIVE
   Exit Function
fi:
  errorcreantlaruta = True
End Function
Sub processar_fitxers_local_drive(vtipus As String, vrutaorigen As String, vrutadestiLOCAL As String, vrutadestiDRIVE As String)
  Dim vnomfitxer As String
  Dim vhihafitxers As Boolean
  Dim vanex As String
  Dim vTipus_original As String
  If errorcreantlaruta(vrutadestiLOCAL, vrutadestiDRIVE) Then GoTo fi
  File3.Path = vrutaorigen
  File3.Pattern = "*.*"
  File3.Refresh
  vTipus_original = vtipus
  
  'vnomfitxer = Dir(vrutaorigen + "*.*")
  'While vnomfitxer <> ""
  For i = 0 To File3.ListCount - 1
       vtipus = vTipus_original
       vnomfitxer = File3.List(i)
       'If DateDiff("s", FileDateTime(vrutaorigen + vnomfitxer), Now) > 10 Then
            vanex = ""
            If vtipus = "SAP" Then
                If Mid(vnomfitxer + "      ", 1, 4) = "CMR_" Then
                     vanex = "CMRs\": vtipus = "CMR"
                     'MsgBox vrutadestiLOCAL + vanex
                End If
                If Mid(vnomfitxer + "      ", 1, 9) = "FraTrans_" Then
                     vanex = "FacturesTransport\": vtipus = "Fra"
                     'MsgBox vrutadestiLOCAL + vanex
                End If
            End If
            vhihafitxers = True
            If Mid(vnomfitxer + "    ", 1, 4) <> "FET_" Then
             escriure_log "(organitzar_fitxers_Escanejats_Expedicions) DRIVE " + vbNewLine + vrutaorigen + vnomfitxer + vbNewLine + vrutadestiDRIVE + vanex, "c:\temp\Log_EnviarMails_servidor.txt"
             Copiar_Fitxer vrutaorigen + vnomfitxer, vrutadestiDRIVE + vanex '+ vnomfitxer
             escriure_log "(organitzar_fitxers_Escanejats_Expedicions) LOCAL " + vbNewLine + vrutaorigen + vnomfitxer + vbNewLine + vrutadestiLOCAL + vanex, "c:\temp\Log_EnviarMails_servidor.txt"
             Copiar_Fitxer vrutaorigen + vnomfitxer, vrutadestiLOCAL + vanex '+ vnomfitxer
             escriure_log "(guardar_registre_fitxers_escanejats)" + vbNewLine + vrutaorigen + vnomfitxer, "c:\temp\Log_EnviarMails_servidor.txt"
             If vtipus <> "Fra" Then guardar_registre_fitxers_escanejats vtipus, vnomfitxer, vrutadestiLOCAL + vanex, vrutadestiDRIVE + vanex
             cambiarnomarxiu vrutaorigen + vnomfitxer, vrutaorigen + "FET_" + vnomfitxer
            End If
       'End If
       
  Next i
  'Wend
  vtipus = vTipus_original
  If vhihafitxers Then
    escriure_log "(organitzar_fitxers_Escanejats_Expedicions) Eliminar: " + vrutaorigen + "*.*", "c:\temp\Log_EnviarMails_servidor.txt"
    eliminar_fitxer vrutaorigen + "FET_*.*"
    wait 2
  End If
fi:
End Sub
Sub assignarCQs()
   Dim rst As Recordset
   Dim vdataPDF As Date
   Dim valb As String
   Dim vnomp As String
   Dim vcodiprov As String
   Dim vnumlot As String
   Dim vvalors As String
   Dim vcamps As String
   
   Set dbcomandes = OpenDatabase(cami)
   Set rst = dbcomandes.OpenRecordset("select * from registre_escanejades_expedicions")
   While Not rst.EOF
      valb = "": vnomp = "": vcodiprov = "": vnumlot = ""
      camps_CQ atrim(rst!nomfitxer), valb, vnomp, vcodiprov, vnumlot
      If vnomp <> "" Then
         rst.Edit
         rst!numalbara = valb
         rst!numlotproveidor = vnumlot
         rst!codiproveidor = vcodiprov
         rst!nomproveidor = vnomp
         rst.Update
      End If
     rst.MoveNext
   Wend
End Sub
Sub guardar_registre_fitxers_escanejats(vtipus As String, vnomfitxer As String, vrutadestiLOCAL As String, vrutadestiDRIVE As String)
   Dim rst As Recordset
   Dim vdataPDF As Date
   Dim valb As String
   Dim vnomp As String
   Dim vcodiprov As String
   Dim vnumlot As String
   Dim vvalors As String
   Dim vcamps As String
   
   Set dbcomandes = OpenDatabase(cami)
   vdataPDF = FileDateTime(vrutadestiLOCAL + vnomfitxer)
   Set rst = dbcomandes.OpenRecordset("select datafitxerPDF from registre_escanejades_expedicions WHERE nomfitxer='" + vnomfitxer + "' and datafitxerPDF=#" + Format(vdataPDF, "mm/dd/yy hh:nn:ss") + "#")
   If Not rst.EOF Then Exit Sub
   
   camps_CQ vnomfitxer, valb, vnomp, vcodiprov, vnumlot
   vcamps = "(tipus,nomfitxer,rutadestiLOCAL,rutadestiDRIVE,datafitxerPDF,nomproveidor,numalbara,codiproveidor,numlotproveidor)"
   vvalors = "('" + vtipus + "','" + vnomfitxer + "','" + vrutadestiLOCAL + "','" + vrutadestiDRIVE + "',#"
   vvalors = vvalors + Format(vdataPDF, "mm/dd/yy hh:nn:ss") + "#,'" + treure_apostruf(vnomp) + "','" + atrim(valb) + "','" + atrim(vcodiprov) + "','" + atrim(vnumlot) + "')"
   
   dbcomandes.Execute "insert into registre_escanejades_expedicions " + vcamps + " VALUES " + vvalors
   
End Sub
Sub camps_CQ(vnomfitxer As String, valb As String, vnomp As String, vcodiprov As String, vnumlot As String)
   If InStr(1, vnomfitxer, "[") = 0 Then Exit Sub
   valb = Mid(vnomfitxer, 1, InStr(1, vnomfitxer, "[") - 1)
   vcodiprov = Mid(vnomfitxer, InStr(1, vnomfitxer, "[") + 1)
   vcodiprov = Mid(vcodiprov, 1, InStr(1, vcodiprov, "]-") - 1)
   vnomp = Mid(vnomfitxer, InStr(1, vnomfitxer, "]-") + 2)
   vnomp = Mid(vnomp, 1, InStr(1, LCase(vnomp), ".pdf"))
   If Mid(valb + "    ", 1, 3) = "CQ_" Then vnumlot = substituir(atrim(valb), "CQ_", ""): valb = ""

End Sub

Sub organitzar_fitxers_etiquetespalets(vrutaorigen As String, vdesti1 As String, vdesti2 As String, Optional vnomtraspas As String)
  Dim vnomfitxer As String
  Dim vany As String
  Dim vsestarexportatalgualacache As String
  If Not existeix(vdesti1) Or Not existeix(vrutaorigen) Or Not existeix(vdesti2) Then Exit Sub
  vsestarexportatalgualacache = UCase(llegir_ini("General", "pujantadrive", vrutaorigen + "organitzar.ini"))
  If bexportarcomandes.Tag = "exportant" Or vsestarexportatalgualacache = "SI" Then Exit Sub
  bexportarcomandes.Tag = "exportant"
  escriure_ini "General", "pujantadrive", "si", vrutaorigen + "\organitzar.ini"
  vnomfitxer = Dir(vrutaorigen + "\*.*", vbDirectory)
  While vnomfitxer <> ""
       If vnomfitxer <> "." And vnomfitxer <> ".." And vnomfitxer <> "organitzar.ini" Then
        etpujantadrive = "Pujant a drive " + IIf(vnomtraspas = "", "EtiquetesBobina", vnomtraspas) + " " + vnomfitxer
            DoEvents
            Copiar_Fitxer vrutaorigen + "\" + vnomfitxer, vdesti1 '+ vnomfitxer
            If sidrivenolocal.Value <> 1 Then
               Copiar_Fitxer vrutaorigen + "\" + vnomfitxer, vdesti2 '+ vnomfitxer
            End If
            borra_carpeta vrutaorigen + "\" + vnomfitxer
        
        vnomfitxer = Dir(vrutaorigen + "\*.*", vbDirectory)
       End If
       vnomfitxer = Dir
  Wend
  escriure_ini "General", "pujantadrive", "no", vrutaorigen + "\organitzar.ini"
  bexportarcomandes.Tag = ""
  etpujantadrive = ""
End Sub
Sub enviar_a_tintes_lalbara(vfitxer As String)
   'enviaremail "controlestoctintes", "Ha arribat un albarà de proveidor de Tintes", "", vfitxer
   enviaremail "copiaalbaratintesexpedicions", "Ha arribat un albarà de proveidor de Tintes", "", vfitxer
End Sub
Sub organitzar_fitxers_comandes(vruta As String, vdesti As String)
  Dim vnomfitxer As String
  Dim vany As String
  Dim vsestarexportatalgualacache As String
  Dim vfile As DirListBox
  If Not existeix(vdesti) Or Not existeix(vruta) Then vstatus = IIf(Not existeix(vdesti), "No existeix vdesti", "") + " - " + IIf(Not existeix(vruta), "No existeix vruta", ""): Exit Sub
  vsestarexportatalgualacache = UCase(llegir_ini("General", "exportantpdfs", vruta + "organitzar.ini"))
  vstatus = atrim(Now) + " | " + bexportarcomandes.Tag + "  -   " + vsestarexportatalgualacache
  If bexportarcomandes.Tag = "exportant" Or vsestarexportatalgualacache = "SI" Then Exit Sub
  DoEvents
  bexportarcomandes.Tag = "exportant"
  Set vfile = Dir1
  escriure_ini "General", "pujantadrive", "si", vruta + "\cache_originals\organitzar.ini"
  vfile.Path = vruta + "cache_originals\"
  vfile.Refresh
  'vnomfitxer = Dir(vruta + "\cache_originals\*.*", vbDirectory)
  For i = 1 To vfile.ListCount - 1
       vnomfitxer = File1.List(i)
       vnomfitxer = substituir(vnomfitxer, rutadelfitxer(vnomfitxer), "")
       If vnomfitxer <> "." And vnomfitxer <> ".." And vnomfitxer <> "organitzar.ini" Then
        etpujantadrive = "Pujant a drive " + vnomfitxer
        If existeix(vdesti + "Originals\") And existeix(vruta) Then
            DoEvents
            Copiar_Fitxer vruta + "cache_originals\" + vnomfitxer, vdesti + "Originals\" + vnomfitxer
            If sidrivenolocal.Value <> 1 Then Copiar_Fitxer vruta + "cache_originals\" + vnomfitxer, vruta '+ vnomfitxer
            borra_carpeta vruta + "cache_originals\" + vnomfitxer
        End If
       ' vnomfitxer = Dir(vruta + "\cache_originals\*.*", vbDirectory)
       End If
  Next i
  'While vnomfitxer <> ""
  '     If vnomfitxer <> "." And vnomfitxer <> ".." And vnomfitxer <> "organitzar.ini" Then
  '      etpujantadrive = "Pujant a drive " + vnomfitxer
  '      If existeix(vdesti + "Originals\") And existeix(vruta) Then
  '          DoEvents
  '          Copiar_Fitxer vruta + "cache_originals\" + vnomfitxer, vdesti + "Originals\" + vnomfitxer
  '          If sidrivenolocal.Value <> 1 Then Copiar_Fitxer vruta + "cache_originals\" + vnomfitxer, vruta '+ vnomfitxer
  '          borra_carpeta vruta + "cache_originals\" + vnomfitxer
  '      End If
  '      vnomfitxer = Dir(vruta + "\cache_originals\*.*", vbDirectory)
  '     End If
   '    vnomfitxer = Dir
  'Wend
  
  'vnomfitxer = Dir(vruta + "\cache_fabricacio\*.*", vbDirectory)
  vfile.Path = vruta + "cache_fabricacio\"
  vfile.Refresh
  For i = 0 To vfile.ListCount - 1
       vnomfitxer = vfile.List(i)
       vnomfitxer = substituir(vnomfitxer, rutadelfitxer(vnomfitxer), "")
       If vnomfitxer <> "." And vnomfitxer <> ".." Then
        etpujantadrive = "Pujant a drive " + vnomfitxer
        If existeix(vdesti + "fulles fabricacio\") And existeix(vruta) Then
            DoEvents
            Copiar_Fitxer vruta + "cache_fabricacio\" + vnomfitxer, vdesti + "fulles fabricacio\" + vnomfitxer, 5
            If sidrivenolocal.Value <> 1 Then Copiar_Fitxer vruta + "cache_fabricacio\" + vnomfitxer, vruta, 5  '+ vnomfitxer
            borra_carpeta vruta + "cache_fabricacio\" + vnomfitxer
        End If
        'vnomfitxer = Dir(vruta + "\cache_fabricacio\*.*", vbDirectory)
       End If
       'vnomfitxer = Dir
  Next i
  'While vnomfitxer <> ""
  '     If vnomfitxer <> "." And vnomfitxer <> ".." Then
  '      etpujantadrive = "Pujant a drive " + vnomfitxer
  '      If existeix(vdesti + "fulles fabricacio\") And existeix(vruta) Then
  '          DoEvents
  '          Copiar_Fitxer vruta + "cache_fabricacio\" + vnomfitxer, vdesti + "fulles fabricacio\" + vnomfitxer, 5
  '          If sidrivenolocal.Value <> 1 Then Copiar_Fitxer vruta + "cache_fabricacio\" + vnomfitxer, vruta, 5  '+ vnomfitxer
  '          borra_carpeta vruta + "cache_fabricacio\" + vnomfitxer
  '      End If
  '      vnomfitxer = Dir(vruta + "\cache_fabricacio\*.*", vbDirectory)
  '     End If
  '     vnomfitxer = Dir
  'Wend
  escriure_ini "General", "pujantadrive", "no", vruta + "\cache_originals\organitzar.ini"
  bexportarcomandes.Tag = ""
  etpujantadrive = ""
End Sub

Sub organitzar_fitxers(vruta As String, vcarpeta As String)
  Dim vnomfitxer As String
  Dim vdirfile As FileListBox
  Dim vdirfile2 As FileListBox
  Dim vany As String
  Set vdirfile = File1
  vdirfile.Path = vruta + vcarpeta
  vdirfile.Pattern = "*.*"
  vdirfile.Refresh
  
 
  For i = 0 To vdirfile.ListCount - 1
      vnomfitxer = vdirfile.List(i)
      vany = Year(FileDateTime(vruta + vcarpeta + vnomfitxer))
       If Not existeix(etrutaalbarans + vcarpeta + vany) Then MkDir etrutaalbarans + vcarpeta + vany
       Copiar_Fitxer vruta + vcarpeta + vnomfitxer, etrutaalbarans + vcarpeta + vany + "\" '+ vnomfitxer
       If existeix(etrutaalbarans + vcarpeta + vany + "\" + vnomfitxer) Then
          If vcarpeta = "Tinta\" Then enviar_a_tintes_lalbara vruta + vcarpeta + vnomfitxer
          If vcarpeta = "Varis\" Then enviaremail "EscanejatExpedicionsCarpetaVaris", "A expedicions s'ha escanejat un document de VARIS.", "", vruta + vcarpeta + vnomfitxer
          eliminar_fitxer vruta + vcarpeta + vnomfitxer
       End If
  Next i
  
  
  'vnomfitxer = Dir(vruta + vcarpeta + "*.*")
  'While vnomfitxer <> ""
  '     vany = Year(FileDateTime(vruta + vcarpeta + vnomfitxer))
  '     If Not existeix(etrutaalbarans + vcarpeta + vany) Then MkDir etrutaalbarans + vcarpeta + vany
  '     Copiar_Fitxer vruta + vcarpeta + vnomfitxer, etrutaalbarans + vcarpeta + vany + "\" '+ vnomfitxer
  '     If existeix(etrutaalbarans + vcarpeta + vany + "\" + vnomfitxer) Then
  '        If vcarpeta = "Tinta\" Then enviar_a_tintes_lalbara vruta + vcarpeta + vnomfitxer
  '       eliminar_fitxer vruta + vcarpeta + vnomfitxer
  '
  '     End If
  '     vnomfitxer = Dir(vruta + vcarpeta + "*.*")
  'Wend
End Sub

Sub comprovar_si_hi_ha_algunenvio(vruta As String)
   Dim vfitxer As String
   vfitxer = Dir(vruta + "\*.vbs")
   If Len(vfitxer) > 10 Then
      r = Shell("c:\windows\system32\cmd.exe /c " + vruta + "\" + vfitxer, vbHide)
      wait 2
      If existeix(vruta + "\" + vfitxer) Then eliminar_fitxer vruta + "\" + vfitxer
   End If
End Sub
Sub comprovarsihihainternet()
   
End Sub
Sub comprovarmodificacionscomandesienviarles()
  Dim rst As Recordset
  obrir_tancar_taules True
  Set rst = dbcomandes.OpenRecordset("SELECT aviscampsmodificats.comanda, Max(aviscampsmodificats.datacanvi) AS ultimcanvi From aviscampsmodificats where not enviat GROUP BY aviscampsmodificats.comanda;")
  While Not rst.EOF
    If DateDiff("n", rst!ultimcanvi, Now) > 59 Then
        enviarcanvisdecomanda rst!comanda
    End If
    rst.MoveNext
  Wend
  Set rst = Nothing
 ' obrir_tancar_taules False
End Sub
Sub enviarcanvisdecomanda(vnumc As Double)
   Dim rst As Recordset
   Dim cos As String
   Dim vdesc As String
   Dim destinatari As String
   Dim rstc As Recordset
   Dim rstmat As Recordset
   
   destinatari = "avisCANVISCOMANDA"
   Set rst = dbcomandes.OpenRecordset("SELECT aviscampsmodificats.comanda, aviscampsmodificats.camp, First(aviscampsmodificats.valorinicial) AS anterior, Last(aviscampsmodificats.valorfinal) AS final From aviscampsmodificats where enviat=false and comanda=" + atrim(vnumc) + " GROUP BY aviscampsmodificats.comanda, aviscampsmodificats.camp;")
   Set rstc = dbcomandes.OpenRecordset("select * from comandesmesextres where comanda=" + atrim(vnumc))
   If Not rstc.EOF Then
     cos = "Client: " + atrim(rstc!nomclient) + Chr(10) + Chr(13)
     cos = cos + "Linia:  " + atrim(rstc!marcailinia) + Chr(10) + Chr(13) + Chr(10) + Chr(13) + Chr(10) + Chr(13)
   End If
   While Not rst.EOF
      vdesc = ""
      If atrim(rst!camp) = "DATAACTIVACIO" Then
         If atrim(rst!anterior) = "" And IsDate(atrim(rst!final)) Then vdesc = " COMANDA ACTIVADA "
         If atrim(rst!final) = "" And IsDate(atrim(rst!anterior)) Then vdesc = " !!!! COMANDA DESACTIVADA !!!!"
      End If
      If atrim(rst!camp) = "MATERIALEX" Then
         Set rstmat = dbcomandes.OpenRecordset("select * from materials ")
         rstmat.FindFirst "codi=" + atrim(cadbl(rst!anterior))
         If Not rstmat.NoMatch Then vdesc = " Material: " + atrim(rstmat!descripcio)
         rstmat.FindFirst "codi=" + atrim(cadbl(rst!final))
         If Not rstmat.NoMatch Then vdesc = vdesc + "   -->  " + atrim(rstmat!descripcio)
         Set rstmat = Nothing
      End If
      If vdesc = "" Then vdesc = UCase(atrim(rst!camp)) + "          Valor anterior: " + atrim(rst!anterior) + "     Valor final: " + atrim(rst!final)
      cos = cos + "Camp modificat:  " + vdesc + Chr(10) + Chr(13)
      rst.MoveNext
      
   Wend
   enviaremail destinatari, "Canvi de la fulla     Comanda: " + atrim(vnumc), cos
   dbcomandes.Execute "update aviscampsmodificats set enviat=true where comanda=" + atrim(vnumc)
   Set rst = Nothing
   Set rstc = Nothing
End Sub
Sub importarelsCSVdelamaquinalaser()
   Dim vruta As String
   Dim vdir As String
   Dim vlinia As String
   Dim rst As Recordset
   Dim vv As String
   Dim dbbaixes As Database
   Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
   Set rst = dbbaixes.OpenRecordset("select * from neteja_aniloxos")
   vruta = "\\serverprodu\Dades\MaquinaLaser"
   vdir = Dir(vruta + "\*.csv")
   While vdir <> ""
      If Mid(vdir, 1, 1) <> "." Then
          rst.FindFirst "fitxercsv='" + atrim(vdir) + "'"
          If rst.NoMatch Then
             Open vruta + "\" + vdir For Input As #1
             Line Input #1, vlinia
             Close #1
             rst.AddNew
             rst!fitxercsv = atrim(vdir)
             vv = Mid(vlinia, 1, InStr(1, vlinia, ";") - 1)
             vlinia = substituir(vlinia, vv + ";", "")
             rst!matricula = vv
             If atrim(rst!matricula) <> "" Then
               vv = Mid(vlinia, 1, InStr(1, vlinia, ";") - 1)
               vlinia = substituir(vlinia, vv + ";", "")
               rst!Datainici = vv
               vv = Mid(vlinia, 1, InStr(1, vlinia, ";") - 1)
               vlinia = substituir(vlinia, vv + ";", "")
               rst!datafi = vv
               vv = Mid(vlinia, 1, InStr(1, vlinia, ";") - 1)
               vlinia = substituir(vlinia, vv + ";", "")
               rst!tipusneteja = vv
               rst!operari = vlinia
               rst.Update
                Else:
                   rst.CancelUpdate
                   GoTo proxim
             End If
          End If
      End If
      FileCopy vruta + "\" + vdir, vruta + "\Actualitzats_a_Producció\" + vdir
proxim:
      eliminar_fitxer vruta + "\" + vdir
      vdir = Dir
   Wend
   Set rst = Nothing
  ' Set dbbaixes = Nothing
End Sub
Sub actualitzant_netejaaniloxos()
   escriure_log "Nejeja anilox importarelsCSVdelamaquinalaser", "c:\temp\Log_EnviarMails_servidor.txt"
   importarelsCSVdelamaquinalaser
   actualitzarCSVnetejalaser_a_aniloxos
End Sub
Sub enviarmissatgenoexisteixmatriculaanilox(rst As Recordset, verrorsmatricula As String)
  '  enviaremail "impresores@inplacsa.com", "    "
  verrorsmatricula = verrorsmatricula + "Matricula: " + atrim(rst!matricula) + "   Data: " + atrim(rst!Datainici) + "   Operari: " + atrim(rst!operari) + Chr(13) + Chr(10)
  
End Sub
Sub actualitzarCSVnetejalaser_a_aniloxos()
   Dim dbbaixes As Database
   Dim rst As Recordset
   Dim rstanilox As Recordset
   Dim verrorsmatricula As String
   Dim vvalues As String
   Set dbcomandes = OpenDatabase(rutadelfitxer(cami) + "comandes.mdb")
   Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
   Set rst = dbbaixes.OpenRecordset("select * from neteja_aniloxos where actualitzataaniloxos=false")
   While Not rst.EOF
     Set rstanilox = dbcomandes.OpenRecordset("select * from aniloxos_informacio where matricula='" + atrim(rst!matricula) + "' order by data")
     If Not rstanilox.EOF Then
         rstanilox.FindFirst "data=#" + Format(rst!Datainici, "mm/dd/yy hh:nn:ss") + "#"
         If rstanilox.NoMatch Then
             vvalues = atrim(rstanilox!idanilox) + ",'" + rstanilox!matricula + "','" + rstanilox!matricula_inplacsa + "',#" + Format(rst!Datainici, "mm/dd/yy hh:nn:ss") + "#,'NETEJA AMB LASER'," + IIf(rstanilox!actiu, "True", "False")
             dbcomandes.Execute "insert into aniloxos_informacio (idanilox,matricula,matricula_inplacsa,data,informacio,actiu) values (" + vvalues + ")"
         End If
           Else: enviarmissatgenoexisteixmatriculaanilox rst, verrorsmatricula
     End If
     rst.Edit
     rst!actualitzataaniloxos = True
     rst.Update
     rst.MoveNext
   Wend
   If verrorsmatricula <> "" Then enviaremail "impresores@inplacsa.com", "Neteja Laser - Matricules d'anilox no trobades.", Chr(13) + Chr(10) + Chr(13) + Chr(10) + verrorsmatricula
   Set rst = Nothing
   'Set dbbaixes = Nothing
End Sub

Sub comprovarsiactualitzaciodeTORERUShapetat()
   Dim vdata As String
   Static vultimadata As String
   vdata = atrim(llegir_ini("Torerus", "horaultimaactualitzacio", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"))
   If vdata = "" Or Not IsDate(vdata) Then Exit Sub
   If DateDiff("n", vdata, Now) > 5 And vultimadata <> vdata Then
       enviaremail "miquel.inplacsa@gmail.com", "ERROR actualització TORERUS", "Mes de 2 minuts d'espera."
       vultimadata = vdata
   End If
End Sub

Sub comprovar_ordres_execucio_servidor()
   Dim rst As Recordset
   Static vincidenciajaenviada As Boolean
   Set rst = db.OpenRecordset("select * from ordres_execucio_servidor order by id")
   While Not rst.EOF
     If rst!funcio = "copiar" Then
        If existeix(rutadelfitxer(rst!desti)) And existeix(rst!Origen) Then
         Copiar_Fitxer rst!Origen, rst!desti
         If existeix(rst!desti) Then rst.Delete: GoTo proxima
        End If
     End If
     If rst!funcio = "eliminar" Then
         If existeix(rst!Origen) Then Kill rst!Origen
         If Not existeix(rst!Origen) Then rst.Delete: GoTo proxima
     End If
     If rst!funcio = "canvinom" Then
         If existeix(rst!Origen) Then
             FileCopy rst!Origen, rst!desti
             Kill rst!Origen
             If existeix(rst!desti) Then rst.Delete: GoTo proxima
         End If
     End If
proxima:
     rst.MoveNext
   Wend
   Set rst = db.OpenRecordset("select * from ordres_execucio_servidor order by id")
   If Not rst.EOF Then
       If Hour(Now) = 0 Then vincidenciajaenviada = False: GoTo fi
       If vincidenciajaenviada = False Then enviaremail "miquel.inplacsa@gmail.com", "Instruccions ordre_execucio_servidor no executades.", "REVISAR LA TAULA ordres_execucio_servidor DE avisosincidencies.mdb"
       vincidenciajaenviada = True
       Else: vincidenciajaenviada = False
   End If
fi:
   Set rst = Nothing
End Sub
Private Sub Timer1_Timer()
   Dim vultimaactualitzacio As String
   Static copiafeta As Boolean
   Static ultimdia As Byte
   Static cont As Byte
   Static vultimtmp As Byte
   Static vultimcalloff As Byte
   Static vultimsensepreu As Byte
   Static vjahaentrat As Boolean
   
   comprovarsihihainternet
   'substituit per l'envio per spooler
   'comprovar_si_hi_ha_algunenvio rutadelfitxer(cami) + "spoolerenviament"
   If vjahaentrat Then Exit Sub
   If cont < 5 Then cont = cont + 1: Exit Sub
   'comprovarsiactualitzaciodeTORERUShapetat
   Me.Tag = ""
   vjahaentrat = True
   comprovar_ordres_execucio_servidor
   
   
   If noenviar.Value = 0 Then
        escriure_log "Timer1 (comprovar_incidencies)", "c:\temp\Log_EnviarMails_servidor.txt"
        comprovar_incidencies
        If cadbl(Format(Now, "hh")) = 1 Then
           If Not copiafeta Then
              If fercopiahistoric Then copiafeta = True
           End If
             Else: copiafeta = False
        End If
   End If
   
   cont = 0
   'fer feines a mitja nit
   escriure_log "Timer1 (ferPing('servidorsap'))", "c:\temp\Log_EnviarMails_servidor.txt"
   If Not ferPing("servidorsap") Then vjahaentrat = False: Exit Sub 'si no hi ha access a servidorsap surtu
   If (ultimdia <> Day(Now) And Hour(Now) = 0) Or checkfertasquesdemitjanit.Value = 1 Then
      escriure_log "Timer1 (mitjanit)", "c:\temp\Log_EnviarMails_servidor.txt"
      Me.Tag = "mitjanit"
      estat.Caption = "Tasques mitjanit..."
      checkfertasquesdemitjanit.Value = 0
      If obrir_tancar_taules(True) Then
       Me.Caption = "Sincronització mestres": DoEvents
        sincronitzar_taulesmestra
        'vincular_factures_amb_albarans
       Me.Caption = "Clients albarans clixes": DoEvents
        revisar_clientsalbaransclixes
       Me.Caption = "Actualitzant netejes d'aniloxos": DoEvents
        actualitzant_netejaaniloxos
       Me.Caption = "Calculant estadistica aniloxos": DoEvents
        calcular_estadisticaaniloxos False 'amb el parametre true no enviarà cap informe a impresores
       Me.Caption = "clients donatsdebaixa amb comandes": DoEvents
        revisar_clients_donatsdebaixa_ambcomandescirculant
       Me.Caption = "Informe tintes noves": DoEvents
        passar_informe_tintesnoves_perrevisar
       Me.Caption = "Resum albarans fotogravadors": DoEvents
        passar_resumalbaransfotogravadors
       Me.Caption = "Comandes desactivades": DoEvents
        passar_informe_comandesdesactivades
       Me.Caption = "call off comandes ja fabricades": DoEvents
        comprovarcalloffdecomandesjafabricades
       Me.Caption = "Estoc minim de llaunes": DoEvents
       ' comprovarestocminimdellaunes
       Me.Caption = "Llaunes dosificadors": DoEvents
        comprovar_lesllaunesdelsdosificadors
       Me.Caption = "comandes sense escanejar": DoEvents
        comprovar_comandesafabricaciosenseescanejar
       comprovar_comandesafabricaciosenseescanejar True
       Me.Caption = "estadistica llaunes": DoEvents
        estadistica_llaunesinplacsa
       ' comprovarestocdeadhesiuamuntadora
       Me.Caption = "Compres amb data d'entrega pasada": DoEvents
        comprovar_compres_datadentregapasada "T"
'desactivat per ordre den rabassedas        comprovar_compres_datadentregapasada "M"
       Me.Caption = "Informe de contenidors per recuperar": DoEvents
        enviarinformedecontenidorsperrecuperar "destinatari2" 'nomes el diumenge
       Me.Caption = "Guardar llista estoc a magatzem": DoEvents
        guardarllistatestocamagatzem
        'Generar els PDF mes petits
       Me.Caption = "Generant els PDF_Editables a mes petits.": DoEvents
        convertirPDFeditablesaPDFpetits
       Me.Caption = "Generant llistat de llaunes 1r de mes"
         llistat_llaunes_1rdemes
       Me.Caption = "Generant llistat de compres pendents 1r de mes"
         llistat_comprespendents_1rdemes
       Me.Caption = "Netejar referencies de tarifes... que no existeixen"
        netejar_referencies_tarifes
       Me.Caption = "Llistat de modificacions de tintes de clixes pendents de revisar. "
        llistat_clixes_tintespendentsderevisar
       Me.Caption = "Actualitzant certificats de Qualitat"
        actualitzar_CQ_lots
       Me.Caption = "Comprovar mesura PVP i mesura quantitat demanada del client coincideixin."
        comprovar_mesuraPVPvsCLIENT
       Me.Caption = "Comprovar firmes pendents de PVP."
       comprovar_firmes_pendents_PVP
        borrar_compres_senselinies
        Me.Caption = "Enviar revisió disposició de materials a la comanda (Doble firma)"
        informe_doblefirma_disposiciomaterials
        Me.Caption = "Enviament d'Incidències": DoEvents
        enviaremail "miquel.inplacsa@gmail.com", "Sincronització SAP (Finalitzada) ", ""
      End If
     ' obrir_tancar_taules False
      ultimdia = Day(Now)
      Me.Tag = ""
      estat.Caption = ""
      escriure_log "Timer1 (mitjanit FI)", "c:\temp\Log_EnviarMails_servidor.txt"
   End If
   'si es hora de generar el temporal de planificació
   If checkplanificaciotmp.Value = 1 Then
       vultimaactualitzacio = llegir_ini("Planificacio", "ultimaactualitzacio", rutadelfitxer(cami) + "\actualitzacioplanificacio.ini")
       If WeekDay(Now, vbMonday) < 6 And Hour(Now) > 5 And Hour(Now) < 18 And (cadbl(Format(Now, "nn")) = 15 Or cadbl(Format(Now, "nn")) = 30 Or cadbl(Format(Now, "nn")) = 45 Or cadbl(Format(Now, "nn")) = 0) Then
          If vultimtmp <> cadbl(Format(Now, "nn")) Then
             If DateDiff("n", vultimaactualitzacio, Now) > 15 And DateDiff("n", vultimaactualitzacio, Now) < 25 Then
                 enviaremail "miquel.inplacsa@gmail.com", "Planificació no s'actualitza correctament. " + atrim(Now), "Planificacio"
             End If
             vultimtmp = cadbl(Format(Now, "nn"))
             
             escriure_log "Timer1 (Generar el temporal de planificació)", "c:\temp\Log_EnviarMails_servidor.txt"
             Shell llegir_ini("General", "rutallistats", "comandes.ini") + "planificacio.exe GENERARFITXERTEMPORAL", vbNormalFocus
             wait 5
          End If
       End If
   End If
   
   'envio un registre de comandes sense preu a les 8 i a les 3pm
   If (Hour(Now) = 15 And Minute(Now) = 0) Or (Hour(Now) = 8 And Minute(Now) = 0) Then
        If Hour(Now) <> vultimsensepreu Then
          mirarcomandesenproducciosensepreu
          mirarcomandesambrefinplacsanovalidades
          vultimsensepreu = Hour(Now)
        End If
   End If
   
   'miro si hi ha calloff pendents a les 8 i a les 12 del migdia també
     'he desactivar l'avís perquè la oana ha dit que no calia tan sovint ara (En teroria arriba el de producció i cada dia a les 12 de la nit)
    'If (Hour(Now) = 12 And Minute(Now) = 0) Or (Hour(Now) = 8 And Minute(Now) = 0) Then
    '   If Hour(Now) <> vultimcalloff Then
    '     vultimcalloff = Hour(Now)
    '     comprovarcalloffdecomandesjafabricades
    '   End If
    ' End If
   
   
   'Actualitza planificacio si algú ho ha demanat
   If llegir_ini("Planificacio", "forzaractualitzacio", rutadelfitxer(cami) + "\actualitzacioplanificacio.ini") = "S" Then
      escriure_log "Timer1 (forzaractualitzacio planificacio)", "c:\temp\Log_EnviarMails_servidor.txt"
      escriure_ini "Planificacio", "forzaractualitzacio", "N", rutadelfitxer(cami) + "\actualitzacioplanificacio.ini"
      Shell llegir_ini("General", "rutallistats", "comandes.ini") + "planificacio.exe GENERARFITXERTEMPORAL", vbNormalFocus
      wait 1
   End If
   
   'tira el llistat de contenidors si algú el demana
   If llegir_ini("accionsglobals", "tirarllistatcontenidors", rutadelfitxer(cami) + "valorsprograma.ini") <> "" Then
      escriure_log "Timer1 (Tira el llistat de contenidors)", "c:\temp\Log_EnviarMails_servidor.txt"
     enviarinformedecontenidorsperrecuperar llegir_ini("accionsglobals", "tirarllistatcontenidors", rutadelfitxer(cami) + "valorsprograma.ini"), True
     escriure_ini "accionsglobals", "tirarllistatcontenidors", "", rutadelfitxer(cami) + "valorsprograma.ini"
   End If
   
   'comprova si exportarcomandes i pdfs a png
   If exportarauto.Value = 1 And bexportarcomandes.BackColor = QBColor(15) Then
      escriure_log "Timer1 (mirarsieshoradexportar)", "c:\temp\Log_EnviarMails_servidor.txt"
      mirarsieshoradexportar
      escriure_log "Timer1 (mirarsieshoradexportar_FI)", "c:\temp\Log_EnviarMails_servidor.txt"
   End If
   If exportarpdfapng.Value = 1 And bexportarcomandes.BackColor = QBColor(15) Then
     escriure_log "Timer1 (mirarsieshoradexportarpdfapng)", "c:\temp\Log_EnviarMails_servidor.txt"
     mirarsieshoradexportarpdfapng
   End If
   
   vjahaentrat = False
End Sub
Sub informe_doblefirma_disposiciomaterials()
   Dim rst As Recordset
   Dim vmsg As String
   Set rst = dbcomandes.OpenRecordset("Select * from referencies_disposiciomaterials where nomverificador='' or nomverificador=null")
   
   While Not rst.EOF
      vmsg = vmsg + atrim(rst!refinplacsa) + " ; "
      rst.MoveNext
   Wend
   If vmsg <> "" Then
       vmsg = "Relació de referencies sense la segona firma: " + vbNewLine + vbNewLine + vmsg
       enviaremailgeneric "amiquel@inplacsa.com", "Relació de disposició de materials a les comandes sense sogona firma. " + Format(Now, "dd/mm/yy"), vmsg
   End If
   Set rst = Nothing
End Sub
Sub comprovar_firmes_pendents_PVP()
  Dim rstf As Recordset
  Dim rstc As Recordset
  Dim vsql As String
  Dim vusuariactual As String
  Dim vmsg As String
  If WeekDay(Now, vbMonday) > 5 Then Exit Sub
  
  'vsql = "SELECT comandes_firmes.comanda, First(comandes_firmes.usuari) AS Pusuari FROM (comandes INNER JOIN comandes_firmes ON comandes.comanda = comandes_firmes.comanda) LEFT JOIN clients ON comandes.client = clients.codi "
  'vsql = vsql + " WHERE (((comandes_firmes.tipus)='PVP') AND ((comandes_firmes.comanda) In (select comanda from comandes where (((comandes.proximaseccio)<>'T') AND ((comandes.producte)<>'PC' And (comandes.producte)<>'PC2' And (comandes.producte)<>'PCP' And (comandes.producte)<>'PCI3' And (comandes.producte)<>'PCP')))) AND ((clients.grupdeclient)<>'ARDO')) "
  'vsql = vsql + " GROUP BY comandes_firmes.comanda Having (((Count(comandes_firmes.usuari)) = 1)) ORDER BY First(comandes_firmes.usuari);"
  'Clipboard.Clear
  'Clipboard.SetText vsql
  vsql = "SELECT comandes_firmes.comanda, First(comandes_firmes.usuari) AS Pusuari, First(clients.grupdeclient) AS nomdelgrupdeclients "
  vsql = vsql + " FROM (comandes INNER JOIN comandes_firmes ON comandes.comanda = comandes_firmes.comanda) LEFT JOIN clients ON comandes.client = clients.codi"
  vsql = vsql + " WHERE (((comandes_firmes.tipus)='PVP') AND ((comandes_firmes.comanda) In (select comanda from comandes where (((comandes.proximaseccio)<>'T') AND ((comandes.producte)<>'PC' And (comandes.producte)<>'PC2' And (comandes.producte)<>'PCP' And (comandes.producte)<>'PCI3' And (comandes.producte)<>'PCP'))))) "
  vsql = vsql + " GROUP BY comandes_firmes.comanda Having (((Count(comandes_firmes.usuari)) = 1)) ORDER BY First(comandes_firmes.usuari);"

  Set rstf = dbcomandes.OpenRecordset(vsql)
  While Not rstf.EOF
       If vusuariactual <> rstf!Pusuari Then
            vusuariactual = rstf!Pusuari
            vmsg = vmsg + vbNewLine + vbNewLine + "Firmes usuari:  " + atrim(vusuariactual) + vbNewLine
       End If
       vmsg = vmsg + "           Comanda: " + atrim(rstf!comanda) + vbNewLine
       rstf.MoveNext
  Wend
  'enviaremailgeneric "miquel.inplacsa@gmail.com", "Llistat de comandes que falta 2a firma PVP. " + Format(Now, "dd/mm/yy"), vmsg
  enviaremailgeneric "incidenciesdePVP", "Llistat de comandes que falta 2a firma PVP. " + Format(Now, "dd/mm/yy"), vmsg
  Set rstf = Nothing
  Set rstc = Nothing
End Sub
Sub borrar_compres_senselinies()
  Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
  dbcompres.Execute "delete * From capcalera WHERE (((capcalera.id) In (SELECT capcalera.id FROM capcalera LEFT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((liniescompra.idliniacompra) Is Null)))) AND ((capcalera.data)<Now()))"
End Sub
Function comprovarrelaciomesuraPVPidemanada(rstc As Recordset) As Boolean
   'relaciomesureslineals
   Dim rst As Recordset
   comprovarrelaciomesuraPVPidemanada = True
   Set rst = dbcomandes.OpenRecordset("select * from mesures where codi=" + atrim(cadbl(rstc!mesurapvp)))
   If Not rst.EOF Then
       If cadbl(rstc!mesuraquantdemanada) <> cadbl(rst!relaciomesureslineals) Then
          comprovarrelaciomesuraPVPidemanada = False
           Else: If cadbl(rst!relaciomesureslineals) = 0 Then comprovarrelaciomesuraPVPidemanada = True
       End If
   End If
   If cadbl(rstc!mesurapvp) = 0 Then comprovarrelaciomesuraPVPidemanada = True
   Set rst = Nothing
End Function
Sub comprovar_mesuraPVPvsCLIENT()
  Dim rst As Recordset
  Dim vmsg As String
  
  Set rst = dbcomandes.OpenRecordset("select * from comandes where producte<>'PC' and producte<>'PCP' and producte<>'PC2' and proximaseccio<>'T'")
  While Not rst.EOF
    If cadbl(rst!mesuraquantdemanada) > 0 Then
        If Not comprovarrelaciomesuraPVPidemanada(rst) Then vmsg = vmsg + "La comanda " + atrim(rst!comanda) + " no coincideix la mesura PVP amb la demanada pel client." + vbNewLine
    End If
    rst.MoveNext
  Wend
  If vmsg <> "" Then enviaremail "comprovarrelaciomesuraPVPidemanada", "Comandes amb mesura de PVP diferent a mesura demanada pel client.", vmsg
  Set rst = Nothing
End Sub
Sub actualitzar_CQ_lots()
  Dim rst As Recordset
  Dim dbcompres As Database
  Dim dbtintes As Database
  Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  Set rst = dbcomandes.OpenRecordset("select distinct numlotproveidor,codiproveidorcomercial from albaransbip where numlotproveidor<>'-' and numlotproveidor <>null and  DateDiff('d',[data],Now())<150")
  While Not rst.EOF
    If atrim(rst!numlotproveidor) <> "" Then
        If calCQ(rst!numlotproveidor, rst!codiproveidorcomercial, dbcompres) Then
            dbcompres.Execute "update albaransbip set cal_CQ_lot=true where numlotproveidor='" + atrim(rst!numlotproveidor) + "' and codiproveidorcomercial=" + atrim(rst!codiproveidorcomercial)
            Else: dbcompres.Execute "update albaransbip set cal_CQ_lot=false where numlotproveidor='" + atrim(rst!numlotproveidor) + "' and codiproveidorcomercial=" + atrim(rst!codiproveidorcomercial)
        End If
    End If
    rst.MoveNext
  Wend
  Set rst = Nothing
  Set dbcompres = Nothing
  Set dbtintes = Nothing
End Sub
Function calCQ(vlot As String, vcodiproveidor As Double, dbcompres As Database) As Boolean
   Dim rst As Recordset
   Dim rstmat As Recordset
   Dim rstprov As Recordset
      'aquesta rutina la faig servir també a palets quan recepciono material SI ES CANVIA AQUI TAMBÉ S'HA DE FER ALLÀ
   calCQ = True
   Set rstprov = dbcomandes.OpenRecordset("SELECT proveidors.tipusCQ, proveidors.dataCQ, proveidors.codi, proveidors_comercial.codicomptable FROM proveidors LEFT JOIN proveidors_comercial ON proveidors.codi = proveidors_comercial.codiproduccio where proveidors_comercial.codicomptable='" + atrim(vcodiproveidor) + "'")
   If rstprov.EOF Then Exit Function
   Set rst = dbcompres.OpenRecordset("SELECT albaransbip.*, liniescompra.tipusmaterialcomprat FROM albaransbip INNER JOIN liniescompra ON albaransbip.idliniacompra = liniescompra.idliniacompra where codiproveidorcomercial=" + atrim(vcodiproveidor) + " and numlotproveidor='" + atrim(vlot) + "'")
   If Not rst.EOF Then
      If rst!tipusmaterialcomprat <> "T" Then
       Set rstmat = dbcomandes.OpenRecordset("select * from materials where codi=" + atrim(rst!article))
       If rstmat.EOF Then Exit Function
       If rstmat!tipusCQ <> "L" Then calCQ = False
       If Not calCQ Then If atrim(rstprov!tipusCQ) = "L" And atrim(rstmat!tipusCQ) <> "N" Then calCQ = True
         Else
           If atrim(rstprov!tipusCQ) <> "L" Then calCQ = False
      End If
   End If
   Set rst = Nothing
   Set rstmat = Nothing
   Set rstprov = Nothing
End Function

Sub llistat_clixes_tintespendentsderevisar()
  Dim rst As Recordset
  Dim dbtintes As Database
  Dim vmsg As String
  Dim vsql As String
  Dim vfitxer As String
  Dim vlinia As String
  Dim vultimn As String
  Dim valbara As String
  Dim vlot As String
  Dim vmatricula As String
  Dim dbcompres As Database
  Dim vrutapdf As String
  ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  
  vfitxer = "C:\temp\cosmissatgerevisiotintesclixes.txt"
  If existeix(vfitxer) Then eliminar_fitxer vfitxer
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  'vsql = "SELECT Modificacions.id_treball, Modificacions.ordre, [marca]+' - '+[linia] AS marcailinia, Modificacions.estatrevisiotintes, InStr(1,[estatrevisiotintes],'OK DISSENY') AS hihaOKDisseny, comandes.comanda, comandes.proximaseccio, clients.nom, InStr(1,[estatrevisiotintes],'+IMP') AS hihaOKIMP, InStr(1,[estatrevisiotintes],'+TIN') AS hihaOKTIN FROM (comandes RIGHT JOIN (Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) ON (comandes.numtreball = Modificacions.id_treball) AND (comandes.numordremodificacio = Modificacions.ordre)) LEFT JOIN clients ON comandes.client = clients.codi WHERE (((InStr(1,[estatrevisiotintes],'OK DISSENY'))=0) AND ((comandes.proximaseccio)='E' Or (comandes.proximaseccio)='I'));"
  'Set rst = dbclixes.OpenRecordset(vsql)
  
  Open vfitxer For Output As #3
  Print #3, "     "
  Print #3, "   Llistat REVISIONS DE TINTES DE CLIXES NOUS/MODIFICATS"
   vlinia = String(70, "=")
   Print #3, vlinia
  Print #3, " "
  Print #3, " "
  vsql = "SELECT Modificacions.id_treball, Modificacions.ordre, [marca]+' - '+[linia] AS marcailinia, Modificacions.estatrevisiotintes, InStr(1,[estatrevisiotintes],'OK DISSENY') AS hihaOKDisseny, comandes.comanda, comandes.proximaseccio, clients.nom, InStr(1,[estatrevisiotintes],'+IMP') AS hihaOKIMP, InStr(1,[estatrevisiotintes],'+TIN') AS hihaOKTIN, InStr(1,[estatrevisiotintes],'DISSENY') AS hihaDISSENY FROM (comandes RIGHT JOIN (Clixes RIGHT JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) ON (comandes.numtreball = Modificacions.id_treball) AND (comandes.numordremodificacio = Modificacions.ordre)) LEFT JOIN clients ON comandes.client = clients.codi WHERE InStr(1,[estatrevisiotintes],'+IMP') =0  and (((InStr(1,[estatrevisiotintes],'OK DISSENY'))=0) AND ((comandes.proximaseccio)='E' Or (comandes.proximaseccio)='I' Or (comandes.proximaseccio) Is Null) AND ((InStr(1,[estatrevisiotintes],'DISSENY'))>0));"

  'vsql = "SELECT Modificacions.id_treball, Modificacions.ordre, [marca]+' - '+[linia] AS marcailinia, Modificacions.estatrevisiotintes, InStr(1,[estatrevisiotintes],'OK DISSENY') AS hihaOKDisseny, comandes.comanda, comandes.proximaseccio, clients.nom, InStr(1,[estatrevisiotintes],'+IMP') AS hihaOKIMP, InStr(1,[estatrevisiotintes],'+TIN') AS hihaOKTIN FROM (comandes RIGHT JOIN (Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) ON (comandes.numtreball = Modificacions.id_treball) AND (comandes.numordremodificacio = Modificacions.ordre)) LEFT JOIN clients ON comandes.client = clients.codi WHERE InStr(1,[estatrevisiotintes],'+IMP') =0 AND  (((InStr(1,[estatrevisiotintes],'OK DISSENY'))=0) AND ((comandes.proximaseccio)='E' Or (comandes.proximaseccio)='I'));"
  Set rst = dbclixes.OpenRecordset(vsql)
  Print #3, "---------  FALTA OK DE IMPRESORES  -----------"
  While Not rst.EOF
    vrutapdf = ruta_documentacio_clixes + "\" + Format(rst!id_treball, "00000") + "\PDF" + Format(rst!id_treball, "00000") + "-" + Format(rst!ordre, "000") + "_PR.pdf"
    If existeix(vrutapdf) Then
        vlinia = justificar(atrim(rst!id_treball), 8, "D") + " " + justificar(atrim(rst!ordre), 4, "E") + "  " + justificar(atrim(rst!marcailinia), 30, "E") + justificar(cadbl(rst!comanda), 10, "D") + "  " + justificar(atrim(rst!nom), 20, "E")
        Print #3, vlinia
    End If
    rst.MoveNext
  Wend
  Print #3, " "
  Print #3, " "
  
  vsql = "SELECT Modificacions.id_treball, Modificacions.ordre, [marca]+' - '+[linia] AS marcailinia, Modificacions.estatrevisiotintes, InStr(1,[estatrevisiotintes],'OK DISSENY') AS hihaOKDisseny, comandes.comanda, comandes.proximaseccio, clients.nom, InStr(1,[estatrevisiotintes],'+IMP') AS hihaOKIMP, InStr(1,[estatrevisiotintes],'+TIN') AS hihaOKTIN, InStr(1,[estatrevisiotintes],'DISSENY') AS hihaDISSENY FROM (comandes RIGHT JOIN (Clixes RIGHT JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) ON (comandes.numtreball = Modificacions.id_treball) AND (comandes.numordremodificacio = Modificacions.ordre)) LEFT JOIN clients ON comandes.client = clients.codi WHERE InStr(1,[estatrevisiotintes],'+TIN') =0  and (((InStr(1,[estatrevisiotintes],'OK DISSENY'))=0) AND ((comandes.proximaseccio)='E' Or (comandes.proximaseccio)='I' Or (comandes.proximaseccio) Is Null) AND ((InStr(1,[estatrevisiotintes],'DISSENY'))>0));"
  Set rst = dbclixes.OpenRecordset(vsql)
  Print #3, "---------  FALTA OK DE TINTES  -----------"
  While Not rst.EOF
    vrutapdf = ruta_documentacio_clixes + "\" + Format(rst!id_treball, "00000") + "\PDF" + Format(rst!id_treball, "00000") + "-" + Format(rst!ordre, "000") + "_PR.pdf"
    If existeix(vrutapdf) Then
      vlinia = justificar(atrim(rst!id_treball), 8, "D") + " " + justificar(rst!ordre, 4, "E") + "  " + justificar(rst!marcailinia, 30, "E") + justificar(cadbl(rst!comanda), 10, "D") + "  " + justificar(atrim(rst!nom), 20, "E")
      Print #3, vlinia
    End If
    rst.MoveNext
  Wend
  Print #3, " "
  Print #3, " "
  
  vsql = "SELECT Modificacions.id_treball, Modificacions.ordre, [marca]+' - '+[linia] AS marcailinia, Modificacions.estatrevisiotintes, InStr(1,[estatrevisiotintes],'OK DISSENY') AS hihaOKDisseny, comandes.comanda, comandes.proximaseccio, clients.nom, InStr(1,[estatrevisiotintes],'+IMP') AS hihaOKIMP, InStr(1,[estatrevisiotintes],'+TIN') AS hihaOKTIN, InStr(1,[estatrevisiotintes],'DISSENY') AS hihaDISSENY FROM (comandes RIGHT JOIN (Clixes RIGHT JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) ON (comandes.numtreball = Modificacions.id_treball) AND (comandes.numordremodificacio = Modificacions.ordre)) LEFT JOIN clients ON comandes.client = clients.codi WHERE InStr(1,[estatrevisiotintes],'+TIN') =0 AND InStr(1,[estatrevisiotintes],'+IMP') =0  and (((InStr(1,[estatrevisiotintes],'OK DISSENY'))=0) AND ((comandes.proximaseccio)='E' Or (comandes.proximaseccio)='I' Or (comandes.proximaseccio) Is Null) AND ((InStr(1,[estatrevisiotintes],'DISSENY'))>0));"
  Set rst = dbclixes.OpenRecordset(vsql)
  Print #3, "---------  FALTA OK DE TINTES I IMPRESORES  -----------"
  While Not rst.EOF
    vrutapdf = ruta_documentacio_clixes + "\" + Format(rst!id_treball, "00000") + "\PDF" + Format(rst!id_treball, "00000") + "-" + Format(rst!ordre, "000") + "_PR.pdf"
    If existeix(vrutapdf) Then
      vlinia = justificar(atrim(rst!id_treball), 8, "D") + " " + justificar(rst!ordre, 4, "E") + "  " + justificar(rst!marcailinia, 30, "E") + justificar(cadbl(rst!comanda), 10, "D") + "  " + justificar(atrim(rst!nom), 20, "E")
      Print #3, vlinia
    End If
    rst.MoveNext
  Wend
  Close #3
  If existeix(vfitxer) Then
      FileCopy vfitxer, "c:\temp\cosmissatge.txt"
      enviaremail "RevisioTintesTreballs", "Llistat REVISIONS DE TINTES DE CLIXES NOUS/MODIFICATS", "c:\temp\cosmissatge.txt"
  End If
fi:
  Set rst = Nothing
 ' Set dbtintes = Nothing
 ' Set dbcompres = Nothing

End Sub
Sub netejar_referencies_tarifes()
   Dim vsql As String
   vsql = "DELETE DISTINCTROW tarifes_referencies.refinplacsa, comandes_extres.refinplacsa, tarifes_referencies.* "
   vsql = vsql + " FROM tarifes_referencies LEFT JOIN comandes_extres ON tarifes_referencies.refinplacsa = comandes_extres.refinplacsa "
   vsql = vsql + " WHERE (((comandes_extres.refinplacsa) Is Null));"
   dbcomandes.Execute vsql
End Sub
Sub llistat_llaunes_1rdemes()
   If cadbl(llegir_ini("accionsglobals", "Llistatllaunes_1r_de_mes", rutadelfitxer(cami) + "valorsprograma.ini")) <> Month(Now) Then
        ShellAndWait """" + llegir_ini("General", "rutallistats", "comandes.ini") + "Manteniment tintes.exe""" + " llistattoteslesllaunes", vbNormalFocus
        If existeix("c:\temp\Llistat_llaunes_primerdemes.pdf") Then
            enviaremail "tintes@inplacsa.com", "Llistat de llaunes a 1r del mes de " + Format(Now, "mmmm"), "Llistat per en Ramon de l'estoc de llaunes de tinta.", "c:\temp\Llistat_llaunes_primerdemes.pdf"
            escriure_ini "accionsglobals", "Llistatllaunes_1r_de_mes", atrim(Month(Now)), rutadelfitxer(cami) + "valorsprograma.ini"
        End If
   End If
End Sub

Sub llistat_comprespendents_1rdemes()
   If cadbl(llegir_ini("accionsglobals", "Comprespendents_1r_de_mes", rutadelfitxer(cami) + "valorsprograma.ini")) <> Month(Now) Then
        ShellAndWait """" + llegir_ini("General", "rutallistats", "comandes.ini") + "compres.exe""" + " llistatcomprespendents", vbNormalFocus
        If existeix("c:\temp\Llistat_comprespendents.pdf") Then
            enviaremail "llistatdecomprespendents1rdemes", "Llistat compres pendents a 1r del mes de " + Format(Now, "mmmm"), "Llistat de les compres pendents de rebre a principi de mes.", "c:\temp\Llistat_comprespendents.pdf"
            escriure_ini "accionsglobals", "Comprespendents_1r_de_mes", atrim(Month(Now)), rutadelfitxer(cami) + "valorsprograma.ini"
        End If
   End If
End Sub
Sub calcular_estadisticaaniloxos(Optional nopassarcorreuaimpresores As Boolean)
   Dim rst As Recordset
   Dim rsta As Recordset
   Dim rstt As Recordset
   Dim vsqlmatricula As String
   Dim vdata As Date
   Dim vtotal As Double
   Dim vtotalparcial As Double
   Dim vobsneteja As String
   Dim vdiesneteja As Double
   Dim venviaremail As Boolean
   Dim vmtrsminimperferneteja As Double
   escriure_log " (Calcular_estadisticaaniloxos) ", "c:\temp\Log_EnviarMails_servidor.txt"
   Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
   obrir_tancar_taules True
   Set rst = dbcomandes.OpenRecordset("select distinct matricula from aniloxos_informacio where actiu=true")
   Open "c:\temp\llistatmissatgeaniloxos.txt" For Output As #5
   While Not rst.EOF
     vdata = 0
     vtotal = 0
     vobsneteja = ""
     vtotalparcial = 0
     'If rst!matricula = "SNR60418-06" Then Stop
     Set rsta = dbcomandes.OpenRecordset("select * from aniloxos_informacio where matricula='" + atrim(rst!matricula) + "' and informacio=""DATA ENTRADA DE L'ANILOX"" order by data desc")
     If Not rsta.EOF Then vdata = rsta!Data
     Set rsta = dbcomandes.OpenRecordset("select * from aniloxos_informacio where matricula='" + atrim(rst!matricula) + "' and informacio='NETEJA AMB LASER' order by data desc")
     If Not rsta.EOF Then vdata = rsta!Data
     If vdata = 0 Then vdata = "01/01/2000"
     If DateDiff("d", vdata, Now) > 200 Then vdata = "01/02/2021"
     Set rsta = dbcomandes.OpenRecordset("select * from aniloxos_informacio where matricula='" + atrim(rst!matricula) + "'")
     If rsta.EOF Then GoTo cont
     vsqlmatricula = "matricula1='" + rst!matricula + "' or matricula2='" + rst!matricula + "' or matricula3='" + rst!matricula + "' or matricula4='" + rst!matricula + "' or "
     vsqlmatricula = vsqlmatricula + " matricula5='" + rst!matricula + "' or matricula6='" + rst!matricula + "' or matricula7='" + rst!matricula + "' or matricula8='" + rst!matricula + "' "
     Set rstt = dbbaixes.OpenRecordset("select * from aniloxtimeline where (" + vsqlmatricula + ")")
     'If rst!matricula = "SNR60418-06" Then Stop
     While Not rstt.EOF
       
       If rstt!matricula1 = rst!matricula Then vtotal = vtotal + cadbl(rstt!totalmetres1)
       If rstt!matricula2 = rst!matricula Then vtotal = vtotal + cadbl(rstt!totalmetres2)
       If rstt!matricula3 = rst!matricula Then vtotal = vtotal + cadbl(rstt!totalmetres3)
       If rstt!matricula4 = rst!matricula Then vtotal = vtotal + cadbl(rstt!totalmetres4)
       If rstt!matricula5 = rst!matricula Then vtotal = vtotal + cadbl(rstt!totalmetres5)
       If rstt!matricula6 = rst!matricula Then vtotal = vtotal + cadbl(rstt!totalmetres6)
       If rstt!matricula7 = rst!matricula Then vtotal = vtotal + cadbl(rstt!totalmetres7)
       If rstt!matricula8 = rst!matricula Then vtotal = vtotal + cadbl(rstt!totalmetres8)
       If rstt!Data > vdata Then
            If rstt!matricula1 = rst!matricula Then vtotalparcial = vtotalparcial + cadbl(rstt!totalmetres1)
            If rstt!matricula2 = rst!matricula Then vtotalparcial = vtotalparcial + cadbl(rstt!totalmetres2)
            If rstt!matricula3 = rst!matricula Then vtotalparcial = vtotalparcial + cadbl(rstt!totalmetres3)
            If rstt!matricula4 = rst!matricula Then vtotalparcial = vtotalparcial + cadbl(rstt!totalmetres4)
            If rstt!matricula5 = rst!matricula Then vtotalparcial = vtotalparcial + cadbl(rstt!totalmetres5)
            If rstt!matricula6 = rst!matricula Then vtotalparcial = vtotalparcial + cadbl(rstt!totalmetres6)
            If rstt!matricula7 = rst!matricula Then vtotalparcial = vtotalparcial + cadbl(rstt!totalmetres7)
            If rstt!matricula8 = rst!matricula Then vtotalparcial = vtotalparcial + cadbl(rstt!totalmetres8)
       End If
       rstt.MoveNext
     Wend
     vdiesneteja = cadbl(rsta!diesneteja)
     'If vdiesneteja = 0 Then vdiesneteja = 30 'poso 30 dies per fer neteja per defecte
     If vdiesneteja = 0 Then GoTo cont 'Si els dies estan a 0 vol dir que es fa neteja cada cop que es treu de màquina
 'poso un minim de 50000 metres per fer neteja per metres sino no val la pena HO HA DIT EN MIQUEL DE IMPRESORES
     vmtrsminimperferneteja = 50000
     If DateDiff("d", vdata, Now) >= vdiesneteja And vtotalparcial > vmtrsminimperferneteja Then vobsneteja = "Neteja - Fa " + atrim(DateDiff("d", vdata, Now)) + " de " + atrim(cadbl(vdiesneteja)) + " dies."
     
  'poso un minim de 50000 metres per fer neteja per metres sino no val la pena HO HA DIT EN MIQUEL DE IMPRESORES
     If cadbl(rsta!metresneteja) > 0 And vtotalparcial > vmtrsminimperferneteja Then If vtotalparcial >= cadbl(rsta!metresneteja) Then vobsneteja = "Neteja - Metres " + atrim(vtotalparcial) + " de " + atrim(cadbl(rsta!metresneteja)) + " Metres."
     
     dbcomandes.Execute "delete * from aniloxos_estadistica where matricula='" + atrim(rsta!matricula) + "'"
     dbcomandes.Execute "insert into aniloxos_estadistica (matricula,metrestotal,metres,observacioneteja) values ('" + atrim(rsta!matricula) + "'," + Trim(Redondejar(vtotal, 0)) + "," + Trim(Redondejar(vtotalparcial, 0)) + ",'" + treure_apostruf(vobsneteja) + "')"
     If vobsneteja <> "" Then venviaremail = True: Print #5, rsta!matricula + " -" + vobsneteja + Chr(13) + Chr(10)
cont:
     rst.MoveNext
   Wend
   Close #5
   If venviaremail And Not nopassarcorreuaimpresores Then
      FileCopy "c:\temp\llistatmissatgeaniloxos.txt", "c:\temp\cosmissatge.txt"
      enviaremail "impresores@inplacsa.com", "Llistat de aniloxos per netejar.", "c:\temp\cosmissatge.txt"
   End If
   escriure_log " (Calcular_estadisticaaniloxos) FI ", "c:\temp\Log_EnviarMails_servidor.txt"
   Set rst = Nothing
   'Set dbbaixes = Nothing
End Sub
Sub guardarllistatestocamagatzem()
  If Not existeix(rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe") Then Exit Sub
   ShellAndWait rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe comandes.ini guardarllistatestoc", vbNormalFocus
  ' MsgBox "fi"
End Sub
Sub estadistica_llaunesinplacsa()
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim vsql As String
  Dim vsql2 As String
  Dim vsqlkgs As String
  Dim dbtintes As Database
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb", , True)
  'comprova si hi ha estadistica d'aquesta setmanta (s'ha demanat que hi hagi un calcul de llaunes d'inplacsa per setmana)
  Set rst = dbtintes.OpenRecordset("select * from estadistica_llaunesinplacsa  WHERE (((DatePart('ww',[dataestadistica]))=DatePart('ww',Now()) And Year([dataestadistica])=Year(Now())));")
  'si aquesta setmana no s 'ha fet l'estadistica la faig ara
  vsql = "SELECT Sum(dadesllaunestotes.capacitatactual) AS Kg, Count(dadesllaunestotes.id) AS Q From dadesllaunestotes Where dadesllaunestotes.capacitatactual>2 and (((dadesllaunestotes.id_bido) = 12 Or (dadesllaunestotes.id_bido) = 18) And ((dadesllaunestotes.activa) <> False) And "
  vsql2 = " GROUP BY dadesllaunestotes.nomproveidor HAVING (((dadesllaunestotes.nomproveidor)='INPLACSA'));"

 
  If rst.EOF Then
     rst.AddNew
     rst!dataestadistica = Now
     ' agrupo de 0 a 10
     vsqlkgs = " ((dadesllaunestotes.capacitatactual) > 0) And ((dadesllaunestotes.capacitatactual) <= 10)) "
     Set rst2 = dbtintes.OpenRecordset(vsql + vsqlkgs + vsql2)
     If Not rst2.EOF Then
        rst![kg-llaunes0-10] = Redondejar(cadbl(rst2!kg), 0)
        rst![unitats-llaunes0-10] = cadbl(rst2!q)
     End If
     ' agrupo de >10 a 15
     vsqlkgs = " ((dadesllaunestotes.capacitatactual) >10) And ((dadesllaunestotes.capacitatactual) <=15)) "
     Set rst2 = dbtintes.OpenRecordset(vsql + vsqlkgs + vsql2)
     If Not rst2.EOF Then
        rst![kg-llaunes10-15] = Redondejar(cadbl(rst2!kg), 0)
        rst![unitats-llaunes10-15] = cadbl(rst2!q)
     End If
     ' agrupo de >15 a 18
     vsqlkgs = " ((dadesllaunestotes.capacitatactual) >15) And ((dadesllaunestotes.capacitatactual) <=18)) "
     Set rst2 = dbtintes.OpenRecordset(vsql + vsqlkgs + vsql2)
     If Not rst2.EOF Then
        rst![kg-llaunes15-18] = Redondejar(cadbl(rst2!kg), 0)
        rst![unitats-llaunes15-18] = cadbl(rst2!q)
     End If
     ' agrupo de >18 a 25
     vsqlkgs = " ((dadesllaunestotes.capacitatactual) > 18)) "
     Set rst2 = dbtintes.OpenRecordset(vsql + vsqlkgs + vsql2)
     If Not rst2.EOF Then
        rst![kg-llaunes18-25] = Redondejar(cadbl(rst2!kg), 0)
        rst![unitats-llaunes18-25] = cadbl(rst2!q)
     End If
     rst.Update
  End If
  Set rst = Nothing
 ' Set dbtintes = Nothing
  
'  If vbidonsde20i25 Then vsqlfiltarvidons = "({tintesreferencies.id_bido}=12 or {tintesreferencies.id_bido}=18) AND "
'  If vnomesdinplacsa Then vsqlfiltarvidons = vsqlfiltarvidons + "({tintesreferencies.nomproveidor}='INPLACSA') and "
'  "{Llaunes.activa} and ({Llaunes.capacitatactual}>" + atrim(vminimkg) + " and {Llaunes.capacitatactual}<" + atrim(vmaximkg) + ")"
  
End Sub
Sub comprovar_comandesafabricaciosenseescanejar(Optional vcomprovarlesT As Boolean)
  Dim rst As Recordset
  Dim rstfirmes As Recordset
  Dim vcarpetaprincipal As String
  Dim vnumc As Double
  Dim v As String
  Dim vcomandessenseescanejar As String
  Dim vmsg As String
  
  If Not existeix(etrutaescanercomandes) Then Exit Sub
  Set dbcomandes = OpenDatabase(cami)
  If vcomprovarlesT Then
      Set rst = dbcomandes.OpenRecordset("SELECT comandes.dataactivacio,comandes.comanda, comandes.proximaseccio, clients.grupdeclient FROM clients RIGHT JOIN comandes ON clients.codi = comandes.client where producte<>'PC' and producte<>'PC2' and producte<>'PCP' and producte<>'PCI3' and client>10 and year(dataactivacio)>year(now)-2 and proximaseccio='T' and ((clients.grupdeclient) Is Null Or (clients.grupdeclient)<>'INPLACSA') ")
        Else
            Set rst = dbcomandes.OpenRecordset("SELECT comandes.dataactivacio,comandes.comanda, comandes.proximaseccio, clients.grupdeclient FROM clients RIGHT JOIN comandes ON clients.codi = comandes.client where producte<>'PC' and producte<>'PC2' and producte<>'PCP' and producte<>'PCI3' and client>10 and year(dataactivacio)>year(now)-2 and proximaseccio<>'E' and proximaseccio<>'T' and proximaseccio<>'P' and ((clients.grupdeclient) Is Null Or (clients.grupdeclient)<>'INPLACSA') ")
  End If
  If Not rst.EOF Then rst.MoveLast: rst.MoveFirst
  While Not rst.EOF
       Me.Caption = atrim(rst.AbsolutePosition) + "/" + atrim(rst.RecordCount): DoEvents
       vnumc = rst!comanda
       Set rstfirmes = dbcomandes.OpenRecordset("select * from comandes_firmes where comanda=" + atrim(vnumc))
       If rstfirmes.EOF Then
            vcarpetaprincipal = "Les_" + atrim(atrim(Int(cadbl(vnumc) / 1000)) + "000")
            v = Dir(etrutaescanercomandes + vcarpetaprincipal + "\" + atrim(vnumc) + "\CM*.*")
            If v = "" Then
                If Work_Days(rst!dataactivacio, Now) > 5 Then
                  vcomandessenseescanejar = vcomandessenseescanejar + " [" + atrim(vnumc) + "]"
                End If
            End If
       End If
       rst.MoveNext
  Wend
  '"incidenciesdePVP"

  vmsg = IIf(vcomprovarlesT, "Comandes passades a T i sense escanejar documentació i sense firmes.", "Comandes a producció sense escanejar documentació")
  If vcomandessenseescanejar <> "" Then enviaremail "ComandesSenseEscanejar", vmsg, "Aquestes comandes no tenen documentació relacionada." + Chr(10) + Chr(13) + Chr(10) + Chr(13) + vcomandessenseescanejar
  Set rst = Nothing
End Sub
Function Work_Days(BegDate As Variant, EndDate As Variant) As Integer
 
 Dim WholeWeeks As Variant
 Dim DateCnt As Variant
 Dim EndDays As Integer
 
 On Error GoTo Err_Work_Days
 
 BegDate = DateValue(BegDate)
 EndDate = DateValue(EndDate)
 WholeWeeks = DateDiff("w", BegDate, EndDate)
 DateCnt = DateAdd("ww", WholeWeeks, BegDate)
 EndDays = 0
 
 Do While DateCnt <= EndDate
 If WeekDay(DateCnt, vbMonday) < 6 Then
 EndDays = EndDays + 1
 End If
 DateCnt = DateAdd("d", 1, DateCnt)
 Loop
 
 Work_Days = WholeWeeks * 5 + EndDays
 
Exit Function
 
Err_Work_Days:
 
 ' If either BegDate or EndDate is Null, return a zero
 ' to indicate that no workdays passed between the two dates.
 
 If err.Number = 94 Then
 Work_Days = 0
 Exit Function
 Else
' If some other error occurs, provide a message.
 MsgBox "Error " & err.Number & ": " & err.Description
 End If
 
End Function
Sub mirarsieshoradexportarpdfapng(Optional vsenseprogramacio As Boolean)
   If bexportarcomandes.Tag = "exportant" Then Exit Sub
  If vsenseprogramacio Then GoTo exportar
  If cadbl(llegir_ini("Exportarpdfsapng", "ultimaexecucio", "enviarservidor.ini")) <> Day(Now) Then
     escriure_ini "Exportarpdfsapng", "ultimaexecucio", atrim(Day(Now)), "enviarservidor.ini"
exportar:
     bexportarcomandes.Tag = "exportant"
     obrir_tancar_taules True
     convertirPDFaPNGdelstreballs
'     obrir_tancar_taules False
     bexportarcomandes.Tag = ""
  End If
   
End Sub
Sub convertirPDFeditablesaPDFpetits()
  Dim rst As Recordset
  Dim id_treball As Double
  Dim vrutapdftreball As String
  Dim vrutapdftreball_Editable As String
  Dim vruta_documentacio_clixes As String
  Dim ordre As Double
  Dim ruta_pngs As String
  Dim vnomfitxerpng As String
 ' Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  estat = "Buscant dades PDF..."
  vruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  DoEvents
  Set rst = dbclixes.OpenRecordset("select * from modificacions where pdfperconvertir=true")
  If rst.EOF Then GoTo fi
  rst.MoveLast
  rst.MoveFirst
  While Not rst.EOF
     id_treball = rst!id_treball
     ordre = rst!ordre
     estat = "comparant PDF... "
     DoEvents
     'vrutapdftreball = vruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordre, "000") + ".pdf"
     vrutapdftreball_Editable = vruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordre, "000") + "_Editable.pdf"
     vrutapdftreball = vruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordre, "000") + ".pdf"
     estat = atrim(rst.AbsolutePosition) + " de " + atrim(rst.RecordCount)
     DoEvents
     If existeix(vrutapdftreball_Editable) Then
       If existeix(vrutapdftreball) Then eliminar_fitxer vrutapdftreball
       If Not existeix(vrutapdftreball) Then
         estat = "Generant PDF del editable del Treball " + atrim(id_treball) + "  " + atrim(rst.AbsolutePosition) + " de " + atrim(rst.RecordCount)
         DoEvents
         generar_PDF vrutapdftreball_Editable, vrutapdftreball
         rst.Edit
         rst!pdfperconvertir = False
         rst.Update
       End If
     End If
     rst.MoveNext
  Wend
fi:
  estat = ""
  Set rst = Nothing
End Sub
Sub convertirPDFaPNGdelstreballs()
  Dim rst As Recordset
  Dim id_treball As Double
  Dim vrutapdftreball As String
  Dim vrutapdftreball_CR As String
  Dim vruta_documentacio_clixes As String
  Dim ordre As Double
  Dim ruta_pngs As String
  Dim vnomfitxerpng As String
  estat = "Buscant dades..."
  DoEvents
  ruta_pngs = llegir_ini("ruta", "ruta_pdf_a_png", rutadelfitxer(cami) + "valorsprograma.ini")
  vruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  Set rst = dbclixes.OpenRecordset("select * from modificacions where datapdf>#" + Format(DateAdd("yyyy", -2, Now), "mm/dd/yy") + "#")
  rst.MoveLast
  rst.MoveFirst
  While Not rst.EOF
     id_treball = rst!id_treball
     ordre = rst!ordre
     estat = "comparant... "
     DoEvents
     'vrutapdftreball = vruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordre, "000") + ".pdf"
     vrutapdftreball_CR = vruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\PDF" + Format(id_treball, "00000") + "-" + Format(ordre, "000") + "_CR.pdf"
     vrutapdftreball = vrutapdftreball_CR
     estat = atrim(rst.AbsolutePosition) + " de " + atrim(rst.RecordCount)
     DoEvents
     If existeix(vrutapdftreball_CR) Then
       If existeix(vrutapdftreball) Then
        vnomfitxerpng = ruta_pngs + "\" + Format(id_treball, "00000") + "-" + Format(ordre, "00") + ".png"
        If Not existeix(vnomfitxerpng) Then
         estat = "Generant PNG Treball " + atrim(id_treball) + "  " + atrim(rst.AbsolutePosition) + " de " + atrim(rst.RecordCount)
         DoEvents
         generar_PNG vrutapdftreball, vnomfitxerpng
        End If
       End If
     End If
     
     rst.MoveNext
  Wend
  estat = ""
  Set rst = Nothing
End Sub
Sub generar_PDF(vfitxerorigen As String, vfitxerdesti As String)
   Dim vgif As String
   vgif = Mid(vfitxerdesti, 1, InStr(1, UCase(vfitxerdesti), ".PDF") - 1) + ".gif"
   ShellAndWait "magick.exe convert -density 100 " + vfitxerorigen + " " + vgif
   ShellAndWait "magick.exe convert " + vgif + " " + vfitxerdesti
  eliminar_fitxer vgif
End Sub

Sub generar_PNG(vfitxerorigen As String, vfitxerdesti As String)
   ShellAndWait "magick.exe convert -density 150 -quality 95 " + vfitxerorigen + " " + vfitxerdesti
End Sub
Sub enviarinformedebobinessensenumerodepalet()
  Dim rst As Recordset
  Dim vmsg As String
  Dim vsql As String
  Exit Sub
  vsql = "SELECT distinct comandes.proximaseccio, bobinesent.entregat, comandes.comanda, bobinesent.numpalet, comandes.client "
  vsql = vsql + " FROM bobinesent INNER JOIN comandes ON bobinesent.comanda = comandes.comanda "
  vsql = vsql + " WHERE (((comandes.proximaseccio)<>'T') AND ((bobinesent.entregat)='' Or (bobinesent.entregat) Is Null) AND ((bobinesent.numpalet)=0) AND ((comandes.client)<>1 And (comandes.client)<>6335));"
 ' obrir_tancar_taules True
  Set rst = dbcomandes.OpenRecordset(vsql)
  While Not rst.EOF
     vmsg = vmsg + "[" + atrim(rst!comanda) + "] "
     rst.MoveNext
  Wend
  If vmsg <> "" Then enviaremail "odamian@gmail.com", "Llistat comandes amb bobines sense numero de palet.", "Comandes amb bobines d'entrega sense numero de palet:" + Chr(13) + Chr(10) + vmsg + Chr(13) + Chr(10) + "S'HAURIA D'ANAR A POSSAR-HI EL PALET CORRESPONENT."
  Set rst = Nothing
End Sub
Sub enviarinformedecontenidorsperrecuperar(vdestinatari As String, Optional noprogramat As Boolean)
  Dim rst As Recordset
  Dim dbtintes As Database
  Dim vmsg As String
  Dim vsql As String
  Dim vfitxer As String
  Dim vlinia As String
  Dim vultimn As String
  Dim valbara As String
  Dim vlot As String
  Dim vmatricula As String
  Dim dbcompres As Database
  
  If Not noprogramat Then If WeekDay(Now, vbMonday) <> 7 Then vdestinatari = ""
  vfitxer = "C:\temp\cosmissatgecontenidors.txt"
  If existeix(vfitxer) Then eliminar_fitxer vfitxer
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
  vsql = "SELECT Llaunes.numllauna, tipusbidons.capacitat, Contenidors_material.descripcio AS nomcontenidor, Llaunes.activa, recuperadorsdecontenidors.nomcomercial FROM recuperadorsdecontenidors RIGHT JOIN (((Llaunes LEFT JOIN tintesreferencies ON Llaunes.id_refproveidor = tintesreferencies.id) LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id) LEFT JOIN Contenidors_material ON Llaunes.idmaterialcontenidor = Contenidors_material.codi) ON recuperadorsdecontenidors.Id = Llaunes.idproveidorrecuperador WHERE (((Contenidors_material.descripcio)<>'') and llaunes.activa=false and llaunes.situacio<>'REC') and ((Llaunes.numllauna) Not In (select numllauna from linies_contenidors)) order by recuperadorsdecontenidors.nomcomercial;"
  Set rst = dbtintes.OpenRecordset(vsql)
  If rst.EOF Then GoTo fi
  Open vfitxer For Output As #3
  Print #3, "     "
  Print #3, "   Llistat de contenidors buits"
  Print #3, "     "
  Print #3, "     "
  vlinia = justificar(" Llauna ", 7, "E") + justificar(" Capacitat", 10, "E") + justificar(" Matricula", 20, "E") + justificar(" Contenidor", 30, "E")
   Print #3, vlinia
   vlinia = String(70, "=")
   Print #3, vlinia
   vultimn = "---"
  ' MsgBox rst.RecordCount
  While Not rst.EOF
    If atrim(rst!nomcomercial) <> vultimn Then
        Print #3, "     "
        vultimn = atrim(rst!nomcomercial)
        Print #3, " Recuperador:  " + atrim(rst!nomcomercial)
    End If
    buscarinformaciodelallauna atrim(rst!numllauna), valbara, vlot, vmatricula, dbtintes, dbcompres
    vlinia = justificar(atrim(rst!numllauna), 8, "E") + justificar("   " + atrim(valbara), 15, "E") + justificar("   " + atrim(vlot), 25, "E") + justificar("   " + atrim(vmatricula), 20, "E") + justificar(treuresimbols(atrim(rst!nomcontenidor)), 50, "E")
    Print #3, vlinia
    rst.MoveNext
  Wend
  Close #3
  If existeix(vfitxer) Then
      FileCopy vfitxer, "c:\temp\cosmissatge.txt"
      If vdestinatari <> "" Then enviaremail vdestinatari, "Llistat de contenidors buits.", "c:\temp\cosmissatge.txt"
      If Not noprogramat Then
         FileCopy vfitxer, "c:\temp\cosmissatge.txt"
         enviaremail "LlistatcontenidorsBuitsTintes", "Llistat de contenidors buits.", "c:\temp\cosmissatge.txt"
      End If
  End If
fi:
  Set rst = Nothing
 ' Set dbtintes = Nothing
 ' Set dbcompres = Nothing
End Sub
Function buscaralbproveidor(valbprov As String, dbcompres As Database) As String
  Dim rst As Recordset
  Set rst = dbcompres.OpenRecordset("select numalbaraprov,data from albaransbip where numlotproveidor ='" + atrim(valbprov) + "'")
  If Not rst.EOF Then buscaralbproveidor = "A: " + atrim(rst!numalbaraprov) '+ " " + atrim(rst!Data)
  Set rst = Nothing
End Function
Sub buscarinformaciodelallauna(vnumllauna As String, valbara As String, vlot As String, vmatricula As String, dbtintes As Database, dbcompres As Database)
   Dim rsthistoria As Recordset
   valbara = ""
   vlot = ""
   vmatricula = ""
   Set rsthistoria = dbtintes.OpenRecordset("SELECT Llaunes.id, Llaunes.numllauna,llaunes.vmatriculacontenidor, historiallauna.data, historiallauna.id, historiallauna.tipusmoviment, historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) LEFT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) LEFT JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent Where (numllauna='" + atrim(vnumllauna) + "') ORDER BY historiallauna.data DESC;")
   While Not rsthistoria.EOF
       If rsthistoria!tipusmoviment = "C" Then vlot = atrim(rsthistoria!numlotbase)
       vmatricula = IIf(IsNull(rsthistoria!vmatriculacontenidor), "", rsthistoria!vmatriculacontenidor)
       rsthistoria.MoveNext
   Wend
   valbara = buscaralbproveidor(vlot, dbcompres)
   vlot = "L: " + vlot
   If vmatricula <> "" Then vmatricula = "M: " + vmatricula
   Set rsthistoria = Nothing
End Sub
Sub exportarinformaciodesactivades()
   Dim vfitxer As String
   Dim rst As Recordset
   Dim vlinia As String
   vfitxer = "C:\temp\cosmissatge.txt"
   If existeix(vfitxer) Then eliminar_fitxer vfitxer
  ' Set dbcomandes = OpenDatabase(cami)
   Set rst = dbcomandes.OpenRecordset("select * from informaciodesactivades where actiu order by data")
   If rst.EOF Then GoTo fi
   Open vfitxer For Output As #3
   vlinia = justificar("  Data ", 7, "E") + justificar(" Comanda/Referència", 30, "E") + justificar("       Nom del client", 30, "E") + justificar("         Descripció", 30, "E")
   Print #3, vlinia
   vlinia = String(70, "=")
   Print #3, vlinia
   While Not rst.EOF
      vlinia = justificar(Format(rst!Data, "dd/mm/yy"), 10, "E") + justificar(treuresimbols(atrim(rst!comandaoreferencia)), 30, "E") + justificar(treuresimbols(atrim(rst!nomclient)), 30, "E") + justificar(treuresimbols(atrim(rst!descripcio)), 60, "E")
      Print #3, vlinia
      rst.MoveNext
   Wend
   Close #3
fi:
   Set rst = Nothing
   'dbcomandes.Close
   'Set dbcomandes = Nothing
   'If existeix(vfitxer) Then obrir_document vfitxer
End Sub
Function justificar(v As String, longitut As Integer, DoE As String, Optional vSimbolRelleno As String) As String
  Dim vcaracter  As String
    v = Mid(v, 1, longitut)
    vcaracter = Mid(atrim(vSimbolRelleno) + " ", 1, 1)
    If DoE = "E" Then
       v = v + String(longitut - Len(v), vcaracter)
      Else: v = String(longitut - Len(v), vcaracter) + v
    End If
    justificar = v
End Function


Sub passar_informe_comandesdesactivades()
  If WeekDay(Now, vbMonday) Mod 2 <> 0 And WeekDay(Now, vbMonday) <> 7 Then
    exportarinformaciodesactivades
    If existeix("C:\temp\cosmissatge.txt") Then
       '"llistatdecomandesdesactivades"
      enviaremail "llistatdecomandesdesactivades", "Llistat de comandes desactivades.", "c:\temp\cosmissatge.txt"
    End If
  End If
End Sub
Sub comprovar_compres_datadentregapasada(vseccio As String)
  Dim rst As Recordset
  Dim vnomllistat As String
  Dim dbcompres As Database
  vnomllistat = "Relació de compres " + IIf(vseccio = "T", " de tinta", "") + " que ja haurien d'haver arribat."
  Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb", , True)
  Set rst = dbcompres.OpenRecordset("select * from [compres pendents de rebre] where " + IIf(vseccio <> "T", "", "tipusmaterialcomprat='T' and") + " dataentrega<format(now,'dd/mm/yy')", , True)
   If rst.EOF Then GoTo fi
   'creo el fitxer de cos de missatge
   Open "c:\temp\cosmissatge.txt" For Output As #2
   Print #2, " "
   Print #2, " "
   Print #2, vnomllistat
   Print #2, " "
   Print #2, " "
   Print #2, "Data_Com Data_prev NºComanda Proveïdor                    Nom del material comprat"
   Print #2, "====================================================================================="
   While Not rst.EOF
     Print #2, atrim(Format(rst!Data, "dd/mm/yy")) + "  " + atrim(Format(rst!dataentrega, "dd/mm/yy")) + " " + atrim(rst!numcomanda) + "  - " + atrim(rst!nomprovcomercial) + "       Nom material: " + atrim(rst!nommaterial)
     rst.MoveNext
   Wend
   Close #2
   If vseccio = "T" Then
      If existeix("c:\temp\cosmissatge2.txt") Then eliminar_fitxer "c:\temp\cosmissatge2.txt"
      If existeix("c:\temp\cosmissatge.txt") Then Copiar_Fitxer "c:\temp\cosmissatge.txt", "c:\temp\cosmissatge2.txt"
      enviaremail "controlestoctintes", vnomllistat, "c:\temp\cosmissatge.txt"
      'l´esaú va dir que també se li envies a en miquel per ordre de l´alicia i en miralles
      If existeix("c:\temp\cosmissatge.txt") Then eliminar_fitxer "c:\temp\cosmissatge.txt"
      If existeix("c:\temp\cosmissatge2.txt") Then Copiar_Fitxer "c:\temp\cosmissatge2.txt", "c:\temp\cosmissatge.txt"
      enviaremail "llistatdetintesquehauriendhaverarribat", vnomllistat, "c:\temp\cosmissatge.txt"
         Else: enviaremail "destinatari2", vnomllistat, "c:\temp\cosmissatge.txt"
   End If
   
fi:
 ' Set dbcompres = Nothing
  Set rst = Nothing
End Sub
Sub comprovarestocminimdellaunes()
  Dim rst As Recordset
  Dim dbtintes As Database
  Dim vmsg As String
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  actualitzar_estocactual dbtintes
  
  Set rst = dbtintes.OpenRecordset("SELECT * from estocsminims")
  vmsg = ""
  While Not rst.EOF
    If cadbl(rst!estocminim) > cadbl(rst!estocactual) Then
      vmsg = vmsg + crearnomdelatinta(dbtintes, rst) + " ---> Actual " + atrim(rst!estocactual) + " Kg / Mínim " + atrim(rst!estocminim) + " Kg (Desitjat " + atrim(cadbl(rst!estocdesitjat)) + " Kg - " + atrim(rst!estocactual) + "Kg = " + atrim(cadbl(rst!estocdesitjat) - cadbl(rst!estocactual)) + " Kg amb " + atrim(rst!descripciobido) + ")" + Chr(10)
    End If
    rst.MoveNext
  Wend
  If vmsg <> "" Then enviaremail "controlestoctintes", "Control estoc mínim de llaunes", vmsg
fi:
  Set rst = Nothing
  'Set dbtintes = Nothing
End Sub
Function crearnomdelatinta(dbtintes As Database, rst As Recordset)
   Dim rstt As Recordset
   Dim vsql As String
   Dim vwhere As String
   Dim rstt2 As Recordset
   If cadbl(rst!codi) > 0 Then
      Set rstt = dbtintes.OpenRecordset("select descripcio from tintes where codi='" + atrim(rst!codi) + "'")
      If Not rstt.EOF Then crearnomdelatinta = atrim(rstt!descripcio)
       Else
         vsql = "SELECT familiestintes.descripcio, subfamiliestintes.descripcio, familiescolors.descripcio, subfamiliescolors.descripcio FROM (((estocsminims INNER JOIN familiestintes ON estocsminims.idfamilia = familiestintes.codi) INNER JOIN subfamiliestintes ON estocsminims.idsubfamilia = subfamiliestintes.codi) INNER JOIN familiescolors ON estocsminims.idfamcolor = familiescolors.codi) INNER JOIN subfamiliescolors ON estocsminims.idsubfamcolor = subfamiliescolors.codi "
         With rst
          vwhere = " where (idfamilia=" + atrim(cadbl(!idfamilia)) + " and idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + "and idfamcolor= " + atrim(cadbl(!idfamcolor)) + " and idsubfamcolor=" + atrim(cadbl(!idsubfamcolor)) + ") "
         End With
          Set rstt = dbtintes.OpenRecordset(vsql + vwhere)
          Set rstt2 = dbtintes.OpenRecordset("select codi,descripcio from tintes " + vwhere)
          If Not rstt.EOF Then crearnomdelatinta = IIf(Not rstt2.EOF, atrim(rstt2!codi) + " - " + atrim(rstt2!descripcio) + " -> ", "") + atrim(rstt![familiestintes.descripcio]) + "  " + atrim(rstt![subfamiliestintes.descripcio]) + "  " + atrim(rstt![familiescolors.descripcio]) + "  " + atrim(rstt![subfamiliescolors.descripcio])
   End If
   
   
   Set rstt = Nothing
End Function
Sub actualitzar_estocactual(dbtintes As Database)
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from estocsminims")
   While Not rst.EOF
      rst.Edit
      rst!estocactual = calcular_estoc_delatinta(dbtintes, rst)
      rst.Update
      rst.MoveNext
   Wend
   Set rst = Nothing
   
End Sub
Function calcular_estoc_delatinta(dbtintes As Database, rst As Recordset) As Double
   Dim rstestoc As Recordset
   Dim vsubconsulta As String
   Dim vwhere As String
   'aquesta funcio també s'utilitza a enviar mail servidor i a manteniment tintes
     ' QUALSEVOL CANVI S'HA D'APLICAR A ALS DOS
   With rst
   If cadbl(!codi) > 0 Then
      vwhere = "codi='" + atrim(rst!codi) + "'"
        Else
         vwhere = " (idfamilia=" + atrim(cadbl(!idfamilia)) + " and idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + "and idfamcolor= " + atrim(cadbl(!idfamcolor)) + " and idsubfamcolor=" + atrim(cadbl(!idsubfamcolor)) + ") "
   End If
   End With
   vsubconsulta = "select idtinta from tintes where " + vwhere
   Set rstestoc = dbtintes.OpenRecordset("SELECT Count(*) AS Tllaunes, Sum(Llaunes.capacitatactual) AS SumaDecapacitatactual, tipusbidons.capacitat FROM Llaunes LEFT JOIN (tipusbidons RIGHT JOIN tintesreferencies ON tipusbidons.id = tintesreferencies.id_bido) ON Llaunes.id_refproveidor = tintesreferencies.id  Where (((Llaunes.capacitatactual) > 0.9) And ((Llaunes.activa) = True))  and tipusbidons.nominterndelbido='" + atrim(rst!descripciobido) + "' and Llaunes.idtinta in (" + vsubconsulta + ")  GROUP BY  tipusbidons.capacitat;")
   Clipboard.SetText "SELECT Count(*) AS Tllaunes, Sum(Llaunes.capacitatactual) AS SumaDecapacitatactual, tipusbidons.capacitat FROM Llaunes LEFT JOIN (tipusbidons RIGHT JOIN tintesreferencies ON tipusbidons.id = tintesreferencies.id_bido) ON Llaunes.id_refproveidor = tintesreferencies.id  Where (((Llaunes.capacitatactual) > 0.9) And ((Llaunes.activa) = True))  and tipusbidons.nominterndelbido='" + atrim(rst!descripciobido) + "' and Llaunes.idtinta in (" + vsubconsulta + ")  GROUP BY  tipusbidons.capacitat;"
   calcular_estoc_delatinta = cadbl(rstestoc!SumaDecapacitatactual)
End Function
Sub passar_resumalbaransfotogravadors()
   Dim rst As Recordset
   Dim vregistre As String
   Dim vnomempresa As String
   Dim vconsulta As String
   Dim vdatai As Date
   Dim vdataf As Date
   Dim vmesanterior As Date
   Dim vultimdiamesanterior As Date
   Dim vlinia As String
   Dim ultimdiaresum As Integer
   ultimdiaresum = cadbl(llegir_ini("General", "ultimdiaresum", "enviarservidor.ini"))
   'If ultimdiaresum = 0 Then ultimdiaresum = 6
   'shauria de passar el llistat a dia 1 i dia 15 pero deixem 5 dies de marge per entrar albarans de fotogravador
   If (Day(Now) > 6 And Day(Now) < 21) Then
     If ultimdiaresum = 6 Then Exit Sub
      Else
        If (Day(Now) > 21) Then
             If ultimdiaresum = 21 Then Exit Sub
        End If
   End If
   If Day(Now) < 21 Then
      escriure_ini "General", "ultimdiaresum", 6, "enviarservidor.ini"
       Else: escriure_ini "General", "ultimdiaresum", 21, "enviarservidor.ini"
   End If
   ultimdiaresum = cadbl(llegir_ini("General", "ultimdiaresum", "enviarservidor.ini"))
   vmesanterior = DateAdd("m", -1, Now)
   vultimdiamesanterior = DateAdd("d", -1, CVDate("01" + Format(Now, "/mm/yy")))
   If ultimdiaresum = 21 Then
        vdatai = CVDate("01/" + Format(Now, "mm/yy"))
        vdataf = CVDate("15/" + Format(Now, "mm/yy"))
   End If
   If ultimdiaresum = 6 Then
        vdatai = CVDate("16/" + Format(vmesanterior, "mm/yy"))
        vdataf = CVDate(vultimdiamesanterior)
   End If
        
   vconsulta = "SELECT Modificacions.id_treball, Modificacions.ordre, fotogravadors.nomfotogravador,Modificacions.codiclientfactclixes, Modificacions.empresafacturadora, Clixes_albarans.data, Clixes_albarans.num_alb, Clixes_detallsalb.descripcio, Clixes_albarans.quantitat, Clixes_albarans.import"
   vconsulta = vconsulta + " FROM ((Clixes_albarans INNER JOIN Modificacions ON (Clixes_albarans.ordremodificacio = Modificacions.ordre) AND (Clixes_albarans.id_treball = Modificacions.id_treball)) INNER JOIN Clixes_detallsalb ON Clixes_albarans.id_detall = Clixes_detallsalb.id_detall) INNER JOIN Fotogravadors ON Modificacions.fotograbador = Fotogravadors.codi "
   vconsulta = vconsulta + " WHERE (((Clixes_albarans.data) >= #" + Format(vdatai, "mm/dd/yy") + " 00:00:00# And (Clixes_albarans.data)<=#" + Format(vdataf, "mm/dd/yy") + " 23:59:59#));"
   'InputBox "a", "a", vconsulta
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
   If existeix("c:\temp\resumalbaransclixes.csv") Then eliminar_fitxer "c:\temp\resumalbaransclixes.csv"
   Open "c:\temp\resumalbaransclixes.csv" For Output As #3
   Set rst = dbclixes.OpenRecordset(vconsulta)
   vlinia = "Treball;Versió;Codi Client Facturació;Inplacsa/Plasel;Nom Fotogravador;Data;NºAlbarà;Descripció;Quantitat;Import"
   Print #3, vlinia
   While Not rst.EOF
      vlinia = atrim(rst!id_treball) + ";" + atrim(rst!ordre) + ";" + atrim(rst!codiclientfactclixes) + ";" + atrim(rst!empresafacturadora) + ";" + atrim(rst!nomfotogravador) + ";" + atrim(rst!Data) + ";" + atrim(rst!num_alb) + ";" + atrim(rst!descripcio) + ";" + atrim(rst!quantitat) + ";" + atrim(rst!import)
      Print #3, vlinia
      rst.MoveNext
   Wend
   vconsulta = "SELECT reposicionsfotogravador.id_treball, reposicionsfotogravador.ordremodificacio, Fotogravadors.nomfotogravador, reposicionsfotogravador.dataalbara, reposicionsfotogravador.num_alb, reposicionsfotogravador.descripcio, reposicionsfotogravador.preu"
   vconsulta = vconsulta + " FROM reposicionsfotogravador INNER JOIN (Modificacions LEFT JOIN Fotogravadors ON Modificacions.fotograbador = Fotogravadors.codi) ON (reposicionsfotogravador.ordremodificacio = Modificacions.ordre) AND (reposicionsfotogravador.id_treball = Modificacions.id_treball) "
   vconsulta = vconsulta + " WHERE (((reposicionsfotogravador.dataalbara) >= #" + Format(vdatai, "mm/dd/yy") + " 00:00:00# And (reposicionsfotogravador.dataalbara)<=#" + Format(vdataf, "mm/dd/yy") + " 23:59:59#));"
   Set rst = dbclixes.OpenRecordset(vconsulta)
   While Not rst.EOF
      vlinia = atrim(rst!id_treball) + ";" + atrim(rst!ordremodificacio) + ";0;N;" + atrim(rst!nomfotogravador) + ";" + atrim(rst!Dataalbara) + ";" + atrim(rst!num_alb) + ";" + atrim(rst!descripcio) + ";1;" + atrim(rst!preu)
      Print #3, vlinia
      rst.MoveNext
   Wend
   Close #3
   Set rst = Nothing
   'Set dbclixes = Nothing
'   MsgBox vlinia
   enviaremail "incidenciesillistatsSAPcomptabilitat", treuresimbols("Resum albarans de fotogravadors periode " + Format(vdatai, "dd/mm/yy") + " - " + Format(vdataf, "dd/mm/yy")), treuresimbols("Adjunt amb aquest E-Mail passem relació dels albarans entrats a disseny per tal de revisar-los amb les factures que enviin els fotogravadors corresponents."), "c:\temp\resumalbaransclixes.csv"
End Sub
Sub passar_informe_tintesnoves_perrevisar()
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim dbtintes As Database
   Dim cosmissatge As String
   Dim destinatari As String
   Dim vformula As String
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
   Set rst = dbtintes.OpenRecordset("SELECT * from tintes_tot where not ok_revisat")
   While Not rst.EOF
     vformula = ""
     Set rst2 = dbtintes.OpenRecordset("SELECT tintesformules.numformula, tintes_tot.codi FROM tintes_tot RIGHT JOIN tintesformules ON tintes_tot.idtinta = tintesformules.idtinta where codi='" + atrim(rst!codi) + "'")
     If Not rst2.EOF Then
       vformula = atrim(rst2!numformula)
        Else: vformula = "(S/F)"
     End If
     cosmissatge = cosmissatge + atrim(rst!codi) + " " + atrim(rst!descripcio) + Chr(10) + " RefColor: " + atrim(rst!referenciacolor) + " Serie: " + atrim(rst!descripcioserie) + Chr(10) + " Fam/Sub: " + atrim(rst!descripciofam) + " - " + atrim(rst!descripciosubfam) + Chr(10) + "Fam/Sub Color: " + atrim(rst!descripciofamcol) + " - " + atrim(rst!descripciosubfamcol) + Chr(10) + "Ref: " + atrim(rst!refproveidor) + Chr(10) + "Formula: " + vformula + Chr(10) + Chr(10)
     rst.MoveNext
   Wend
   If cosmissatge <> "" Then
    destinatari = llegir_ini("destinataris", "destinatarirevisartintes", "enviarservidor.ini")
    enviaremail destinatari, "Informe de tintes creades noves que s'han de revisar", cosmissatge
    dbtintes.Execute "update tintes set ok_revisat=true"
   End If
   Set rst = Nothing
   Set rst2 = Nothing
   'Set dbtintes = Nothing
End Sub
Sub revisar_metresicanutorebobinadora_ambcomandescirculant()
   Dim rst As Recordset
   Dim cosmissatge As String
   Dim destinatari As String
   obrir_tancar_taules True
   'Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda, comandes.proximaseccio, clients_codisSAP.codiSAP FROM (comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda) LEFT JOIN clients_codisSAP ON comandes_extres.codicomptable = clients_codisSAP.codiSAP Where (((comandes.comanda) > 165000) And ((comandes.proximaseccio) <> 'T') And ((clients_codisSAP.codiSAP) Is Null)) ORDER BY clients_codisSAP.codiSAP")
   Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda,comandes.tubbase, comandes.mtrslinbob, clients.nom, comandes.refclient, comandes.marcailinia, InStr(1,[productes].[ruta],'R') AS hihareb, comandes.proximaseccio FROM (comandes INNER JOIN clients ON comandes.client = clients.codi) INNER JOIN productes ON comandes.producte = productes.codi where InStr(1,[productes].[ruta],'R')>0 and proximaseccio<>'E' and proximaseccio<>'P' and proximaseccio<>'V' and proximaseccio<>'T'")
   While Not rst.EOF
     If cadbl(rst!tubbase) = 0 Or cadbl(rst!mtrslinbob) = 0 Then
       If cosmissatge = "" Then cosmissatge = "Error en el camp Canutu o Metres bobina de la secció de Rebobinadora" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
       cosmissatge = cosmissatge + "Comanda: " + atrim(rst!comanda) + " - " + atrim(rst!nom) + " Ref_Cli: " + atrim(rst!refclient) + " Texte: " + atrim(rst!marcailinia) + Chr(10) + Chr(13)
     End If
     rst.MoveNext
   Wend
   If cosmissatge <> "" Then
    enviaremail "incidencies@inplacsa.com", "Comandes en fabricació sense Canuto o metres bobina a REBOBINADORA (Arreglar-ho)", cosmissatge
   End If
   Set rst = Nothing
End Sub
Sub revisar_clients_donatsdebaixa_ambcomandescirculant()
   Dim rst As Recordset
   Dim cosmissatge As String
   Dim destinatari As String
   
   'Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda, comandes.proximaseccio, clients_codisSAP.codiSAP, clients.nom, comandes_extres.codicomptable, comandes.producte FROM ((comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda) LEFT JOIN clients_codisSAP ON comandes_extres.codicomptable = clients_codisSAP.codiSAP) LEFT JOIN clients ON comandes.client = clients.codi WHERE (((comandes.comanda)>165000) AND ((comandes.proximaseccio)<>'T') AND ((clients_codisSAP.codiSAP) Is Null) AND ((comandes.producte)<>'PC' And (comandes.producte)<>'PC2' And (comandes.producte)<>'PCP')) ORDER BY clients_codisSAP.codiSAP;")
   
   
   Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda, comandes.proximaseccio, clients_codisSAP.codiSAP, clients.nom, comandes_extres.codicomptable, comandes.producte, Clients_envios.empresa FROM (((comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda) LEFT JOIN clients_codisSAP ON comandes_extres.codicomptable = clients_codisSAP.codiSAP) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN Clients_envios ON comandes.direnvio = Clients_envios.id WHERE dataactivacio <>null and (((comandes.comanda)>165000) AND ((comandes.proximaseccio)<>'T') AND ((clients_codisSAP.codiSAP) Is Null) AND ((comandes.producte)<>'PC' And (comandes.producte)<>'PC2' And (comandes.producte)<>'PCP') AND ((Clients_envios.empresa)='INPLACSA')) ORDER BY clients_codisSAP.codiSAP;")
   While Not rst.EOF
     If atrim(cadbl(rst!codisap)) > 0 Then
        cosmissatge = cosmissatge + "El codicomptable " + atrim(rst!codisap) + " relacionat amb la comanda " + atrim(rst!comanda) + " de " + atrim(rst!nom) + " ara ja no està operatiu al SAP." + Chr(10)
          Else: cosmissatge = cosmissatge + "La comanda " + atrim(rst!comanda) + " de " + atrim(rst!nom) + " no te codi comptable relacionat, penseu a arregar-ho." + Chr(10)
     End If
     rst.MoveNext
   Wend
   
   Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda, comandes.proximaseccio, Clients_CodisSAPPlasel.codiSAP, clients.nom, comandes_extres.codicomptable, comandes.producte, Clients_envios.empresa FROM (((comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda) LEFT JOIN Clients_CodisSAPPlasel ON comandes_extres.codicomptable = Clients_CodisSAPPlasel.codiSAP) LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN Clients_envios ON comandes.direnvio = Clients_envios.id WHERE (((comandes.comanda)>165000) AND ((comandes.proximaseccio)<>'T') AND ((Clients_CodisSAPPlasel.codiSAP) Is Null) AND ((comandes.producte)<>'PC' And (comandes.producte)<>'PC2' And (comandes.producte)<>'PCP') AND ((Clients_envios.empresa)='PLASEL')) ORDER BY Clients_CodisSAPPlasel.codiSAP;")
   While Not rst.EOF
     If atrim(cadbl(rst!codisap)) > 0 Then
        cosmissatge = cosmissatge + "El codicomptable " + atrim(rst!codisap) + " DE PLASEL relacionat amb la comanda " + atrim(rst!comanda) + " de " + atrim(rst!nom) + " ara ja no està operatiu al SAP." + Chr(10)
          Else: cosmissatge = cosmissatge + "La comanda " + atrim(rst!comanda) + " DE PLASEL de " + atrim(rst!nom) + " no te codi comptable relacionat, penseu a arregar-ho." + Chr(10)
     End If
     rst.MoveNext
   Wend
   
   
   If cosmissatge <> "" Then
    destinatari = llegir_ini("destinataris", "destinatariclientdesvinculatsSAP", "enviarservidor.ini")
   ' destinatari = "miquel.inplacsa@gmail.com"
    enviaremail destinatari, "Comandes operatives amb codis comptables eliminats del SAP (Arreglar-ho)", cosmissatge
   End If
   Set rst = Nothing
End Sub
Sub mirarsieshoradexportar()
  
  'If (cadbl(Format(Now, "hh")) >= 22 Or cadbl(Format(Now, "hh")) < 6) Or Format(Now, "w") = 1 Or Format(Now, "w") = 7 Then
  If cadbl(llegir_ini("Exportarcomandes", "ultimaexecucio", "enviarservidor.ini")) <> Day(DateAdd("n", -15, Now)) Then
     escriure_ini "Exportarcomandes", "ultimaexecucio", atrim(Day(Now)), "enviarservidor.ini"
     If bexportarcomandes.Tag = "exportant" Then Exit Sub
      
     ' generarllistatdefetes
     If bexportarcomandes.Tag <> "exportant" Then exportarlescomandes
       Else:
           bexportarcomandes.Tag = ""
           'If existeix("c:\ordprog2.ini") Then FileCopy "c:\ordprog2.ini", "c:\ordprog.ini": Kill "c:\ordprog2.ini"
           'Set dbtmp = Nothing
  End If
End Sub
Sub generarllistatdefetes()
  Dim d As String
  'aixó borra tota la llista de comandes ja exportades i fa buscar dins la carpeta de pdf i genera una llista
    'de amb totes les comandes que hi ha creades
  Me.Caption = " Creant llistat de carpetes creades"
  DoEvents
  d = Dir(llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\*.*", vbDirectory)
  dbtmp.Execute "delete * from comandesexportades"
  While d <> ""
     If cadbl(d) > 0 Then dbtmp.Execute "insert into comandesexportades (comanda,data) values (" + atrim(cadbl(d)) + ",now)"
     d = Dir
  Wend
End Sub
Function fercopiahistoric() As Boolean
  Dim cami As String
  On Error GoTo no
  cami = llegir_ini("General", "cami", "comandes.ini")
  borrarhistoric
  db.Close
  Me.Tag = "copiant"
  FileCopy rutadelfitxer(cami) + "avisosincidencies.mdb", rutadelfitxer(cami) + "historicavisosincidencies.mdb"
  fercopiahistoric = True
  Set db = DBEngine.OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
  Me.Tag = ""
  Exit Function
no:
   fercopiahistoric = False
   Set db = DBEngine.OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
   Me.Tag = ""
End Function
Sub borrarhistoric()
  On Error GoTo fi
  eliminar_fitxer rutadelfitxer(cami) + "historicavisosincidencies.mdb"
fi:
End Sub
Sub comprovar_torerus()
   Dim vhorainici As Date
   Dim vultimresultat As String
   If llegir_ini("Torerus", "Generartorerus", rutadelfitxer(cami) + "valorsprograma.ini") = "Si" Then
    escriure_ini "Torerus", "Generartorerus", "Processant", rutadelfitxer(cami) + "valorsprograma.ini"
    etpujantadrive = "PROCESSANT TABLET TORERUS": DoEvents
    escriure_log "INICI - PROCESSANT TABLET TORERUS", "c:\temp\Log_EnviarMails_servidor.txt"
    Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe comandes.ini temporalTORERUS", vbNormalFocus
    vhorainici = Now
    wait 2
    
    
    While DateDiff("n", vhorainici, Now) < 2 And vultimresultat <> "OK"
      wait 1
      vultimresultat = llegir_ini("Torerus", "ultimresultat", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini")
      etpujantadrive = "PROCESSANT TABLET TORERUS"
      DoEvents
    Wend
    etpujantadrive = ""
    If DateDiff("n", vhorainici, Now) >= 3 Then KillProcess "palets.exe"
    escriure_ini "Torerus", "Generartorerus", "No", rutadelfitxer(cami) + "valorsprograma.ini"
    escriure_log "FI - PROCESSANT TABLET TORERUS", "c:\temp\Log_EnviarMails_servidor.txt"
  End If
  Exit Sub
End Sub
Sub comprovar_llistattubos()
  Dim vinici As String
  vinici = llegir_ini("Llistattubos", "horainici", rutadelfitxer(cami) + "valorsprograma.ini")
  If vinici <> "" Then
     etpujantadrive = "PROCESSANT LLISTAT DE TUBOS": DoEvents
     ShellAndWait rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "desembolicar bobines.exe llistattubos", vbNormalFocus
     etpujantadrive = "": DoEvents
     escriure_ini "Llistattubos", "horainici", "", rutadelfitxer(cami) + "valorsprograma.ini"
  End If
End Sub
Sub comprovar_incidencies()
  Dim rst As Recordset
  Dim cami As String
  On Error GoTo fi
  If Me.Tag <> "" Then Exit Sub
  estat.Caption = "Comprovant..."
  DoEvents
  cami = llegir_ini("General", "cami", "comandes.ini")
  
  comprovar_torerus
  comprovar_llistattubos
  
  Set rst = db.OpenRecordset("select * from avisos_baixes where not enviat")
  If Not rst.EOF Then
     While Not rst.EOF
       estat.Caption = "Enviant..."
       DoEvents
       enviar_incidencia rst
       rst.MoveNext
     Wend
  End If
  
  Set rst = db.OpenRecordset("select * from envios_mails where not enviat")
  If Not rst.EOF Then
     While Not rst.EOF
       estat.Caption = "Enviant mails..."
       DoEvents
       enviar_mailsgenerals rst
       rst.MoveNext
     Wend
  End If
  
  'db.Close
  Set rst = Nothing
  'Set db = Nothing
  estat.Caption = "Comprovació acabada."
  wait (1)
  estat.Caption = ""
  Exit Sub
fi:
  
End Sub
Sub enviar_mailsgenerals(rst As Recordset)
  Dim destinatari As String
  Dim i As Byte
  destinatari = llegir_ini("destinataris", rst!destinatari, "enviarservidor.ini")
  If destinatari = "{[}]" Then destinatari = rst!destinatari
  If InStr(1, destinatari, "@") > 0 Then
      enviaremail destinatari, rst!assumpte, atrim(rst!cos) + atrim(rst!cos2) + atrim(rst!cos3) + atrim(rst!cos4) + atrim(rst!cos5) + atrim(rst!cos6), , rst!ID
  End If
  db.Execute "update envios_mails set enviat=true where id=" + atrim(rst!ID)
  'rst.Edit
  'rst!enviat = True
  'rst.Update
End Sub


Sub enviar_incidencia(rst As Recordset)
Dim destinatari As String
Dim i As Byte
  'For i = 1 To 10
    'destinatari = llegir_ini("destinataris", "destinatari" + atrim(i), "enviarservidor.ini")
    destinatari = llegir_ini("destinataris", "destinatari1", "enviarservidor.ini")
    If InStr(1, destinatari, "@") > 0 Then
      enviaremail destinatari, crearasumpte(rst), crearcos(rst)
    End If
  'Next i
  db.Execute "update avisos_baixes set enviat=true where id=" + atrim(rst!ID)
End Sub

Function crearasumpte(rst As Recordset) As String
   crearasumpte = atrim(rst!comanda) + " - " + atrim(rst!avis)
End Function
Function crearcos(rst As Recordset) As String
   For i = 0 To rst.Fields.Count - 1
     If rst.Fields(i).Name <> "Id" And rst.Fields(i).Name <> "enviat" Then
       crearcos = crearcos + rst.Fields(i).Name + " = " + treure_apostruf(atrim(rst.Fields(i))) + Chr(10) + Chr(13)
     End If
   Next i
End Function

Function enviaremail(sSendTo As String, sSubject As String, sText As String, Optional adjunt As String, Optional vidavis As Long) As Boolean
  Dim usuarim As String
  Dim contrasenyam As String
  Dim destinatari As String
  Dim vnomcarpeta As String
  Dim vadjunt As String
  Dim vadjunt2 As String
  Dim vadjunt3 As String
  Dim vv As String
  vv = llegir_ini("destinataris", sSendTo, "enviarservidor.ini")
  If vv = "{[}]" Then
    sSendTo = sSendTo
     Else: sSendTo = vv
  End If
  vadjunt = adjunt
  vnomcarpeta = "\\serverprodu\Dades\progcomandes\dades\spoolerenviament\" + nomordinador + "_" + Format(Now, "yymmdd_hhnnss")
  usuarim = llegir_ini("dadesservidor", "usrsmtp", "enviarservidor.ini")
  contrasenyam = llegir_ini("dadesservidor", "passsmtp", "enviarservidor.ini")
  If usuarim = "{[}]" Or contrasenyam = "{[}]" Then
      escriure_ini "dadesservidor", "usrsmtp", " ", "enviarservidor.ini"
      escriure_ini "dadesservidor", "passsmtp", " ", "enviarservidor.ini"
      MsgBox "L'usuari o la contrasenya no estan entrades", vbCritical, "Error": Exit Function
  End If
  If Not existeix(vnomcarpeta) Then MkDir vnomcarpeta
  escriure_ini "Capcalera", "apuntperenviar", "No", vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "data", Now, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "nomordinador", nomordinador, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "usuari", usuarim, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "contrasenya", contrasenyam, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "destinatari", sSendTo, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "remitent", "incidencies@inplacsa.com", vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "assumpte", treure_apostruf(sSubject), vnomcarpeta + "\dadesmail.txt"
  'vadjunt = vadjunt3
  If vadjunt <> "" Then
    vadjunt3 = vadjunt
    vadjunt2 = vadjunt + "|"
    While vadjunt2 <> ""
      vadjunt = Mid(vadjunt2, 1, InStr(1, vadjunt2, "|") - 1)
      
      Copiar_Fitxer vadjunt, vnomcarpeta
      vadjunt2 = substituirtot(vadjunt2, vadjunt + "|", "")
    Wend
    vadjunt = substituirtot(vadjunt3, rutadelfitxer(vadjunt), vnomcarpeta + "\")
    escriure_ini "Capcalera", "adjunt", vadjunt, vnomcarpeta + "\dadesmail.txt"
  End If
  If LCase(sText) <> "c:\temp\cosmissatge.txt" Then
      If Not existeix(sText) Then
        Open "c:\temp\cosmissatge.txt" For Output As #2
        Print #2, sText
        passarliniesdavisosalfitxertxt vidavis
        Close #2
         Else
           If existeix("c:\temp\cosmissatge.txt") Then Kill "c:\temp\cosmissatge.txt"
           Copiar_Fitxer sText, "c:\temp\cosmissatge.txt"
       End If
   End If
   Copiar_Fitxer "c:\temp\cosmissatge.txt", vnomcarpeta
   If existeix("c:\temp\cosmissatge.txt") Then eliminar_fitxer "c:\temp\cosmissatge.txt"
   escriure_ini "Capcalera", "apuntperenviar", "Si", vnomcarpeta + "\dadesmail.txt"
   wait 1
   
End Function


Sub passarliniesdavisosalfitxertxt(vidavis As Long)
    Dim rst As Recordset
    Dim v As String
    If vidavis = 0 Then Exit Sub
    Set rst = db.OpenRecordset("select * from envios_mails_linies where id_envio=" + atrim(vidavis))
    If Not rst.EOF Then
       Print #2, ""
       Print #2, ""
    End If
    While Not rst.EOF
      v = atrim(rst!descripcio)
      If Len(v) > 0 Then
        'If InStr(1, v, Chr(10)) = 0 Then v = v + Chr(10)
        Print #2, v
      End If
      rst.MoveNext
    Wend
    Set rst = Nothing
End Sub
Function substituirtot(ByVal cadena As String, buscar As String, canviar As String) As String
   Dim comença As Integer
   Dim acaba As Integer
   If buscar = canviar Then GoTo fi
   While InStr(1, cadena, buscar) > 0
    comença = InStr(1, cadena, buscar) - 1
    
    If comença < 1 And InStr(1, cadena, buscar) <> 1 Then substituirtot = cadena: Exit Function
    acaba = comença + Len(buscar) + 1
    cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   Wend
fi:
   substituirtot = cadena
   'MsgBox linia
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




Private Sub Timercadaminut_Timer()
   Static vesticdins As Boolean
 '  eterrorservidorsap.Visible = IIf(Not ferPing("servidorsap"), True, False)
   escriure_ini "General", "dataprogramafuncionant", Trim(Now), "enviarservidor.ini"
   If vesticdins Then Exit Sub
   vesticdins = True
   comprovar_error_taules
   'Static vcadahora As Integer
   escriure_log "Timercadaminut_Timer (comprovar_albarans_tintes_nous)", "c:\temp\Log_EnviarMails_servidor.txt"
   
    comprovar_albarans_tintes_nous
   
    'comprovar si hi ha modificacions de comandes per passarles a l'alicia
   escriure_log "Timercadaminut_Timer (comprovarmodificacionscomandesienviarles)", "c:\temp\Log_EnviarMails_servidor.txt"
    comprovarmodificacionscomandesienviarles

  ' vcadahora = vcadahora + 1
   'cada hora actualitzar la neteja d'anilox i l'estadistica de les impresores
   If Minute(Now) = 0 Then
     escriure_log "Timercadaminut_Timer (actualitzant_netejaaniloxos)", "c:\temp\Log_EnviarMails_servidor.txt"
     actualitzant_netejaaniloxos
     
   End If
   
   'horari laboral entre setmana i cada 4 hores
   If (Hour(Now) > 7 And Hour(Now) < 21) And Format(Now, "w", vbMonday) < 6 And (Minute(Now) = 0 And Hour(Now) Mod 4 = 0) Then
     escriure_log "Timercadaminut_Timer (horari laboral entre setmana i cada 4 hores)", "c:\temp\Log_EnviarMails_servidor.txt"
     revisar_metresicanutorebobinadora_ambcomandescirculant
     enviarinformedebobinessensenumerodepalet
     revisarsihihacomandesalallistadeimpresiosensepackinglist
     calcular_credit_tots_clients
     Revisarcomandaacabadaimpresoressinohihaliniesdefuncionamentabaixes
     GRUPS_revisarsihihaprous_metres_assignats
     revisarBaixesPDFdelesseccions
     actualitzar_CQ_lots
     EnviarPaletsSenseImpost
     'revisarPackinglistDescuadrats
   End If
   vesticdins = False
End Sub
Sub EnviarPaletsSenseImpost()
   Dim rst As Recordset
   Dim vpaletsavisats As String
   Dim vmsg As String
   Dim vtipus As String
   Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
   vpaletsavisats = llegir_ini("General", "Paletsavisatssenseimpost", rutadelfitxer(cami) + "valorsprograma.ini")
   If vpaletsavisats = "{[}]" Then vpaletsavisats = ""
   Set rst = dbstocks.OpenRecordset("SELECT materials.tanpercentimpostenvasos, proveidors.tipusproveidorIMPOST AS tipusp, * FROM (palets LEFT JOIN materials ON palets.codimatprognou = materials.codi) LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE (((palets.teimpost)=False) AND ((palets.dataaltapalet)>DateAdd('d',-30,Now()) And (palets.dataaltapalet)<DateAdd('n',-30,Now())));")
   While Not rst.EOF
     vtipus = atrim(rst!tipusp)
     If rst!tanpercentimpostenvasos > 0 And (vtipus = "Importació" Or vtipus = "Intracomunitari" Or vtipus = "Espanyol") Then
        If InStr(1, vpaletsavisats, atrim(rst!idpalet)) = 0 Then
            vmsg = vmsg + vbNewLine + "Palet: " + atrim(rst!idpalet) + "  ---  Creat el dia: " + atrim(rst!dataaltapalet) + " SENSE IMPOST"
            vpaletsavisats = vpaletsavisats + " " + atrim(rst!idpalet)
        End If
     End If
     rst.MoveNext
   Wend
   If vmsg <> "" Then
        vmsg = " LLISTAT DE PALETS CREATS SENSE IMPOST" + vbNewLine + "==============================" + vbNewLine + vbNewLine + vmsg
        enviaremail "odamian@inplacsa.com;jmiralles@inplacsa.com;miquel.inplacsa@gmail.com;amiquel@inplacsa.com", "Palets creats sense IMPOST. " + atrim(Now), vbNewLine + vmsg
        
        If Len(vpaletsavisats) > 254 Then vpaletsavisats = Mid(vpaletsavisats, 120)
        escriure_ini "General", "Paletsavisatssenseimpost", vpaletsavisats, rutadelfitxer(cami) + "valorsprograma.ini"
   End If
   Set rst = Nothing
End Sub
Sub revisarPackinglistDescuadrats()
   Dim rst As Recordset
   Dim rste As Recordset
   Dim rstp As Recordset
   Dim rsth As Recordset
   Dim vmetresactuals As Double
   Dim vmetresassignats As Double
   Dim vmsg As String
   Dim vdetall As String
   Dim vnumc As Double
   Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
   Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
   Set rst = dbcomandes.OpenRecordset("select * from comandes where proximaseccio='I'")
   While Not rst.EOF
      vnumc = rst!comanda
      Set rste = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(rst!comanda))
      If Not rste.EOF Then
         vmetresassignats = cadbl(rste!metresassignatspackinglist)
         If vmetresassignats = 0 Then GoTo proxim
         Set rstp = dbstocks.OpenRecordset("select metres as Tmetres from historic_packinglist where comanda='" + atrim(rst!comanda) + "'")
         If Not rstp.EOF Then GoTo proxim
         Set rstp = dbstocks.OpenRecordset("select sum(metres) as Tmetres from parcials where comanda='" + atrim(rst!comanda) + "'")
         If Not rstp.EOF Then
             vmetresactuals = cadbl(rstp!tmetres)
         End If
         'If vmetresactuals < vmetresassignats Then
         '    vmsg = vmsg + "La comanda " + atrim(rst!comanda) + " tenia assignats " + atrim(vmetresassignats) + " metres i ara té " + atrim(vmetresactuals) + " metres." + vbNewLine
         'End If
         If nohihaproumetresassignats(vnumc, vmetresactuals, vdetall) And Not comandacomençada(vnumc) Then
             vmsg = vmsg + "La comanda " + atrim(rst!comanda) + " tenia assignats " + atrim(vmetresassignats) + " metres i ara no hi ha prous metres a les bobines assignades." + vbNewLine + vdetall
         End If
      End If
proxim:
      rst.MoveNext
   Wend
'   MsgBox vmsg
   If vmsg <> "" Then enviaremail "PackingListDescuadrat", "Diferencies Assignats i Packinglist.", vbNewLine + vbNewLine + vmsg
   Set rst = Nothing
   Set rste = Nothing
   Set rstp = Nothing
   Set rsth = Nothing
End Sub
Function comandacomençada(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbbaixes.OpenRecordset("select * from impressores where tipus='F' and comanda=" + atrim(vnumc))
   If Not rst.EOF Then comandacomençada = True
   Set rst = Nothing
End Function
Function nohihaproumetresassignats(vnumc As Double, vmetresactuals As Double, vdetall As String) As Boolean
   Dim rstp As Recordset
   Dim vmetresassignats As Double
   Dim vmetresdescuadrats As Double
   vdetall = ""
   Set rstp = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(vnumc))
   If Not rstp.EOF Then vmetresassignats = cadbl(rstp!metresassignatspackinglist)
   If vmetresassignats = 0 Then GoTo fi
   Set rstp = dbstocks.OpenRecordset("select sum(metres) as Tmetres from parcials where utilitzada=false and comanda='" + atrim(vnumc) + "'")
   If Not rstp.EOF Then vmetresactuals = cadbl(rstp!tmetres)
   
   If vmetresactuals < vmetresassignats Then
       nohihaproumetresassignats = True
   End If
   If Not nohihaproumetresassignats Then
     Set rstp = dbstocks.OpenRecordset("select * from parcials where not utilitzada and comanda='" + atrim(vnumc) + "'")
     vmetresactuals = 0
     While Not rstp.EOF
        vmetresdescuadrats = bobinesdentrada.calcular_mtrsdispreals(rstp!idpalet, rstp!idbobina) - rstp!metres
        If vmetresdescuadrats < 0 Then
                nohihaproumetresassignats = True
                vdetall = vdetall + "     Bobina: " + atrim(rstp!idpalet) + "/" + atrim(rstp!idbobina) + " no té prous metres. En falten " + atrim(vmetresdescuadrats * -1) + " Mtrs." + vbNewLine
        End If
        vmetresactuals = vmetresactuals + rstp!metres
        rstp.MoveNext
     Wend
   End If
fi:
   Set rstp = Nothing
End Function


Sub revisarBaixesPDFdelesseccions()
   Dim vsetmanaactual As Double
   vsetmanaactual = cadbl(llegir_ini("accionsglobals", "SetmanarevisioPDFbaixes", rutadelfitxer(cami) + "valorsprograma.ini"))
   'Me.Caption = "Comprovant els PDF de baixes de seccions"
   If WeekDay(Now, vbMonday) = 5 And vsetmanaactual <> Format(Now, "ww") And Hour(Now) > 13 Then
      comprovar_PDF_baixesseccions
      escriure_ini "accionsglobals", "SetmanarevisioPDFbaixes", atrim(Format(Now, "ww")), rutadelfitxer(cami) + "valorsprograma.ini"
   End If
End Sub
Sub Revisarcomandaacabadaimpresoressinohihaliniesdefuncionamentabaixes()
   Dim vsql As String
   Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
   vsql = "SELECT impressorestot.comanda FROM comandes RIGHT JOIN (impressorestot LEFT JOIN impressores ON impressorestot.comanda = impressores.comanda) ON comandes.comanda = impressorestot.comanda WHERE (((impressores.comanda) Is Null) AND ((comandes.proximaseccio)<>'T' And (comandes.proximaseccio)<>'V' And (comandes.proximaseccio)<>'P'));"
   dbbaixes.Execute "update impressorestot set acavada=0 where comanda in (" + vsql + ")"
  ' Set dbbaixes = Nothing
End Sub

Sub GRUPS_revisarsihihaprous_metres_assignats()
     Dim rst As Recordset
     Dim vmetresassignats As Double
     Dim vmetresnecessaris As Double
     Dim vm As Double
     Dim vmsg As String
     Dim vmsgexp As String
     Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
     Set rst = dbstocks.OpenRecordset("select * from grupsdepalets")
     While Not rst.EOF
       vmetresassignats = 0: vmetresnecessaris = 0: vm = 0
       hihaprousmetresestoc cadbl(rst!numerogrup), vmetresassignats, vmetresnecessaris
       If vmetresnecessaris > vmetresassignats Then vmsg = vmsg + "Del Grup " + atrim(rst!numerogrup) + " hi ha " + atrim(vmetresassignats) + " metres assignats i es necessiten " + atrim(vmetresnecessaris) + " metres." + Chr(13) + Chr(110)
       vm = vmetresassignats
       vmetresassignats = 0: vmetresnecessaris = 0
       hihaprousmetresestoc cadbl(rst!numerogrup), vmetresassignats, vmetresnecessaris, IIf(atrim(rst!seccio) = "I", "IMP", "LAM")
       If vmetresnecessaris > vmetresassignats Then vmsgexp = vmsgexp + "Del Grup " + atrim(rst!numerogrup) + " hi ha " + atrim(vm) + " metres assignats però falta baixar " + atrim(vmetresnecessaris - vmetresassignats) + "metres i es necessiten " + atrim(vmetresnecessaris) + " metres per totes les comandes." + Chr(13) + Chr(110)
       rst.MoveNext
     Wend
     Set rst = Nothing
     If vmsg <> "" Then enviaremail "compres@inplacsa.com", "ATENCIÓ - Falten bobines en algun GRUP DE PALETS.", vmsg
     If vmsgexp <> "" Then enviaremail "expedicions@inplacsa.com", "ATENCIÓ - Falta baixar bobines d'algun GRUP DE PALETS.", vmsgexp
End Sub
Function hihaprousmetresestoc(vnumestoc As Double, vmetresassignats As Double, vmetresnecessaris As Double, Optional vAIMP As String)
  Dim rstopcions As Recordset
  Dim vsql As String
  hihaprousmetresestoc = True
  Set rstopcions = dbstocks.OpenRecordset("SELECT opcionsdajust.grupdestoc as GrupEstoc, Sum(comandes.cantitatex) AS Tmetres FROM opcionsdajust LEFT JOIN comandes ON opcionsdajust.comanda = comandes.comanda Where (((comandes.proximaseccio) = 'I' )) GROUP BY opcionsdajust.grupdestoc;")
  rstopcions.FindFirst "GrupEstoc=" + atrim(vnumestoc)
  If Not rstopcions.NoMatch Then
      vmetresnecessaris = cadbl(rstopcions!tmetres)
      If vAIMP <> "" Then
         'Clipboard.Clear
         'Clipboard.SetText "SELECT parcials.comanda, Sum(parcials.metres) AS Tmetres FROM Bobines RIGHT JOIN parcials ON (Bobines.Idbobina = parcials.idbobina) AND (Bobines.Idpalet = parcials.idpalet) Where (((Bobines.Sit) LIKE '" + vAIMP + "*')) GROUP BY parcials.comanda HAVING (((parcials.comanda)='" + atrim(vnumestoc) + "'));"
         'Set rstopcions = dbstocks.OpenRecordset("SELECT parcials.comanda, Sum(parcials.metres) AS Tmetres FROM Bobines RIGHT JOIN parcials ON (Bobines.Idbobina = parcials.idbobina) AND (Bobines.Idpalet = parcials.idpalet) Where (((Bobines.Sit) LIKE '" + vAIMP + "*')) GROUP BY parcials.comanda HAVING (((parcials.comanda)='" + atrim(vnumestoc) + "'));")
            If vAIMP = "LAM" Then vsql = "SELECT Sum([COMANDES].[cantitatex]) AS tmetres FROM ((opcionsdajust LEFT JOIN comandes ON opcionsdajust.comanda = comandes.comanda) LEFT JOIN productes ON comandes.producte = productes.codi) LEFT JOIN comandes AS comandes_1 ON comandes.linkcomanda1 = comandes_1.comanda WHERE (((comandes.proximaseccio)<>'T') AND ((opcionsdajust.grupdestoc)=" + atrim(cadbl(rstopcions!grupestoc)) + ") AND ((comandes.producte)='PC' Or (comandes.producte)='PCP' Or (comandes.producte)='PC2') AND ((comandes_1.proximaseccio)='E' Or (comandes_1.proximaseccio)='I' Or (comandes_1.proximaseccio)='L'));"
            If vAIMP = "IMP" Then vsql = "SELECT Sum([COMANDES].[cantitatex]) AS tmetres FROM ((opcionsdajust LEFT JOIN comandes ON opcionsdajust.comanda = comandes.comanda) LEFT JOIN productes ON comandes.producte = productes.codi) LEFT JOIN comandes AS comandes_1 ON comandes.linkcomanda1 = comandes_1.comanda WHERE (comandes.proximaseccio='E' Or comandes.proximaseccio='I')  AND (opcionsdajust.grupdestoc=" + atrim(cadbl(rstopcions!grupestoc)) + ");"
            Set rstopcions = dbstocks.OpenRecordset(vsql)
              Else:
               Set rstopcions = dbstocks.OpenRecordset("SELECT Parcials.comanda, Sum(Parcials.metres) AS Tmetres From parcials GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + atrim(vnumestoc) + "'));")
      End If
      If Not rstopcions.EOF Then vmetresassignats = cadbl(rstopcions!tmetres)
  End If
  Set rstopcions = Nothing
End Function
Sub revisarsihihacomandesalallistadeimpresiosensepackinglist()
     Dim rst As Recordset
     Dim vmsg As String
     Dim vestat As String
     Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
     Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
     Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
     Set rst = dbbaixes.OpenRecordset("select comanda from impresores_ordreimpresio")
     While Not rst.EOF
        vestat = estatdelacompra(rst!comanda)
        If vestat = "Reservat" Or vestat = "" Then
            vmsg = vmsg + " - " + atrim(rst!comanda)
        End If
        rst.MoveNext
     Wend
     If vmsg <> "" Then
         enviaremail "destinatari1", "Comandes sense packing-list a la llista d'Impresores", Chr(13) + Chr(10) + "Comandes: " + vmsg
     End If
     Set rst = Nothing
    ' Set dbstocks = Nothing
  '   Set dbcompres = Nothing
   '  Set dbbaixes = Nothing
End Sub
Function estatdelacompra(numc As Double) As String
   Dim rstc As Recordset
   Set rstc = dbstocks.OpenRecordset("select * from parcials where comanda='" + atrim(numc) + "'", dbOpenSnapshot, dbReadOnly)
   If Not rstc.EOF Then
      estatdelacompra = "Packing-List"
     Else
        'Set dbcomandes = OpenDatabase(cami)
        Set rstc = dbcomandes.OpenRecordset("select * from comandes_extres where comanda=" + atrim(numc), dbOpenSnapshot, dbReadOnly)
        If Not rstc.EOF Then If rstc!assignarstock Then estatdelacompra = "Estoc"
   End If
   If estatdelacompra = "" Then
       Set rstc = dbstocks.OpenRecordset("select * from percomandaoclient where numcomanda=" + atrim(numc), dbOpenSnapshot, dbReadOnly)
       If Not rstc.EOF Then estatdelacompra = "Reservat"
   End If
   If estatdelacompra = "" Then
    Set rstc = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, capcalera.nomprov,capcalera.dataentrega, comandesxlinia.numcomanda FROM (capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra) RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra WHERE (((comandesxlinia.numcomanda)=" + atrim(numc) + "));", dbOpenSnapshot, dbReadOnly)
    If Not rstc.EOF Then
       estatdelacompra = "Compra: " + atrim(rstc![capcalera.numcomanda]) + " Entrega: " + atrim(rstc!dataentrega) + "  " + atrim(rstc!nomprov)
    End If
   End If
   
   Set rstc = Nothing
End Function
Sub comprovar_error_taules()
   If donaerrorlabasededades("comandes.mdb") Then
       If llegir_ini("ErrorBD", "Comandes", "enviarservidor.ini") <> "S" Then
          'enviar email
          enviaremail "miquel.inplacsa@gmail.com", "Error BD Comandes    " + atrim(Now), ""
          escriure_ini "ErrorBD", "Comandes", "S", "enviarservidor.ini"
       End If
       Else: escriure_ini "ErrorBD", "Comandes", "N", "enviarservidor.ini"
   End If
   If donaerrorlabasededades("baixes.mdb") Then
       If llegir_ini("ErrorBD", "Baixes", "enviarservidor.ini") <> "S" Then
          'enviar email
          enviaremail "miquel.inplacsa@gmail.com", "Error BD Baixes    " + atrim(Now), ""
          escriure_ini "ErrorBD", "Baixes", "S", "enviarservidor.ini"
       End If
       Else: escriure_ini "ErrorBD", "Baixes", "N", "enviarservidor.ini"
   End If
  If donaerrorlabasededades("palets.mdb") Then
       If llegir_ini("ErrorBD", "Comandes", "enviarservidor.ini") <> "S" Then
          'enviar email
          enviaremail "miquel.inplacsa@gmail.com", "Error BD Comandes    " + atrim(Now), ""
          escriure_ini "ErrorBD", "Comandes", "S", "enviarservidor.ini"
       End If
       Else: escriure_ini "ErrorBD", "Comandes", "N", "enviarservidor.ini"
   End If
End Sub
Function donaerrorlabasededades(vnombd As String) As Boolean
   Dim dberror As Database
   On Error GoTo errorbd
   Set dberror = OpenDatabase(rutadelfitxer(cami) + vnombd, , True)
   Set dberror = Nothing
   Exit Function
errorbd:
   donaerrorlabasededades = True
   
End Function

Private Sub TimerSAP_Timer()
   Dim sincronitzar As Boolean
   sincronitzar = IIf(llegir_ini("General", "sincronitzarsap", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini") = "Si", True, False)
   If sincronitzar Then
      escriure_log "TimerSAP_Timer (Sincronitzant SAP)", "c:\temp\Log_EnviarMails_servidor.txt"
      If obrir_tancar_taules(True) Then
        sincronitzar_taulesmestra
        escriure_ini "General", "sincronitzarsap", "No", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini"
        escriure_ini "General", "sincronitzarsapusuari", "", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini"
      End If
     ' obrir_tancar_taules False
   End If
   sincronitzar = IIf(llegir_ini("General", "calcularestadisticaaniloxos", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini") = "Si", True, False)
   If sincronitzar Then
     escriure_log "TimerSAP_Timer (calcular_estadisticaaniloxos)", "c:\temp\Log_EnviarMails_servidor.txt"
     calcular_estadisticaaniloxos True
     escriure_ini "General", "calcularestadisticaaniloxos", "No", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini"
   End If
   
End Sub




    

