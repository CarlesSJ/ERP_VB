VERSION 5.00
Begin VB.Form competreb 
   Caption         =   "Comprobar Ok Et. Reb."
   ClientHeight    =   6195
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   2385
   ControlBox      =   0   'False
   Icon            =   "Revisaretreb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   2385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Ok dels operaris"
      Height          =   450
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5685
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   495
      Top             =   2190
   End
   Begin VB.CheckBox norecordar 
      Caption         =   "No m'ho recordis mes."
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   5445
      Width           =   2085
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   45
      Top             =   525
   End
   Begin VB.ListBox etxactivar 
      Height          =   5130
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tancar"
      Height          =   390
      Left            =   1080
      TabIndex        =   0
      Top             =   5745
      Width           =   1140
   End
   Begin VB.Menu m_opcions 
      Caption         =   "Opcions"
      Begin VB.Menu m_historialsoks 
         Caption         =   "Historial Oks"
      End
      Begin VB.Menu ensenya_totes 
         Caption         =   "Ensenya totes."
      End
      Begin VB.Menu m_Pararcomprovacioitancar 
         Caption         =   "Parar i Tancar"
      End
   End
End
Attribute VB_Name = "competreb"
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




Private Sub Command1_Click()
  Me.Hide
 Call PonerSystray

End Sub

Private Sub Command2_Click()
  ShellandWait "notepad.exe " + App.Path + "\etokoperaris.txt", , 1
  If MsgBox("Arxivo aquests oks a l'historial?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
      ShellandWait "c:\windows\system32\cmd.exe /c type " + Chr$(34) + App.Path + "\etokoperaris.txt" + Chr$(34) + " >> " + Chr$(34) + App.Path + "\etokophistoric.txt" + Chr$(34), , 1
      Kill App.Path + "\etokoperaris.txt"
      Command2.Visible = False
  End If
End Sub

Private Sub ensenya_totes_Click()
  ensenya_totes.Checked = Not ensenya_totes.Checked
  comprovar_et_reb
End Sub

Private Sub Form_DblClick()
 'Timer1_Timer
End Sub

Private Sub Form_Initialize()
 Me.Hide
Call PonerSystray
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
 If EstaCorriendo("comprovaretreb.exe") Then End
 arguments = ObtenerLíneaComando
 If Trim(arguments(1)) <> "" Then cami = Trim(arguments(1))
 cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
   DoEvents
   Me.Visible = False
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
  Shell "notepad.exe " + App.Path + "\etokophistoric.txt", vbNormalFocus
End Sub

Private Sub m_Pararcomprovacioitancar_Click()
  End
End Sub

Private Sub Timer1_Timer()
  Static cont As Byte
  If cont < 21 And cont > 0 Then cont = cont + 1: Exit Sub
  comprovar_et_reb
  If etxactivar.ListCount > 0 And norecordar.Value = 0 Then
   Me.WindowState = vbNormal
   Me.Left = Screen.Width - Me.Width
   Me.Top = Screen.Height - Me.Height - 350
            Call SetForegroundWindow(Me.hwnd)
            Me.Show
    
  End If
  If Timer1.Interval = 5000 Then Timer1.Interval = 60000
  cont = 1
End Sub
Sub comprovar_et_reb()
   Dim rst As Recordset
   Dim direnvio As Recordset
   Dim rstp As Recordset
   Dim afegir As Boolean
   etxactivar.Clear
   Set db = DBEngine.OpenDatabase(cami)
   Set rst = db.OpenRecordset("select comanda,impressio,producte,etrebvistiplau from comandes where producte not in ('PC','PC2') and proximaseccio in ('L','I','R') " + IIf(ensenya_totes.Checked, "", " and not etrebvistiplau"), , dbRunAsync)
   afegir = False
   While Not rst.EOF
    Set rstp = db.OpenRecordset("select ruta from productes where codi='" + atrim(rst!producte) + "'")
    If Not rstp.EOF Then
     If InStr(1, rstp!ruta, "I") Then
       If rst!impressio <> "R" Then
             afegir = True
           Else
              rst.Edit
              rst!etrebvistiplau = True
              rst.Update
       End If
         Else: afegir = True
     End If
    End If
    If afegir Then
     'etxactivar.AddItem Trim(rst!comanda)
     rst.Edit
     rst!etrebvistiplau = True
     rst.Update
    End If
    rst.MoveNext
    DoEvents
   Wend
   If Not ensenya_totes.Checked Then
    Set rst = db.OpenRecordset("select comanda,direnvio from comandes where  comanda >145000 and  proximaseccio<>'T' order by comanda", , dbRunAsync)
    
    While Not rst.EOF
     Set direnvio = db.OpenRecordset("select id from clients_envios where codi>0 and id=" + atrim(cadbl(rst!direnvio)))
     'direnvio.FindFirst "id=" + atrim(cadbl(rst!direnvio))
     If direnvio.EOF Then
       etxactivar.AddItem "Falta DirEnvio: " + Trim(rst!comanda)
     End If
     rst.MoveNext
     DoEvents
    Wend
    'comprovo direccions d'envio noves que encara no s'ha fet cap etiqueta
    Set direnvio = db.OpenRecordset("select verificatclientnouxretiquetes,codi  from clients_envios where not verificatclientnouxretiquetes")
    While Not direnvio.EOF
      If cadbl(direnvio!codi) > 0 Then
        etxactivar.AddItem "DirEnvio nou: Client " + atrim(direnvio!codi)
      End If
      direnvio.MoveNext
    Wend
   End If
  If existeix(App.Path + "\etokoperaris.txt") Then
     Command2.Visible = True
  Else: Command2.Visible = False
  End If
   db.Close
   Set rst = Nothing
   Set db = Nothing
End Sub

Private Sub Timer2_Timer()
  If etxactivar.ListCount > 0 And norecordar.Value = 0 Then FlashWindow Me.hwnd, 1
  
End Sub

  

