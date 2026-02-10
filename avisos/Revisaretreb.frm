VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form competreb 
   Caption         =   "Avisos"
   ClientHeight    =   6195
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   3885
   ControlBox      =   0   'False
   Icon            =   "Revisaretreb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Height          =   420
      Index           =   0
      Left            =   1470
      Picture         =   "Revisaretreb.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Imprimir avisos"
      Top             =   5715
      Width           =   975
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   3360
      Top             =   5700
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Borrar"
      Height          =   420
      Left            =   255
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5715
      Width           =   975
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
      Height          =   4935
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   315
      Width           =   3795
   End
   Begin VB.CommandButton tancar 
      Caption         =   "Amagar"
      Height          =   420
      Left            =   2685
      TabIndex        =   0
      Top             =   5715
      Width           =   975
   End
   Begin VB.Label etiqueta 
      Height          =   255
      Left            =   105
      TabIndex        =   5
      Top             =   0
      Width           =   3660
   End
   Begin VB.Menu m_opcions 
      Caption         =   "Opcions"
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
 

End Sub

Private Sub Command2_Click()
  
  If InputBox("Segur que vols borrar aquests avisos?" + Chr(13) + Chr(10) + "Escriu [ELIMINAR] per borrar-los", "Atenció") = "ELIMINAR" Then
      'Kill rutadelfitxer(cami) + "aviscampsmodificats.txt"
      'etxactivar.Clear
      borrartotselsavisos
  End If
End Sub
Sub borrartotselsavisos()
  Dim i As Integer
  For i = 0 To etxactivar.ListCount - 1
       dbcomandes.Execute "insert into historic_aviscampsmodificats select * from aviscampsmodificats where comanda=" + atrim(etxactivar.ItemData(i))
       dbcomandes.Execute "delete * from aviscampsmodificats where comanda=" + atrim(etxactivar.ItemData(i))
  Next i
  etxactivar.Clear
End Sub
Private Sub ensenya_totes_Click()
End Sub

Private Sub Command9_Click(Index As Integer)
   comprovar_avisos
   
   llistaravisos
   If MsgBox("Assegura que s'hagi imprès bé el llistat d'Avisos abans de borrar-los." + Chr(10) + "Vols borrar tots els avisos?", vbCritical + vbYesNo + vbDefaultButton2, "Impresio i borrat d'avisos") = vbYes Then
       borrartotselsavisos
   End If
End Sub

Private Sub etxactivar_Click()
   Dim rst As Recordset
   etiqueta = ""
   Set rst = dbcomandes.OpenRecordset("select * from aviscampsmodificats where comanda=" + atrim(etxactivar.ItemData(etxactivar.ListIndex)))
   If Not rst.EOF Then
     etiqueta = Format(rst!datacanvi, "dd/mm/yy") + " - " + UCase(atrim(rst!ordinadorcanvi))
   End If
   Set rst = Nothing
End Sub

Private Sub etxactivar_DblClick()
  If MsgBox("Vols borrar aquesta linia?" + Chr(10) + Chr(13) + etxactivar.Text, vbCritical + vbYesNo, "Atenció") = vbYes Then
      dbcomandes.Execute "insert into historic_aviscampsmodificats select * from aviscampsmodificats where comanda=" + atrim(etxactivar.ItemData(etxactivar.ListIndex))
      dbcomandes.Execute "delete * from aviscampsmodificats where comanda=" + atrim(etxactivar.ItemData(etxactivar.ListIndex))
      etxactivar.RemoveItem (etxactivar.ListIndex)
     
      'carregarlallista
  End If
  'guardar_llista etxactivar.Tag
End Sub

Private Sub Form_Activate()
Form_Resize
End Sub
Sub llistaravisos()
  Dim rst As Recordset
 Set rst = dbcomandes.OpenRecordset("select * from aviscampsmodificats")
  If rst.EOF Then Exit Sub
  llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistatavisoscampsmodificats.rpt"
  If Not existeix("c:\ordprog.ini") Then llistat.Destination = crptToPrinter
  llistat.DataFiles(0) = cami
'  llistat.SelectionFormula = "{diferenciescomandaitreball.comanda}=" + atrim(numc)
  llistat.Action = 1
End Sub

Private Sub Form_Click()
  MsgBox Environ("computername")
End Sub

Private Sub Form_DblClick()
 'Timer1_Timer
 comprovar_avisos
 
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
    etxactivar.Width = competreb.Width - 300
    tancar.Left = competreb.Width - 1440
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
 If EstaCorriendo("avisos.exe") Then End
 arguments = ObtenerLíneaComando
 If Trim(arguments(1)) <> "" Then cami = Trim(arguments(1))
 If cami = "" Then
   cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
 End If
 Set dbcomandes = OpenDatabase(cami)
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


Private Sub m_Pararcomprovacioitancar_Click()
  End
End Sub

Private Sub tancar_Click()
 Me.Hide
 Call PonerSystray
End Sub

Private Sub Timer1_Timer()
  Static cont As Byte
  If cont < 21 And cont > 0 Then cont = cont + 1: Exit Sub
  comprovar_avisos
  If etxactivar.ListCount > 0 And norecordar.Value = 0 Then
    colocarfinestreabaix
  End If
  'If Timer1.Interval = 5000 Then Timer1.Interval = 60000
  cont = 1
End Sub
Sub colocarfinestreabaix()

 Me.WindowState = vbNormal
   Me.Left = Screen.Width - Me.Width
   Me.Top = Screen.Height - Me.Height - 350
            Call SetForegroundWindow(Me.hwnd)
            Me.Show
End Sub
Sub comprovar_avisos()
 'If existeix(rutadelfitxer(cami) + "aviscampsmodificats.txt") Then
     carregarlallista 'rutadelfitxer(cami) + "aviscampsmodificats.txt"
 'End If
End Sub
Sub carregarlallista()
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select * from aviscampsmodificats")
   etxactivar.Clear
   While Not rst.EOF
     etxactivar.AddItem atrim(rst!comanda) + " " + atrim(rst!descripciomodificacio)
     etxactivar.ItemData(etxactivar.NewIndex) = rst!comanda
     rst.MoveNext
   Wend
   If etxactivar.ListCount > 0 And norecordar.Value = 0 Then
     Form_Resize
     colocarfinestreabaix
   End If
End Sub
Sub carregarlallista_Vella(nomfitxer As String)
   Dim linia As String
   etxactivar.Clear
   etxactivar.Tag = nomfitxer
   Open nomfitxer For Input As #1
   While Not EOF(1)
    Line Input #1, linia
    If atrim(linia) <> "" Then
     etxactivar.AddItem linia
     etxactivar.ItemData(etxactivar.NewIndex) = cadbl(Mid(linia, 9, 8))
    End If
   Wend
   Close #1
End Sub

Sub guardar_llista(nomfitxer As String)
   Dim linia As String
   Dim i As Integer
   Kill nomfitxer
   Open nomfitxer For Output As #1
   For i = 0 To etxactivar.ListCount - 1
    Print #1, etxactivar.List(i)
   Next i
   Close #1
   
End Sub
Private Sub Timer2_Timer()
  If etxactivar.ListCount > 0 And norecordar.Value = 0 Then FlashWindow Me.hwnd, 1
  
End Sub

  

