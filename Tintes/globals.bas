Attribute VB_Name = "Module2"
Const LOCALE_SDECIMAL = &HE
Const LOCALE_STHOUSAND = &HF

Const GW_HWNDNEXT = 2
 Declare Function PostMessage Lib "User" (ByVal hwnd _
      As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, _
      lParam As Any) As Integer
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle _
    As Long, ByVal dwMilliseconds As Long) As Long
 Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
 Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
  ByVal wCmd As Long) As Long
 Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
 Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
  (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
 Declare Function GetWindowThreadProcessId Lib "user32" _
  (ByVal hwnd As Long, lpdwprocessid As Long) As Long

Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function apiSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

       Private Type OPENFILENAME
         lStructSize As Long
         hwndOwner As Long
         hInstance As Long
         lpstrFilter As String
         lpstrCustomFilter As String
         nMaxCustFilter As Long
         nFilterIndex As Long
         lpstrFile As String
         nMaxFile As Long
         lpstrFileTitle As String
         nMaxFileTitle As Long
         lpstrInitialDir As String
         lpstrTitle As String
         flags As Long
         nFileOffset As Integer
         nFileExtension As Integer
         lpstrDefExt As String
         lCustData As Long
         lpfnHook As Long
         lpTemplateName As String
       End Type
       Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
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



 
Public Function GetTCursX() As Long
    Dim pt As POINTAPI
    GetCaretPos pt
    GetTCursX = pt.x
End Function
 
Public Function GetTCursY() As Long
    Dim pt As POINTAPI
    GetCaretPos pt
    GetTCursY = pt.y
End Function
  Function treure_apostrof(nomf As String) As String
    While InStr(1, nomf, "'") <> 0
       nomf = Mid(nomf, 1, InStr(1, nomf, "'") - 1) + "´" + Mid(nomf, InStr(1, nomf, "'") + 1)
    Wend
    treure_apostrof = nomf
  End Function
  Function existeixlataula(basededades As String, nomtaula As String) As Boolean
     Dim dbexist As Database
     Dim rstexist As Recordset
     existeixlataula = True
     On Error GoTo noexisteix
     Set dbexist = DBEngine.OpenDatabase(basededades, , True)
     Set rstexist = dbexist.OpenRecordset(nomtaula)
     Set rstexist = Nothing
     Exit Function
noexisteix:
      existeixlataula = False
  End Function
Sub esperarunaestona()
  Dim valesp As Double
  missatge.Show
  missatge.etimissatge.caption = "Creando listado, Espere ..."
  DoEvents
  valesp = cadbl(llegir_ini("General", "tempsesperallistat", "ferral.ini"))
  If valesp = 0 Then valesp = 1500: escriure_ini "General", "tempsesperallistat", "1500", "ferral.ini"
  For i = 1 To valesp * 100
   DoEvents
  Next i
  Unload missatge
End Sub

Function obre_fitxer(dirinici As String, flags As Double) As String
      Dim OpenFile As OPENFILENAME
      Dim lReturn As Long
      Dim sFilter As String
      OpenFile.lStructSize = Len(OpenFile)
      OpenFile.hwndOwner = frmclixes.hwnd
      OpenFile.hInstance = App.hInstance
      sFilter = "*.*"
      'sFilter = ""
      OpenFile.lpstrFilter = sFilter
      OpenFile.nFilterIndex = 1
      OpenFile.lpstrFile = String(257, 0)
      OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
      OpenFile.lpstrFileTitle = OpenFile.lpstrFile
      OpenFile.nMaxFileTitle = OpenFile.nMaxFile
      OpenFile.lpstrInitialDir = dirinici
      OpenFile.lpstrTitle = "Tria el fitxer..."
      OpenFile.flags = flags
      lReturn = GetOpenFileName(OpenFile)
      If lReturn = 0 Then
            obre_fitxer = ""
        Else
            obre_fitxer = atrim(OpenFile.lpstrFile)
            If InStr(1, obre_fitxer, "'") > 0 Then MsgBox "Aquest nom de fitxer conté un APOSTROF substituiu-lo per un accent+espai i torneu-lo a Linkar": obre_fitxer = ""
            
      End If
End Function


Sub assignardecimalipunt()
  Dim LocalID As Long
  LocalID = GetUserDefaultLCID()
  SetLocaleInfo LocalID, LOCALE_SDECIMAL, ","
  SetLocaleInfo LocalID, LOCALE_STHOUSAND, "."
End Sub

'Sub wait(segonsespera As Byte)
'  horaentradawait = Now
'  While DateDiff("s", horaentradawait, Now) < segonsespera
'    DoEvents
'  Wend
'End Sub

Public Function Redondejar(dblnToR As Double, Optional intCntDec As Integer) As Double
   
    Dim dblPot As Double
    Dim dblF As Double
    
    If dblnToR < 0 Then dblF = -0.5 Else: dblF = 0.5
    dblPot = 10 ^ intCntDec
    Redondejar = Fix(dblnToR * dblPot * (1 + 1E-16) + dblF) / dblPot

End Function
Sub Main()
  
  If EstaCorriendo(App.EXEName) Then
     MsgBox2 "El programa de Manteniment de Tintes ja està obert.", 3, "Atenció", vbCritical
     If MsgBox("Vols tancart l'instància anterior?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then KillProcess App.EXEName
     End
  End If
  formtintes.Show
End Sub
Public Sub KillProcess(ByVal sProcessName As String)
    ' Kill process using Visual Basic 6 0 and WMI.
    ' The full .exe name (including the .exe) is supplied, but no path.
    ' Example: KillProcess "excel.exe"
    ' BE CAREFUL:  No prompt for saving takes place.
    '              ALSO, it kills all occurrences.
    Dim oWMI As Object
    Dim ret As Long
    Dim oServices As Object
    Dim oService As Object
    Dim sServiceName As String
    Dim bFoundOne As Boolean
    '
    On Error Resume Next
        sProcessName = LCase$(sProcessName)
        Set oWMI = GetObject("WinMgmts:")
        Set oServices = oWMI.InstancesOf("win32_process")
        '
        Do
            For Each oService In oServices
                sServiceName = LCase$(Trim$(CStr(oService.Name)))
                If sServiceName = sProcessName Then
                    ret = oService.Terminate
                    bFoundOne = True
                End If
            Next oService
            If Not bFoundOne Then Exit Do
            If err Then Exit Do
            bFoundOne = False
        Loop
    On Error GoTo 0
End Sub

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


Sub sonar_sirena(v As String)

End Sub
