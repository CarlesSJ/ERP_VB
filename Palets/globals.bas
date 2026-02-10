Attribute VB_Name = "Module2"
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess _
    As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle _
    As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const LOCALE_SDECIMAL = &HE
Const LOCALE_STHOUSAND = &HF
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
  missatge.etimissatge.Caption = "Creando listado, Espere ..."
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
      OpenFile.hwndOwner = formcomandes.hwnd
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
  If vfiltrebobinesdesdeimpresores Then Exit Sub
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


Public Sub concatenar_fitxers(fgsp As String, fseidor As String)
    Dim vlinia As String
    Open fseidor For Append As 3
    Open fgsp For Input As 4
    While Not EOF(4)
       Line Input #4, vlinia
       Print #3, vlinia
    Wend
    Close 3
    Close 4
End Sub

Function nomproveidor(vidpalet As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("SELECT Palets.Idpalet, proveidors.nom FROM (Palets INNER JOIN materials ON Palets.codimatprognou = materials.codi) INNER JOIN proveidors ON materials.proveidor = proveidors.codi WHERE Palets.Idpalet=" + atrim(vidpalet) + ";")
   If Not rst.EOF Then nomproveidor = atrim(rst!nom)
   Set rst = Nothing
End Function

Sub SendKeys(vtecla As String)
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

WshShell.SendKeys vtecla
Set WshShell = Nothing

End Sub
