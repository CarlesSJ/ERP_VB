Attribute VB_Name = "Module1"
Global fitxerini As String
Global db As Database
Global rs As Recordset
Global cami As String
Global dbcomandes As Database


Declare Function apiGetPrivateProfileString Lib "kernel32" _
       Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
       As String, ByVal lpKeyName As Any, ByVal lpDefault As _
       String, ByVal lpReturnedString As String, ByVal nSize As _
       Long, ByVal lpFileName As String) As Long

Private Declare Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long
Private Const STATUS_PENDING = &H103&
Private Const PROCESS_QUERY_INFORMATION = &H400
Function llegir_ini(ByVal Ap As String, ByVal cl As String, ByVal ini As String) As String
  Dim va As String
  Dim r As Integer
  cl = Trim(cl)
  va = Space$(255)
  r = apiGetPrivateProfileString(Ap, cl, "{[}]", va, 255, ini)
  If Mid(va, 1, 4) <> "{[}]" Then
     va = Mid(va, 1, Len(Trim(va)) - 1)
   Else: va = "{[}]"
  End If
  llegir_ini = va
End Function


Public Function ShellandWait(ExeFullPath As String, Optional TimeOutValue As Long = 0, Optional focus As Byte) As Boolean
 Dim lInst As Long
Dim lStart As Long
Dim lTimeToQuit As Long
Dim sExeName As String
 Dim lProcessId As Long
Dim lExitCode As Long
Dim bPastMidnight As Boolean
On Error GoTo ErrorHandler
lStart = CLng(Timer)
sExeName = ExeFullPath
'Deal with timeout being reset at Midnight
If TimeOutValue > 0 Then
      If lStart + TimeOutValue < 86400 Then
        lTimeToQuit = lStart + TimeOutValue
        Else
          lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
      End If
 End If
lInst = Shell(sExeName, IIf(focus = 1, vbNormalFocus, vbMinimizedNoFocus))
lProcessId = OpenProcess(PROCESS_QUERY_INFORMATION, False, lInst) 'Optenemos el ProcessID
  Do 'Aqui se genera un ciclo hasta que el proceso sea distinto de pendiente, o sea, Alla terminado.
    Call GetExitCodeProcess(lProcessId, lExitCode) ' Optenemos el si hay exits code o todavia esta en ejecucion (pending)
    DoEvents
    If TimeOutValue And Timer > lTimeToQuit Then
    If bPastMidnight Then
     If Timer < lStart Then Exit Do
       Else
         Exit Do ' Se sale del ciclo si se acavo el tiemo de espera
     End If
    End If
  Loop While lExitCode = STATUS_PENDING
ShellandWait = True
ErrorHandler:
ShellandWait = False
Exit Function
End Function
Function existeix(nomfitxer As String) As Boolean
Dim a As Integer
On Error GoTo err:
 a = GetAttr(nomfitxer)
 existeix = True
 Exit Function
err:
 existeix = False
End Function
Public Function rutadelfitxer(cam As String) As String
   Dim C As Byte
   C = 0
   While InStr(C + 1, cam, "\") <> 0
    C = InStr(C + 1, cam, "\")
   Wend
   If C = 0 Then C = Len(cam)
   rutadelfitxer = Mid(cam, 1, C)
End Function

