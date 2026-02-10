Attribute VB_Name = "Module1"
Const LOCALE_SDECIMAL = &HE
Const LOCALE_STHOUSAND = &HF
Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long
Private Const STATUS_PENDING = &H103&
Private Const PROCESS_QUERY_INFORMATION = &H400



Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function EnumProcesses Lib "PSAPI.DLL" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
'Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
'Private Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Const PROCESS_VM_READ = &H10
Global vperforat As Boolean
Global contadorverificacio As Double
Global etiquetesean13 As Boolean
Global noespota0 As Boolean
Global rstconsulta As Recordset
Global idiomaclient As String
Global rstopcionset As Recordset
Global tempseditant As Date
Global pescanutu As Double
Global fitxerini As String
Global nomfitxertemporal As String
Global dbtemp As Database
Global camistocks As String
Global ncilindre As Double
Global horaapretada As Date
Global ntintes As Byte
Global nample As Double
Global refclient As String
Global micrescomanda As String
Global mesuraespcomanda As String
Global codibarras As String
Global comandaclient As String
Global idteclat As Variant
Global rsttmp As Recordset
Global campcontrol As Control
Global dbbaixes As Database
Global dbtmp As Database
Global dbtmpb As Database
Global bdllistat As Database
Global dbstocks As Database
Global rststocks As Recordset
Global idfuncionament As Long
Global tipus As String
Global queryorder As String
Global querywhere As String
Global cami As String
Global buscant As Boolean
Global seleccioret As Byte
Global ultimcontrol As Control
Global llocform As Byte
Global taulapos(15) As Variant
Global ruta As String
Global ruta_relativa_docs As String
Global ruta_relativa_client As String
Global tecla As Integer
Global i As Integer
Global r As String
Global pr As String
Global sa As String
Global re As String
Global espessor As String
Global controlcanviat As Control
Global colorcanviat As String
Global camicomandes As String
Global canvissortirseccio As Boolean
Global recordsourcetotals As String
Global iniconfigreixa As String
Global muntadora As Boolean
Global colorrisc As String
Global vlink1 As Double
Global vlink2 As Double
Global vlink3 As Double
Global nummaq As Byte
Global numop As Byte
Global amplereb As Double
Global numpalet As Byte
Global rstpespalet As Recordset
Global arguments As Variant

Global Const topeform = 5
Global Const colorfonscontrolactiu = 10


'Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal _
'       lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
'       lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long'


Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
                
Private Declare Function SendMessage Lib "user32" _
                Alias "SendMessageA" _
                (ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long


Public Enum StyleInputBox
    SNone
    SPassword
    SNumber
    SLowerCase
    SUpperCase
End Enum

Private Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long
                
Private Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
                (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function SetTimer Lib "user32" _
                (ByVal hwnd As Long, _
                ByVal nIDEvent As Long, _
                ByVal uElapse As Long, _
                ByVal lpTimerFunc As Long) As Long
                
Private Declare Function KillTimer Lib "user32" _
                (ByVal hwnd As Long, _
                ByVal nIDEvent As Long) As Long

Private Const GWL_STYLE = (-16)

' constantes con los estilos de
' controles 'EDIT'
Private Const ES_UPPERCASE = &H8
Private Const ES_LOWERCASE = &H10
Private Const ES_PASSWORD = &H20
Private Const ES_NUMBER = &H2000

' mensaje para establecer el caracter que se mostrará
' como máscara para el InputBoxEx tipo contraseña
Private Const EM_SETPASSWORDCHAR = &HCC
' constante que contiene el carácter que se mostrará
' (este valor puede ser cualquier otro, en este caso
' he escogido el típico asterisco)
Private Const KEY_MASK = 42& ' "*"
' mensaje para establecer el número máximo de
' caracteres permitidos
Private Const EM_LIMITTEXT = &HC5

Private SInputBox As StyleInputBox
Private hInputBox As Long
Private cChar As Long
Private Tm As Long


Const VK_TAB = &H9
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2

Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


' Estructura SHFILEOPSTRUCT o para usar con el Api
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

'Declaración Api SHFileOperation
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
                                                (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Constantes
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_SILENT = &H4


'

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim buffer As Integer
Declare Function apiGetPrivateProfileString Lib "kernel32" _
       Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
       As String, ByVal lpKeyName As Any, ByVal lpDefault As _
       String, ByVal lpReturnedString As String, ByVal nSize As _
       Long, ByVal lpFileName As String) As Long
Declare Function apiWritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
        As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
        ByVal lpFileName As String) As Long

Public Sub Copiar_Fitxer(ByVal Origen As String, ByVal Destino As String, Optional opcions As Long)

Dim t_Op As SHFILEOPSTRUCT
  If IsNull(opcions) Then opcions = FOF_ALLOWUNDO
    With t_Op
        .hwnd = 0
        .wFunc = FO_COPY
        .pFrom = Origen & vbNullChar & vbNullChar
        .pTo = Destino & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_SILENT
    End With
    If opcions = 5 Then
    With t_Op
        .hwnd = 0
        .wFunc = FO_COPY
        .pFrom = Origen & vbNullChar & vbNullChar
        .pTo = Destino & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
        
    End With
    End If
    If opcions = 6 Then
    With t_Op
        .hwnd = 0
        .wFunc = FO_COPY
        .pFrom = Origen & vbNullChar & vbNullChar
        .pTo = Destino & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With
    End If

    ' Se ejecuta la función Api pasandole la estructura
    SHFileOperation t_Op
    
    
End Sub

Sub obrir_document(nomfitxer As String)
 Dim vaaa As Integer
 vaaa = ShellExecute(Screen.ActiveForm.hwnd, "Open", nomfitxer, "", "", 1)
 r = ""
End Sub

Function MsgBox2( _
                Prompt As String, _
                Optional SecondsToWait, _
                Optional Title, _
                Optional Buttons) As VbMsgBoxResult


Dim wShell As Object 'New wshShell


'    Set wShell = CreateObject("WScript.Shell")
'    MsgBox2 = wShell.PopUp(Prompt, SecondsToWait, Title, Buttons)
 '   Set wShell = Nothing
  Load avis
  avis.missatge = Prompt
  avis.Caption = Title
  If InStr(1, Screen.ActiveForm.Caption, "Rebobinadores") > 0 Then avis.Show: wait 5
  
  Unload avis
End Function

Function MsgBoxandWait( _
                Prompt As String, _
                Optional SecondsToWait, _
                Optional Title, _
                Optional Buttons) As VbMsgBoxResult
    Dim wShell As Object 'New wshShell
    Set wShell = CreateObject("WScript.Shell")
   ' MsgBoxandWait = wShell.PopUp(Prompt, SecondsToWait, Title, Buttons) no funciona hi ha algun problema de microsoft
    CreateObject("WScript.Shell").Run "mshta.exe vbscript:close(CreateObject(""WScript.Shell"").Popup(""" + Prompt + """," + atrim(SecondsToWait) + ",""" + Title + """))"
    wait cadbl(SecondsToWait)
    Set wShell = Nothing
End Function

Sub controldeteclat()
  Dim intascii
  
  For intascii = 1 To 255
   buffer = 0
   buffer = GetAsyncKeyState(intascii)
   If buffer <> 0 Then
     tecla = intascii
   End If
   'Screen.ActiveForm.Caption = tecla
  Next
  'If tecla = 110 Then
   '   SendKeys "{BACKSPACE}": SendKeys ","
   '   tecla = 0
   'End If
End Sub
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
Function llegir_ini2(ByVal Ap As String, ByVal cl As String, ByVal ini As String) As String
  Dim va As String
  Dim r As Integer
  cl = Trim(cl)
  va = Space$(255)
  r = apiGetPrivateProfileString(Ap, cl, "{[}]", va, 255, ini)
  If Mid(va, 1, 4) <> "{[}]" Then
     va = Mid(va, 1, Len(RTrim(va)) - 1)
   Else: va = ""
  End If
  llegir_ini2 = va
End Function

Sub escriure_ini(Ap As String, cl As String, tex As String, ini As String)

  Dim r As Integer
  cl = Trim(cl)
  r = apiWritePrivateProfileString(Ap, cl, tex, ini)
End Sub
Function atrim(valo As Variant) As String
On Error Resume Next
  If IsNull(valo) Then valo = ""
  atrim = Trim(valo)
End Function
Function cadbl(ByVal valo As Variant) As Double
  If Not IsNumeric(valo) Then valo = 0
  cadbl = CDbl(valo)
End Function


Function treure_apostruf(ByVal N As String) As String
   While InStr(N, "'")
     N = Mid(N, 1, InStr(1, N, "'") - 1) + "´" + Mid(N, InStr(1, N, "'") + 1)
   Wend
   If N = "{[}]" Then N = ""
   treure_apostruf = N
End Function
Function posar_apostruf(ByVal N As String) As String
   While InStr(N, "´")
     N = Mid(N, 1, InStr(1, N, "´") - 1) + "'" + Mid(N, InStr(1, N, "´") + 1)
   Wend
   posar_apostruf = N
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
'Function textestattaula(estatval As Byte) As String
'  Dim texteestat As String
'  If estatval = 0 Then textestat = ""
'  If estatval = 1 Then texteestat = "Editant..."
'  If estatval = 2 Then texteestat = "Afegint..."
'  If buscant Then texteestat = "Buscant..."
'  textestattaula = texteestat
'End Function

Sub centerscreen(fo As Form)
On Error GoTo fi
'center screen
    fo.Top = (Screen.Height / 2) - (fo.Height / 2)
    fo.Left = (Screen.Width / 2) - (fo.Width / 2)
  'fins aqui
fi:
End Sub

Sub ratoli(estat As String)
  
  If estat = "espera" Then Screen.MousePointer = 11
  If estat = "normal" Then Screen.MousePointer = 0
  DoEvents
End Sub

Function tamany_camp(camp As Object) As Integer

  Select Case camp.Type
     Case 4 ' enter llarg
       tamany_camp = 7
     Case 2 'byte
       tamany_camp = 3
     Case 7 'double
       tamany_camp = 10
     Case 10 'texte
       tamany_camp = camp.Size
     Case 8  'data
       tamany_camp = 10
  End Select
  
End Function

Function format_camp(camp As Object) As String
  Select Case camp.Type
     Case 4 ' enter llarg
       format_camp = "#,##0"
     Case 2 'byte
       format_camp = "##0"
     Case 7 'double
       format_camp = "#,##0.00"
     Case 10 'texte
       format_camp = "@"
     Case 8  'data
       format_camp = "dd/mm/yyyy"
  End Select
  
End Function

Function mascara_camp(camp As Object) As String
  Select Case camp.Type
     Case 4 ' enter llarg
       mascara_camp = "#####"
     Case 2 'byte
       mascara_camp = "###"
     Case 7 'double
       mascara_camp = "##########"
     Case 10 'texte
       mascara_camp = ">A"
     Case 8  'data
       mascara_camp = "99/99/99"
  End Select
  
End Function

Sub canviarelscolorsdelscontrolsalentrar()
'canvia el color dels controls al entrar a dins
 On Error Resume Next
 If controlcanviat Is Nothing Then Set controlcanviat = Screen.ActiveControl: colorcanviat = controlcanviat.BackColor

 If controlcanviat.Name <> Screen.ActiveControl.Name Then
    controlcanviat.BackColor = colorcanviat
    If TypeOf controlcanviat Is MaskEdBox Then controlcanviat.Text = passaradecimal(controlcanviat.Text)
    Set controlcanviat = Screen.ActiveControl
    colorcanviat = Screen.ActiveControl.BackColor
    If TypeOf Screen.ActiveControl Is TextBox Or TypeOf Screen.ActiveControl Is MaskEdBox Or TypeOf Screen.ActiveControl Is ComboBox Then
     Screen.ActiveControl.BackColor = QBColor(colorfonscontrolactiu)
     Screen.ActiveControl.SelStart = 0
     Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
    End If
 End If
End Sub
Function passaradecimal(valormasked As String) As String
   Dim valmas As String
   valmas = valormasked
   While InStr(1, valmas, ".") > 1
     valmas = Mid(valmas, 1, InStr(1, valmas, ".") - 1) + "," + Mid(valmas, InStr(1, valmas, ".") + 1)
   Wend
   If IsNumeric(valmas) Then valormasked = valmas
   passaradecimal = valormasked
End Function
Function passaradecimalpunt(valormasked As Variant) As String
   Dim valmas As String
   valmas = valormasked
   While InStr(1, valmas, ",") > 1
     valmas = Mid(valmas, 1, InStr(1, valmas, ",") - 1) + "." + Mid(valmas, InStr(1, valmas, ",") + 1)
   Wend
   valormasked = valmas
   passaradecimalpunt = valormasked
End Function
Function passaradecimalpunt2(valormasked As Variant) As String
   Dim valmas As String
   valmas = valormasked
   While InStr(1, valmas, ",") > 1
     valmas = Mid(valmas, 1, InStr(1, valmas, ",") - 1) + "." + Mid(valmas, InStr(1, valmas, ",") + 1)
   Wend
  ' valormasked = valmas
   passaradecimalpunt2 = valmas
End Function

Sub modificar_estat_comanda(comanda As Long, ruta As String, seccio As String, proxima As Boolean)

Dim rutav As String
rutav = "EILRSV"
If Mid(ruta, Len(ruta), 1) = "T" Then
   ruta = Mid(ruta, 1, Len(ruta) - 1) + "V"
  Else: ruta = ruta + "V"
End If
'rutav = ruta
If InStr(1, rutav, sa) < InStr(1, rutav, seccio) Then
   dbtmp.Execute "update comandes set seccioactual='" + seccio + "' where comanda=" + atrim(comanda)
End If
If proxima Then
    'dbtmp.Execute "update comandes set proximaseccio='" + Mid(ruta, InStr(1, ruta, seccio) + 1, 1) + "' where comanda=" + atrim(comanda)
    If InStr(1, rutav, pr) <= InStr(1, rutav, Mid(ruta, InStr(1, ruta, seccio) + 1, 1)) Then
     If Mid(ruta, Len(ruta), 1) = "V" Then ruta = ruta + "T"
     dbtmp.Execute "update comandes set proximaseccio='" + Mid(ruta, InStr(1, ruta, seccio) + 1, 1) + "' where comanda=" + atrim(comanda)
    End If
  Else:
     If seccio = "V" Then seccio = "P": ruta = ruta + "P"
     If InStr(1, rutav, pr) <= InStr(1, rutav, seccio) Then
       dbtmp.Execute "update comandes set proximaseccio='" + seccio + "' where comanda=" + atrim(comanda)
     End If
End If

End Sub
 
 'Sub controlar_fiseccio(seccio As String, ruta As String, Optional hihabobines As Boolean)
 ' If canvissortirseccio Then
 '   If hihabobines Then
 '     modificar_estat_comanda entradabaixes.comanda, ruta + "T", seccio, 1
 '    Else: modificar_estat_comanda entradabaixes.comanda, ruta + "T", seccio, 0
 '    End If
 ' End If
 ' If Not canvissortirseccio Then End
 ' entradabaixes.Visible = True
 'End Sub
Sub obrir_baixes()
  Shell llegir_ini("General", "rutaprogbaixes", "comandes.ini"), vbNormalFocus
End Sub
'Function carpeta_del_client() As String
'
'    formcomandes.data1tmp.DatabaseName = cami
'    formcomandes.data1tmp.RecordSource = "select * from carpeta_client where codiclient=" + atrim(cadbl((formcomandes.Data1.Recordset!client)))
'    formcomandes.data1tmp.Refresh
'    If Not formcomandes.data1tmp.Recordset.EOF Then carpeta_del_client = formcomandes.data1tmp.Recordset!nomcarpeta
'
'End Function
Sub wait(segonsespera As Byte)
  horaentradawait = Now
  While DateDiff("s", horaentradawait, Now) < segonsespera
    DoEvents
  Wend
End Sub


Public Function InputBoxEx( _
                Prompt, _
                Optional Title, _
                Optional Default, _
                Optional XPos, _
                Optional YPos, _
                Optional HelpFile, _
                Optional Context, _
                Optional Style As StyleInputBox = SNone, _
                Optional MaxChar As Long) As String
              
    ' si no hay ningún otro InputBoxEx abierto...
    If hInputBox = 0 Then
       ' Creamos un timer que se ejecutará a la décima de segundo
       Tm = SetTimer(0&, 0&, 100, AddressOf TimerProc)
    
       SInputBox = Style
       cChar = MaxChar
       ' llamamos al InputBox de manera normal
       On Error GoTo AnularTimer
       InputBoxEx = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)
    End If
    
    Exit Function
    
AnularTimer:
    ' si ha habido algún error, se cancela la operación
    Call KillTimer(0&, Tm)
    MsgBox "Error: " & err.Number & vbCrLf & err.Description
    
End Function

Private Sub TimerProc( _
                     ByVal hwnd As Long, _
                     ByVal uMsg As Long, _
                     ByVal idEvent As Long, _
                     ByVal dwTime As Long)
Dim hEdit As Long
Dim CurStyle As Long
    
    ' localizamos el manipulador de la ventana activa
    ' (se supone que es la ventana del InputBox)
    hInputBox = GetForegroundWindow
    ' localizamos el manipulador de la caja de texto
    ' del InputBox
    hEdit = FindWindowEx(hInputBox, 0&, "EDIT", vbNullString)
    
    ' obtenemos los estilos de la caja de texto ...
    CurStyle = GetWindowLong(hEdit, GWL_STYLE)
    
    Select Case SInputBox
        Case SPassword ' tipo password
            ' le decimos a la caja de texto cuál será el carácter
            ' que aparecerá en vez de lo que teclee el usuario
            Call SendMessage(hEdit, EM_SETPASSWORDCHAR, KEY_MASK, 0&)
            ' y le añadimos el estilo de introducción de contraseñas
            CurStyle = CurStyle Or ES_PASSWORD
        Case SNumber ' tipo número
            CurStyle = CurStyle Or ES_NUMBER
        Case SLowerCase ' tipo minúsculas
            CurStyle = CurStyle Or ES_LOWERCASE
        Case SUpperCase ' tipo mayúsculas
            CurStyle = CurStyle Or ES_UPPERCASE
    End Select
    
    If cChar > 0 Then
        Call SendMessage(hEdit, EM_LIMITTEXT, cChar, 0&)
    End If
    ' cambiamos el estilo
    Call SetWindowLong(hEdit, GWL_STYLE, CurStyle)
    ' desactivamos el timer para que sólo se ejecute esta vez
    Call KillTimer(0&, Tm)
    hInputBox = 0
    
End Sub
'---------------------------------------------------------


Public Function ShellandWait(ExeFullPath As String, Optional TimeOutValue As Long = 0) As Boolean
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
lInst = Shell(sExeName, vbMinimizedNoFocus)
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

 
 'Public Function EstaCorriendo(ByVal NombreDelProceso As String) As Boolean
 '   Const MAX_PATH As Long = 260
 '   Dim con As Byte
 '   Dim lProcesses() As Long, lModules() As Long, N As Long, lRet As Long, hProcess As Long
 '   Dim sName As String
 '   NombreDelProceso = UCase$(NombreDelProceso)
 '   ReDim lProcesses(1023) As Long
 '   con = 0
 '   If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
 '       For N = 0 To (lRet \ 4) - 1
 '           hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(N))
 '           If hProcess Then
 '               ReDim lModules(1023)
 '               If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
 '                   sName = String$(MAX_PATH, vbNullChar)
 '                   GetModuleBaseName hProcess, lModules(0), sName, MAX_PATH
 '                   sName = Left$(sName, InStr(sName, vbNullChar) - 1)
'
'                    If Len(sName) = Len(NombreDelProceso) Then
'                        If NombreDelProceso = UCase$(sName) Then con = con + 1
'                        If con = 2 Then EstaCorriendo = True: Exit Function
''                    End If
 '               End If
 '           End If
 '           CloseHandle hProcess
 '       Next N
 '   End If
'End Function




Public Function Redondejar(dblnToR As Double, Optional intCntDec As Integer) As Double
   
    Dim dblPot As Double
    Dim dblF As Double
    
    If dblnToR < 0 Then dblF = -0.5 Else: dblF = 0.5
    dblPot = 10 ^ intCntDec
    Redondejar = Fix(dblnToR * dblPot * (1 + 1E-16) + dblF) / dblPot
   
End Function

Sub assignardecimalipunt()
  Dim LocalID As Long
  If Not existeix("c:\ordprog.ini") And nummaq > 0 And InStr(1, UCase(Environ("computername")), "EXP") = 0 Then
    LocalID = GetUserDefaultLCID()
    SetLocaleInfo LocalID, LOCALE_SDECIMAL, "."
    SetLocaleInfo LocalID, LOCALE_STHOUSAND, ","
  End If
End Sub

  
  
Function rutadelfitxer(cam As String) As String
   Dim c As Byte
   c = 0
   While InStr(c + 1, cam, "\") <> 0
    c = InStr(c + 1, cam, "\")
   Wend
   If c = 0 Then c = Len(cam)
   rutadelfitxer = Mid(cam, 1, c)
End Function
Function substituirtot(cadena As String, buscar As String, canviar As String) As String
   cadena = " " + cadena
   comença = 1
   While InStr(comença, cadena, buscar) > 0
    
    comença = InStr(comença, cadena, buscar) - 1
    If comença < 1 Then substituirtot = cadena: Exit Function
    
    acaba = comença + Len(buscar) + 1
    cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
    comença = acaba
   Wend
   If Mid(cadena, 1, 1) = " " Then cadena = Mid(cadena, 2)
   substituirtot = cadena
   'MsgBox linia
End Function
Function treuresimbols(desc As String) As String
'   desc = substituir(desc, ":", "_")
   desc = substituirtot(desc, "'", "´")
   desc = substituirtot(desc, "|", "_")
   desc = substituirtot(desc, ";", "_")
   treuresimbols = desc
End Function
Sub enviaremailgeneric(destinatari As String, assumpte As String, cos As String)
   Dim dbenvio As Database
   
   If atrim(cos) = "" Then Exit Sub
   Set dbenvio = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
   dbenvio.Execute "insert into envios_mails (data,destinatari,assumpte,cos) values (now,'" + destinatari + "','" + treuresimbols(assumpte) + "','" + treuresimbols(cos) + "')"
   Set dbenvio = Nothing
End Sub

Function substituir(cadena As String, buscar As String, canviar As String) As String
   comença = InStr(1, cadena, buscar) - 1
   If comença < 1 Then substituir = cadena: Exit Function
   acaba = comença + Len(buscar) + 1
   cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   substituir = cadena
   'MsgBox linia
End Function

