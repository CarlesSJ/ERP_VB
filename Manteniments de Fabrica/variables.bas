Attribute VB_Name = "Module1"
Public Const SWP_NOMOVE = 2
      Public Const SWP_NOSIZE = 1
      Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
      Public Const HWND_TOPMOST = -1
      Public Const HWND_NOTOPMOST = -2


Public Type TipusVrisc
 nomdelclient As String
 creditsap As Double
 creditgastatsap As Double
 valorestoc As Double
 valorpendent As Double
 valorproduccio As Double
 valordelsclixes As Double
 valoralbaranspendentsSAP As Double 'valor albarans pendents de facturar a SAP
 valordiferencial As Double
 comandesazero As String
 comandesazerodetall As String
 comandesazeroTotalKg As Double
 comandesestoc As String
 comandespendent As String
 comandesproduccio As String
 albaransSAP As String 'llista dalbarans pendents de facturar a SAP
 treballspendents As String
End Type
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess _
    As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle _
    As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Global arguments As Variant
Global vtreballbuscatsubbusqueda As String
Global campsestat(50, 4) As String
Global dbcontrolcanvis As Database
Global rstcontrolcanvis As Recordset
Global rstcontrolcanvis_extres As Recordset
Global carregant As Boolean
Global dbtemp As Database
Global dbclixes As Database
Global dbclixesnous As Database
Global dbplanificacio As Database
Global dbsap As Database
Global ruta_documentacio_clixes As String
Global horaentradawait As Date
Global cridacomandes As Boolean
Global imprimircomandes As Boolean
Global exportarcomandes As Boolean
Global numcomanda As String
Global campscontrolalicia(20, 2) As String
Global dbconsulta As Database
Global rstconsulta As Recordset
Global rstopcionset As Recordset
Global idiomaclient As String
Global idiomaclientclixes As String
Global taula_tmp As String
Global Idioma As String
Global fitxerini As String
Global llistadecampsvalids As String
Global dbenvios As Database
Global rsttmp As Recordset
Global campcontrol As Control
Global dbbaixes As Database
Global db As Database
Global dbtmp As Database
Global dbtmpb As Database
Global dbtarifes As Database
Global dbcomandes As Database
Global bdllistat As Database
Global dbcompres As Database
Global dbstocks As Database
Global rststocks As Recordset
Global idfuncionament As Long
Global tipus As String
Global queryorder As String
Global querywhere As String
Global cami As String
Global camiclixes As String
Global camistock As String
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
Global duplicant As Boolean

Global Const topeform = 5
Global Const colorfonscontrolactiu = 10


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
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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
 'Shell "c:\windows\system32\cmd.exe /c start " + nomfitxer
 If InStr(1, nomfitxer, ".doc") Then If Not existeix(nomfitxer) Then nomfitxer = nomfitxer + "x"
 vaaa = ShellExecute(Screen.ActiveForm.hwnd, "Open", nomfitxer, "", "", 1)
 r = ""
End Sub

Sub imprimir_document(nomfitxer As String)
 Dim vaaa As Integer
 If Not existeix(nomfitxer) Then nomfitxer = nomfitxer + "x"
 vaaa = ShellExecute(Screen.ActiveForm.hwnd, "Print", nomfitxer, "", "", 1)
 r = ""
 wait 1
End Sub

Function MsgBox2( _
                Prompt As String, _
                Optional SecondsToWait, _
                Optional Title, _
                Optional Buttons) As VbMsgBoxResult


Dim wShell As Object 'New wshShell



  Load avis
  avis.missatge = Prompt
  avis.Caption = Title
  avis.Show
  wait 5
  Unload avis
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


Function treure_apostruf(ByVal n As String) As String
   While InStr(n, "'")
     n = Mid(n, 1, InStr(1, n, "'") - 1) + "´" + Mid(n, InStr(1, n, "'") + 1)
   Wend
   If n = "{[}]" Then n = ""
   treure_apostruf = n
End Function
Function posar_apostruf(ByVal n As String) As String
   While InStr(n, "´")
     n = Mid(n, 1, InStr(1, n, "´") - 1) + "'" + Mid(n, InStr(1, n, "´") + 1)
   Wend
   posar_apostruf = n
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
Function textestattaula(estatval As Byte) As String
  Dim texteestat As String
  If estatval = 0 Then texteestat = ""
  If estatval = 1 Then texteestat = "Editant..."
  If estatval = 2 Then texteestat = "Afegint..."
  If buscant Then texteestat = "Buscant..."
  textestattaula = texteestat
End Function

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
       tamany_camp = 12
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
Function noesungrupdecontrols()
   noesungrupdecontrols = True
   On Error GoTo fi
   If Screen.ActiveControl.Index <> -1 Then noesungrupdecontrols = False
fi:
   
End Function
Sub canviarelscolorsdelscontrolsalentrar()
'canvia el color dels controls al entrar a dins
 On Error Resume Next
 If controlcanviat Is Nothing Then Set controlcanviat = Screen.ActiveControl: colorcanviat = controlcanviat.BackColor
 
 If Screen.ActiveForm.HelpContextID <> 100 And controlcanviat.Name <> Screen.ActiveControl.Name And noesungrupdecontrols Then
 
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
   If Mid(valormasked + " ", 1, 1) = "." Then valormasked = "0" + valormasked
   valmas = valormasked
   While InStr(1, valmas, ".") > 1
     valmas = Mid(valmas, 1, InStr(1, valmas, ".") - 1) + "," + Mid(valmas, InStr(1, valmas, ".") + 1)
   Wend
   If IsNumeric(valmas) Then valormasked = valmas
   passaradecimal = valormasked
End Function
Function passaradecimalpunt(valormasked As String) As String
   Dim valmas As String
   valmas = valormasked
   While InStr(1, valmas, ",") > 1
     valmas = Mid(valmas, 1, InStr(1, valmas, ",") - 1) + "." + Mid(valmas, InStr(1, valmas, ",") + 1)
   Wend
   valormasked = valmas
   passaradecimalpunt = valormasked
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
    If taulamaquines(seccio) <> "" And seccio <> "S" Then dbtmpb.Execute "update " + taulamaquines(seccio) + " set acavada=1 where comanda=" + atrim(comanda)
  Else:
     If seccio = "V" Then seccio = "P": ruta = ruta + "P"
     If InStr(1, rutav, pr) <= InStr(1, rutav, seccio) Then
       dbtmp.Execute "update comandes set proximaseccio='" + seccio + "' where comanda=" + atrim(comanda)
     End If
End If

End Sub
Function taulamaquines(seccio As String) As String
   Select Case seccio
      Case "I"
         taulamaquines = "Impressorestot"
      Case "L"
         taulamaquines = "Laminadorestot"
      Case "R"
         taulamaquines = "Rebobinadorestot"
      Case "S"
         taulamaquines = "Soldadorestot"
   End Select
End Function
 
 Sub controlar_fiseccio(seccio As String, ruta As String, Optional hihabobines As Boolean)
  If canvissortirseccio Then
   If MsgBox("Vols fer els canvis d'estat de comanda.", vbInformation + vbYesNo, "Canvis proxima seccio") = vbYes Then
    If hihabobines Then
      modificar_estat_comanda entradabaixes.comanda, ruta + "T", seccio, 1
       Else: modificar_estat_comanda entradabaixes.comanda, ruta + "T", seccio, 0
     End If
   End If
  End If
  If Not canvissortirseccio Then End
  entradabaixes.Visible = True
 End Sub
Sub obrir_baixes()
  Shell llegir_ini("General", "rutaprogbaixes", fitxerini), vbNormalFocus
End Sub
Function carpeta_del_client() As String
 Dim rstc As Recordset
    'formcomandes.data1tmp.DatabaseName = cami
    'formcomandes.data1tmp.RecordSource = "select * from carpeta_client where codiclient=" + atrim(cadbl((formcomandes.data1.Recordset!client)))
    Set rstc = formcomandes.Data1.Database.OpenRecordset("select * from carpeta_client where codiclient=" + atrim(cadbl((formcomandes.Data1.Recordset!client))))
    'formcomandes.data1tmp.Refresh
    If Not rstc.EOF Then
        carpeta_del_client = rstc!nomcarpeta
       'Else: MsgBox "No s'ha trobat la carpeta del client PROVA DE REINDEXAR AL MENU UTIL REINDEXAR CLIENTS."
    End If
    Set rstc = Nothing
End Function
Sub wait(segonsespera As Byte)
'  horaentradawait = Now
'  While DateDiff("s", horaentradawait, Now) < segonsespera
'    DoEvents
'  Wend
Sleep segonsespera * 1000
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




Public Function Redondejar(dblnToR As Double, Optional intCntDec As Integer) As Double
   
    Dim dblPot As Double
    Dim dblF As Double
    
    If dblnToR < 0 Then dblF = -0.5 Else: dblF = 0.5
    dblPot = 10 ^ intCntDec
    Redondejar = Fix(dblnToR * dblPot * (1 + 1E-16) + dblF) / dblPot

End Function
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



Function substituirtots(cadena As String, buscar As String, canviar As String) As String
  comença = 1
  While InStr(comença, cadena, buscar) > 0
   comença = InStr(1, cadena, buscar) - 1
   If comença < 1 Then GoTo fi
   acaba = comença + Len(buscar) + 1
   cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   comença = comença + Len(buscar) + 1
  Wend
fi:
  substituirtots = cadena
   
   'MsgBox linia
End Function

Function treuresimbols(desc As String) As String
   desc = substituir(desc, ":", "_")
   desc = substituir(desc, "'", "´")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ";", "_")
   treuresimbols = desc
End Function
Function descripciomaterial(rstmat As Recordset) As String
  Dim desc As String
  Dim rstfam As Recordset
  If rstmat.EOF Then Exit Function
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)))
  If Not rstfam.EOF Then desc = desc + atrim(rstfam!descripcio)
  'Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliesmaterials where codi=" + atrim(cadbl(rstmat!subfamilia)))
  'If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  'Set rstfam = dbtmpb.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(rstmat!familiacol)))
  'If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  'Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliescolorants where codi=" + atrim(cadbl(rstmat!subfamiliacol)))
  'If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  'Set rstfam = dbtmpb.OpenRecordset("select descripcio from familiesaditius where codi=" + atrim(cadbl(rstmat!familiaad)))
  'If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  'Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliesaditius where codi=" + atrim(cadbl(rstmat!subfamiliaad)))
  'If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  descripciomaterial = desc
End Function
Function af(v As Variant) As String
  v = atrim(v)
  If Len(v) > 1 Then
     v = " + " + v
    Else: v = ""
  End If
  af = v
End Function
 Sub calcular_credit_delclient(vcodiclient As Double, vrisc As TipusVrisc)
  Dim rst As Recordset
  Dim vvalorcomanda As Double
  Dim rstclixes As Recordset
  Dim rstvendes As Recordset
  Dim vsql As String
  Dim vrisc_buida As TipusVrisc
  Dim vquantitatservida As Double
  Dim viva As Double
  If vcodiclient = 0 Then Exit Sub
  vrisc = vrisc_buida
  vsql = "SELECT Modificacions.id_treball, Modificacions.ordre, Modificacions.codiclientfactclixes, pressupostos.preu, pressupostos.comfirmat, pressupostos.lotambelqueshafacturat"
  vsql = vsql + " FROM pressupostos RIGHT JOIN Modificacions ON (pressupostos.id_treball = Modificacions.id_treball) AND (pressupostos.ordremodificacio = Modificacions.ordre)"
  vsql = vsql + " WHERE (((Modificacions.codiclientfactclixes)<>'' And (Modificacions.codiclientfactclixes)<>'0') AND ((pressupostos.comfirmat)=True) AND ((pressupostos.lotambelqueshafacturat)=0)) "
  Set rstclixes = dbclixes.OpenRecordset(vsql)
  Set rst = dbtmp.OpenRecordset("select * from clients_codisSAP where codiSAP=" + atrim(vcodiclient))
  If Not rst.EOF Then viva = IIf(UCase(Mid(atrim(rst!nif) + "   ", 1, 2)) = "ES", 21, 0)
  Set rst = dbtmp.OpenRecordset("select * from credit_clients_inp where cardcode='" + atrim(vcodiclient) + "'")
  If Not rst.EOF Then
      vrisc.creditsap = cadbl(rst!creditline)
      vrisc.creditgastatsap = cadbl(rst!balance)
      vrisc.nomdelclient = atrim(rst!cardname)
  End If
  
  Set rst = dbtmp.OpenRecordset("select comanda,codicomptable,proximaseccio,pvp,tubbaseext,rebkilos,cantitatsol from comandesmesextres where producte<>'PC' and producte<>'PC2' and producte<>'PCP' and proximaseccio<>'T' and codicomptable=" + atrim(vcodiclient) + " order by proximaseccio")
  vrisc.comandespendent = "] "
  vrisc.comandesproduccio = "] "
  vrisc.comandesestoc = "] "
  While Not rst.EOF
     vdesccomanda = ""
     vquantitatservida = 0
     'Set rstvendes = dbtmp.OpenRecordset("select sum(quantitat) as quantitatentregada from liniesalbara where lotinplacsa=" + atrim(rst!comanda))
     Set rstvendes = dbtmp.OpenRecordset("SELECT Sum(quantitat) AS quantitatentregada, First([dataenvioasap]) AS vdataenvioasap FROM liniesalbara LEFT JOIN capcaleraalbara ON liniesalbara.numalbara = capcaleraalbara.numalbara WHERE liniesalbara.lotinplacsa=" + atrim(rst!comanda))
     If Not rstvendes.EOF Then
        If Not IsNull(rstvendes!vdataenvioasap) Then vendesvquantitatservida = cadbl(rstvendes!quantitatentregada)
     End If
     If cadbl(rst!rebkilos) > 0 Then
        vdesccomanda = atrim(rst!rebkilos) + "Kg "
         Else:
            If cadbl(rst!cantitatsol) > 0 Then
               vdesccomanda = atrim(rst!cantitatsol) + "Un "
                Else: vdesccomanda = "--------"
            End If
     End If
     If cadbl(rst!pvp) = 0 Then
        vrisc.comandesazeroTotalKg = vrisc.comandesazeroTotalKg + cadbl(rst!rebkilos)
        vrisc.comandesazero = vrisc.comandesazero + IIf(vrisc.comandesazero = "", "", ",") + vdesccomanda ': GoTo cont
        vrisc.comandesazerodetall = vrisc.comandesazerodetall + "[" + atrim(rst!proximaseccio) + "]" + atrim(rst!comanda) + " -> " + vdesccomanda + "    " ': GoTo cont
     End If
     vvalorcomanda = Redondejar(cadbl(rst!pvp) * (cadbl(rst!tubbaseext) - vquantitatservida), 2)
     If vvalorcomanda < 0 Then vvalorcomanda = 0
     vdesccomanda = "  " + atrim(rst!comanda) + "-" + vdesccomanda + IIf(vquantitatservida > 0, "P", "") + "(" + Format(vvalorcomanda, "#,##0") + "€)"
     If rst!proximaseccio = "E" Then
       vrisc.valorpendent = vrisc.valorpendent + vvalorcomanda
       If cadbl(rst!pvp) = 0 Then
          vrisc.comandespendent = vdesccomanda + vrisc.comandespendent
           Else: vrisc.comandespendent = vrisc.comandespendent + vdesccomanda
       End If
     End If
     If rst!proximaseccio = "V" Or rst!proximaseccio = "P" Then
       vrisc.valorestoc = vrisc.valorestoc + vvalorcomanda
       If cadbl(rst!pvp) = 0 Then
           vrisc.comandesestoc = vdesccomanda + vrisc.comandesestoc
            Else: vrisc.comandesestoc = vrisc.comandesestoc + vdesccomanda
        End If
     End If
     If rst!proximaseccio <> "E" And rst!proximaseccio <> "V" And rst!proximaseccio <> "P" Then
         vrisc.valorproduccio = vrisc.valorproduccio + vvalorcomanda
         'vrisc.comandesproduccio = vrisc.comandesproduccio + IIf(vrisc.comandesproduccio = "", "", ",") + vdesccomanda
         If cadbl(rst!pvp) = 0 Then
            vrisc.comandesproduccio = vdesccomanda + vrisc.comandesproduccio
             Else: vrisc.comandesproduccio = vrisc.comandesproduccio + vdesccomanda
         End If
     End If
cont:
     rst.MoveNext
  Wend
  vrisc.comandespendent = "[ " + vrisc.comandespendent
  vrisc.comandesproduccio = "[ " + vrisc.comandesproduccio
  vrisc.comandesestoc = "[ " + vrisc.comandesestoc
  Set rst = Nothing
  
  'albarans pujats a sap sense facturar
  vsql = "SELECT capcaleraalbara.codiclient, capcaleraalbara.numalbaraSAP, capcaleraalbara.numfacturaSAP, liniesalbara.lotinplacsa, [quantitat]*[preuvenda] AS total, liniesalbara.preuvenda, liniesalbara.kgimpostenvasos,liniesalbara.eurokg_impost, comandes.rebkilos, comandes.cantitatsol "
  vsql = vsql + " FROM (capcaleraalbara RIGHT JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara) LEFT JOIN comandes ON liniesalbara.lotinplacsa = comandes.comanda"
  vsql = vsql + " WHERE (((capcaleraalbara.codiclient)=" + atrim(vcodiclient) + ") AND ((capcaleraalbara.numalbaraSAP) Is Not Null And (capcaleraalbara.numalbaraSAP)<>0) AND ((capcaleraalbara.numfacturaSAP) Is Null or (capcaleraalbara.numfacturaSAP)=0) AND ((liniesalbara.preuvenda)>0));"

 ' Clipboard.Clear
 ' Clipboard.SetText vsql
  Set rst = dbtmp.OpenRecordset(vsql)
  While Not rst.EOF
    vdesccomanda = ""
    If cadbl(rst!rebkilos) > 0 Then
        vdesccomanda = atrim(rst!rebkilos) + "Kg "
         Else: If cadbl(rst!cantitatsol) > 0 Then vdesccomanda = atrim(rst!cantitatsol) + "Un "
    End If
'    MsgBox rst!lotinplacsa
    vrisc.albaransSAP = vrisc.albaransSAP + atrim(rst!lotinplacsa) + "-" + vdesccomanda + "(" + Format(Redondejar(rst!total, 0), "#,##0") + "€) "
    vrisc.valoralbaranspendentsSAP = vrisc.valoralbaranspendentsSAP + cadbl(Redondejar(rst!total + (cadbl(rst!kgimpostenvasos) * cadbl(rst!eurokg_impost)), 0))
    rst.MoveNext
  Wend
  Set rst = Nothing
  
  'clixes
  While Not rstclixes.EOF
     If cadbl(rstclixes!codiclientfactclixes) = vcodiclient Then
       vrisc.valordelsclixes = vrisc.valordelsclixes + cadbl(rstclixes!preu)
       vrisc.treballspendents = vrisc.treballspendents + " Tr:" + atrim(rstclixes!id_treball) + "/" + atrim(rstclixes!ordre) + " (" + Format(rstclixes!preu, "#,##0") + "€) "
     End If
     rstclixes.MoveNext
  Wend
   vrisc.valordelsclixes = vrisc.valordelsclixes * (1 + (viva / 100))
   vrisc.valoralbaranspendentsSAP = vrisc.valoralbaranspendentsSAP * (1 + (viva / 100))
   vrisc.valorestoc = vrisc.valorestoc * (1 + (viva / 100))
   vrisc.valorproduccio = vrisc.valorproduccio * (1 + (viva / 100))
   vrisc.valorpendent = vrisc.valorpendent * (1 + (viva / 100))
   
  vrisc.valordiferencial = vrisc.creditsap - vrisc.creditgastatsap - vrisc.valordelsclixes - vrisc.valorestoc - vrisc.valorpendent - vrisc.valorproduccio - vrisc.valoralbaranspendentsSAP
  Set rstvendes = Nothing
  Set rstclixes = Nothing
  Set rst = Nothing
End Sub
Sub informe_credit_unclient(vcodicomptable As String)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vrisc As TipusVrisc
  calcular_credit_delclient cadbl(vcodicomptable), vrisc
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "fullcontrolderisc.rpt", 1)
  'oreport.Database.Tables.Item(1).Location = ""
  oreport.FormulaFields.GetItemByName("comandesestoc").Text = "'" + vrisc.comandesestoc + "'"
  oreport.FormulaFields.GetItemByName("comandespendent").Text = "'" + atrim(vrisc.comandespendent) + "'"
  oreport.FormulaFields.GetItemByName("comandesproduccio").Text = "'" + atrim(vrisc.comandesproduccio) + "'"
  oreport.FormulaFields.GetItemByName("treballspendents").Text = "'" + atrim(vrisc.treballspendents) + "'"
  oreport.FormulaFields.GetItemByName("nomdelclient").Text = "'" + vcodicomptable + " - " + treure_apostruf(vrisc.nomdelclient) + "'"
  oreport.FormulaFields.GetItemByName("riscsap").Text = Redondejar(vrisc.creditsap, 0)
  oreport.FormulaFields.GetItemByName("riscsapconsumit").Text = Redondejar(vrisc.creditgastatsap, 0)
  oreport.FormulaFields.GetItemByName("valorestoc").Text = Redondejar(vrisc.valorestoc, 0)
  oreport.FormulaFields.GetItemByName("valorpendent").Text = Redondejar(vrisc.valorpendent, 0)
  oreport.FormulaFields.GetItemByName("valorproduccio").Text = Redondejar(vrisc.valorproduccio, 0)
  oreport.FormulaFields.GetItemByName("valordelsclixes").Text = Redondejar(vrisc.valordelsclixes, 0)
  oreport.FormulaFields.GetItemByName("valordiferencial").Text = Redondejar(vrisc.valordiferencial, 0)
  oreport.FormulaFields.GetItemByName("valoralbaransSAP").Text = Redondejar(vrisc.valoralbaranspendentsSAP, 0)
  oreport.FormulaFields.GetItemByName("comandesSAPpendentsfacturar").Text = "'" + atrim(vrisc.albaransSAP) + "'"
   
  Load veurereport
  veurereport.CRViewer.ReportSource = oreport
  veurereport.CRViewer.DisplayGroupTree = False
  veurereport.CRViewer.ViewReport
  veurereport.WindowState = 2
  veurereport.Show 1
End Sub
Function isloaded(vnomform As String) As Boolean
  Dim f
  For Each f In Forms
   If UCase(f.Name) = UCase(vnomform) Then
         isloaded = True
   End If
  Next
End Function
Function substituir(cadena As String, buscar As String, canviar As String, Optional vcanvis As Long) As String
   If atrim(buscar) = atrim(canviar) Then GoTo fi
   cadena = "  " + cadena
   While InStr(1, cadena, buscar) > 0
    comença = InStr(1, cadena, buscar) - 1
    If comença < 1 Then substituir = cadena: Exit Function
    acaba = comença + Len(buscar) + 1
    cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
    vcanvis = vcanvis + 1
   Wend
fi:
   substituir = atrim(cadena)
   'MsgBox linia
End Function

