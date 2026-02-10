Attribute VB_Name = "Module1"
Global hwd As Long
Global inicidragover As Date
Global ultimtinter As Byte
Global gravant As Boolean
Global id_treball As Long
Global ordremodificacio As Long
Global numcomanda As String
Global dbconsulta As Database
Global rstconsulta As Recordset
Global rstopcionset As Recordset
Global nummanteniment As Long
Global taula_tmp As String
Global Idioma As String
Global fitxerini As String
Global llistadecampsvalids As String
Global dbenvios As Database
Global rsttmp As Recordset
Global campcontrol As Control
Global dbbaixes As Database
Global dbcomandes As Database
Global dbmanteniments As Database
Global dbclixesvells As Database
Global bdllistat As Database
Global dbstocks As Database
Global rststocks As Recordset
Global idfuncionament As Long
Global tipus As String
Global queryorder As String
Global querywhere As String
Global cami As String
Global camiclixes As String
Global buscant As Boolean
Global seleccioret As Byte
Global ultimcontrol As Control
Global llocform As Byte
Global taulapos(15) As Variant
Global ruta As String
Global ruta_documentacio_clixes As String
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

Global Const topeform = 5
Global Const colorfonscontrolactiu = 10


'Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal _
'       lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
'       lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long'


Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
                
Private Declare Function SendMessage Lib "user32" _
                Alias "SendMessageA" _
                (ByVal hWnd As Long, _
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
                (ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long
                
Private Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
                (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function SetTimer Lib "user32" _
                (ByVal hWnd As Long, _
                ByVal nIDEvent As Long, _
                ByVal uElapse As Long, _
                ByVal lpTimerFunc As Long) As Long
                
Private Declare Function KillTimer Lib "user32" _
                (ByVal hWnd As Long, _
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
 Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


' Estructura SHFILEOPSTRUCT o para usar con el Api
Private Type SHFILEOPSTRUCT
    hWnd As Long
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
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Constantes
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40



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

Public Sub Copiar_Fitxer(ByVal origen As String, ByVal Destino As String, Optional opcions As Long)

Dim t_Op As SHFILEOPSTRUCT
  If IsNull(opcions) Then opcions = FOF_ALLOWUNDO
    With t_Op
        .hWnd = 0
        .wFunc = FO_COPY
        .pFrom = origen & vbNullChar & vbNullChar
        .pTo = Destino & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With

    ' Se ejecuta la función Api pasandole la estructura
    SHFileOperation t_Op
    
    
End Sub

Sub obrir_document(nomfitxer As String)
 Dim vaaa As Integer
 'If existeix(nomfitxer) Then MsgBox "No trobo el fitxer:" + Chr(10) + Chr(13) + nomfitxer
 vaaa = ShellExecute(Screen.ActiveForm.hWnd, "Open", nomfitxer, "", "", 1)
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
  If estatval = 0 Then textestat = ""
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


Function passaradecimal(valormasked As String) As String
   Dim valmas As String
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
  Else:
     If seccio = "V" Then seccio = "P": ruta = ruta + "P"
     If InStr(1, rutav, pr) <= InStr(1, rutav, seccio) Then
       dbtmp.Execute "update comandes set proximaseccio='" + seccio + "' where comanda=" + atrim(comanda)
     End If
End If

End Sub
 
 Sub controlar_fiseccio(seccio As String, ruta As String, Optional hihabobines As Boolean)
  If canvissortirseccio Then
    If hihabobines Then
      modificar_estat_comanda entradabaixes.comanda, ruta + "T", seccio, 1
     Else: modificar_estat_comanda entradabaixes.comanda, ruta + "T", seccio, 0
     End If
  End If
  If Not canvissortirseccio Then End
  entradabaixes.Visible = True
 End Sub
Sub obrir_baixes()
  Shell llegir_ini("General", "rutaprogbaixes", fitxerini), vbNormalFocus
End Sub
Function carpeta_del_client(codiclient As Double) As String
   Dim rstc As Recordset
    
    Set rstc = dbcomandes.OpenRecordset("select * from carpeta_client where codiclient=" + atrim(codiclient))
    If Not rstc.EOF Then
        carpeta_del_client = rstc!nomcarpeta
    End If
    
End Function
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
                     ByVal hWnd As Long, _
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

Function rutadelfitxer(cam As String) As String
   Dim c As Byte
   c = 0
   While InStr(c + 1, cam, "\") <> 0
    c = InStr(c + 1, cam, "\")
   Wend
   If c = 0 Then c = Len(cam)
   rutadelfitxer = Mid(cam, 1, c)
End Function

Function ObtenerLíneaComando(Optional MaxArgs)
    'Declara las variables.
    Dim c, LíneaComando, LonLínComando, ArgIn, i, NúmArgs
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
        c = Mid(LíneaComando, i, 1)
        'Comprueba espacio o tabulación.
        If (c <> " " And c <> vbTab) Then
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

ArgArray(NúmArgs) = ArgArray(NúmArgs) + c
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


