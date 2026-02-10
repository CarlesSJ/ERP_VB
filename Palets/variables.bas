Attribute VB_Name = "Module1"
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Type POINTAPI
   X_Pos As Long
   Y_Pos As Long
End Type
Global vfiltrebobinesdesdeimpresores As Boolean
Global arguments As Variant
Global vexportantelllistat As Boolean
Global selecciofam As String
Global nomfiltrefam As String
Global selecciomicres As Double
Global filtrarprestatge As String
Global nomfitxertemporal As String
Global metresareservar As Double
Global numcomanda As String
Global ultimcolor As Long
Global dbconsulta As Database
Global dbcompres As Database
Global rstcompres As Recordset
Global rstconsulta As Recordset
Global taula_tmp As String
Global fitxerini As String
Global rsttmp As Recordset
Global dbtmp As Database
Global dbtmpb As Database
Global dbtemp As Database
Global dbllistat As Database
Global dbstocks As Database
Global dbbaixes As Database
Global dbcomandes As Database
Global dbtintes As Database
Global rstllistat As Recordset
Global criteridebusqueda As String
Global tipus As String

Global cami As String
Global camistock As String
Global buscant As Boolean
Global seleccioret As Byte
Global ruta As String
Global ruta_relativa_docs As String
Global tecla As Integer
Global i As Integer
Global r As String

Global sa As String

Global espessor As String
Global controlcanviat As Control
Global colorcanviat As String
Global canvissortirseccio As Boolean
Global iniconfigreixa As String

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

 Sub Copiar_Fitxer(ByVal Origen As String, ByVal Destino As String, Optional opcions As Long)

Dim t_Op As SHFILEOPSTRUCT
  If IsNull(opcions) Then
     opcions = FOF_ALLOWUNDO
      Else: opcions = opcions + FOF_ALLOWUNDO
  End If
    With t_Op
        .hwnd = 0
        .wFunc = FO_COPY
        .pFrom = Origen & vbNullChar & vbNullChar
        .pTo = Destino & vbNullChar & vbNullChar
        .fFlags = opcions
    End With

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
  If Not IsNumeric(valo) Or atrim(valo) = "" Then valo = 0
  cadbl = CDbl((atrim(valo)))
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
Function simboldecimal() As String
   vsimboldecimal = "."
   vsimbolmiler = ","
   If InStr(1, Trim(CDbl(1 / 2)), ",") Then vsimbolmiler = ".": vsimboldecimal = ","
   simboldecimal = vsimboldecimal
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
   Dim vs As String
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
Function carpeta_del_client() As String
 
    formcomandes.data1tmp.DatabaseName = cami
    formcomandes.data1tmp.RecordSource = "select * from carpeta_client where codiclient=" + atrim(cadbl((formcomandes.Data1.Recordset!client)))
    formcomandes.data1tmp.Refresh
    If Not formcomandes.data1tmp.Recordset.EOF Then
        carpeta_del_client = formcomandes.data1tmp.Recordset!nomcarpeta
       Else: MsgBox "No s'ha trobat la carpeta del client PROVA DE REINDEXAR AL MENU UTIL REINDEXAR CLIENTS."
    End If
    
End Function
Sub wait(segonsespera As Byte)
  horaentradawait = Now
  While DateDiff("s", horaentradawait, Now) < segonsespera
    DoEvents
  Wend
End Sub


 Function InputBoxEx( _
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
Sub imprimirllistatreferencies()
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "referenciesenestoc.rpt", 1)
  oreport.Database.Tables.Item(1).Location = camistock
  
  oreport.DiscardSavedData
  'If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1
   ' Else
   '   oreport.PrintOut False, 1
 ' End If
  
  End Sub
Sub generarfitxertemporal()
  If Not existeix("c:\temp\llistatpalets") Then MkDir "c:\temp\llistatpalets"
  nomfitxertemporal = "c:\temp\llistatpalets\temporal_" + Trim(format(Now, "hhnnss")) + ".mdb"
  On Error Resume Next
  Kill "c:\temp\llistatpalets\temporal_*.*"
End Sub
'---------------------------------------------------------
Sub Main()

Dim rst As Recordset
Dim llistat As CrystalReport
Dim numc As Double
arguments = ObtenerLíneaComando
'If llegir_ini("baixes", "imprimirpackinglist", "comandes.ini") = "1" Then escriure_ini "baixes", "imprimirpackinglist", "2", "comandes.ini"
'    wait 1
'End If
If App.PrevInstance And arguments(2) = "" Then MsgBox "El programa ja està obert.", vbCritical, "Atenció": End

fitxerini = "comandes.ini"
If atrim(arguments(1)) <> "" Then fitxerini = atrim(arguments(1))
'On Error Resume Next
generarfitxertemporal
DBEngine.CreateDatabase nomfitxertemporal, dbLangGeneral, DatabaseTypeEnum.dbVersion30

'On Error GoTo 0

  If llegir_ini("General", "parar", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini") = "si" Then MsgBox "Ara no es pot entrar al programa s'està actualitzant, espera 5 MINUTS, Gràcies", vbCritical, "Actualització": End
  If Not existeix("c:\ordprog.ini") Or cadbl(arguments(2)) <> 0 Then assignardecimalipunt
  cami = llegir_ini("General", "cami", fitxerini)
  
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  
  '"c:\misdoc~1\commandes\comandes.mdb"
  If existeix("c:\ordprog.ini") Then cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
  hora = Now
  'centerscreen Me
  camistock = rutadelfitxer(cami) + "palets.mdb"
  
  

  obrir_dbllistats
  crear_taules_tmp
  Set rst = dbtmp.OpenRecordset("select * from parcials where comanda is null")
  If Not rst.EOF Then
   dbtmp.Execute "update parcials set comanda='0' where comanda is null"
  End If
  If arguments(2) = "FiltrarBobinesImpresores" Then
     vfiltrebobinesdesdeimpresores = True
     DoEvents
     obrir_dbllistats
     crear_taules_tmp
     assignarmat.Show 1
  End If
  
  If arguments(2) = "temporalTORERUS" Then
       generarfitxertemporalpelstorerus
  End If
  If arguments(2) = "comprant" Then wait 2
  Set rst = dbtmp.OpenRecordset("select * from pendentsdereservar where not reservar and not entrat")
  If Not rst.EOF And arguments(2) = "comprant" Then
     dbtmp.Execute "update pendentsdereservar set entrat=true where not entrat"
     DoEvents
     Form1.Hide
     DoEvents
     assignarmat.Show
     DoEvents
     numc = rst!comanda
     
     assignarmat.carregarperreservar numc: Exit Sub
  End If
  Set rst = Nothing
  If arguments(2) = "llistatreferencies" Then imprimirllistatreferencies: End
  If arguments(2) = "guardarllistatestoc" Then Form1.Visible = False:   vexportantelllistat = True
  
  If arguments(2) = "assignamaterials" Then
   'assignamaterial_Click
   Form1.Hide
   assignarmat.Show
   DoEvents
   
   DoEvents
   
    Else:
      If cadbl(arguments(2)) <> 0 Then
          Load Form1
          imprimir_packinglist cadbl(arguments(2)), Form1.llistat, False
          'If cadbl(arguments(2)) < 0 Then
          '  wait (2)
          '  MsgBox "Prem Acceptar per tancar.", vbInformation, "PackingList"
          'End If
          If cadbl(arguments(2)) > 0 Then
             Unload Form1
             End
            Else: esperarqueacavilinforme
          End If
         Else: Form1.Show
      End If
  End If
End Sub
Sub importacio_mdb_tablet()
  Dim dbtablet As Database
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim rstcanvislocals As Recordset
  Dim vusuariTORERU As Byte
  Dim i As Byte
  Dim vQcanvis As Double
  
  If Not existeix(rutadelfitxer(cami) + "Torerus_Tablet.mdb") Then Exit Sub
  
  'faig una copia per si de cas així puc revisar si hi ha un error
  If existeix(rutadelfitxer(cami) + "Torerus_Tablet_ultim.mdb") Then Kill rutadelfitxer(cami) + "Torerus_Tablet_ultim.mdb"
  FileCopy rutadelfitxer(cami) + "Torerus_Tablet.mdb", rutadelfitxer(cami) + "Torerus_Tablet_ultim.mdb"
  vusuariTORERU = cadbl(llegir_ini("Torerus", "usuariTORERUS", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"))
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
  Set dbtablet = OpenDatabase(rutadelfitxer(cami) + "Torerus_Tablet.mdb")
  Set rst = dbtablet.OpenRecordset("select * from canvissituacio order by data asc")
  Set rstcanvislocals = dbtmp.OpenRecordset("select * from canvissituacio")
  vQcanvis = 0
  If Not rst.EOF Then rst.MoveLast: rst.MoveFirst: vQcanvis = rst.RecordCount
  dbstocks.Execute "insert into torerus_sincronitzacions (usuari,canvis_bobines) VALUES (" + atrim(vusuariTORERU) + "," + atrim(vQcanvis) + ")"
  While Not rst.EOF
    'asseguro que aquest canvi no estigui ja fet anteriorment amb data i bobina per si hi hagues algun error
    If rst!canvifet Then GoTo proxim
    rstcanvislocals.FindFirst "data=#" + atrim(format(rst!data, "mm/dd/yy hh:nn:ss")) + "#"
    If Not rstcanvislocals.NoMatch Then
        If rst!bobina = rstcanvislocals!bobina Then GoTo proxim
    End If
    'modifico la ubicacio de la bobina
    dbtmp.Execute "update bobines set sit='" + atrim(rst!sitdesti) + "' where trim(idpalet)+'/'+trim(idbobina)='" + atrim(rst!bobina) + "'"
    'copio el registre a la taula de canvis
    rstcanvislocals.AddNew
    For i = 1 To rst.Fields.Count - 1
        rstcanvislocals.Fields(i) = rst.Fields(i)
    Next i
    'marco els canvi fet per no tornar a fer-lo a les dues taules
    rstcanvislocals!canvifet = True
    rstcanvislocals.Update
    rst.Edit
    rst!canvifet = True
    rst.Update
    
proxim:
    rst.MoveNext
  Wend
 'passo els canvis de la revisió de la taula temporal dels torerus a la REAL
   'només es el camp revisattoreru SI NO ESTÀ ENVIAT PER NO FER CANVIS INVOLUNTARIS
  Set rst = dbtablet.OpenRecordset("select distinct comanda,numalbara,numpalet,revisatTORERU from bobinesent where modificat=true")
  While Not rst.EOF
    Set rst2 = dbstocks.OpenRecordset("select * from linies_expedicions where comanda=" + atrim(rst!comanda) + " and albara=" + atrim(rst!numalbara))
    If Not rst2.EOF Then
        If rst2!enviat = False Then
             dbtmpb.Execute "update bobinesent set revisatTORERU='" + atrim(rst!revisatTORERU) + "',usuariTORERU=" + Trim(vusuariTORERU) + " where comanda=" + atrim(rst!comanda) + " and numpalet=" + atrim(cadbl(rst!numpalet))
        End If
    End If
    rst.MoveNext
  Wend
'posso les modificacions de diametres de bobina
  If existeixlataula(rutadelfitxer(cami) + "Torerus_Tablet.mdb", "comprovacio_diametres_picus") Then
    Set rst = dbtablet.OpenRecordset("select * from comprovacio_diametres_picus where not actualitzat")
    While Not rst.EOF
     dbstocks.Execute "insert into comprovacio_diametres_picus (numpalet,bobina,data, diametre,diametreanterior,actualitzat,metresnous) values (" + atrim(atrim(rst!numpalet)) + "," + atrim(rst!bobina) + ",#" + atrim(rst!data) + "#," + passaradecimalpunt(atrim(rst!diametre)) + "," + passaradecimalpunt(atrim(rst!diametreanterior)) + ",false," + passaradecimalpunt(atrim(IIf(IsNull(rst!metresnous), 0, rst!metresnous))) + ")"
     rst.MoveNext
    Wend
    wait 3
    ajustar_picos_alesbobines
  End If
fi:
  Set dbtablet = Nothing
  Set rst = Nothing
  Set rst2 = Nothing
  wait 2
  Kill rutadelfitxer(cami) + "Torerus_Tablet.mdb"
End Sub
Sub ajustar_picos_alesbobines()
   Dim rst As Recordset
   Set dbstocks = dbtmp
   Set rst = dbstocks.OpenRecordset("select * from comprovacio_diametres_picus where not actualitzat")
   While Not rst.EOF
     If actualitzar_metresxrdiametre(rst) Then rst.Edit: rst!actualitzat = True: rst.Update
     rst.MoveNext
   Wend
End Sub
Function actualitzar_metresxrdiametre(rstpicus As Recordset) As Boolean
    Dim rst As Recordset
    Dim vmetresbob As Double
    Dim vValues As String
    Dim vmetresactualitzar As Double
    dbstocks.Execute "delete * from parcials where comanda='444' and idpalet=" + atrim(rstpicus!numpalet) + " and idbobina=" + atrim(rstpicus!bobina)
    vmetresbob = bobinesdentrada.calcular_mtrsdispreals(rstpicus!numpalet, rstpicus!bobina)
    vmetresactualitzar = vmetresbob - rstpicus!metresnous
    vValues = "(" + atrim(rstpicus!numpalet) + "," + atrim(rstpicus!bobina) + ",True,'444',now,0,'T','Actualització diametre Torerus.')"
    dbstocks.Execute "insert into parcials (idpalet,idbobina,utilitzada,comanda,data,operari,seccio,observacions) values " + vValues
    Set rst = dbstocks.OpenRecordset("select * from parcials where comanda='444' and idpalet=" + atrim(rstpicus!numpalet) + " and idbobina=" + atrim(rstpicus!bobina))
    If Not rst.EOF Then
       rst.Edit: rst!metres = cadbl(rst!metres) + vmetresactualitzar: rst.Update
       bobinesdentrada.actualitzar_metres_disponibles rst!idpalet, rst!idbobina
       actualitzar_metresxrdiametre = True
    End If
    Set rst = Nothing
End Function

Function eliminarfitxertemporal(vnom As String) As Boolean
  On Error GoTo fi
  eliminarfitxertemporal = True
  Kill vnom
  Exit Function
fi:
   eliminarfitxertemporal = False
End Function
Sub generarfitxertemporalpelstorerus()
  Dim vnomtmp As String
  Dim vnomtmpfinal As String
  Dim dbvnomtmp As Database
  Dim vsql As String
gravarLOG "TORERUS", "*******************************"
gravarLOG "TORERUS", "INICI DEL PROCES D'ACTUALITZACIO DE LES TABLETS"
  ratoli "espera"
  escriure_ini "Torerus", "ultimresultat", "", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
gravarLOG "TORERUS", "Inici importaciótablet"
  importacio_mdb_tablet
gravarLOG "TORERUS", "Fi importaciótablet"
  
  vnomtmp = "c:\temp\temporal_TORERUS.mdb"
  'vnomtmpfinal = "c:\temp\TORERUS.mdb"
  vnomtmpfinal = rutadelfitxer(cami) + "TORERUS.mdb"
    'BORRO ELS TEMPORALS
gravarLOG "TORERUS", "Inici eliminartemporal"
  If existeix(vnomtmpfinal) Then
     If Not eliminarfitxertemporal(vnomtmpfinal) Then
         escriure_ini "Torerus", "ultimresultat", "Error eliminant fitxer temporal.", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
         gravarLOG "TORERUS", "ERROR ELIMINANT FITXER TEMPORAL" + vnomtmpfinal
         GoTo fi
     End If
  End If
gravarLOG "TORERUS", "Inici eliminartemporal2"
  If existeix(vnomtmp) Then Kill vnomtmp
gravarLOG "TORERUS", "Fi eliminartemporal"
gravarLOG "TORERUS", "Crear Mdb temporal"
  DBEngine.CreateDatabase vnomtmp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
    '------Impresora PUJAR ------
gravarLOG "TORERUS", "Inici PujarIMP"
  Form1.generartaulallistatperpujarIMP
  dbllistat.Execute "select * into llistatperpujarIMP IN '" + vnomtmp + "' from llistatperpujar"
gravarLOG "TORERUS", "Fi pujarIMP"
    '------Impresora BAIXAR ------
gravarLOG "TORERUS", "Inici BaixarIMP"
  Form1.generartaulallistatperbaixarIMP
  dbllistat.Execute "select * into llistatperbaixarIMP IN '" + vnomtmp + "' from llistatperpujar"
gravarLOG "TORERUS", "Fi BaixarIMP"
    '------Laminadora BAIXAR --------------
gravarLOG "TORERUS", "Inici BaixarLAM"
  formmourebobines.carregar_bobinespermoure "L"
  dbllistat.Execute "select * into llistatperbaixarLAM IN '" + vnomtmp + "' from llistatperpujar where (Mid(UCase([Sit]),1,3))<>'LAM'"
gravarLOG "TORERUS", "Fi BaixarLAM"
  '------Laminadora PUJAR ------
gravarLOG "TORERUS", "Inici PujarLAM"
  Form1.generartaulallistatperpujarLAM
  dbllistat.Execute "select * into llistatperpujarLAM IN '" + vnomtmp + "' from llistatperpujar"
gravarLOG "TORERUS", "Fi BaixarLAM"
     '-------passo les bobines dels grups
gravarLOG "TORERUS", "Inici passarGrups"
  Set dbvnomtmp = OpenDatabase(vnomtmp)
  dbtmp.Execute "select TOP 1 * into bobinesgrups IN '" + vnomtmp + "' from parcials"
  dbvnomtmp.Execute "delete * from bobinesgrups"
  dbtmp.Execute "insert into bobinesgrups IN '" + vnomtmp + "' select * from parcials where metres>0 and (cdbl(comanda)>2000 and cdbl(comanda)<3000)"
  dbtmp.Execute "select * into nomgrups IN '" + vnomtmp + "' from grupsdepalets"
  dbtmp.Execute "select * into productes IN '" + vnomtmp + "' from productes"
  passar_seccioSiR_comagrups vnomtmp, dbvnomtmp
gravarLOG "TORERUS", "Fi passarGrups"
gravarLOG "TORERUS", "Inici Passar operaris i contrasenyes"
  'passo els operaris
  dbcomandes.Execute "select * into operaris IN'" + vnomtmp + "' from operaris"
  dbcomandes.Execute "select * into operaris_contrasenyes IN'" + vnomtmp + "' from operaris_contrasenyes"
gravarLOG "TORERUS", "Fi Passar operaris i contrasenyes"
gravarLOG "TORERUS", "Passo parcials,canvissituacio,foratsnous,prestatgesnous i materials"
  'passo les situacions de les bobines concatenat amb mida,micres i nom del material
  '''''dbstocks.Execute "select * into bobines IN'" + vnomtmp + "' from bobines"
  vsql = " (Bobines LEFT JOIN Palets ON Bobines.Idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi "
  dbstocks.Execute "SELECT Bobines.*, Palets.Ample, Palets.micres, Palets.grmsm2, materials.descripcio into bobines IN'" + vnomtmp + "' from " + vsql
  dbstocks.Execute "select * into parcials IN'" + vnomtmp + "' from parcials"
  dbstocks.Execute "select * into CanvisSituacio IN'" + vnomtmp + "' from CanvisSituacio where 1=2"
  dbstocks.Execute "select * into foratsnous IN'" + vnomtmp + "' from foratsnous"
  dbstocks.Execute "select * into PrestatgesNous IN'" + vnomtmp + "' from PrestatgesNous"
  dbstocks.Execute "select * into Materials IN'" + vnomtmp + "' from materials"
  dbstocks.Execute "select * into palets IN'" + vnomtmp + "' from palets"
  dbstocks.Execute "select * into comprovacio_diametres_picus IN'" + vnomtmp + "' from comprovacio_diametres_picus"
'PASSO LES BOBINES D'ENTREGA PER PODER FER EL CONTROL D'ENTREGA DE MATERIAL DESDE LA TABLE
  dbcomandes.Execute "select * into bobinesent IN'" + vnomtmp + "' from bobinesent where entregat<>'S' or entregat=null and (numalbara<>null and numalbara>0)"
'PASSO LES LINIES D'ENTREGUES PREVISTES
  dbstocks.Execute "SELECT * into linies_expedicions IN '" + vnomtmp + "' From linies_expedicions WHERE (((linies_expedicions.enviat)=False) AND ((linies_expedicions.albara)>0));"
 
  gravarLOG "TORERUS", "FI DE - Passo parcials,canvissituacio,foratsnous,prestatgesnous i materials"
  gravarLOG "TORERUS", "Copio el Temporal com a definitiu."
   '------ COPIAR FITXER TEMPORAL A DEFINITIU
   Set dbvnomtmp = Nothing
  FileCopy vnomtmp, vnomtmpfinal
    'BORRO EL TEMPORAl
gravarLOG "TORERUS", "Elimino el temporal."
gravarLOG "TORERUS", "FI DEL PROCES D'ACTUALITZACIO DE LES TABLETS"
  If existeix(vnomtmp) Then Kill vnomtmp
  escriure_ini "Torerus", "ultimresultat", "OK", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
fi:
  ratoli "normal"
  escriure_ini "Torerus", "horaultimaactualitzacio", "", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
gravarLOG "TORERUS", "*************************************************"
Set dbllistat = Nothing
Set dbvnomtmp = Nothing
Set dbtmp = Nothing
Set dbstocks = Nothing
  End
End Sub
Sub passar_seccioSiR_comagrups(vnomtmp As String, dbvnomtmp As Database)
   Dim vsubsql As String
   'Rebobinadora
   vsubsql = "SELECT id FROM productes RIGHT JOIN (Parcials_DBL LEFT JOIN comandes ON Parcials_DBL.comandaDBL = comandes.comanda) ON productes.codi = comandes.producte WHERE (((comandes.proximaseccio)<>'T' And (comandes.proximaseccio)<>'V') AND ((Parcials_DBL.utilitzada)=False) AND ((Mid([ruta]+' ',2,1))='R'));"
   dbtmp.Execute "INSERT into bobinesgrups IN '" + vnomtmp + "' SELECT * from parcials where metres>0 and parcials.id in (" + vsubsql + ")"
   dbtmp.Execute "update bobinesgrups in '" + vnomtmp + "' set comanda='5' where id in (" + vsubsql + ")"
   'Soldadora
   vsubsql = "SELECT id FROM productes RIGHT JOIN (Parcials_DBL LEFT JOIN comandes ON Parcials_DBL.comandaDBL = comandes.comanda) ON productes.codi = comandes.producte WHERE (((comandes.proximaseccio)<>'T' And (comandes.proximaseccio)<>'V') AND ((Parcials_DBL.utilitzada)=False) AND ((Mid([ruta]+' ',2,1))='S'));"
   dbtmp.Execute "INSERT into bobinesgrups IN '" + vnomtmp + "' SELECT * from parcials where metres>0 and parcials.id in (" + vsubsql + ")"
   dbtmp.Execute "update bobinesgrups in '" + vnomtmp + "' set comanda='6' where id in (" + vsubsql + ")"
   dbvnomtmp.Execute "insert into nomgrups (numerogrup,nomdelgrup,seccio) values (5,'REB','R')"
   dbvnomtmp.Execute "insert into nomgrups (numerogrup,nomdelgrup,seccio) values (6,'SOL','S')"
   
End Sub
Sub gravarLOG(vetiqueta As String, vdescripcio As String)
   Dim vfitxer As String
   Exit Sub  'desactivat perque ja no utilitzo el log SI CAL ES TORNA A ACTIVAR
   vfitxer = rutadelfitxer(cami) + "Log_Torerus.txt"
   If Not existeix(vfitxer) Then
      Open vfitxer For Output As 9
       Else: Open vfitxer For Append As 9
   End If
   Print #9, format(Now, "dd/mm/yy hh:nn:ss") + " " + justificar(atrim(vetiqueta), 10, "E") + " -" + atrim(vdescripcio)
   Close 9
End Sub
Function justificar(v As String, longitut As Integer, DoE As String) As String
    v = Mid(v, 1, longitut)
    If DoE = "E" Then
       v = v + Space(longitut - Len(v))
      Else: v = Space(longitut - Len(v)) + v
    End If
    justificar = v
End Function
Sub esperarqueacavilinforme()
   While cadbl(llegir_ini("baixes", "imprimirpackinglist", "comandes.ini")) > 0
           DoEvents
   Wend
   End
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
 Sub obrir_dbllistats()
  Dim taulatemp As String
  taulatemp = nomfitxertemporal
  If Not existeix(taulatemp) Then
     DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  End If
  Set dbllistat = DBEngine.OpenDatabase(taulatemp)
  Set dbtmp = OpenDatabase(camistock)
  Set dbtmpb = OpenDatabase(cami)

End Sub
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

 Sub crear_taules_tmp()
  Dim camps(100, 2) As String
  taula_tmp = "Packinglistxcomanda"
  taula_tmp2 = "assignaciomaterial"
  taula_tmp3 = "etiquetapalet"
  taula_tmp4 = "reservamaterial"
  On Error Resume Next
   dbllistat.Execute "drop table " + taula_tmp
   dbllistat.Execute "drop table " + taula_tmp2
   dbllistat.Execute "drop table " + taula_tmp3
   dbllistat.Execute "drop table " + taula_tmp4
   dbllistat.Execute "delete * from " + taula_tmp
   dbllistat.Execute "delete * from " + taula_tmp2
   dbllistat.Execute "delete * from " + taula_tmp3
   dbllistat.Execute "drop table " + taula_tmp4
  i = 1
  camps(i, 1) = "Palet": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "bobina": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "material": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "nommaterial": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "familia": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "proveidor": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "reserva": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "numlot": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "datarec": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "tractat": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "situacio": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "observacionsP": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "observacionsB": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "numpaletprov": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "ample": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "plegat": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "solapa": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "metres": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "obert": camps(i, 2) = "string(1)": i = i + 1
  camps(i, 1) = "microperforat": camps(i, 2) = "bit": i = i + 1
  camps(i, 1) = "semielaborat": camps(i, 2) = "string(1)": i = i + 1
  camps(i, 1) = "carestractat": camps(i, 2) = "string(1)": i = i + 1
  camps(i, 1) = "micres": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "kilos": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "kilosprov": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "tipusbobina": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "estoc": camps(i, 2) = "longbinary": i = i + 1
  camps(i, 1) = "impostenvasos": camps(i, 2) = "bit": i = i + 1
  camps(i, 1) = "orcomassignacio": camps(i, 2) = "double": i = i + 1

  
    
  dbllistat.Execute ("create table " + taula_tmp + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbllistat.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
'  Set rstllistat = dbllistat.OpenRecordset(taula_tmp)
  
   i = 1
   camps(i, 1) = "seleccionat": camps(i, 2) = "bit": i = i + 1
  camps(i, 1) = "Palet": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "bobina": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "resto": camps(i, 2) = "bit": i = i + 1
  camps(i, 1) = "datarec": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "families": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "material": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "proveidor": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "ample": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "micres": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "mtrsassignats": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "mtrsdiferencia": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "mtrsdisponibles": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "metres": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "kilos": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "observacionsB": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "familia": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "reserva": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "comanda": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "numlot": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "tractat": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "situacio": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "observacionsP": camps(i, 2) = "string": i = i + 1
  
  camps(i, 1) = "numpaletprov": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "plegat": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "solapa": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "codimat": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "parcial": camps(i, 2) = "bit": i = i + 1
  camps(i, 1) = "impostenvasos": camps(i, 2) = "bit": i = i + 1
  
  
  
  dbllistat.Execute ("create table " + taula_tmp2 + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbllistat.Execute ("alter table " + taula_tmp2 + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  
  
   i = 1
   camps(i, 1) = "seleccionat": camps(i, 2) = "bit": i = i + 1
  camps(i, 1) = "Palet": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "bobina": camps(i, 2) = "integer": i = i + 1
  camps(i, 1) = "material": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "familia": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "proveidor": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "reserva": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "numlot": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "datarec": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "tractat": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "situacio": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "observacionsP": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "observacionsB": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "numpaletprov": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "numbobinaprov": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "ample": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "plegat": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "solapa": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "espesor": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "metres": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "resto": camps(i, 2) = "bit": i = i + 1
  camps(i, 1) = "numpaletgrosbmp": camps(i, 2) = "longbinary": i = i + 1
  camps(i, 1) = "codidebarres": camps(i, 2) = "longbinary": i = i + 1
  camps(i, 1) = "materialdelicat": camps(i, 2) = "byte": i = i + 1
  
  
  dbllistat.Execute ("create table " + taula_tmp3 + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbllistat.Execute ("alter table " + taula_tmp3 + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  
  i = 1
   camps(i, 1) = "ample": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "reservat": camps(i, 2) = "double": i = i + 1
  'camps(i, 1) = "perreservar": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "disponible": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "saldoterra": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "compratlk": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "compratep": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "saldocomprat": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "saldototal": camps(i, 2) = "double": i = i + 1
  
  camps(i, 1) = "idreserva": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "idpalet": camps(i, 2) = "long": i = i + 1
  camps(i, 1) = "estareservat": camps(i, 2) = "bit": i = i + 1
  
  
  
  
  dbllistat.Execute ("create table " + taula_tmp4 + " (id integer)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbllistat.Execute ("alter table " + taula_tmp4 + " add column " + camps(i, 1) + " " + camps(i, 2))
       camps(i, 1) = ""
        Else: i = 1000
    End If
  Next i
  
  
  
End Sub

 Sub imprimir_packinglist(numcomanda As Double, llistat As CrystalReport, Optional gravardata As Boolean, Optional taulaparcials As String)
  Dim dadescomanda As String
  Dim stock As Boolean
  Dim rstpalet As Recordset
  Dim rstpro As Recordset
  Dim rstmat As Recordset
  Dim rstbobina As Recordset
  Dim rstmaterial As Recordset
  Dim rstparcials As Recordset
  Dim rstopcions As Recordset
  Dim dataimpresio As String
  Dim rstgrups As Recordset
  Dim nomdelgrup As String
  Dim vnumerodelgrup As Double
  Dim nomdelclient As String
  Dim mtrsimpresos As Double
  Dim mtrsdolents As Double
  Dim mtrsajust As Double
  Dim mtrsajust11111 As Double
  Dim perpantalla As Boolean
  Dim i As Long
  Dim vdatainicireport As Date
  vdatainicireport = Now
  
  Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
  llistat.Reset
  perpantalla = False
  If numcomanda < 0 Then
     numcomanda = numcomanda * -1
     perpantalla = True
  End If
  If numcomanda < 1 Then Exit Sub
  If atrim(taulaparcials) = "" Then taulaparcials = "parcials"
  obrir_dbllistats
  
'crear_taules_tmp
  dbllistat.Execute "delete * from packinglistxcomanda"
  Set rstllistat = dbllistat.OpenRecordset("Packinglistxcomanda")
  Set rstparcials = dbtmp.OpenRecordset("select * from " + taulaparcials + " where metres>0 and comanda='" + Trim(numcomanda) + "'")
  nomdelgrup = ""
  nomdelclient = ""
  If numcomanda < 10000 Then
     Set rstgrups = dbtmp.OpenRecordset("select * from grupsdepalets where numerogrup=" + atrim(numcomanda))
     If Not rstgrups.EOF Then nomdelgrup = rstgrups!nomdelgrup: vnumerodelgrup = cadbl(rstgrups!numerogrup)
     
  End If
  Set rstopcions = dbtmp.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(numcomanda))
  Set rstmat = dbtmpb.OpenRecordset("SELECT comandes.proximaseccio as seccio, comandes_extres.assignarstock as estoc, comandes.comanda FROM comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda WHERE (comandes.comanda=" + atrim(numcomanda) + ");")

'gravo la data d'impresio
  If gravardata Then
     dbtmpb.Execute "insert into comandes_extres (comanda,dataimpresiopacking) values (" + atrim(numcomanda) + ",NOW)"
     dbtmpb.Execute "update comandes_extres set dataimpresiopacking=NOW where comanda=" + atrim(numcomanda)
     wait 1
  End If
  


  If Not rstmat.EOF Then If rstmat!estoc And (Not comandafeta(numcomanda) Or taulaparcials <> "parcials") Then stock = True: GoTo stock
  
  If rstparcials.EOF Then MsgBox "No hi ha cap bobina assignada a aquesta comanda.", vbInformation, "Bobines": Exit Sub
  While Not rstparcials.EOF
    Set rstpalet = dbtmp.OpenRecordset("select * from Palets where idpalet=" + atrim(cadbl(rstparcials!idpalet)) + " order by datarec")
    If Not rstpalet.EOF Then
     Set rstbobina = dbtmp.OpenRecordset("select * from bobines where idpalet=" + atrim(cadbl(rstparcials!idpalet)) + " and idbobina=" + atrim(cadbl(rstparcials!idbobina)))
     Set rstmaterial = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstpalet!codimatprognou)))
     If Not rstmaterial.EOF Then Set rstpro = dbtmpb.OpenRecordset("select * from proveidors where codi=" + atrim(cadbl(rstmaterial!proveidor)))
     guardar_registre_taulatmp rstpalet, rstpro, rstbobina, rstmaterial, rstparcials
    End If
    rstparcials.MoveNext
  Wend
  calculartotalsmetres mtrsimpresos, mtrsajust, mtrsdolents, mtrsajust11111, numcomanda
stock:
  If stock Then
    If Not rstopcions.EOF Then
       Set rstgrups = dbtmp.OpenRecordset("select * from grupsdepalets where numerogrup=" + atrim(cadbl(rstopcions!grupdestoc)))
       If Not rstgrups.EOF Then
          nomdelgrup = "ESTOC - " + atrim(rstopcions!grupdestoc)
          guardarpaletexemple rstgrups!paletexemple
       End If
    End If
  End If
  'rstllistat.AddNew: Form1.copiafoto llegir_ini("General", "rutallistats", "comandes.ini") + "estoc.bmp", rstllistat!estoc: rstllistat.Update
  Set rstmat = dbtmpb.OpenRecordset("SELECT comandes.comanda, clients.nom as nomclient FROM comandes INNER JOIN clients ON comandes.client = clients.codi WHERE (((comandes.comanda)=" + atrim(numcomanda) + "));")
  If Not rstmat.EOF Then nomdelclient = treure_apostruf(rstmat!nomclient)
  Set rstmat = dbtmpb.OpenRecordset("select dataimpresiopacking from comandes_extres where comanda=" + atrim(numcomanda))
  
  If gravardata And Not rstmat.EOF Then
    i = 0
    Do
     Set rstmat = dbtmpb.OpenRecordset("select dataimpresiopacking from comandes_extres where comanda=" + atrim(numcomanda))
     i = i + 1
    Loop Until IsDate(rstmat!dataimpresiopacking) Or i > 10000
     Else: Set rstmat = dbtmpb.OpenRecordset("select dataimpresiopacking,assignarstock from comandes_extres where comanda=" + atrim(numcomanda))
  End If
  If Not rstmat.EOF Then
    If IsDate(rstmat!dataimpresiopacking) Then
            dataimpresio = format(rstmat!dataimpresiopacking, "Long Date", vbMonday) + " " + format(rstmat!dataimpresiopacking, "hh:nn")
          Else: dataimpresio = ""
    End If
  End If
 dbllistat.Close
 wait 1
   'imprimir llistat
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "packinglistxcomanda.rpt"
 If stock Then llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "packinglistxcomandaESTOC.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.DiscardSavedData = True
 llistat.Formulas(1) = "data1aimpresio='Creació: " + dataimpresio + "'"
 llistat.Formulas(0) = "comanda='" + atrim(numcomanda) + "'"
 llistat.Formulas(12) = "comanda_format='" + format(numcomanda, "#,##0") + "'"
 llistat.Formulas(2) = "dataimpresio='" + format(Now, "long date", vbMonday) + " " + format(Now, "hh:nn") + "'"
 llistat.Formulas(3) = "nomdelgrup='" + nomdelgrup + "'"
 llistat.Formulas(4) = "nomdelclient='" + nomdelclient + "'"
 llistat.Formulas(5) = "texteajust='" + treure_apostruf(textedajust(cadbl(numcomanda))) + "'"
 llistat.Formulas(6) = "msgestoc='" + IIf(stock, generardadescomanda(cadbl(numcomanda)), "") + "'"
 
 llistat.Formulas(7) = "historic='" + "PACKING-LIST ABANS DE PRODUIR" + "'"
 llistat.Formulas(8) = "mtrsimpresos='" + format(mtrsimpresos, "#,##0") + "'"
 llistat.Formulas(9) = "mtrsajust='" + format(mtrsajust, "#,##0") + "'"
 llistat.Formulas(10) = "mtrsdolents='" + format(mtrsdolents, "#,##0") + "'"
 llistat.Formulas(11) = "mtrsajust11111=" + atrim(mtrsajust11111)
 llistat.Formulas(13) = "llistacomandesassignades='" + llistacomandesassignades(vnumerodelgrup) + "'"
 If hihabobinamaterialespecial(numcomanda) Then
     llistat.Formulas(14) = "etiquetamatespecial='MATERIAL ESPECIAL EMBOLICAR AMB BOBINA 53647-X'"
       Else: llistat.Formulas(14) = "etiquetamatespecial=''"
 End If
 If InStr(1, taulaparcials, "historic") = 0 Then
    llistat.Formulas(7) = "historic='" + "" + "'"
 End If
 llistat.Formulas(15) = "comandaacabada='" + mirarsiestafabricada(numcomanda) + "'"
 llistat.Formulas(16) = ""
 wait 1
 DoEvents
 llistat.WindowState = crptMaximized
' llistat.PageZoom 500
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If Form1.mllistaperpantalla.Checked Or perpantalla Then llistat.Destination = crptToWindow
 DoEvents
 'imprimir_reportv9 llistat, llistat.ReportFileName, IIf(llistat.Destination = crptToWindowTrue, True, False)
 llistat.WindowTitle = "Packing-List"
 llistat.Action = 1
 While FindWindow(vbNullString, "Packing-List") > 0 And DateDiff("s", vdatainicireport, Now) < 60
   DoEvents
 Wend
 escriure_ini "baixes", "imprimirpackinglist", "0", "comandes.ini"
 llistat.Formulas(4) = ""
 llistat.Formulas(5) = ""
 llistat.Formulas(6) = ""
 Set rstgrups = Nothing
 obrir_dbllistats
 ' Set rstllistat = Nothing
 ' Set rstllistat = Nothing
  'Set dbllistat = Nothing
'  Set dbtmp = Nothing
'  Set dbtmpb = Nothing
End Sub
Sub imprimir_reportv9(vllistat As CrystalReport, vnomfitxerRPT As String, vpantalla As Boolean)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim camp As TextObject
  Dim f  As OLEObject
  Dim vformula As String
  Dim i As Byte
  Dim vcopies As Byte
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(vnomfitxerRPT, 1)
  For i = 1 To oreport.Database.Tables.Count
    oreport.Database.Tables.Item(i).Location = vllistat.DataFiles(0)
  Next i
  'oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + UCase(atrim(numllauna)) + "'"
  'oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  'oreport.DiscardSavedData
  convertirformules oreport, vllistat
'  oreport.DisplayProgressDialog = FalsE
  If Not vpantalla Then
        For i = 1 To vllistat.PrinterCopies
           oreport.PrintOut False
           wait 1
        Next i
     Else
        Load veurereport
        veurereport.CRViewer.ReportSource = oreport
        veurereport.CRViewer.DisplayGroupTree = False
        veurereport.CRViewer.ViewReport
        veurereport.Show 1
  End If
End Sub
Sub convertirformules(oreport As CRAXDDRT.Report, vllistat As CrystalReport)
  Dim i As Byte
  Dim vn As String
  Dim vv As String
  Dim v As String
  i = 0
  While vllistat.Formulas(i) <> ""
     v = vllistat.Formulas(i)
     vn = Mid(v, 1, InStr(1, v, "=") - 1)
     vv = Mid(v, InStr(1, v, "=") + 1)
     oreport.FormulaFields.GetItemByName(vn).Text = vv
     i = i + 1
  Wend
End Sub

Function mirarsiestafabricada(vnumc As Double) As String
  Dim rst As Recordset
  Dim vacabada As Boolean
  Dim vproducte As String
  
  Set rst = dbbaixes.OpenRecordset("select producte from comandes where comanda=" + atrim(vnumc))
  If rst.EOF Then Exit Function
  vproducte = rst!producte
  If Mid(vproducte, 1, 2) = "PC" Then
     Set rst = dbbaixes.OpenRecordset("select acavada from laminadorestot where comanda=" + atrim(IIf(vproducte = "PC", vnumc - 1, vnumc + 1)))
     If rst.EOF Then
          vacabada = False
           Else: If rst!acavada Then vacabada = True
     End If
     GoTo fi
  End If
  Set rst = dbbaixes.OpenRecordset("select acavada from impressorestot where comanda=" + atrim(vnumc))
  If rst.EOF Then
      Set rst = dbbaixes.OpenRecordset("select acavada from REBOBINADORESTOT where comanda=" + atrim(vnumc))
      If rst.EOF Then
          vacabada = False
           Else: If rst!acavada Then vacabada = True
     End If
     GoTo fi
      Else
        If rst!acavada Then vacabada = True
        GoTo fi
  End If
fi:
  If Not vacabada Then mirarsiestafabricada = " (COMANDA NO ACABADA)"
  Set rst = Nothing
End Function
Function hihabobinamaterialespecial(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("SELECT Parcials.comanda, materials.materialdelicat FROM (Parcials LEFT JOIN Palets ON Parcials.idpalet = Palets.Idpalet) LEFT JOIN materials ON Palets.codimatprognou = materials.codi WHERE (((Parcials.comanda)='" + atrim(vnumc) + "') AND ((materials.materialdelicat)=True));")
   If Not rst.EOF Then hihabobinamaterialespecial = True
   Set rst = Nothing
End Function
Function llistacomandesassignades(vgrup As Double) As String
   If vgrup = 0 Then Exit Function
   llistacomandesassignades = atrim(vgrup)
End Function
Function comandesassignadesalgrup(vgrup As Double) As String
Dim rstm As Recordset
  Dim rstmpc As Recordset
  Dim rstll As Recordset
  Dim nommaterial As String
  Dim nomfamilia As String
  Dim vComandes As String
  Dim rstmat As Recordset
  
  Set rstm = dbtmp.OpenRecordset("select * from grupsdepalets where numerogrup=" + atrim(vgrup))
  While Not rstm.EOF
    vsqlLAM = "SELECT [COMANDES].[cantitatex] AS smetres,comandes.comanda FROM ((opcionsdajust LEFT JOIN comandes ON opcionsdajust.comanda = comandes.comanda) LEFT JOIN productes ON comandes.producte = productes.codi) LEFT JOIN comandes AS comandes_1 ON comandes.linkcomanda1 = comandes_1.comanda WHERE (((comandes.proximaseccio)<>'T') AND ((opcionsdajust.grupdestoc)=" + atrim(cadbl(rstm!numerogrup)) + ") AND ((comandes.producte)='PC' Or (comandes.producte)='PCP' Or (comandes.producte)='PC2') AND ((comandes_1.proximaseccio)='E' Or (comandes_1.proximaseccio)='I' Or (comandes_1.proximaseccio)='L'));"
    vsqlIMP = "SELECT [COMANDES].[cantitatex] AS smetres,comandes.comanda FROM ((opcionsdajust LEFT JOIN comandes ON opcionsdajust.comanda = comandes.comanda) LEFT JOIN productes ON comandes.producte = productes.codi) LEFT JOIN comandes AS comandes_1 ON comandes.linkcomanda1 = comandes_1.comanda WHERE (comandes.proximaseccio='E' Or comandes.proximaseccio='I')  AND (opcionsdajust.grupdestoc=" + atrim(cadbl(rstm!numerogrup)) + ");"
   
    vImp_o_Lam = IIf(rstm!seccio = "I", vsqlIMP, IIf(rstm!seccio = "L", vsqlLAM, ""))
    'posso el valor de metres de comandes assignades a aquest grup en el camp preucompra que es el que faig servir en el llistat per sumar els metres
    Set rstmpc = dbtmp.OpenRecordset(vImp_o_Lam)
    If Not rstmpc.EOF Then
     Set rstll = dbllistat.OpenRecordset("Select * from llistatinventari where metres>0 and preucompra=" + atrim(cadbl(rstm!numerogrup)))
     If Not rstll.EOF Then comandesassignadesalgrup = " [" + atrim(rstmpc!comanda) + "]->" + atrim(cadbl(rstmpc!smetres)) + "_Mtrs "
    End If
 
    rstm.MoveNext
  Wend
End Function
Sub guardarpaletexemple(numpalet As Double)
  Dim rstpalet As Recordset
  Dim rstbobina As Recordset
  Dim rstmaterial As Recordset
  Dim rstpro As Recordset
  Set rstpalet = dbtmp.OpenRecordset("select * from Palets where idpalet=" + atrim(numpalet) + " order by datarec")
    If Not rstpalet.EOF Then
     Set rstbobina = dbtmp.OpenRecordset("select * from bobines where idpalet=" + atrim(cadbl(numpalet)) + " and idbobina=1")
     Set rstmaterial = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstpalet!codimatprognou)))
     If Not rstmaterial.EOF Then Set rstpro = dbtmpb.OpenRecordset("select * from proveidors where codi=" + atrim(cadbl(rstmaterial!proveidor)))
     guardar_registre_taulatmp rstpalet, rstpro, rstbobina, rstmaterial
    End If
    
End Sub
Sub calculartotalsmetres(impresos As Double, ajust As Double, dolents As Double, ajust11 As Double, numc As Double)
   
   Dim rstmtrs As Recordset
   Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", "comandes.ini"))
   Set rstmtrs = dbbaixes.OpenRecordset("SELECT impressores.comanda, Sum(bobinesimp.metres) AS total FROM impressores INNER JOIN bobinesimp ON impressores.Id = bobinesimp.controlid GROUP BY impressores.comanda HAVING (((impressores.comanda)=" + atrim(numc) + "));")
   If Not rstmtrs.EOF Then
       impresos = cadbl(rstmtrs!total)
   End If
   
      Set rstmtrs = dbbaixes.OpenRecordset("SELECT  Sum(tmetresdolents) AS total FROM impressorestot where comanda=" + atrim(numc) + ";")
   If Not rstmtrs.EOF Then
       dolents = cadbl(rstmtrs!total)
   End If
   
   Set rstmtrs = dbtmp.OpenRecordset("SELECT Parcials.comanda, Sum(Parcials.metres) AS total, Parcials.orcomassignacio From parcials GROUP BY Parcials.comanda, Parcials.orcomassignacio HAVING (((Parcials.comanda)='" + atrim(numc) + "') AND ((Parcials.orcomassignacio)='500'));")
   If Not rstmtrs.EOF Then
        ajust = cadbl(rstmtrs!total)
   End If
    
   Set rstmtrs = dbbaixes.OpenRecordset("select paletprova,metresprova,paletprova2,metresprova2 from impressores where (paletprova=11111) and comanda=" + atrim(numc))
   While Not rstmtrs.EOF
      ajust11 = ajust11 + cadbl(rstmtrs!metresprova)
      rstmtrs.MoveNext
   Wend
   Set rstmtrs = dbbaixes.OpenRecordset("select paletprova,metresprova,paletprova2,metresprova2 from impressores where paletprova2=11111 and comanda=" + atrim(numc))
   While Not rstmtrs.EOF
     ajust11 = ajust11 + cadbl(rstmtrs!metresprova2)
     rstmtrs.MoveNext
   Wend
   Set rstmtrs = Nothing
   End Sub
Function generardadescomanda(numc As Double) As String
    Dim rstd As Recordset
    Dim rstmat As Recordset
    Dim nommaterial As String
    Dim descmicres As String
    Set rstd = dbtmpb.OpenRecordset("select * from comandes where comanda=" + atrim(numc))
    If Not rstd.EOF Then
        Set rstmat = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(rstd!materialex))
        If Not rstmat.EOF Then
          nommaterial = descripciomaterial(rstmat)
        End If
        r = ""
        descmicres = assignarmat.micresmaterial(cadbl(rstd!mesuraesp), rstd!espessor, atrim(rstd!tubolam))
        generardadescomanda = nommaterial + " ## Mida: " + atrim(rstd!ampleesq) + "/" + atrim(rstd!plegatesq) + "+" + atrim(rstd!solapa) + " Espesor: " + descmicres + r
    End If
End Function
Function comandafeta(numc As Double) As Boolean

  Dim rstpar As Recordset
  Set rstpar = dbtmp.OpenRecordset("select * from parcials where comanda='" + atrim(numc) + "'")
  comandafeta = False
  While Not rstpar.EOF
    If rstpar!utilitzada Then comandafeta = True
    rstpar.MoveNext
  Wend

End Function


Function textedajust(numc As Double) As String
  Dim rstopcions As Recordset
  Dim rstgrup As Recordset
  Dim t As String
  Dim sisaj As Byte
   Set rstopcions = dbtmp.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(numc))
   If Not rstopcions.EOF Then
     sisaj = atrim(cadbl(rstopcions!sistemadajust))
     If sisaj > 0 Then
      t = atrim(cadbl(rstopcions!mtrsajust)) + " Mtrs D'AJUST.  "
      If sisaj = 1 Then t = t + " S'HA D'UTILITZAR MATERIAL PER LLENÇAR."
      If sisaj = 2 Then
        If cadbl(rstopcions!grupdestoc) > 0 Then
           Set rstgrup = dbtmp.OpenRecordset("select nomdelgrup from grupsdepalets where numerogrup=" + atrim(cadbl(rstopcions!grupdestoc)))
           If Not rstgrup.EOF Then
            t = t + " S'HA D'UTILITZAR MATERIAL D'ESTOC DEL " + UCase(rstgrup!nomdelgrup)
           End If
        End If
      End If
      If sisaj = 3 And cadbl(rstopcions!paletajust) > 0 Then t = t + " S'HA D'UTILITZAR EL PALET " + atrim(rstopcions!paletajust) + " BOB: " + atrim(rstopcions!bobinaajust)
      
      
     End If
   End If
   textedajust = t
End Function
Sub guardar_registre_taulatmp(rstpalet As Recordset, rstpro As Recordset, rstbobina As Recordset, rstmaterial As Recordset, Optional rstparcials As Recordset)
   Dim kg As Double
   Dim ample As Double
   Dim tipusbobina As String
   Dim micres As Double
   assignarmat.actualitzar_metres_disponibles rstbobina!idpalet, rstbobina!idbobina
   rstllistat.AddNew
   rstllistat!palet = rstpalet!idpalet
   rstllistat!bobina = rstbobina!idbobina
   rstllistat!ample = rstpalet!ample
   rstllistat!plegat = rstpalet!plegat
   rstllistat!solapa = rstpalet!solapa
   rstllistat!obert = rstpalet!obert
   rstllistat!microperforat = rstpalet!microperforat
   rstllistat!semielaborat = rstpalet!semielaborat
   rstllistat!carestractat = rstpalet!carestractat
   If cadbl(rstpalet!micres) > 0 Then
       rstllistat!micres = atrim(rstpalet!micres) + " µ"
       micres = cadbl(rstpalet!micres)
      Else:
        If cadbl(rstpalet!grmsm2) > 0 Then
           rstllistat!micres = atrim(rstpalet!grmsm2) + " Gr/m²"
           micres = cadbl(rstpalet!grmsm2) * -1
        End If
   End If
   rstllistat!numpaletprov = atrim(rstpalet!numpaletpro)
   rstllistat!numlot = atrim(rstpalet!numlot)
   If Not rstmaterial.EOF Then
     rstllistat!nommaterial = atrim(rstmaterial!descripcio)
     rstllistat!material = descripciomaterial(rstmaterial)
     rstllistat!familia = rstmaterial!refproducte
     If Not rstpro.EOF Then rstllistat!proveidor = "[" + Mid(atrim(rstpro!tipusproveidorIMPOST) + "   ", 1, 3) + "] " + rstpro!nom
   End If
   
   rstllistat!tractat = rstpalet!tractat
   rstllistat!datarec = rstpalet!dataactivacio
   rstllistat!situacio = rstbobina!sit
   rstllistat!reserva = rstbobina!numcomrev
   
   'rstllistat!comanda = rstbobina!numcom
   On Error Resume Next
   rstllistat!metres = rstparcials!metres
   rstllistat!orcomassignacio = cadbl(rstparcials!orcomassignacio)
   On Error GoTo 0
   
   'persaber els grams mt2
   
   'kg = ((cadbl(rstmaterial!grmcm3) / 0.000001) * (cadbl(rstpalet!micres) * 0.000001) / 1000)
   ample = cadbl(rstpalet!ample)
   'If (rstpalet!semielaborat <> "L") Then ample = cadbl(rstpalet!ample) * 2 + cadbl(rstpalet!solapa)
   'ample = ample / 100
'   kg = demetresakilos(ample, cadbl(rstmaterial!grmcm3), cadbl(micres), atrim(rstpalet!semielaborat), cadbl(rstpalet!solapa))
   If cadbl(rstllistat!metres) >= 0 Then
     kg = compramat.conversiokilos(cadbl(rstpalet!codimatprognou), cadbl(rstpalet!ample), cadbl(rstllistat!metres), cadbl(micres), atrim(rstpalet!semielaborat), cadbl(rstpalet!solapa))
    Else: kg = 0
   End If
   'rstllistat!kilos = kg * ample * rstparcials!metres
   'rstllistat!mtrsdisponibles = rstbobina!disponible
   rstllistat!kilos = kg
   rstllistat!kilosprov = (rstbobina!pesdelproveidor / rstbobina!mts) * rstllistat!metres
   tipusbobina = IIf(assignarmat.esparcial(rstpalet!idpalet, rstbobina!idbobina), "PARCIAL", "")
   tipusbobina = IIf(assignarmat.esrestu(rstpalet!idpalet, rstbobina!idbobina), "RESTE", tipusbobina)
   If tipusbobina = "" Then tipusbobina = "Q"
   rstllistat!tipusbobina = tipusbobina
   rstllistat!impostenvasos = packinglistmirarsielpaletteimpostdenvasos(rstpalet!idpalet, rstpalet!teimpost)
   rstllistat!observacionsp = rstpalet!observ
   rstllistat!observacionsb = ""
   
   'HEM TRET LA OBRS DE BOBINA PER PROBLEMES DE PARACIALS AL IMPORTAR atrim (rstbobina!Obser)
     
   
   'rstllistat!resto = resto
   rstllistat.Update
End Sub
Function packinglistmirarsielpaletteimpostdenvasos(vnumpalet As Double, vteimpost As Boolean) As Boolean
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select numpalet from albaransbip where kgimpostenvasos>0 and kgimpostenvasos<>null and numpalet=" + atrim(vnumpalet), , ReadOnly)
  If Not rst.EOF Then
     packinglistmirarsielpaletteimpostdenvasos = True
       Else: If vteimpost Then packinglistmirarsielpaletteimpostdenvasos = True
  End If
End Function
Function demetresakilos(ample As Double, grmcm3 As Double, micres As Double, semielaborat As String, solapa As Double) As Double
  Dim kg As Double
  kg = ((grmcm3 / 0.000001) * (micres * 0.000001) / 1000)
  ample = cadbl(ample)
  If (semielaborat = "T") Then ample = ample * 2 + solapa
  ample = ample / 100
  demetresakilos = kg
End Function
Function descripciomaterial(rstmat As Recordset) As String
  Dim desc As String
  Dim rstfam As Recordset
  If rstmat.EOF Then Exit Function
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)))
  If Not rstfam.EOF Then desc = desc + atrim(rstfam!descripcio)
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliesmaterials where codi=" + atrim(cadbl(rstmat!subfamilia)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(rstmat!familiacol)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliescolorants where codi=" + atrim(cadbl(rstmat!subfamiliacol)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from familiesaditius where codi=" + atrim(cadbl(rstmat!familiaad)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliesaditius where codi=" + atrim(cadbl(rstmat!subfamiliaad)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
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
Function cabool(valor As Variant) As Boolean
  If IsNull(valor) Then valor = False
  If valor = "" Then valor = False
  If valor = "Sí" Or valor = "S" Then valor = True
  If valor = "No" Or valor = "N" Then valor = False
  If valor = "1" Or valor = "-1" Then valor = True
  If valor = "0" Then valor = False
  If valor Then
    cabool = True
   Else: cabool = False
  End If
End Function

Function generarfitxeralbaratxt(rstalb As Recordset) As String
   generarfitxeralbaratxt = "CAL-" + atrim(rstalb!codiproveidorcomercial) + "-" + format(rstalb!data, "dd") + "-" + format(rstalb!data, "mm") + "-" + format(rstalb!data, "yyyy") + "-" + comprespalets.albprovsensebarres(rstalb!numalbaraprov) + ".txt"
End Function
Function comprovaraccessabip() As Boolean
  If existeix("\\servidorsap\SEIDOR_COMUNICADOR") Then
     comprovaraccessabip = True
    Else:
    '  Shell "c:\windows\system32\cmd.exe /c \\servidorsap\seidor_comunicador /user:Administrador Ipc123 /persistent:yes", vbMaximizedFocus
    '  If existeix("\\servidorsap\SEIDOR_COMUNICADOR") Then
    '          comprovaraccessabip = True
    '     Else: comprovaraccessabip = False
    '  End If
    comprovaraccessabip = False
  End If
End Function
Public Function Redondejar(dblnToR As Double, Optional intCntDec As Integer) As Double
   
    Dim dblPot As Double
    Dim dblF As Double
    
    If dblnToR < 0 Then dblF = -0.5 Else: dblF = 0.5
    dblPot = 10 ^ intCntDec
    Redondejar = Fix(dblnToR * dblPot * (1 + 1E-16) + dblF) / dblPot

End Function
Public Function nomordinador() As String
   nomordinador = Environ("computername")
End Function
Function substituir(cadena As String, buscar As String, canviar As String) As String
   comença = InStr(1, cadena, buscar) - 1
   If comença < 1 Then substituir = cadena: Exit Function
   acaba = comença + Len(buscar) + 1
   cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   substituir = cadena
   'MsgBox linia
End Function
Function substituirtots(cadena As String, buscar As String, canviar As String) As String
  comença = 1
  cadena = " " + cadena
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
Sub enviaremailgeneric(destinatari As String, assumpte As String, cos As String)
   Dim dbenvio As Database
   If atrim(cos) = "" Then Exit Sub
   Set dbenvio = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
   dbenvio.Execute "insert into envios_mails (data,destinatari,assumpte,cos) values (now,'" + destinatari + "','" + treuresimbols(assumpte) + "','" + treuresimbols(cos) + "')"
   Set dbenvio = Nothing
End Sub


