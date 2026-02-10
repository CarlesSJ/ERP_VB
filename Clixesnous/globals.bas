Attribute VB_Name = "Module2"
Global arguments As Variant
Const LOCALE_SDECIMAL = &HE
Const LOCALE_STHOUSAND = &HF

Const GW_HWNDNEXT = 2
 Declare Function PostMessage Lib "User" (ByVal hWnd _
      As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, _
      lParam As Any) As Integer

 Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
 Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, _
  ByVal wCmd As Long) As Long
 Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
 Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
  (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
 Declare Function GetWindowThreadProcessId Lib "user32" _
  (ByVal hWnd As Long, lpdwprocessid As Long) As Long

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
    X As Long
    Y As Long
End Type

 

 
Public Function GetTCursX() As Long
    Dim pt As POINTAPI
    GetCaretPos pt
    GetTCursX = pt.X
End Function
 
Public Function GetTCursY() As Long
    Dim pt As POINTAPI
    GetCaretPos pt
    GetTCursY = pt.Y
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
      OpenFile.hwndOwner = frmclixes.hWnd
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

Sub imprimir_word(nomfitxer As String)
  Dim objWord As New Word.Application
  If Not existeix(nomfitxer) Then Exit Sub
  objWord.visible = False
  On Error Resume Next
  objWord.Documents.Open filename:=nomfitxer, ConfirmConversions:=False, _
        ReadOnly:=True, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto
  objWord.PrintOut
  wait 2
  objWord.Quit SaveChanges:=wdDoNotSaveChanges
  Set objWord = Nothing
  On Error GoTo 0
End Sub
Sub obrir_word(nomfitxer As String)
  Dim objWord As New Word.Application
  objWord.visible = True
  objWord.Documents.Open filename:=nomfitxer, ConfirmConversions:=False, _
        ReadOnly:=True, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto
  Set objWord = Nothing
End Sub


Sub calcular_mtrsminut(vidtreball As Long, vordre As Long, v1fw As Double, v2fw As Double, v1f2 As Double, v2f2 As Double)
  Dim rst As Recordset
  Dim vsql As String
  vsql = "SELECT impressores.numeromaquina as nummaq, Avg(impressores.mtrsminut) AS mitjana, Max(impressores.mtrsminut) AS maxim FROM comandes INNER JOIN impressores ON comandes.comanda = impressores.comanda Where "
  vsql = vsql + " (((comandes.numtreball) = " + atrim(vidtreball) + ") And ((comandes.numordremodificacio) = " + atrim(vordre) + ") And ((impressores.tipus) = 'F')) GROUP BY impressores.numeromaquina HAVING (((Avg(impressores.mtrsminut))<>0));"
  Set rst = dbbaixes.OpenRecordset(vsql)
  While Not rst.EOF
    If rst!nummaq = 7 Then v1fw = cadbl(rst!mitjana): v2fw = cadbl(rst!maxim)
    If rst!nummaq = 9 Then v1f2 = cadbl(rst!mitjana): v2f2 = cadbl(rst!maxim)
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
