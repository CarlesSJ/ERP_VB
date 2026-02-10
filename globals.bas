Attribute VB_Name = "Module2"
Const LOCALE_SDECIMAL = &HE
Const LOCALE_STHOUSAND = &HF
Const LOCALE_SMONDECIMALSEP = &H16
Const LOCALE_SMONTHOUSANDSEP = &H17
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
     Set dbexist = Nothing
     Exit Function
noexisteix:
      existeixlataula = False
      Set dbexist = Nothing
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
  If Not existeix("c:\ordprog.ini") And nummaq > 0 Then
    LocalID = GetUserDefaultLCID()
    SetLocaleInfo LocalID, LOCALE_SDECIMAL, ","
    SetLocaleInfo LocalID, LOCALE_STHOUSAND, "."
    SetLocaleInfo LocalID, LOCALE_SMONDECIMALSEP, ","
    SetLocaleInfo LocalID, LOCALE_SMONTHOUSANDSEP, "."
  End If
End Sub


'Sub wait(segonsespera As Byte)
'  horaentradawait = Now
'  While DateDiff("s", horaentradawait, Now) < segonsespera
'    DoEvents
'  Wend
'End Sub

Function buscarelnomreal(nomfitxer As String) As String
  If InStr(1, nomfitxer, ".lnk") > 0 Then
    Set ObjShell = CreateObject("WScript.shell")
    Set objLink = ObjShell.CreateShortCut(nomfitxer)
    buscarelnomreal = objLink.targetpath
      Else: buscarelnomreal = nomfitxer
  End If
End Function
Sub imprimir_word2(nomfitxer As String, Optional veurel As Boolean)
  Dim objWord As Word.Application
 ' MsgBox "creant"
  Set objWord = CreateObject("Word.Application")
 ' MsgBox "fi crear"
  If Not existeix(nomfitxer) Then
    If Not exportarcomandes Then MsgBox nomfitxer, vbCritical, "FITXER NO TROBAT"
    Exit Sub
  End If
  objWord.Visible = veurel
  nomfitxer = buscarelnomreal(nomfitxer)
  'On Error Resume Next
  'If existeix("c:\ordprog.ini") Then MsgBox nomfitxer
  'objWord.Documents.Open filename:=nomfitxer, ConfirmConversions:=False, _
  '      ReadOnly:=True, AddToRecentFiles:=False, PasswordDocument:="", _
  '      PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
  '      WritePasswordTemplate:="", Format:=wdOpenFormatAuto
  Set objDoc = objWord.Documents.Add()
  objWord.PrintOut 0, , , , , , , , , , , , nomfitxer
  If cadbl(llegir_ini("General", "esperaimpresiocomandes", fitxerini)) > 0 Then
    wait cadbl(llegir_ini("General", "esperaimpresiocomandes", fitxerini))
      Else:
      wait 3
      escriure_ini "General", "esperaimpresiocomandes", "0", fitxerini
  End If
  'wait 5
  objWord.Quit SaveChanges:=wdDoNotSaveChanges
  
  Set objWord = Nothing
  'On Error GoTo 0
End Sub
Sub imprimir_word(nomfitxer As String, Optional veurel As Boolean)
  Dim v As String
  Dim vnomfitxer As String
  vnomfitxer = nomfitxer
  If llegir_ini("General", "exportant", fitxerini) <> "1" Then
   If InStr(1, vnomfitxer, ".doc") > 0 And InStr(1, vnomfitxer, ".docx") = 0 And Not existeix(vnomfitxer + "x") Then
     '  If Not existeix("c:\temp\docx") Then MkDir "c:\temp\docx"
     '  v = "c:\temp\docx\" + Format(Now, "yymmddhhnnss") + ".doc"
     '  FileCopy vnomfitxer, v
     '  guardar_doc_a_docx v
     '  If existeix(v) Then vnomfitxer = v
     MsgBox "El fitxer word es de la versió anterior primer s'ha de convertir", vbCritical, "Atenció"
     obrir_document vnomfitxer
     MsgBox "FES ACCEPTAR QUAN HAGIS CANVIAT EL FORMAT DEL FITXER.", vbExclamation, "ATENCIÓ"
     If Not existeix(nomfitxer + "x") Then
           MsgBox "No s'ha trobat el fitxer convertit.", vbCritical, "Error": Exit Sub
            Else: Kill nomfitxer
     End If
   End If
  End If
  imprimir_document vnomfitxer
  
End Sub
Sub guardar_doc_a_docx(nomfitxer As String)
  Dim objWord As Word.Application
 ' MsgBox "creant"
  Set objWord = CreateObject("Word.Application")
 ' MsgBox "fi crear"
  If Not existeix(nomfitxer) Then Exit Sub
  objWord.Visible = False
  nomfitxer = buscarelnomreal(nomfitxer)
  'On Error Resume Next
  'If existeix("c:\ordprog.ini") Then MsgBox nomfitxer
  'objWord.Documents.Open filename:=nomfitxer, ConfirmConversions:=False, _
  '      ReadOnly:=True, AddToRecentFiles:=False, PasswordDocument:="", _
  '      PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
  '      WritePasswordTemplate:="", Format:=wdOpenFormatAuto
  Set objDoc = objWord.Documents.Add()
  objWord.Documents.Open nomfitxer
  'atrim(Mid(nomfitxer, 1, InStr(1, nomfitxer, ".doc")) + "docx")
  
  objWord.ActiveDocument.SaveAs nomfitxer, 12, AddToRecentFiles:=False
  If existeix(atrim(Mid(nomfitxer, 1, InStr(1, nomfitxer, ".doc")) + "docx")) Then nomfitxer = atrim(Mid(nomfitxer, 1, InStr(1, nomfitxer, ".doc")) + "docx")
  objWord.Quit SaveChanges:=wdDoNotSaveChanges
  Set objWord = Nothing
  'On Error GoTo 0
End Sub
Sub obrir_word(nomfitxer As String)
  Dim objWord As New Word.Application
  objWord.Visible = True
  objWord.Documents.Open FileName:=nomfitxer, ConfirmConversions:=False, _
        ReadOnly:=True, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto
  
  Set objWord = Nothing
End Sub

Public Function rutadelfitxer(cam As String) As String
   Dim c As Byte
   c = 0
   While InStr(c + 1, cam, "\") <> 0
    c = InStr(c + 1, cam, "\")
   Wend
   If c = 0 Then c = Len(cam)
   rutadelfitxer = Mid(cam, 1, c)
End Function

Public Function nomordinador() As String
   nomordinador = Environ("computername")
   'nomordinador = "ORD_ALICIAM"
  ' nomordinador = "ORD_OANA"
'    nomordinador = "ORD_JOSEPM"
  ' nomordinador = "ORDINADOR_LP"
End Function
Function justificar(v As String, longitut As Integer, Optional DoE As String) As String
    v = Mid(v, 1, longitut)
    If DoE <> "D" Then
       v = v + Space(longitut - Len(v))
      Else: v = Space(longitut - Len(v)) + v
    End If
    justificar = v
End Function
Sub calcular_mtrsminut(vidtreball As Long, vordre As Long, v1fw As Double, v2fw As Double, v1f2 As Double, v2f2 As Double)
  Dim rst As Recordset
  Dim vsql As String
  vsql = "SELECT impressores.numeromaquina as nummaq, Avg(impressores.mtrsminut) AS mitjana, Max(impressores.mtrsminut) AS maxim FROM comandes INNER JOIN impressores ON comandes.comanda = impressores.comanda Where "
  vsql = vsql + " (((comandes.numtreball) = " + atrim(vidtreball) + ") And ((comandes.numordremodificacio) = " + atrim(vordre) + ") And ((impressores.tipus) = 'F')) GROUP BY impressores.numeromaquina HAVING (((Avg(impressores.mtrsminut))<>0));"
  Set rst = dbbaixes.OpenRecordset(vsql)
  While Not rst.EOF
    If rst!nummaq = 7 Then v1fw = Redondejar(cadbl(rst!mitjana), 0): v2fw = Redondejar(cadbl(rst!maxim), 0)
    If rst!nummaq = 9 Then v1f2 = Redondejar(cadbl(rst!mitjana), 0): v2f2 = Redondejar(cadbl(rst!maxim), 0)
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub

'========== Codigo realizado por CULD ==========
'============= culd_@hotmail.com ===============
'La funcion "EAN13_Valido" devuelve si el codigo
'control del EAN13 es VALIDO...
'El algoritmo utilizado es el descrito en la
'siguiente pagina web
'http://latecladeescape.com/w0/recetas-algoritmicas/validar-codigos-ean.html
'La function "EAN13_Control" devuelve el numero de
'control correspondiente para un codigo EAN13 de
'12 digitos (asi devuelve el control que seria el 13)
'===============================================
 
Public Function EAN13_Valido(Codigo As String) As Boolean
'Variables a utilizar
Dim x As Integer
Dim SumaPar As Integer
Dim SumaImpar As Integer
Dim Resto As Integer
Dim Control As Integer
 
'Comprobar que el código tiene 13 dígitos. De no ser así, no es correcto.
If Len(Codigo) <> 13 Then
    EAN13_Valido = False
    Exit Function
End If
 
'Sumar los dígitos de lugares pares por un lado y los de los impares por otro, pero sin incuir el último dígito.
For x = 1 To 12
    If x Mod 2 = 0 Then
        SumaPar = SumaPar + CInt(Mid(Codigo, x, 1))
    Else
        SumaImpar = SumaImpar + CInt(Mid(Codigo, x, 1))
    End If
Next x
 
'multiplicar la suma de los pares por 3.
SumaPar = SumaPar * 3
 
'Sumar el resultado de los pares y el de los impares y hallar el resto de la división por 10.
Resto = (SumaPar + SumaImpar) Mod 10
 
'Realizar la operación 10 menos ese resto y ese es el dígito de control
Control = 10 - Resto
 
'Si como resultado sale 10, entenderemos que el dígito de control es 0.
If Control = 10 Then
    If CInt(Right(Codigo, 1)) = 0 Then
        EAN13_Valido = True
        Exit Function
    Else
        EAN13_Valido = False
        Exit Function
    End If
End If
 
'Comprobar que el dígito de control que hemos calculado y el último dígito del código EAN coinciden
If CInt(Right(Codigo, 1)) = Control Then
    EAN13_Valido = True
    Exit Function
Else
    EAN13_Valido = False
    Exit Function
End If
End Function
 
Public Function EAN13_Control(Codigo As String) As Integer
'Variables a utilizar
Dim x As Integer
Dim SumaPar As Integer
Dim SumaImpar As Integer
Dim Resto As Integer
Dim Control As Integer
 
'Comprobar que el código tiene 12 dígitos. De no ser así, no es correcto.
'devuelvo un numero mayor a 9
If Len(Codigo) <> 12 Then
    EAN13_Control = 10
    Exit Function
End If
 
'Sumar los dígitos de lugares pares por un lado y los de los impares por otro, pero sin incuir el último dígito.
For x = 1 To 12
    If x Mod 2 = 0 Then
        SumaPar = SumaPar + CInt(Mid(Codigo, x, 1))
    Else
        SumaImpar = SumaImpar + CInt(Mid(Codigo, x, 1))
    End If
Next x
 
'multiplicar la suma de los pares por 3.
SumaPar = SumaPar * 3
 
'Sumar el resultado de los pares y el de los impares y hallar el resto de la división por 10.
Resto = (SumaPar + SumaImpar) Mod 10
 
'Realizar la operación 10 menos ese resto y ese es el dígito de control
Control = 10 - Resto
 
'Si como resultado sale 10, entenderemos que el dígito de control es 0.
'de lo contrario, el control es el numero que salio
If Control = 10 Then
    EAN13_Control = 0
Else
    EAN13_Control = Control
End If
End Function



Public Function Enviar_Mail_CDO(SerVidor_SMTP As String, _
                             Para As String, _
                             De As String, _
                             Asunto As String, _
                             Mensaje As String, _
                             Optional Path_Adjunto As String, _
                             Optional Puerto As String = "465", _
                             Optional Usuario As String, _
                             Optional Password As String, _
                             Optional Usar_Autentificacion As Boolean = True, _
                             Optional Usar_SSL As Boolean = True) As Boolean
      
    'Dim Obj_Email As CDO.Message
    Screen.MousePointer = vbHourglass
      
    ' Variable de objeto Cdo.Message
    Set Obj_Email = CreateObject("CDO.Message")
    'Set Obj_Email = New CDO.Message
            
      
    ' Crea un Nuevo objeto CDO.Message
    'Set Obj_Email = New CDO.Message
      
    ' Indica el servidor Smtp para poder enviar el Mail ( puede ser el nombre _
      del servidor o su dirección IP )
      Obj_Email.Configuration.Fields(cdoSMTPServer) = SerVidor_SMTP
    Obj_Email.Configuration.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SerVidor_SMTP
      
      
    Obj_Email.Configuration.Fields(cdoSendUsingMethod) = 2
'    Obj_Email.Configuration.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
      
    ' Puerto. Por defecto se usa el puerto 25, en el caso de Gmail se usan los puertos _
      465 o  el puerto 587 ( este último me dio error )
      
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(Puerto)
  
      
    ' Indica el tipo de autentificación con el servidor de correo _
     El valor 0 no requiere autentificarse, el valor 1 es con autentificación
    Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/" & _
                "configuration/smtpauthenticate") = Abs(Usar_Autentificacion)
      
      
      
        ' Tiempo máximo de espera en segundos para la conexión
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
  
      
    ' Configura las opciones para el login en el SMTP
    If Usar_Autentificacion Then
  
    ' Id de usuario del servidor Smtp ( en el caso de gmail, debe ser la dirección de correro _
     mas el @gmail.com )
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendusername") = Usuario
  
    ' Password de la cuenta
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Password
  
    ' Indica si se usa SSL para el envío. En el caso de Gmail requiere que esté en True
    Obj_Email.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = Usar_SSL
      
    End If
      
  
    ' *********************************************************************************
    ' Estructura del mail
    '**********************************************************************************
      
    ' Dirección del Destinatario
    Obj_Email.To = Para
      
    ' Dirección del remitente
    Obj_Email.From = De
      
    ' Asunto del mensaje
    Obj_Email.Subject = Asunto
      
    ' Cuerpo del mensaje
    If InStr(1, Mensaje, "cosmissatge.txt") > 0 Then
        Obj_Email.TextBody = CreateObject("Scripting.FileSystemObject").OpenTextFile(Mensaje, 1).ReadAll
       Else: Obj_Email.TextBody = Mensaje
    End If
    'Ruta del archivo adjunto
      
    If Path_Adjunto <> vbNullString Then
        Obj_Email.AddAttachment (Path_Adjunto)
    End If
      
    ' Actualiza los datos antes de enviar
    Obj_Email.Configuration.Fields.Update
      
   ' On Error Resume Next
    ' Envía el email
    formspooler.listlog.AddItem "Preparant...(Obj_Email.Send)"
    If Not existeix(Path_Adjunto) And Path_Adjunto <> "" Then formspooler.listlog.AddItem "Preparant...(Obj_Email.Send) no existeix adjunt"
    Obj_Email.Send
    formspooler.listlog.AddItem "Preparant...(Obj_Email.Send) Acabat "
    wait 2
    If err.Number = 0 Then
       Enviar_Mail_CDO = True
     Else
        Usuario = err.Description
        listlog.AddItem "Preparant...(" + err.Description + ")"
        'MsgBox Usuario
    End If
      
    ' Descarga la referencia
    If Not Obj_Email Is Nothing Then
        Set Obj_Email = Nothing
    End If
      
    On Error GoTo 0
    Screen.MousePointer = vbNormal
  
End Function

Sub cambiarnomarxiu(vnomarxiu As String, vnomarxiunou As String)
    'Elimina la carpeta sin necesidad de eliminar los ficheros en ella contenidos
   ' MsgBox "Eliminar " + strRuta
    Dim FSO As Object
    Dim i As Byte
    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo error
    If existeix(vnomarxiunou) Then GoTo fi
    If existeix(vnomarxiu) Then FSO.movefile vnomarxiu, vnomarxiunou
fi:
    Set FSO = Nothing
    Exit Sub
error:
    escriure_log "(cambiarnomarxiu)" + vbNewLine + vrutaorigen + vnomarxiu + " --->  " + vnomarxiunou, "c:\temp\Log_EnviarMails_servidor.txt"
End Sub
Sub borra_carpeta(strRuta As String)
    'Elimina la carpeta sin necesidad de eliminar los ficheros en ella contenidos
   ' MsgBox "Eliminar " + strRuta
    Dim FSO As Object
    Dim i As Byte
    escriure_log "(Borrar Carpeta) Inici " + strRuta + vbNewLine, "c:\temp\Log_EnviarMails_servidor.txt"
    'quitamos la posible última barra \ de la ruta
    If Right(strRuta, 1) = "\" Then strRuta = Left(strRuta, Len(strRuta) - 1)
    'llamamos al script de FileSystem
    Set FSO = CreateObject("Scripting.FileSystemObject")
    'y acabamos borrando la carpeta y todo su contenido
    On Error GoTo fi

    If existeix(strRuta) Then
      If FSO.FolderExists(strRuta) Then
          FSO.DeleteFolder strRuta, True
          wait 1
          If FSO.FolderExists(strRuta) Then wait 1: FSO.DeleteFolder strRuta, True
      End If
      
      If existeix(strRuta) Then borrar_fitxer strRuta: wait 1
      If existeix(strRuta) Then
        For i = 1 To 2
         If existeix(strRuta) Then wait 1: borrar_fitxer strRuta
        Next i
      End If
    End If
fi:
    Set FSO = Nothing
    'MsgBox "Fi de Eliminar " + strRuta
    escriure_log "(Borrar Carpeta) Fi " + strRuta + vbNewLine, "c:\temp\Log_EnviarMails_servidor.txt"
End Sub
Sub borrar_fitxer(vfitxer As String)
  Dim FSO As Object
  escriure_log "(Borrar fitxer) Inici " + vfitxer + vbNewLine, "c:\temp\Log_EnviarMails_servidor.txt"
  Set FSO = CreateObject("Scripting.FileSystemObject")
  On Error Resume Next
  If existeix(vfitxer) Then FSO.DeleteFile vfitxer, True  'Kill vfitxer
  Set FSO = Nothing
  escriure_log "(Borrar fitxer) Fi " + vfitxer + vbNewLine, "c:\temp\Log_EnviarMails_servidor.txt"
End Sub
Sub enviaremailgeneric(destinatari As String, assumpte As String, vcos As String)
   Dim dbenvio As Database
   Dim cos1 As String
   Dim cos2 As String
   Dim cos3 As String
   Dim cos4 As String
   Dim cos5 As String
   Dim cos6 As String
   Dim rst As Recordset
   If atrim(vcos) = "" Then Exit Sub
   cos1 = Mid(vcos, 1, 255)
   If Len(vcos) > 255 Then cos2 = Mid(vcos, 256, 510)
   If Len(vcos) > 510 Then cos3 = Mid(vcos, 511, 765)
   If Len(vcos) > 765 Then cos4 = Mid(vcos, 766, 1019)
   If Len(vcos) > 1019 Then cos5 = Mid(vcos, 1020, 1275)
   If Len(vcos) > 1275 Then cos6 = Mid(vcos, 1276, 1500) + vbNewLine + "... E-MAIL INCOMPLERT"
   'If Len(vcos) > 1529 Then cos6 = Mid(vcos, 1530, 1750) + Chr(13) + Chr(10) + "... E-MAIL INCOMPLERT"
   
   Set dbenvio = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
   dbenvio.Execute "insert into envios_mails (data,destinatari,assumpte,cos,cos2,cos3,cos4,cos5,cos6) values (now,'" + destinatari + "','" + treuresimbols(assumpte) + "','" + treuresimbols(cos1) + "','" + treuresimbols(cos2) + "','" + treuresimbols(cos3) + "','" + treuresimbols(cos4) + "','" + treuresimbols(cos5) + "','" + treuresimbols(cos6) + "')"
   Set rst = dbenvio.OpenRecordset("select * from envios_mails order by id desc")
   If Not rst.EOF Then dbenvio.Execute "update envios_mails_linies set id_envio=" + atrim(rst!ID) + " where id_envio=0"
   'If Len(vcos) > 1785 Then enviaremailgeneric destinatari, assumpte, Mid(vcos, 1785)
   Set rst = Nothing
   Set dbenvio = Nothing
End Sub
Function enviaremailgenericambadjunt(sSendTo As String, sSubject As String, sText As String, vadjunt As String, Optional noensenyarinterficie As Boolean) As Boolean
   Dim usuarim As String
   Dim contrasenyam As String
   Dim destinatari As String
   Dim vnomcarpeta As String
  
   vnomcarpeta = "\\serverprodu\Dades\progcomandes\dades\spoolerenviament\" + nomordinador + "_" + Format(Now, "yymmdd_hhnnss")
   If Not existeix(vnomcarpeta) Then MkDir vnomcarpeta
   escriure_ini "Capcalera", "apuntperenviar", "No", vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "data", Now, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "nomordinador", nomordinador, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "usuari", usuarim, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "contrasenya", contrasenyam, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "destinatari", sSendTo, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "remitent", usuarim, vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "assumpte", treure_apostruf(sSubject), vnomcarpeta + "\dadesmail.txt"
   escriure_ini "Capcalera", "adjunt", vnomcarpeta + "\" + substituir(atrim(vadjunt), rutadelfitxer(vadjunt), ""), vnomcarpeta + "\dadesmail.txt"
   Copiar_Fitxer vadjunt, vnomcarpeta
   Open "c:\temp\cosmissatge.txt" For Output As #2
   Print #2, sText
   Close #2
   Copiar_Fitxer "c:\temp\cosmissatge.txt", vnomcarpeta
   Kill "c:\temp\cosmissatge.txt"
  
   escriure_ini "Capcalera", "apuntperenviar", "Si", vnomcarpeta + "\dadesmail.txt"
End Function
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
Function treuresimbolsnovalidsnomfitxer(desc As String) As String
   desc = substituir(desc, "\", "_")
   desc = substituir(desc, "/", "_")
   desc = substituir(desc, "|", "_")
   desc = substituir(desc, ":", ";")
   desc = substituir(desc, "?", "¿")
   desc = substituir(desc, "*", "x")
   desc = substituir(desc, """", "'")
   desc = substituir(desc, ">", "+")
   desc = substituir(desc, "<", "-")
   treuresimbolsnovalidsnomfitxer = desc
End Function

Sub eliminar_fitxer(vnomfitxer As String)
  Dim vlinia As String
   On Error GoTo Errors
   If InStr(1, vnomfitxer, "*") Then
      Kill vnomfitxer
        Else: If existeix(vnomfitxer) Then Kill vnomfitxer
   End If
   Exit Sub
Errors:
  If Not existeix("c:\temp\errors_eliminarfitxers.txt") Then
     Open "c:\temp\errors_eliminarfitxers.txt" For Output As #8
      Else: Open "c:\temp\errors_eliminarfitxers.txt" For Append As #8
  End If
  vlinia = Trim(Format(Now, "dd/mm/yy hh:nn:ss")) & " - " & vnomfitxer + "(" + err.Description + ")"
  Print #8, vlinia
 Close #8

End Sub
Sub escriure_log(vmissatge As String, vfitxer As String)
  Static vdins As Date
  
  If vdins <> 0 Then
     While DateDiff("s", vdins, Now) < 2
      DoEvents
     Wend
  End If
  vdins = Now
  If Not existeix(vfitxer) Then
     Open vfitxer For Output As #10
      Else: Open vfitxer For Append As #10
  End If
  vlinia = Trim(Format(Now, "dd/mm/yy hh:nn:ss")) & " - " & vmissatge
  Print #10, vlinia
 Close #10
 If FileLen(vfitxer) > 1000000 Then
    Copiar_Fitxer vfitxer, rutadelfitxer(vfitxer) + "Historic_" + substituir(atrim(vfitxer), rutadelfitxer(vfitxer), "")
    eliminar_fitxer (vfitxer)
 End If
 vdins = 0
End Sub
Sub SendKeys(vtecla As String)
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

WshShell.SendKeys vtecla
Set WshShell = Nothing

End Sub

