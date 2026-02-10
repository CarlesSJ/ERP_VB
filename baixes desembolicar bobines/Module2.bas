Attribute VB_Name = "Module2"
Global arguments As Variant
Function enviaremail(sSendTo As String, sSubject As String, sText As String, Optional adjunt As String, Optional vidavis As Long, Optional adjunt2 As String, Optional adjunt3 As String) As Boolean
  Dim usuarim As String
  Dim contrasenyam As String
  Dim destinatari As String
  Dim vnomcarpeta As String
  Dim vadjunt As String
  Dim vadjunt2 As String
  Dim vadjunt3 As String
  Dim vv As String
  vv = llegir_ini("destinataris", sSendTo, "enviarservidor.ini")
  If vv = "{[}]" Then
    sSendTo = sSendTo
     Else: sSendTo = vv
  End If
  vadjunt = adjunt
  vadjunt2 = adjunt2
  vadjunt3 = adjunt3
  vnomcarpeta = "\\serverprodu\Dades\progcomandes\dades\spoolerenviament\" + nomordinador + "_" + Format(Now, "yymmdd_hhnnss")
'  usuarim = llegir_ini("dadesservidor", "usrsmtp", "enviarservidor.ini")
'  contrasenyam = llegir_ini("dadesservidor", "passsmtp", "enviarservidor.ini")
  If usuarim = "{[}]" Or contrasenyam = "{[}]" Then
      escriure_ini "dadesservidor", "usrsmtp", " ", "enviarservidor.ini"
      escriure_ini "dadesservidor", "passsmtp", " ", "enviarservidor.ini"
      MsgBox "L'usuari o la contrasenya no estan entrades", vbCritical, "Error": Exit Function
  End If
  If Not existeix(vnomcarpeta) Then MkDir vnomcarpeta
  escriure_ini "Capcalera", "apuntperenviar", "No", vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "data", Now, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "nomordinador", nomordinador, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "usuari", usuarim, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "contrasenya", contrasenyam, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "destinatari", sSendTo, vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "remitent", "incidencies@inplacsa.com", vnomcarpeta + "\dadesmail.txt"
  escriure_ini "Capcalera", "assumpte", treure_apostruf(sSubject), vnomcarpeta + "\dadesmail.txt"
  If existeix(vadjunt) Then
     Copiar_Fitxer vadjunt, vnomcarpeta
     vadjunt = substituirtot(vadjunt, rutadelfitxer(vadjunt), vnomcarpeta + "\")
     escriure_ini "Capcalera", "adjunt", vadjunt, vnomcarpeta + "\dadesmail.txt"
  End If
  If existeix(vadjunt2) Then
     Copiar_Fitxer vadjunt2, vnomcarpeta
     vadjunt2 = substituirtot(vadjunt2, rutadelfitxer(vadjunt2), vnomcarpeta + "\")
     escriure_ini "Capcalera", "adjunt2", vadjunt2, vnomcarpeta + "\dadesmail.txt"
  End If
  If existeix(vadjunt3) Then
     Copiar_Fitxer vadjunt3, vnomcarpeta
     vadjunt3 = substituirtot(vadjunt3, rutadelfitxer(vadjunt3), vnomcarpeta + "\")
     escriure_ini "Capcalera", "adjunt3", vadjunt3, vnomcarpeta + "\dadesmail.txt"
  End If
  
  If LCase(sText) <> "c:\temp\cosmissatge.txt" Then
        Open "c:\temp\cosmissatge.txt" For Output As #2
        Print #2, sText
        passarliniesdavisosalfitxertxt vidavis
        Close #2
   End If
   Copiar_Fitxer "c:\temp\cosmissatge.txt", vnomcarpeta
   If existeix("c:\temp\cosmissatge.txt") Then Kill "c:\temp\cosmissatge.txt"
   escriure_ini "Capcalera", "apuntperenviar", "Si", vnomcarpeta + "\dadesmail.txt"
   wait 1
   
End Function
Public Function nomordinador() As String
   nomordinador = Environ("computername")
End Function
Sub passarliniesdavisosalfitxertxt(vidavis As Long)
    Dim rst As Recordset
    Dim v As String
    If vidavis = 0 Then Exit Sub
    Set rst = db.OpenRecordset("select * from envios_mails_linies where id_envio=" + atrim(vidavis))
    If Not rst.EOF Then
       Print #2, ""
       Print #2, ""
    End If
    While Not rst.EOF
      v = atrim(rst!descripcio)
      If Len(v) > 0 Then
        'If InStr(1, v, Chr(10)) = 0 Then v = v + Chr(10)
        Print #2, v
      End If
      rst.MoveNext
    Wend
    Set rst = Nothing
End Sub

