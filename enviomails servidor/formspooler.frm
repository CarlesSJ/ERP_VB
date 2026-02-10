VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formspooler 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spooler Email (Temps de refresc cada 10 segons)"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9795
   ControlBox      =   0   'False
   Icon            =   "formspooler.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox checknoenviar 
      Caption         =   "No enviar emails"
      Height          =   195
      Left            =   6375
      TabIndex        =   6
      Top             =   15
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Amagar"
      Height          =   360
      Left            =   8805
      TabIndex        =   4
      Top             =   45
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Interval        =   9000
      Left            =   1575
      Top             =   2610
   End
   Begin TabDlg.SSTab pestanyes 
      Height          =   6810
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   12012
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Bandeja Sortida"
      TabPicture(0)   =   "formspooler.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "etstatus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "llistasortida"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Dir1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "llistadirs"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Bandeja Enviats"
      TabPicture(1)   =   "formspooler.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dataenviats"
      Tab(1).Control(1)=   "reixaenviats"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Log"
      TabPicture(2)   =   "formspooler.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "listlog"
      Tab(2).ControlCount=   1
      Begin VB.ListBox llistadirs 
         Height          =   450
         Left            =   3210
         TabIndex        =   8
         Top             =   1245
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.DirListBox Dir1 
         Height          =   5040
         Left            =   5550
         TabIndex        =   7
         Top             =   825
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.ListBox listlog 
         Height          =   5910
         Left            =   -74805
         TabIndex        =   5
         Top             =   510
         Width           =   9045
      End
      Begin VB.Data dataenviats 
         Caption         =   "dataenviats"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   -71175
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from registreenviament order by data Desc"
         Top             =   4530
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSDBGrid.DBGrid reixaenviats 
         Bindings        =   "formspooler.frx":05DE
         Height          =   5970
         Left            =   -74850
         OleObjectBlob   =   "formspooler.frx":05F4
         TabIndex        =   3
         Top             =   495
         Width           =   9000
      End
      Begin VB.ListBox llistasortida 
         Height          =   5910
         Left            =   135
         TabIndex        =   1
         Top             =   570
         Width           =   9165
      End
      Begin VB.Label etstatus 
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   6495
         Width           =   9000
      End
   End
End
Attribute VB_Name = "formspooler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSTab1_DblClick()

End Sub

Private Sub Command1_Click()
  Me.Visible = False
End Sub

Sub borrarfitxersmalcopiats(vubicacio As String)
  Dim vdir As String
  Dim vdata As Date
  vdir = Dir(vubicacio + "\*.*", vbArchive)
  Me.Caption = ""
  While vdir <> ""
    If vdir <> "." And vdir <> ".." Then
      vdata = FileDateTime(vubicacio + "\" + vdir)
      If DateDiff("s", vdata, Now) > 20 Then Kill vubicacio + "\" + vdir
    End If
    vdir = Dir(, vbArchive)
  Wend
End Sub
Sub carregar_pendents()
 Dim vubicaciomails As String
 Dim vdir As String
 Dim vdata As String
 Dim vdestinatari As String
 Dim vassumpte As String
 
 escriure_log "(carregar_pendents) Envio EMAILS" + vbNewLine, "c:\temp\Log_EnviarMails_servidor.txt"
 If llistasortida.Tag = "1" Then Exit Sub
 llistasortida.Tag = "1"
 etstatus.Caption = "Comprovant bandeja de sortida..."
 llistasortida.Clear
 llistadirs.Clear
 vubicaciomails = "C:\Dades\progcomandes\dades\spoolerenviament"
 If Not existeix(vubicaciomails) Then vubicaciomails = "\\serverprodu\Dades\progcomandes\dades\spoolerenviament"
 
 Dir1.Path = vubicaciomails
 Dir1.Refresh
 
 For i = 0 To Dir1.ListCount - 1
     If existeix(Dir1.List(i) + "\dadesmail.txt") Then
      If llegir_ini("Capcalera", "apuntperenviar", Dir1.List(i) + "\dadesmail.txt") = "Si" Then
       vdata = llegir_ini("Capcalera", "data", Dir1.List(i) + "\dadesmail.txt")
       If IsDate(vdata) Then
          vdestinatari = llegir_ini("Capcalera", "destinatari", Dir1.List(i) + "\dadesmail.txt")
          vassumpte = llegir_ini("Capcalera", "assumpte", Dir1.List(i) + "\dadesmail.txt")
          llistasortida.AddItem IIf(InStr(1, Dir1.List(i), "#Error") > 0, "#Error " + justificar(vdata, 20), "") + " " + justificar(vdestinatari, 30) + " " + vassumpte
          llistadirs.AddItem Dir1.List(i)
            Else
              borra_carpeta Dir1.List(i)
              If existeix(Dir1.List(i)) Then llistasortida.AddItem "#Error eliminant " + Dir1.List(i)
       End If
      End If
          Else
            wait 1
            If Not existeix(Dir1.List(i) + "\dadesmail.txt") Then
                borra_carpeta Dir1.List(i)
                If existeix(Dir1.List(i)) Then llistasortida.AddItem "#Error eliminant " + Dir1.List(i)
            End If
     End If
 Next i
 
 
 
' vdir = Dir(vubicaciomails + "\*.", vbDirectory)
' While vdir <> ""
'   If vdir <> "." And vdir <> ".." Then
'     vdata = llegir_ini("Capcalera", "data", vubicaciomails + "\" + vdir + "\dadesmail.txt")
'     If vdata <> "{[}]" Then
'      vdestinatari = llegir_ini("Capcalera", "destinatari", vubicaciomails + "\" + vdir + "\dadesmail.txt")
'      vassumpte = llegir_ini("Capcalera", "assumpte", vubicaciomails + "\" + vdir + "\dadesmail.txt")
'      llistasortida.AddItem justificar(vdata, 20) + " " + justificar(vdestinatari, 30) + " " + vassumpte
'     End If
'   End If
'   On Error GoTo cont
'   vdir = Dir(, vbDirectory)
'
'   DoEvents
' Wend

cont:
 etstatus.Caption = ""
 DoEvents
 llistasortida.Tag = ""
 escriure_log "(carregar_pendents) FI" + vbNewLine, "c:\temp\Log_EnviarMails_servidor.txt"
End Sub
Sub enviar_pendents()
 Dim vubicaciomails As String
 Dim vdir As String
 Dim vdata As String
 Dim vdestinatari As String
 Dim vassumpte As String
 Dim vcont As Long
 Static vnohihahaguterrorslultimcop As Boolean
inici:
  escriure_log "Inici Enviar_pendents.", "c:\temp\Log_EnviarMails_servidor.txt"
 listlog.AddItem "Inici Enviant..."
 vubicaciomails = "\\serverprodu\Dades\progcomandes\dades\spoolerenviament"
 escriure_log "Enviar_pendents -ubicacio.", "c:\temp\Log_EnviarMails_servidor.txt"
 Dir1.Path = vubicaciomails
 Dir1.Refresh
 vcont = 0
 While vcont < (Dir1.ListCount)
    If InStr(1, Dir1.List(vcont), "#Error#") = 0 Then
      escriure_log "Enviar_pendents -agafar dades capçalera.", "c:\temp\Log_EnviarMails_servidor.txt"
      If llegir_ini("Capcalera", "apuntperenviar", Dir1.List(vcont) + "\dadesmail.txt") = "Si" Then
       vdata = llegir_ini("Capcalera", "data", Dir1.List(vcont) + "\dadesmail.txt")
       listlog.AddItem "ENVIANT: " + Dir1.List(vcont) + "\dadesmail.txt"
       If IsDate(vdata) Then
        listlog.AddItem "Preparant..."
        escriure_log "Enviar_pendents -entrar preparar_envio.", "c:\temp\Log_EnviarMails_servidor.txt"
        preparar_enviomail Dir1.List(vcont)
        listlog.AddItem "Enviat..."
        vnohihahaguterrorslultimcop = True
        If InStr(1, etstatus, "#Error") > 0 Then
          listlog.AddItem "Error enviament..."
          listlog.AddItem etstatus
          vnohihahaguterrorslultimcop = False
        End If
         Else
           borra_carpeta Dir1.List(vcont)
           GoTo inici
       End If
      End If
         Else
            vdir = Dir1.List(vcont)
            If vnohihahaguterrorslultimcop Then Copiar_Fitxer vdir, substituir(vdir, "#Error#", "")
    End If
    vcont = vcont + 1
 Wend
 

cont:
  escriure_log "Fi Enviar_pendents.", "c:\temp\Log_EnviarMails_servidor.txt"

End Sub
Sub preparar_enviomail(vdir As String)
  Dim vdata As String
  Dim vdestinatari As String
  Dim vassumpte As String
  Dim vusuari As String
  Dim vcontrassenya As String
  Dim vadjunt As String
  Dim vadjunt2 As String
  Dim vadjunt3 As String
  Dim vremitent As String
  Dim vpassword As String
  Dim vresp As Boolean
  Dim vvalors As String
  Dim v As String
  
  escriure_log "Enviar_pendents -llegint dades preparar envio.", "c:\temp\Log_EnviarMails_servidor.txt"
  If llegir_ini("Capcalera", "apuntperenviar", vdir + "\dadesmail.txt") = "No" Then
     etstatus = "#NoApuntperenviar enviant: " + vassumpte: DoEvents
     vdata = llegir_ini("Capcalera", "data", vdir + "\dadesmail.txt")
     If DateDiff("n", vdata, Now) > 0 Then borra_carpeta vdir
     Exit Sub
  End If
  vdata = llegir_ini("Capcalera", "data", vdir + "\dadesmail.txt")
  vremitent = llegir_ini("Capcalera", "remitent", vdir + "\dadesmail.txt")
  vdestinatari = llegir_ini("Capcalera", "destinatari", vdir + "\dadesmail.txt")
  vassumpte = llegir_ini("Capcalera", "assumpte", vdir + "\dadesmail.txt")
  vusuari = llegir_ini("Capcalera", "usuari", vdir + "\dadesmail.txt")
  vcontrassenya = llegir_ini("Capcalera", "contrasenya", vdir + "\dadesmail.txt")
  If vusuari = "" Then
    vusuari = llegir_ini("dadesservidor", "usrsmtp", "enviarservidor.ini")
    vcontrassenya = llegir_ini("dadesservidor", "passsmtp", "enviarservidor.ini")
  End If
  vadjunt = llegir_ini("Capcalera", "adjunt", vdir + "\dadesmail.txt")
  If Not existeix(vadjunt) Then
       vadjunt = ""
         Else
           v = vadjunt:  vadjunt = substituir(vadjunt, "&", "_"):  Copiar_Fitxer v, vadjunt
  End If
  'If vadjunt = "{[}]" Then vadjunt = ""
  vadjunt2 = llegir_ini("Capcalera", "adjunt2", vdir + "\dadesmail.txt")
  If Not existeix(vadjunt2) Then
      vadjunt2 = ""
       Else:
         v = vadjunt2:  vadjunt2 = substituir(vadjunt2, "&", "_"):  Copiar_Fitxer v, vadjunt2
  End If
  'If vadjunt2 = "{[}]" Then vadjunt2 = ""
  vadjunt3 = llegir_ini("Capcalera", "adjunt3", vdir + "\dadesmail.txt")
  If Not existeix(vadjunt3) Then
       vadjunt3 = ""
         Else:
             v = vadjunt3:  vadjunt3 = substituir(vadjunt3, "&", "_"):  Copiar_Fitxer v, vadjunt3
  End If
  'If vadjunt3 = "{[}]" Then vadjunt3 = ""
  vassumpte = substituir(vassumpte, "&", "_")
  etstatus = "Enviant: " + vdestinatari + " --> " + vassumpte
  listlog.AddItem "Preparant...(Enviar_Mail_CDO)"
  
  If vusuari = "{[}]" Then
    vusuari = llegir_ini("dadesservidor", "usrsmtp", "enviarservidor.ini")
    vcontrassenya = llegir_ini("dadesservidor", "passsmtp", "enviarservidor.ini")
    vremitent = vusuari
  End If
  
  'vresp = Enviar_Mail_CDO("smtp.gmail.com", vdestinatari, vremitent, vassumpte, vdir + "\cosmissatge.txt", vadjunt, , vusuari, vcontrassenya, True, True)
  escriure_log "Enviar_pendents -enviar switchmail.", "c:\temp\Log_EnviarMails_servidor.txt"
  vresp = form1.enviaremailswitchmail(vremitent, vdestinatari, vassumpte, vdir + "\cosmissatge.txt", vadjunt, , vusuari, vcontrassenya, vadjunt2, vadjunt3)
  listlog.AddItem "Preparant...(FI DE Enviar_Mail_CDO) " + err.Description
  wait 1
  If vresp Then
    etstatus = "Passant a Enviats: " + vdestinatari + " --> " + vassumpte
    listlog.AddItem "Enviat...(Borrar acrpeta) " + vdir
    borra_carpeta vdir
    vvalors = "('" + vdata + "','" + vdestinatari + "','" + vremitent + "','" + vassumpte + "',now)"
    escriure_log "Envio EMAILS " + vvalors + vbNewLine, "c:\temp\Log_EnviarMails_servidor.txt"
     Else:
       vvalors = "('" + vdata + "','" + vdestinatari + "','" + vremitent + "','Error: " + vusuari + " / " + vassumpte + "',now)"
       etstatus = "#Error enviant: " + vusuari + "  " + vassumpte: DoEvents
       escriure_log "ERROR ENVIO EMAIL " + vvalors + vbNewLine, "c:\temp\Log_EnviarMails_servidor.txt"
       'If Not existeix(vdir + "#Error#") Then MkDir vdir + "#Error#"
       If Not existeix(vdir + "#Error#") Then
          If existeix("c:\temp\registreemail.txt") Then Copiar_Fitxer "c:\temp\registreemail.txt", vdir
          Copiar_Fitxer vdir, vdir + "#Error#"
       End If
       borra_carpeta vdir
  End If
  On Error GoTo 0
  escriure_log "Enviar_pendents -gravant resultat envio.", "c:\temp\Log_EnviarMails_servidor.txt"
  db.Execute "insert into registreenviament (dataentradacua,destinatari,remitent,assumpte,data) values " + vvalors
  wait 1
  dataenviats.Refresh
End Sub
Sub possar_carpetes_Error_a_normal_perprovardereenviaralaproxima()

End Sub

Private Sub Form_Load()
  dataenviats.DatabaseName = rutadelfitxer(cami) + "avisosincidencies.mdb"
  pestanyes.Tab = 0
  
End Sub

Private Sub llistasortida_Click()
   Dim v As String
   Dim vfitxererror As String
   
   v = llistasortida.List(llistasortida.ListIndex)
   If InStr(1, v, "#Error") > 0 Then
      vfitxererror = llistadirs.List(llistasortida.ListIndex) + "\registreemail.txt"
      If existeix(vfitxererror) Then obrir_document vfitxererror
      
   End If
End Sub

Private Sub pestanyes_Click(PreviousTab As Integer)
  'If pestanyes.Tab = 0 Then carregar_pendents
  If pestanyes.Tab = 1 Then dataenviats.Refresh
End Sub
Sub provar_enviaments_ERRORs()
  Dim i As Byte
  Dim v As String
  Dim vfitxererror As String
  If llistasortida.ListCount = 0 Then Exit Sub
  For i = 0 To llistasortida.ListCount - 1
    vfitxererror = llistadirs.List(i)
    If InStr(1, vfitxererror, "#Error#") > 0 Then
         v = vfitxererror
         v = substituir(v, "#Error#", "")
         If existeix(vfitxererror) Then Copiar_Fitxer vfitxererror, v
         If existeix(vfitxererror) Then borra_carpeta vfitxererror
         If Hour(Now) = 0 Then enviaremailgeneric "miquel.inplacsa@gmail.com", "Reenviament email amb error " + atrim(Now), "Carpeta Error: " + vfitxererror + vbNewLine + "Carpeta sense error: " + v
    End If
  Next i
End Sub
Private Sub Timer1_Timer()
  Static vcontador1hora As Integer
  vcontador1hora = vcontador1hora + 1
  If form1.Tag <> "" Then Exit Sub
  If Timer1.Tag = "1" Then Exit Sub
  If etpujantadrive <> "" Then Exit Sub
  Timer1.Tag = "1"
  listlog.Clear
  listlog.AddItem "Buscant pendents... " + Format(Now, "hh:nn:ss")
  'wait 1
  DoEvents
  carregar_pendents
  If checknoenviar.Value <> 0 Then GoTo fi
  
  If llistasortida.ListCount > 0 Then
       enviar_pendents
       carregar_pendents
  End If
  If vcontador1hora >= 300 Then provar_enviaments_ERRORs: vcontador1hora = 0
fi:
  listlog.AddItem "Procés acabat..."
  Timer1.Tag = ""
End Sub
