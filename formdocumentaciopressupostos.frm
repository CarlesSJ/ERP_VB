VERSION 5.00
Begin VB.Form formdocumentaciopressupostos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Escanejar documentació dels pressupostos."
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12255
   Icon            =   "formdocumentaciopressupostos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framecanvinom 
      Caption         =   "Canvi de nom del pressupost escanejat"
      Height          =   5550
      Left            =   465
      TabIndex        =   14
      Top             =   90
      Visible         =   0   'False
      Width           =   11970
      Begin VB.CommandButton bbuscarclient 
         Height          =   360
         Left            =   8580
         Picture         =   "formdocumentaciopressupostos.frx":048A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Buscar el client per relacionar-lo amb el pressupost"
         Top             =   960
         Width           =   585
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cancelar"
         Height          =   1170
         Left            =   9810
         Picture         =   "formdocumentaciopressupostos.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4200
         Width           =   1890
      End
      Begin VB.ComboBox chihamostra 
         Height          =   315
         ItemData        =   "formdocumentaciopressupostos.frx":12B9
         Left            =   1350
         List            =   "formdocumentaciopressupostos.frx":12C3
         TabIndex        =   9
         Top             =   4560
         Width           =   900
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   9225
         Top             =   330
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Acceptar"
         Height          =   1170
         Left            =   9840
         Picture         =   "formdocumentaciopressupostos.frx":12CF
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1260
         Width           =   1890
      End
      Begin VB.TextBox cnumpressupost 
         Height          =   330
         Left            =   1335
         TabIndex        =   8
         Top             =   4125
         Width           =   1995
      End
      Begin VB.ComboBox cimpres 
         Height          =   315
         ItemData        =   "formdocumentaciopressupostos.frx":15A5
         Left            =   1365
         List            =   "formdocumentaciopressupostos.frx":15AF
         TabIndex        =   7
         Top             =   3660
         Width           =   1695
      End
      Begin VB.TextBox cproducte 
         Height          =   330
         Left            =   1350
         TabIndex        =   3
         Top             =   1560
         Width           =   795
      End
      Begin VB.TextBox cmat3 
         Height          =   330
         Left            =   1365
         TabIndex        =   6
         Top             =   3180
         Width           =   7170
      End
      Begin VB.TextBox cmat2 
         Height          =   330
         Left            =   1365
         TabIndex        =   5
         Top             =   2640
         Width           =   7170
      End
      Begin VB.TextBox cmat1 
         Height          =   330
         Left            =   1380
         TabIndex        =   4
         Top             =   2130
         Width           =   7170
      End
      Begin VB.TextBox cdata 
         BackColor       =   &H005C31DD&
         Height          =   330
         Left            =   1395
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox cnomclient 
         Height          =   330
         Left            =   1380
         TabIndex        =   2
         Top             =   990
         Width           =   7170
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Hi ha mostra:"
         Height          =   300
         Left            =   90
         TabIndex        =   24
         Top             =   4605
         Width           =   1245
      End
      Begin VB.Label etnomfitxer 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005C31DD&
         Height          =   390
         Left            =   135
         TabIndex        =   23
         Top             =   5010
         Width           =   7725
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº  pressupost:"
         Height          =   300
         Left            =   105
         TabIndex        =   22
         Top             =   4185
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Imprès/Anònim:"
         Height          =   300
         Left            =   105
         TabIndex        =   21
         Top             =   3705
         Width           =   1245
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Producte:"
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1245
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Material 3:"
         Height          =   300
         Left            =   135
         TabIndex        =   19
         Top             =   3240
         Width           =   1245
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Material 2:"
         Height          =   300
         Left            =   135
         TabIndex        =   18
         Top             =   2700
         Width           =   1245
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Material 1:"
         Height          =   300
         Left            =   150
         TabIndex        =   17
         Top             =   2190
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data pressupost:"
         Height          =   300
         Left            =   165
         TabIndex        =   16
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom del Client:"
         Height          =   300
         Left            =   150
         TabIndex        =   15
         Top             =   990
         Width           =   1245
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Eliminar "
      Height          =   1050
      Left            =   10140
      Picture         =   "formdocumentaciopressupostos.frx":15C3
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1740
      Width           =   1740
   End
   Begin VB.ListBox llista 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4905
      Left            =   120
      TabIndex        =   12
      Top             =   225
      Width           =   9810
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   10755
      Top             =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Passar-ho al drive"
      Height          =   1005
      Left            =   10095
      Picture         =   "formdocumentaciopressupostos.frx":1ACD
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4455
      Width           =   1740
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Canviar el nom"
      Height          =   1050
      Left            =   10140
      Picture         =   "formdocumentaciopressupostos.frx":2B97
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   420
      Width           =   1740
   End
End
Attribute VB_Name = "formdocumentaciopressupostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vrutafitxers As String

Private Sub bbuscarclient_Click()
  Dim vcodiclient As Double
  Dim vnomclient As String
  triar_client vcodiclient, vnomclient
  If vcodiclient > 0 Then cnomclient = atrim(vcodiclient) + "-" + vnomclient
End Sub
Sub triar_client(vcodiclient As Double, vnomclient As String)
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,nom  from clients"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 7200
  formseleccio.Width = 9000
  formseleccio.Left = (Screen.Width / 2) - (formseleccio.Width / 2)
  formseleccio.Top = (Screen.Height / 2) - (formseleccio.Height / 2)
  formseleccio.Show 1
  formseleccio.Left = (Screen.Width / 2) - (formseleccio.Width / 2)
  formseleccio.Top = (Screen.Height / 2) - (formseleccio.Height / 2)
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           vnomclient = formseleccio.DBGrid2.Columns("nom")
           vcodiclient = cadbl(formseleccio.DBGrid2.Columns("codi"))
        End If
   End If
    If seleccioret = 9 Then
        vnomclient = ""
        vcodiclient = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub cdata_LostFocus()
    cdata = substituirtots(cdata, ".", "/")
    If Not IsDate(cdata) Then MsgBox "Error, aquesta data no es vàlida.", vbCritical, "Error": cdata = "": Exit Sub
    generar_nomfitxer
End Sub

Private Sub chihamostra_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub chihamostra_LostFocus()
generar_nomfitxer
End Sub

Private Sub cimpres_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cimpres_LostFocus()
generar_nomfitxer
End Sub

Private Sub cmat1_LostFocus()
generar_nomfitxer
End Sub

Private Sub cmat2_LostFocus()
generar_nomfitxer
End Sub

Private Sub cmat3_LostFocus()
generar_nomfitxer
End Sub

Private Sub cnomclient_LostFocus()
generar_nomfitxer
End Sub

Private Sub cnumpressupost_LostFocus()
  generar_nomfitxer
End Sub
Sub generar_nomfitxer()
    Dim vnomfitxer As String
    Dim vfaltaalgun As Boolean
    
    If atrim(cnomclient) <> "" Then
         cnomclient.BackColor = &H80FF80  'verd
         vnomfitxer = atrim(cnomclient) + " _PRCLESC_ "
          Else: vfaltaalgun = True: cnomclient.BackColor = &H5C31DD 'vermell
    End If
    
    If IsDate(cdata) Then
         cdata.BackColor = &H80FF80  'verd
         vnomfitxer = vnomfitxer + Format(cdata, "yy_mm_dd") + " "
          Else: vfaltaalgun = True: cdata.BackColor = &H5C31DD 'vermell
    End If
    
    If atrim(cproducte) <> "" Then
         cproducte.BackColor = &H80FF80  'verd
         vnomfitxer = vnomfitxer + UCase(cproducte) + " "
          Else: vfaltaalgun = True: cproducte.BackColor = &H5C31DD 'vermell
    End If
    
    If atrim(cmat1) <> "" Then
         cmat1.BackColor = &H80FF80  'verd
         vnomfitxer = vnomfitxer + cmat1 + " "
          Else: vfaltaalgun = True: cmat1.BackColor = &H5C31DD 'vermell
    End If
    
    If atrim(cmat2) <> "" Then
         cmat2.BackColor = &H80FF80  'verd
         vnomfitxer = vnomfitxer + "+ " + cmat2 + " "
          Else: vfaltaalgun = True: cmat2.BackColor = &H5C31DD 'vermell
    End If
    
    If atrim(cmat3) <> "" Then
         cmat3.BackColor = &H80FF80  'verd
         vnomfitxer = vnomfitxer + "+ " + cmat3 + " "
          Else: vfaltaalgun = True: cmat3.BackColor = &H5C31DD 'vermell
    End If
    
    If cimpres <> "" Then
         cimpres.BackColor = &H80FF80  'verd
         vnomfitxer = vnomfitxer + cimpres + " "
          Else: vfaltaalgun = True: cimpres.BackColor = &H5C31DD 'vermell
    End If
    
    If cnumpressupost <> "" Then
         cnumpressupost.BackColor = &H80FF80  'verd
         vnomfitxer = vnomfitxer + cnumpressupost + " "
          Else: vfaltaalgun = True: cnumpressupost.BackColor = &H5C31DD 'vermell
    End If
    If chihamostra <> "" Then
         chihamostra.BackColor = &H80FF80  'verd
         vnomfitxer = vnomfitxer + "MOSTRA_" + UCase(chihamostra) + " "
          Else: vfaltaalgun = True: chihamostra.BackColor = &H5C31DD 'vermell
    End If
    
    If vfaltaalgun Then
       etnomfitxer.Tag = ""
       Command4.Enabled = False
         Else: etnomfitxer.Tag = "OK": Command4.Enabled = True
    End If
    etnomfitxer = vnomfitxer
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_LostFocus()
generar_nomfitxer
End Sub

Private Sub Command1_Click()
  Dim vnom As String
  Dim vnomclient As String
  If framecanvinom.Visible Then framecanvinom.Visible = False: Exit Sub
  If llista.ListIndex = -1 Then
    If llista.ListCount > 0 Then llista.ListIndex = 0
  End If
  refrescar_llistafitxers
  If Mid(llista.Text + "     ", 1, 3) = "***" Then MsgBox "Aquest fitxer està obert espera tanca'l i torna-ho a provar.", vbCritical, "Error": Exit Sub
  If llista.Text = "" Then Exit Sub
  framecanvinom.Left = 60
  framecanvinom.Top = 60
  framecanvinom.Visible = True
  Timer2.Enabled = True
  borrarcamps
  cdata.SetFocus
  
End Sub
Sub borrarcamps()
   cdata = ""
   cnomclient = ""
   cproducte = ""
   cmat1 = ""
   cmat2 = ""
   cmat3 = ""
   cimpres = ""
   cnumpressupost = ""
   
   generar_nomfitxer
End Sub
Sub refrescar_llistafitxers()
  ' File1.path = ""
  ' File1.Refresh
  ' File1.path = vrutafitxers
  ' File1.Pattern = "*.pdf"
  ' File1.Refresh
  Static jasocdins As Boolean
  Dim i As Integer
  Dim j As Integer
  Dim v As String
  Dim vindex As String
  If jasocdins Then Exit Sub
  jasocdins = True
  llista.Tag = ""
  If llista.ListIndex > -1 Then vindex = llista.Text
  v = Dir(vrutafitxers + "\*.pdf")
  i = 0
  
  While v <> ""
     If v <> "." And v <> ".." Then
        If llista.List(i) <> "" Then llista.RemoveItem i
        If ArchivoEnUso(vrutafitxers + "\" + v) Then v = "***" + v + " (Fitxer obert)": llista.Tag = "*"
        llista.AddItem v, i
        If vindex = v Then llista.ListIndex = i
     End If
     v = Dir
     i = i + 1
  Wend
  For j = i To llista.ListCount
   If j < llista.ListCount Then llista.RemoveItem j
  Next j
  jasocdins = False
  
End Sub

Private Sub Command2_Click()
   Dim vruta As String
   On Error GoTo errorgeneral
   If llista.ListCount = 0 Then Exit Sub
   If UCase(llegir_ini("General", "pujantadrive", vrutafitxers + "\cache\organitzar.ini")) = "SI" Then MsgBox "Traspasant fitxers al drive espera uns segons i torna-ho a provar", vbInformation + "Atenció": Exit Sub
   If llista.Tag = "*" Then MsgBox "Hi ha algun fitxer obert i no podré copiar-los primer tancal's i despres passa-ho a DRIVE.", vbCritical, "Error": Exit Sub
   escriure_ini "General", "pujantadrive", "si", vrutafitxers + "\cache\organitzar.ini"
   vruta = crear_ruta_pressupostosescanejats
   If Not existeix(vruta) Then Exit Sub
   Copiar_Fitxer vrutafitxers + "\*.pdf", vrutafitxers + "\cache\" + atrim(Year(Now)) + " Presupuestos"
   Kill vrutafitxers + "\*.pdf"
   escriure_ini "General", "pujantadrive", "no", vrutafitxers + "\cache\organitzar.ini"
   refrescar_llistafitxers
   Exit Sub
errorgeneral:
   escriure_ini "General", "pujantadrive", "no", vrutafitxers + "\cache\organitzar.ini"
    MsgBox "Hi ha hagut un error al copiar els fitxers." + Chr(10) + "O POTSER TENS OBERT ALGUN PDF DE LA LLISTA.", vbCritical, "Error"
    On Error GoTo 0
End Sub
Public Function ArchivoEnUso(ByVal sFileName As String) As Boolean
    Dim filenum As Integer, errnum As Integer
    
    
    On Error Resume Next ' Turn error checking off.
    filenum = FreeFile() ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open sFileName For Input Lock Read As #filenum
    Close filenum ' Close the file.
    errnum = err ' Save the error number that occurred.
    On Error GoTo 0 ' Turn error checking back on.
    
    ' Check to see which error occurred.
    Select Case errnum
    
    ' No error occurred.
    ' File is NOT already open by another user.
    Case 0
    ArchivoEnUso = False
    
    ' Error number for «Permission Denied.»
    ' File is already opened by another user.
    Case 70
    ArchivoEnUso = True
    
    ' Another error occurred.
    Case Else
    Error errnum
End Select
End Function
Function crear_ruta_pressupostosescanejats() As String
   Dim ruta_documentacio_pressupostos As String
   ruta_documentacio_pressupostos = llegir_ini("ruta", "ruta_documentacio_pressupostos", rutadelfitxer(cami) + "valorsprograma.ini")
   ruta_documentacio_pressupostos = ruta_documentacio_pressupostos + "\Escanejats"
   On Error Resume Next
   MkDir ruta_documentacio_pressupostos
   r = ruta_documentacio_pressupostos + "\" + Trim(Year(Now)) + Chr(32) + "Presupuestos" + ""
   MkDir r
   crear_ruta_pressupostosescanejats = r
  
   On Error GoTo 0
End Function

Private Sub Command3_Click()
   If llista.Text = "" Then Exit Sub
   If existeix(vrutafitxers + "\" + llista.Text) Then
       If MsgBox("Segur que vols eliminar-lo?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
            Kill vrutafitxers + "\" + llista.Text
            refrescar_llistafitxers
       End If
   End If
   
End Sub

Private Sub Command4_Click()
vnom = vrutafitxers + "\" + llista.Text
  'vnomclient = InputBox("Entra el nom DEL CLIENT d'aquest pressupost." + Chr(10) + "EVITA POSSAR SIMBOLS ESPECIALS \/:*?""""<>| ", "Canvi de nom")
  
 ' If atrim(vnomclient) = "" Then Exit Sub
  'vdata = InputBox("Entra la data d'aquest pressupost. dd/mm/yy" + Chr(10) + " ", "Canvi de nom")
  'vdata = substituirtots(atrim(vdata), ".", "/")
  'If Not IsDate(vdata) Then MsgBox "Aquesta data no es correcte.", vbCritical, "Error": Exit Sub
  'vdata = Format(vdata, "yy_mm_dd")
  'v = vnomclient + " " + vdata
  If etnomfitxer.Tag <> "OK" Then MsgBox "El nom del fitxer no es correcte falta algun camp.", vbCritical, "Error": Exit Sub
  v = etnomfitxer
  If v <> "" Then
     If InStr(1, v, ".pdf") = 0 Then v = v + ".pdf"
     
     If existeix(vrutafitxers + "\" + v) Then MsgBox "Aquest nom de fitxer ja existeix", vbCritical, "Error": Exit Sub
     On Error GoTo errorfitxer
     FileCopy vnom, vrutafitxers + "\" + v
     wait 1
     Kill vnom
  End If
  refrescar_llistafitxers
  Timer2.Enabled = False
  framecanvinom.Visible = False
  Exit Sub
errorfitxer:
  MsgBox "No s'ha pogut canviar el nom assegura que no estigui obert i que el nom que has possat no inclou caràcters no vàlids.", vbCritical, "Error"
  Timer2.Enabled = False
  framecanvinom.Visible = False
End Sub

Private Sub Command5_Click()
  refrescar_llistafitxers
  Timer2.Enabled = False
  framecanvinom.Visible = False
End Sub

Private Sub cproducte_LostFocus()
generar_nomfitxer
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
  vrutafitxers = rutadelfitxer(cami) + "Cache_escanejarpressupostos"
  refrescar_llistafitxers
      
End Sub

Private Sub llista_DblClick()
  obrir_document vrutafitxers + "\" + llista.Text
End Sub

Private Sub Timer1_Timer()
  Dim vnomfitxer As String
  Dim i As Integer
  If llista.ListIndex <> -1 Then
     vnomfitxer = llista.List(llista.ListIndex)
  End If
  refrescar_llistafitxers
 ' For i = 0 To File1.ListCount - 1
 '   If File1.List(i) = vnomfitxer Then File1.ListIndex = i: GoTo fi
 ' Next i
fi:
End Sub

Private Sub Timer2_Timer()
   generar_nomfitxer
End Sub
