VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form formpressupost 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pressupost"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7860
   Icon            =   "formpressupost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9930
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bdesbloquejarsap 
      Height          =   375
      Left            =   5700
      Picture         =   "formpressupost.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Desmarcar per tornar-lo a facturar."
      Top             =   45
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton bborrar 
      Height          =   375
      Left            =   7410
      Picture         =   "formpressupost.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Eliminar el pressupost"
      Top             =   30
      Width           =   435
   End
   Begin VB.CommandButton imprimir 
      Height          =   375
      Left            =   5100
      Picture         =   "formpressupost.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Imprimir el Pressupost"
      Top             =   30
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   5970
      Picture         =   "formpressupost.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Tornar a crear la descripció"
      Top             =   4605
      Width           =   315
   End
   Begin VB.TextBox ctinters 
      BackColor       =   &H00EAD9CE&
      DataField       =   "tinters"
      DataSource      =   "datapressupost"
      Height          =   315
      Left            =   3630
      TabIndex        =   26
      Top             =   4170
      Width           =   360
   End
   Begin VB.TextBox cdesarroll 
      BackColor       =   &H00EAD9CE&
      DataField       =   "desarroll"
      DataSource      =   "datapressupost"
      Height          =   315
      Left            =   2760
      TabIndex        =   24
      Top             =   4170
      Width           =   690
   End
   Begin VB.TextBox ccilindre 
      BackColor       =   &H00EAD9CE&
      DataField       =   "cilindre"
      DataSource      =   "datapressupost"
      Height          =   315
      Left            =   1905
      TabIndex        =   22
      Top             =   4170
      Width           =   690
   End
   Begin VB.TextBox cbandes 
      BackColor       =   &H00EAD9CE&
      DataField       =   "bandes"
      DataSource      =   "datapressupost"
      Height          =   315
      Left            =   1275
      TabIndex        =   20
      Top             =   4170
      Width           =   420
   End
   Begin VB.TextBox cample 
      BackColor       =   &H00EAD9CE&
      DataField       =   "amplelamina"
      DataSource      =   "datapressupost"
      Height          =   315
      Left            =   210
      TabIndex        =   18
      Top             =   4170
      Width           =   690
   End
   Begin VB.Timer timerdrag 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   6345
      Top             =   7470
   End
   Begin VB.TextBox npressupost 
      Alignment       =   2  'Center
      BackColor       =   &H00EAD9CE&
      DataField       =   "numpressupost"
      DataSource      =   "datapressupost"
      Height          =   285
      Left            =   3540
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   105
      Width           =   1560
   End
   Begin MSMAPI.MAPIMessages MiMAPIMessages 
      Left            =   3435
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MiMAPISession 
      Left            =   2745
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   6960
      Picture         =   "formpressupost.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Enviar per mail amb pdf."
      Top             =   30
      Width           =   435
   End
   Begin VB.Data datapressupost 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1110
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   660
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox cidioma 
      DataField       =   "idioma"
      DataSource      =   "datapressupost"
      Height          =   285
      Left            =   3915
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox cpreu 
      BackColor       =   &H00EAD9CE&
      DataField       =   "preu"
      DataSource      =   "datapressupost"
      Height          =   315
      Left            =   6315
      TabIndex        =   4
      Top             =   4590
      Width           =   990
   End
   Begin VB.TextBox cdescripcio 
      BackColor       =   &H00EAD9CE&
      DataField       =   "descripcio"
      DataSource      =   "datapressupost"
      Height          =   720
      Left            =   195
      MaxLength       =   254
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4575
      Width           =   5760
   End
   Begin VB.CommandButton bang 
      BackColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   1395
      Picture         =   "formpressupost.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   -15
      Width           =   660
   End
   Begin VB.CommandButton besp 
      BackColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   750
      Picture         =   "formpressupost.frx":3042
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   -15
      Width           =   660
   End
   Begin VB.CommandButton bcat 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   660
      Left            =   105
      MaskColor       =   &H000000FF&
      Picture         =   "formpressupost.frx":3EDC
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   -15
      Width           =   660
   End
   Begin VB.TextBox cdata 
      BackColor       =   &H00EAD9CE&
      DataField       =   "data"
      DataSource      =   "datapressupost"
      Height          =   285
      Left            =   1965
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1665
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   225
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "formpressupost.frx":4E4E
      Top             =   9420
      Width           =   7545
   End
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "formpressupost.frx":4EDF
      Top             =   5340
      Width           =   7545
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   165
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "formpressupost.frx":5025
      Top             =   2520
      Width           =   7545
   End
   Begin VB.TextBox linia2 
      BackColor       =   &H00EAD9CE&
      DataField       =   "linia2"
      DataSource      =   "datapressupost"
      Height          =   300
      Left            =   3765
      TabIndex        =   2
      Top             =   2025
      Width           =   3795
   End
   Begin VB.TextBox linia1 
      BackColor       =   &H00EAD9CE&
      DataField       =   "linia1"
      DataSource      =   "datapressupost"
      Height          =   300
      Left            =   3765
      TabIndex        =   0
      Top             =   1695
      Width           =   3795
   End
   Begin VB.Image Image3 
      Height          =   1680
      Left            =   -15
      Picture         =   "formpressupost.frx":509D
      Top             =   690
      Visible         =   0   'False
      Width           =   8505
   End
   Begin VB.Label etfacturat 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   225
      Left            =   2115
      TabIndex        =   31
      Top             =   390
      Width           =   5655
   End
   Begin VB.Label msgerrorvalors 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Atenció els valors introduïts són diferents que els del treball, reviseu que tot sigui correcte."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   4050
      TabIndex        =   28
      Top             =   3945
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tinters"
      Height          =   285
      Left            =   3555
      TabIndex        =   27
      Top             =   3915
      Width           =   660
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Desarroll"
      Height          =   285
      Left            =   2775
      TabIndex        =   25
      Top             =   3930
      Width           =   930
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cilindre"
      Height          =   285
      Left            =   1995
      TabIndex        =   23
      Top             =   3930
      Width           =   660
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Bandes"
      Height          =   285
      Left            =   1215
      TabIndex        =   21
      Top             =   3930
      Width           =   705
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ample Lam"
      Height          =   285
      Left            =   165
      TabIndex        =   19
      Top             =   3930
      Width           =   930
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Presupuesto:"
      Height          =   285
      Left            =   2250
      TabIndex        =   17
      Top             =   135
      Width           =   1605
   End
   Begin VB.Label enviat 
      BackStyle       =   0  'Transparent
      Caption         =   "Enviat --->"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5640
      TabIndex        =   15
      Top             =   345
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Euros"
      Height          =   285
      Left            =   7320
      TabIndex        =   12
      Top             =   4650
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   2685
      Left            =   45
      Picture         =   "formpressupost.frx":2BBD9
      Top             =   7395
      Width           =   3135
   End
   Begin VB.Label data 
      BackStyle       =   0  'Transparent
      Caption         =   "Cassà de la Selva,"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   1
      Top             =   1995
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   75
      Picture         =   "formpressupost.frx":47337
      Top             =   720
      Width           =   7845
   End
End
Attribute VB_Name = "formpressupost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bang_Click()
  cidioma = "ang"
End Sub

Private Sub bborrar_Click()
   If UCase(InputBox("Escriu [eliminar] per eliminar aquest pressupost.", "Eliminar")) = "ELIMINAR" Then
     If datapressupost.Recordset.EditMode > 0 Then datapressupost.Recordset.CancelUpdate
     datapressupost.Recordset.Delete
     Unload formpressupost
   End If
End Sub

Private Sub bcat_Click()
  cidioma = "cat"
End Sub

Private Sub bdesbloquejarsap_Click()
   If cadbl(InputBox("Per desbloquejar aquest pressupost escriu el numero de treball.", "Desfacturar pressupost")) = id_treball Then
     dbclixes.Execute "update pressupostos set datafacturacio=null,lotambelqueshafacturat=0 where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio)
     MsgBox "Hauràs de passar els albarans del proveïdor afectats amb aquest pressupost a no facturats per poder tornar a facturar els clixes.", vbInformation, "Atenció"
     Unload formpressupost
   End If
End Sub

Private Sub besp_Click()
  cidioma = "esp"
End Sub

Private Sub cample_Change()
totsplens
End Sub

Private Sub cbandes_Change()
totsplens
End Sub

Private Sub ccilindre_Change()
totsplens
End Sub

Private Sub cdesarroll_Change()
totsplens
End Sub

Private Sub cidioma_Change()
   bcat.BackColor = QBColor(15)
   besp.BackColor = QBColor(15)
   bang.BackColor = QBColor(15)
   
   Select Case cidioma
      Case "cat"
        bcat.BackColor = QBColor(12)
      Case "esp"
        besp.BackColor = QBColor(12)
      Case "ang"
        bang.BackColor = QBColor(12)
   End Select
End Sub

Private Sub Command1_Click()
  If MsgBox("Aixó torna a generar la descripció." + Chr(10) + " ATENCIÓ EL QUE HI HAGI ARA S'ESBORRARÀ", vbCritical + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbNo Then Exit Sub
  cdescripcio = ".-" + formclixes.marcaproducte + " - " + formclixes.liniaproducte + " (" + ctinters + " " + idiomatintes + "):"
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command7_Click()
'   enviarpressupost
   If msgerrorvalors.visible Then MsgBox "Hi ha diferencies amb els valors entrats al treball i els del pressupost, assegura que el que envies es correcte.", vbCritical, "Atenció"
   enviarpressupost_gmail
End Sub
Sub borrarelstemporals()
   On Error Resume Next
   MkDir "c:\temp\clixespressupostos"
   MkDir "c:\temp\clixescomandes"
   Kill "c:\temp\clixespressupostos\clixespressupost*.*"
   Kill "c:\temp\clixespressupostos\carta*.*"
   Kill "c:\temp\clixescomandes\clixescomandaproveidor*.*"
End Sub
Sub enviarpressupost_gmail()
Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim fitxerpdftemporal As String
  Dim email As String
  Dim cosmissatge As String
  If datapressupost.Recordset.EditMode > 0 Then datapressupost.Recordset.Update
  ratoli "espera"
  wait 4
  borrarelstemporals
  fitxerpdftemporal = "c:\temp\clixespressupostos\clixespressupost_" + atrim(datapressupost.Recordset!id_treball) + ".pdf"
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "clixespressupost_" + atrim(datapressupost.Recordset!Idioma) + ".rpt", 1)
  oreport.DiscardSavedData
  oreport.ExportOptions.DiskFileName = fitxerpdftemporal
  oreport.ExportOptions.PDFExportAllPages = True
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  wait 3
  oreport.Database.Tables.Item(1).Location = camiclixes
  oreport.RecordSelectionFormula = "{pressupostos.numpressupost}=" + atrim(datapressupost.Recordset!numpressupost)
  oreport.DiscardSavedData
  'oreport.ExportOptions.DestinationType = crEDTEMailMAPI
  
  oreport.Export False
  
    datapressupost.Recordset.Edit
    datapressupost.Recordset!enviat = True
    datapressupost.Recordset.Update
    copiarelpdfalacarpetadeltreball fitxerpdftemporal, datapressupost.Recordset!numpressupost
    If atrim(datapressupost.Recordset!empresafacturadora) = "P" Then sieselprimerpressupostdeplaselcopiarlacartainformativaiavisar cadbl(datapressupost.Recordset!numpressupost)
  If datapressupost.Recordset.EditMode > 0 Then datapressupost.Recordset.Edit
  Shell "c:\windows\system32\cmd.exe /c start mailto:"
  wait 2
  idp = ShellExecute(Me.hWnd, "Open", "c:\windows\explorer.exe", " " + "c:\temp\clixespressupostos", "", 1)
  
  ratoli "normal"
End Sub
Sub sieselprimerpressupostdeplaselcopiarlacartainformativaiavisar(vnumpressupost As Double)
    Dim vsql As String
    Dim rst As Recordset
    
    vsql = "SELECT Clixes.codiclienttemporal, pressupostos.empresafacturadora, pressupostos.numpressupost, pressupostos.enviat "
    vsql = vsql + " FROM Clixes INNER JOIN pressupostos ON Clixes.id_treball = pressupostos.id_treball "
    'vsql = vsql + " WHERE (((pressupostos.empresafacturadora)='P') AND ((pressupostos.numpressupost)=" + atrim(vnumpressupost) + ") AND ((pressupostos.enviat)=True));"
    vsql = vsql + " WHERE (((pressupostos.numpressupost)=" + atrim(vnumpressupost) + ") );"
    Set rst = dbclixes.OpenRecordset(vsql)
    If rst.EOF Then Exit Sub
    vsql = "SELECT Clixes.codiclienttemporal, pressupostos.empresafacturadora, pressupostos.numpressupost, pressupostos.enviat "
    vsql = vsql + " FROM Clixes INNER JOIN pressupostos ON Clixes.id_treball = pressupostos.id_treball "
    vsql = vsql + " WHERE (((pressupostos.empresafacturadora)='P') AND (clixes.codiclienttemporal=" + atrim(rst!codiclienttemporal) + " AND pressupostos.enviat=True));"
    'Clipboard.Clear
    'Clipboard.SetText vsql
    Set rst = dbclixes.OpenRecordset(vsql)
    If rst.EOF Then
      MsgBox "Aquest client es el primer cop se li passa pressupost amb PLASEL, s'ha copiat la carta de presentació de l'empresa a la carpeta temporal junt amb el pressupost ADJUNTALA SI HO CREUS NECESSARI.", vbInformation, "ATENCIÓ"
      Copiar_Fitxer rutadelfitxer(cami) + "\Carta presentació empresa PLASEL.pdf", "c:\temp\clixespressupostos"
    End If
End Sub
Sub imprimirpressupost()
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  If datapressupost.Recordset.EditMode > 0 Then datapressupost.Recordset.Update
  ratoli "espera"
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "clixespressupost_" + atrim(datapressupost.Recordset!Idioma) + ".rpt", 1)
  oreport.DiscardSavedData
  oreport.Database.Tables.Item(1).Location = camiclixes
  oreport.RecordSelectionFormula = "{pressupostos.numpressupost}=" + atrim(datapressupost.Recordset!numpressupost)
  oreport.DiscardSavedData
  'oreport.ExportOptions.DestinationType = crEDTEMailMAPI
  oreport.PrintOut False, 1

  ratoli "normal"
End Sub

Sub enviarpressupost()
Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim fitxerpdftemporal As String
  Dim email As String
  Dim cosmissatge As String
  If datapressupost.Recordset.EditMode > 0 Then datapressupost.Recordset.Update
  ratoli "espera"
  borrarelstemporals
  fitxerpdftemporal = "c:\temp\clixespressupost_" + atrim(datapressupost.Recordset!id_treball) + ".pdf"
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "clixespressupost_" + atrim(datapressupost.Recordset!Idioma) + ".rpt", 1)
  oreport.DiscardSavedData
  oreport.ExportOptions.DiskFileName = fitxerpdftemporal
  oreport.ExportOptions.PDFExportAllPages = True
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  wait 3
  oreport.Database.Tables.Item(1).Location = camiclixes
  oreport.RecordSelectionFormula = "{pressupostos.numpressupost}=" + atrim(datapressupost.Recordset!numpressupost)
  oreport.DiscardSavedData
  'oreport.ExportOptions.DestinationType = crEDTEMailMAPI
  
  oreport.Export False
  email = " "
  cosmissatge = " " + Chr(10) + " " + Chr(10) + "Dpt. Màrqueting" + Chr(10) + "Telf: +34 972 460 190" + Chr(10) + "e-mail: mkinplacsa@inplacsa.com" + Chr(10) + "web: www.inplacsa.com"
  'cosmissatge = " hola"
  If enviaremail(email, possarasumpte + atrim(datapressupost.Recordset!numpressupost), cosmissatge, fitxerpdftemporal) Then
    datapressupost.Recordset.Edit
    datapressupost.Recordset!enviat = True
    datapressupost.Recordset.Update
    copiarelpdfalacarpetadeltreball fitxerpdftemporal, datapressupost.Recordset!numpressupost
  End If
  If datapressupost.Recordset.EditMode > 0 Then datapressupost.Recordset.Edit
  ratoli "normal"
End Sub
Function possarasumpte() As String
  If cidioma = "esp" Then possarasumpte = "Presupuesto de Clichés  Nº: "
  If cidioma = "cat" Then possarasumpte = "Pressupost de Clixes  Nº: "
  If cidioma = "ang" Then possarasumpte = "Plate Offer  Nº: "
End Function
Sub copiarelpdfalacarpetadeltreball(fitxerpdf As String, nump As Long)
  Dim rutacarpetatreball As String
  Dim fitxerdesti As String
  formclixes.crearruta ruta_documentacio_clixes + "\" + Format(id_treball, "00000")
  rutacarpetatreball = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Arxiu_documentacio_relacionada" + "\v" + atrim(ordremodificacio)
  formclixes.crearruta rutacarpetatreball
  fitxerdesti = rutacarpetatreball + "\clixespressupost_" + atrim(nump) + ".pdf"
  If existeix(fitxerdesti) Then Kill fitxerdesti
  Copiar_Fitxer fitxerpdf, fitxerdesti
End Sub

Function enviaremail_no(sSendTo As String, sSubject As String, sText As String, adjunt As String) As Boolean
   
    On Error GoTo ErrHandler
     If MiMAPISession.SessionID <> 0 Then MiMAPISession.SignOff
     
     MiMAPISession.SignOn
    With MiMAPISession
        If .SessionID = 0 Then
            .DownLoadMail = False
            .LogonUI = True
            .SignOn
            .NewSession = True
            MAPIMessages1.SessionID = .SessionID
        End If
    End With
    MiMAPIMessages.SessionID = MiMAPISession.SessionID
    With MiMAPIMessages
        .Compose
        .RecipAddress = sSendTo
        .AddressResolveUI = True
        
        .ResolveName
        .MsgSubject = sSubject
        .MsgNoteText = sText
        
        
    MiMAPIMessages.AttachmentPathName = adjunt
        
        .Send True
    End With
    enviaremail_no = True
    Exit Function
ErrHandler:
    'MsgBox err.Description
    MsgBox "Error enviant el mail.", vbCritical, "Atenció"
    enviaremail_no = False
End Function

Private Sub cpreu_Change()
  comprovarquelesvalorsestiguinbe
End Sub
Sub comprovarquelesvalorsestiguinbe()
   If Not formclixes.valorspressupostcorrectes(cadbl(cample), cadbl(cbandes), cadbl(ccilindre), cadbl(cdesarroll), cadbl(ctinters)) Then
       msgerrorvalors.visible = True
         Else: msgerrorvalors.visible = False
   End If
End Sub
Private Sub cpreu_GotFocus()
  If cadbl(cample) = 0 Or cadbl(cbandes) = 0 Or cadbl(ccilindre) = 0 Or cadbl(cdesarroll) = 0 Or cadbl(ctinters) = 0 Then
      MsgBox "Abans de possar el preu has d'emplenar els camps d'ample, bandes, cilindre, desarroll i tinters.", vbCritical, "Atenció"
  End If
End Sub

Private Sub ctinters_Change()
  totsplens
End Sub
Sub totsplens()
  If cadbl(cample) = 0 Or cadbl(cbandes) = 0 Or cadbl(ccilindre) = 0 Or cadbl(cdesarroll) = 0 Or cadbl(ctinters) = 0 Then
      DoEvents
        Else: comprovarquelesvalorsestiguinbe
  End If
End Sub
Private Sub datapressupost_Reposition()
    enviat.visible = False
    If Not datapressupost.Recordset.EOF Then
      If datapressupost.Recordset!enviat Then enviat.visible = True: enviat.caption = "Enviat -->"
      If datapressupost.Recordset!comfirmat Then enviat.visible = True: enviat.caption = "Confirmat"
      If datapressupost.Recordset!empresafacturadora = "P" Then Image3.visible = True: Image1.visible = False
      If datapressupost.Recordset!empresafacturadora = "I" Then Image3.visible = False: Image1.visible = True
    End If
End Sub

Private Sub Form_Activate()
  If atrim(datapressupost.Recordset!empresafacturadora) <> atrim(formclixes.modificacions.Recordset!empresafacturadora) Then
      'canviardempresafacturadora
    If datapressupost.Recordset.EditMode = 0 Then datapressupost.Recordset.Edit
    datapressupost.Recordset!empresafacturadora = "I"
    datapressupost.Recordset.Update
  End If
End Sub
Sub canviardempresafacturadora()
    If MsgBox("L'empresa escullida per fer albarans d'aquest clixes no correspon amb l'escullida per fer pressupostos, vols canviar-lo?", vbCritical + vbDefaultButton1 + vbYesNo, "Atenció") = vbYes Then
        datapressupost.Recordset!empresafacturadora = atrim(formclixes.modificacions.Recordset!empresafacturadora)
        datapressupost.Recordset.Update
        Unload formpressupost
        'datapressupost.Recordset.Move 0
        'datapressupost.Recordset.Edit
    End If

End Sub

Private Sub Form_Load()
    datapressupost.DatabaseName = camiclixes
    datapressupost.RecordSource = "select * from pressupostos where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio)
    datapressupost.Refresh
    If datapressupost.Recordset.EOF Then
      datapressupost.Recordset.AddNew
      datapressupost.Recordset!id_treball = id_treball
      datapressupost.Recordset!ordremodificacio = ordremodificacio
      datapressupost.Recordset!numpressupost = proximpressupost
      datapressupost.Recordset!empresafacturadora = atrim(formclixes.modificacions.Recordset!empresafacturadora)
      If atrim(datapressupost.Recordset!empresafacturadora) = "" Then datapressupost.Recordset!empresafacturadora = "I"
      carregarcampsdetallpressupost
      cidioma = "esp"
      cdata = Format(Now, "dd/mm/yyyy")
      linia2 = "Att. Sr. "
      cdescripcio = ".-" + formclixes.marcaproducte + " - " + formclixes.liniaproducte + " (" + ctinters + " " + idiomatintes + "):"
      datapressupost.Recordset.Update
      datapressupost.Recordset.Bookmark = datapressupost.Recordset.LastModified
    End If
    If Not datapressupost.Recordset!enviat Then   ' si no està enviat al client actualitzo els parametres
       carregarcampsdetallpressupost
       cdescripcio = ".-" + formclixes.marcaproducte + " - " + formclixes.liniaproducte + " (" + ctinters + " " + idiomatintes + "):"
    End If
    comprovarquelesvalorsestiguinbe
    If IsDate(datapressupost.Recordset!datafacturacio) Then
      etfacturat = "Albarant dia " + atrim(Format(datapressupost.Recordset!datafacturacio, "dd/mm/yy")) + " amb el lot " + atrim(datapressupost.Recordset!lotambelqueshafacturat)
      bdesbloquejarsap.visible = True
    End If
    datapressupost.Recordset.Edit
End Sub
Sub carregarcampsdetallpressupost()
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
   cbandes = cadbl(formclixes.Text10)
   cample = cadbl(formclixes.camplelamina)
   cdesarroll = cadbl(formclixes.Text5)
   ctinters = cadbl(formclixes.Text19)
   If Not rst.EOF Then ccilindre = cadbl(rst!cilindre)
   Set rst = Nothing
End Sub
Function idiomatintes() As String
   If cidioma = "esp" Then idiomatintes = "Tintas"
   If cidioma = "cat" Then idiomatintes = "Tintes"
   If cidioma = "ang" Then idiomatintes = "Inks"
End Function
Function proximpressupost() As Long
  Dim rst As Recordset
  Set rst = datapressupost.Database.OpenRecordset("select numpressupost from pressupostos order by numpressupost desc")
  If Not rst.EOF Then
    proximpressupost = cadbl(rst!numpressupost) + 1
   Else: proximpressupost = cadbl(Format(Now, "yy") + "0001")
  End If
End Function
Private Sub t1_Click()

End Sub

Private Sub Form_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Not data.GetFormat(1) Then
     If InStr(1, data.Files(1), ".pdf") = 0 Then MsgBox "El fitxer que arrastris ha de ser un PDF.", vbCritical, "Atenció": Exit Sub
    inicidragover = DateAdd("s", 1, inicidragover)
     formclixes.obrirtemporalclixes True
     Copiar_Fitxer data.Files(1), "c:\temp\tmpclixes\dragover.pdf"
     timerdrag.Enabled = True
       Else: inicidragover = 0
  End If
End Sub

Private Sub Form_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
 If inicidragover = 0 Or DateDiff("s", inicidragover, Now) > 10 Then inicidragover = Now
  If data.GetFormat(1) And DateDiff("s", inicidragover, Now) = 2 Then
     inicidragover = DateAdd("s", 1, inicidragover)
     formclixes.obrirtemporalclixes False
     timerdrag.Enabled = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If datapressupost.Recordset.EditMode > 0 Then datapressupost.Recordset.Update
End Sub

Private Sub imprimir_Click()
  imprimirpressupost
End Sub

Private Sub timerdrag_Timer()
   esperar10segonsaveuresientraelfitxer
End Sub

Sub esperar10segonsaveuresientraelfitxer()
   Dim fitxer As String
   Dim nump As Long
   If DateDiff("s", inicidragover, Now) < 12 Then
      fitxer = formclixes.mirarsihihaalgualtemp
         Else: inicidragover = 0: timerdrag.Enabled = False
   End If
   If fitxer <> "" Then
    If MsgBox("Segur que vols guardar l'OK d'aquest pressupost?", vbExclamation + vbYesNo + vbDefaultButton2, "Ok Pressupost") = vbYes Then
      AppActivate Me.caption
      nump = datapressupost.Recordset!numpressupost
      guardarokpressupost "c:\temp\tmpclixes\" + fitxer, nump
      datapressupost.Recordset.Edit
      datapressupost.Recordset!comfirmat = True
      datapressupost.Recordset.Update
      datapressupost.RecordSource = "select * from pressupostos where numpressupost=" + atrim(nump)
      datapressupost.Refresh
    End If
    If existeix("c:\temp\tmpclixes\" + fitxer) Then Kill "c:\temp\tmpclixes\" + fitxer
   End If
End Sub

Sub guardarokpressupost(origen As String, nump As Long)
  Dim rutacarpetatreball As String
  Dim fitxerdesti As String
  rutacarpetatreball = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Arxiu_documentacio_relacionada" + "\v" + atrim(ordremodificacio)
  fitxerdesti = rutacarpetatreball + "\clixespressupost_" + atrim(nump) + ".pdf"
  If existeix(fitxerdesti) Then Kill fitxerdesti
  fitxerdesti = rutacarpetatreball + "\CLIXESPRESSUPOST_" + atrim(nump) + "_OK.pdf"
  If existeix(fitxerdesti) Then Kill fitxerdesti
  Copiar_Fitxer origen, fitxerdesti
End Sub
