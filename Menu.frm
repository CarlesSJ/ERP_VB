VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu de Comandes"
   ClientHeight    =   8520
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   11955
   DrawMode        =   3  'Not Merge Pen
   DrawStyle       =   4  'Dash-Dot-Dot
   FillColor       =   &H00FF0000&
   FillStyle       =   7  'Diagonal Cross
   Icon            =   "Menu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   Picture         =   "Menu.frx":0442
   ScaleHeight     =   8520
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Tag             =   "1"
   Begin VB.CommandButton Command5 
      Height          =   375
      Left            =   1335
      Picture         =   "Menu.frx":0FC4
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Linkar documentació dels pressupostos."
      Top             =   60
      Width           =   780
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5235
      Top             =   780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   660
      Top             =   2715
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   7725
      TabIndex        =   8
      Top             =   630
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   510
      Top             =   1605
   End
   Begin VB.Frame reindexant 
      Caption         =   "Reindexant les carpetes dels clients"
      Height          =   2820
      Left            =   2100
      TabIndex        =   4
      Top             =   1980
      Visible         =   0   'False
      Width           =   5925
      Begin VB.ListBox duplicats 
         Height          =   1230
         ItemData        =   "Menu.frx":144E
         Left            =   120
         List            =   "Menu.frx":1455
         TabIndex        =   7
         Top             =   1380
         Width           =   5745
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   240
         Left            =   5550
         TabIndex        =   6
         Top             =   135
         Width           =   330
      End
      Begin VB.Label etiqueta 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   630
         Left            =   75
         TabIndex        =   5
         Top             =   615
         Width           =   5760
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1635
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   510
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.DirListBox directoris 
      Height          =   315
      Left            =   1290
      TabIndex        =   3
      Top             =   2085
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2625
      Top             =   1500
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   -45
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   -30
      Width           =   12015
      Begin VB.CommandButton Command9 
         Caption         =   "Manteniments"
         Height          =   375
         Left            =   9945
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Escanejar CQ de Lots"
         Top             =   75
         Width           =   1425
      End
      Begin VB.CommandButton Command8 
         Height          =   375
         Left            =   4290
         Picture         =   "Menu.frx":146B
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Escanejar CQ de Lots"
         Top             =   75
         Width           =   1005
      End
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   3240
         Picture         =   "Menu.frx":18B5
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Escanejar CQ de Lots"
         Top             =   75
         Width           =   1005
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   2205
         Picture         =   "Menu.frx":1D02
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Escanejar Albarans Proveidor"
         Top             =   75
         Width           =   1005
      End
      Begin VB.CommandButton Command4 
         Height          =   375
         Left            =   555
         Picture         =   "Menu.frx":2151
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Linkar documentació de les comandes."
         Top             =   75
         Width           =   780
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   75
         Picture         =   "Menu.frx":26DB
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Linkar pressupostos"
         Top             =   75
         Width           =   435
      End
      Begin VB.CommandButton sortirs 
         Height          =   375
         Left            =   11460
         Picture         =   "Menu.frx":2C65
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Sortir del Programa"
         Top             =   75
         Width           =   435
      End
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   11475
      Picture         =   "Menu.frx":31EF
      Top             =   525
      Width           =   300
   End
   Begin VB.Label nomusuari 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Std"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10455
      TabIndex        =   9
      Top             =   1035
      Width           =   1335
   End
   Begin VB.Label hora 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8520
      TabIndex        =   2
      Top             =   555
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Image Image1 
      Height          =   8145
      Left            =   -15
      Picture         =   "Menu.frx":370F
      Stretch         =   -1  'True
      Top             =   465
      Width           =   12000
   End
   Begin VB.Menu entradadatos 
      Caption         =   "Entrada de Dades"
      Begin VB.Menu ta 
         Caption         =   "Taules Auxiliars"
         Begin VB.Menu mpalets 
            Caption         =   "Palets"
            Begin VB.Menu mtipuspalets 
               Caption         =   "Tipus de Palets"
            End
            Begin VB.Menu malcadespalets 
               Caption         =   "Alçades Palets"
            End
            Begin VB.Menu mtipusproteccions 
               Caption         =   "Tipus Proteccions"
            End
            Begin VB.Menu membanonims 
               Caption         =   "Embalatges Anónims"
            End
            Begin VB.Menu mcertqualitat 
               Caption         =   "Certificat qualitat"
            End
            Begin VB.Menu mguardarmostres 
               Caption         =   "Guardar mostres"
            End
            Begin VB.Menu mconosprotectors 
               Caption         =   "Conos Protectors"
            End
            Begin VB.Menu mtipuspaperfrontal 
               Caption         =   "Tipus Paper Frontal"
            End
         End
         Begin VB.Menu smcompres 
            Caption         =   "Compres"
            Begin VB.Menu mpeucomanda 
               Caption         =   "Missatges peu de Comanda"
            End
         End
         Begin VB.Menu madesiusmuntadora 
            Caption         =   "Adhesius Muntadora"
            Begin VB.Menu madhesius 
               Caption         =   "Adhesius"
            End
            Begin VB.Menu mantfamadhesius 
               Caption         =   "Manteniment Famillies "
            End
            Begin VB.Menu mantsubfamiliesadhesius 
               Caption         =   "Manteniment Subfamilies"
            End
         End
         Begin VB.Menu mcilindres 
            Caption         =   "Cilindres Impresores"
         End
         Begin VB.Menu maniloxs 
            Caption         =   "Aniloxs d'Impresores"
         End
         Begin VB.Menu mavariaimpresores 
            Caption         =   "Tipificacions avaria impresores"
         End
         Begin VB.Menu mpeuimprenta 
            Caption         =   "Peu Imprenta i data"
         End
         Begin VB.Menu mbobinesembolicades 
            Caption         =   "Bobines embolicades"
         End
         Begin VB.Menu mtipusetreb 
            Caption         =   "Tipus d'etiquetes Rebobinadora"
         End
         Begin VB.Menu tipussoldadures 
            Caption         =   "Tipus Soldadures"
         End
         Begin VB.Menu transportistes 
            Caption         =   "Transportistes"
         End
         Begin VB.Menu tubbase 
            Caption         =   "Tubs Base"
         End
         Begin VB.Menu mtarifesref 
            Caption         =   "Tarifes per referencia"
         End
         Begin VB.Menu unitats 
            Caption         =   "Unitats Preus"
         End
         Begin VB.Menu mant_operaris 
            Caption         =   "Operaris"
         End
         Begin VB.Menu unitatslineals 
            Caption         =   "Unitats Lineals"
         End
         Begin VB.Menu altesformpag 
            Caption         =   "Formes de Pagament"
         End
         Begin VB.Menu tipusentregues 
            Caption         =   "Tipus Entregues"
         End
         Begin VB.Menu mcoleslam 
            Caption         =   "Coles de Laminadora"
            Begin VB.Menu adhesius 
               Caption         =   "Adhesius"
            End
            Begin VB.Menu mantfamres 
               Caption         =   "Manteniment families de coles"
            End
            Begin VB.Menu mantfamendur 
               Caption         =   "Manteniment subfamilies de coles"
            End
         End
         Begin VB.Menu camises 
            Caption         =   "Camises"
         End
         Begin VB.Menu representants 
            Caption         =   "Representants"
         End
         Begin VB.Menu accessoris 
            Caption         =   "Accessoris"
         End
         Begin VB.Menu maccessorisSol 
            Caption         =   "Accessoris Soldadora"
         End
         Begin VB.Menu mtipusiva 
            Caption         =   "Canviar Tipus d'Iva"
         End
      End
      Begin VB.Menu men_materials 
         Caption         =   "Materials"
         Begin VB.Menu materials 
            Caption         =   "Materials"
         End
         Begin VB.Menu aditius 
            Caption         =   "Aditius"
         End
         Begin VB.Menu colorants 
            Caption         =   "Colorants"
         End
         Begin VB.Menu mtractamentcares 
            Caption         =   "Tractament cares materials"
         End
         Begin VB.Menu msubstancies 
            Caption         =   "Substancies"
         End
         Begin VB.Menu mfam 
            Caption         =   "Families Materials"
            Begin VB.Menu familiesmaterials 
               Caption         =   "Families Materials"
            End
            Begin VB.Menu familiescolorants 
               Caption         =   "Families Colorants"
            End
            Begin VB.Menu famaditius 
               Caption         =   "Families Aditius"
            End
            Begin VB.Menu subfammaterials 
               Caption         =   "SubFamilies Materials"
               Visible         =   0   'False
            End
            Begin VB.Menu subfamcol 
               Caption         =   "SubFamilies Colorants"
               Visible         =   0   'False
            End
            Begin VB.Menu subfaditius 
               Caption         =   "SubFamilies Aditius"
               Visible         =   0   'False
            End
         End
      End
      Begin VB.Menu maquines 
         Caption         =   "Màquines"
      End
      Begin VB.Menu altaproductes 
         Caption         =   "Productes"
      End
      Begin VB.Menu clients 
         Caption         =   "Clients"
      End
      Begin VB.Menu clientseguiment 
         Caption         =   "Clients - Seguiment"
      End
      Begin VB.Menu manproveidors 
         Caption         =   "Proveïdors"
      End
      Begin VB.Menu mantenimentdecalloffs 
         Caption         =   "Manteniment de Call-Offs"
      End
   End
   Begin VB.Menu execcomandes 
      Caption         =   "Comandes"
   End
   Begin VB.Menu baixes 
      Caption         =   "Baixes"
      Begin VB.Menu Baixesmanteniment 
         Caption         =   "Baixes Oficina"
         Index           =   1
      End
      Begin VB.Menu baixesmaquines 
         Caption         =   "Baixes Maquines"
         Begin VB.Menu baixesmaqimpresores 
            Caption         =   "Impresores"
         End
         Begin VB.Menu mordreimpresio 
            Caption         =   "Impresores Ordre d'impresio"
         End
         Begin VB.Menu Baixesmaqlaminadores 
            Caption         =   "Laminadores"
         End
         Begin VB.Menu baixarebobinadora 
            Caption         =   "Rebobinadora"
         End
         Begin VB.Menu mmmuntadora 
            Caption         =   "Muntadora"
            Begin VB.Menu mmuntadora 
               Caption         =   "Baixes muntadora"
            End
            Begin VB.Menu mllistamuntadorapendent 
               Caption         =   "Llista clixes pendents de muntar."
            End
            Begin VB.Menu mbaixesmuntadoraentredates 
               Caption         =   "Control baixes muntadora entre dates"
            End
         End
      End
      Begin VB.Menu mbaixescostos 
         Caption         =   "Baixes Costos"
      End
      Begin VB.Menu ajba 
         Caption         =   "Ajustos Baixes"
         Begin VB.Menu mtm 
            Caption         =   "Manteniment Tolerancies Maquina"
         End
         Begin VB.Menu mantobser 
            Caption         =   "Manteniment Observacions"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mclixes 
      Caption         =   "Clixes Nous"
   End
   Begin VB.Menu mrepasdeclixes 
      Caption         =   "Repàs Clixes"
   End
   Begin VB.Menu m_palets 
      Caption         =   "Palets"
   End
   Begin VB.Menu mcompres 
      Caption         =   "Compres"
   End
   Begin VB.Menu mplanificacio 
      Caption         =   "Planificacio"
   End
   Begin VB.Menu albaransdevenda 
      Caption         =   "Vendes"
      Begin VB.Menu malbarans 
         Caption         =   "Programa Albarans"
      End
      Begin VB.Menu massignartranspaenvio 
         Caption         =   "Assignar transport a envio"
      End
      Begin VB.Menu menviamentipaqueteria 
         Caption         =   "Registre enviaments i paqueteria"
      End
   End
   Begin VB.Menu mtintes 
      Caption         =   "Tintes"
   End
   Begin VB.Menu llistats 
      Caption         =   "Llistats"
      Begin VB.Menu mllenvxrfacturaCSV 
         Caption         =   "Llistat Impost Envasos per factura.(CSV)"
      End
      Begin VB.Menu mllistatimpenvalb 
         Caption         =   "Llistat Impost Envasos per Albarà.(CSV)"
      End
      Begin VB.Menu mconsums 
         Caption         =   "Consums Inplacsa"
         Begin VB.Menu llperfamiliaentredates 
            Caption         =   "Per familia entre dates"
         End
      End
      Begin VB.Menu mconsulsreferenciescli 
         Caption         =   "Consums del client (Referencies)"
      End
      Begin VB.Menu mc 
         Caption         =   "Comandes"
         Begin VB.Menu mrelaciocomandes 
            Caption         =   "Relació de referències gastades per client"
         End
      End
      Begin VB.Menu mrelaciorefpes 
         Caption         =   "Llistat relació Referencia->pes mtr,bob,palet"
      End
      Begin VB.Menu llistatproduccions 
         Caption         =   "Llistat de Produccions"
      End
      Begin VB.Menu mllistatcredit 
         Caption         =   "Llistat credit clients"
         Begin VB.Menu mexpCSVcrèdit 
            Caption         =   "Exportació CSV cronologia Crèdit consumit."
         End
         Begin VB.Menu mllistatcreditEXCEL 
            Caption         =   "NOU LLISTAT CREDIT DE TOTS ELS CLIENTS (EXCEL)"
         End
         Begin VB.Menu mnomesunclient 
            Caption         =   "Només un client"
         End
         Begin VB.Menu mtotselsclients 
            Caption         =   "Tots els clients (filtre)"
         End
      End
      Begin VB.Menu llistatconsumsinstorics 
         Caption         =   "Llistat de Consums Historics"
      End
      Begin VB.Menu llistatdestatcomandescrops 
         Caption         =   "Llistat estat comandes Crop's"
      End
      Begin VB.Menu mllistatcropspartits 
         Caption         =   "Llistat Contractes Partits >=2018"
      End
      Begin VB.Menu llistatestocgeneral 
         Caption         =   "Llistat GENERAL STOCK OVERVIEW"
      End
      Begin VB.Menu mlotstiadhesius 
         Caption         =   "Traçabilitat de Lots a(Tintes,Adhesius,Canutus,Bosses)"
         Begin VB.Menu tintescomandesafectades 
            Caption         =   "Lots tintes-Comandes afectades"
         End
         Begin VB.Menu adhesiuscomandesafectades 
            Caption         =   "Lots adhesius-Comades afectades"
         End
         Begin VB.Menu lotscanutuscomandes 
            Caption         =   "Lots canutus-Comandes afectades"
         End
         Begin VB.Menu bossescomandesafectades 
            Caption         =   "Lots bosses-Comandes afectades"
         End
         Begin VB.Menu llaunesalacomandax 
            Caption         =   "Llaunes utilitzades a la comanda X"
         End
      End
      Begin VB.Menu mllistatrefenestoc 
         Caption         =   "Llistat referencies en estoc"
      End
      Begin VB.Menu Llistatagrupacioardo 
         Caption         =   "Llistat agrupació ARDO (Comanda,PVP)"
      End
   End
   Begin VB.Menu utils 
      Caption         =   "Utils"
      Begin VB.Menu mimpostenvasos 
         Caption         =   "Manteniment impost envasos."
      End
      Begin VB.Menu mrevcq 
         Caption         =   "Revisio Albarans i CQs Qualitat"
      End
      Begin VB.Menu mcontrolprl 
         Caption         =   "Control PRL"
      End
      Begin VB.Menu mrevisarescaneig 
         Caption         =   "Revisar escaneig desde expedicions"
      End
      Begin VB.Menu comandesacabases 
         Caption         =   "Passar Fills de Comandes a acabades"
      End
      Begin VB.Menu indexarcarpetesclients 
         Caption         =   "Indexar Carpetes dels Clients"
      End
      Begin VB.Menu escullirsegonaimpresoracomandes 
         Caption         =   "Escullir segona impresora de comandes"
      End
      Begin VB.Menu mvalorimpostenvasos 
         Caption         =   "Possar el valor Impost Envasos"
         Visible         =   0   'False
      End
      Begin VB.Menu mcomprovarcomandessensetemperatures 
         Caption         =   "Comprovar comandes sense fitxer de temperatures a impresores"
      End
      Begin VB.Menu mborrarcomanda 
         Caption         =   "Forçar borrar comanda "
      End
      Begin VB.Menu mactualitzarsap 
         Caption         =   "Actualitzar SAP"
      End
   End
   Begin VB.Menu sortir 
      Caption         =   "Sortir"
   End
   Begin VB.Menu mclixesnous 
      Caption         =   "Clixes"
      Visible         =   0   'False
   End
   Begin VB.Menu mavisosseccions 
      Caption         =   "Avisos Seccions"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbvendes As Database
Dim dbtintes As Database
Dim cridaraniloxos As Boolean

Private Sub adhesiuscomandesafectades_Click()
Dim numerodelot As String
  Dim db As Database
  Dim db2 As Database
  Dim were As String
  Dim rsttmp2 As Recordset
  Dim rstclient As Recordset
  Dim taulatemp As String
  numerodelot = InputBox("Entra el numero de lot de Cola que vols buscar:", "Lot de Tinta")
  taulatemp = "c:\temporal.mdb"
  ratoli "espera"
  'Me.Caption = "Processant... "
  If Not existeix(taulatemp) Then DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  Set db = OpenDatabase(cami)
  Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
  On Error Resume Next
  Set db2 = OpenDatabase(taulatemp)
  db2.Execute ("drop table llistatlots")
  db2.Execute ("create table llistatlots (comanda double,pantone string,codiclient double,nomclient string)")
  On Error GoTo 0
  Set rsttmp2 = db2.OpenRecordset("llistatlots")
  were = "(lot1='" + numerodelot + "')"
  were = were + " or (lot2='" + numerodelot + "')"
  Set rsttmp = dbbaixes.OpenRecordset("select * from laminadoresadhesius where " + were)
  While Not rsttmp.EOF
    Set rstclient = db.OpenRecordset("select client from comandes where comanda=" + atrim(cadbl(rsttmp!comanda)))
    If Not rstclient.EOF Then
      Set rstclient = db.OpenRecordset("select codi,nom from clients where codi=" + atrim(cadbl(rstclient!client)))
      If Not rstclient.EOF Then
        rsttmp2.AddNew
        rsttmp2!comanda = rsttmp!comanda
        rsttmp2!codiclient = rstclient!codi
        rsttmp2!nomclient = rstclient!nom
        rsttmp2.Update
      End If
    End If
    rsttmp.MoveNext
  Wend
  r = "Comandes afectades pel lot de Cola: " + numerodelot
  llistat.DataFiles(0) = taulatemp
  llistat.WindowState = crptMaximized
  llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "comandesxlot.rpt"
  llistat.Formulas(0) = "nomdelllistat=" + "'" + r + "'"
  llistat.Action = 1
  
  
  ratoli "normal"
  Set db = Nothing
  Set db2 = Nothing
  'Set dbbaixes = Nothing
  Set rsttmp = Nothing
  Set rsttmp2 = Nothing
  Set rstclient = Nothing
End Sub

Private Sub baixarebobinadora_Click()
  If Not existeix("c:\windows\system32\MSCOMM32.OCX") Then
      Copiar_Fitxer "\\serverprodu\dades\progcomandes\aplicacio\instalaciocomandes\mscom*.*", "c:\windows\system32"
  End If
  If Not existeix("c:\windows\system32\foxitreaderocx.ocx") And Not existeix("c:\windows\syswow64\foxitreaderocx.ocx") Then
    Copiar_Fitxer "\\serverprodu\Dades\progcomandes\aplicacio\PDF\ocx\*.*", "c:\windows\system32"
    Copiar_Fitxer "\\serverprodu\Dades\progcomandes\aplicacio\PDF\ocx\*.*", "c:\windows\syswow64"
  End If
  Shell "\\serverprodu\dades\progcomandes\aplicacio\baixesrebobinadora.exe", vbNormalFocus
End Sub

Private Sub baixescostos_Click()
  
End Sub

Private Sub Baixesmanteniment_Click(Index As Integer)
obrir_baixes
End Sub

Private Sub baixesmaqimpresores_Click()
   Shell "\\serverprodu\dades\progcomandes\aplicacio\baixesimpresoramaquina.exe", vbNormalFocus
End Sub

Private Sub Baixesmaqlaminadores_Click()
Shell "\\serverprodu\dades\progcomandes\aplicacio\baixeslaminadora.exe", vbNormalFocus
End Sub

Private Sub bossescomandesafectades_Click()
  Dim numerodelot As String
  Dim db As Database
  Dim db2 As Database
  Dim were As String
  Dim rsttmp2 As Recordset
  Dim rstclient As Recordset
  Dim taulatemp As String
  numerodelot = InputBox("Entra el numero de lot del bosses que vols buscar:", "Lots de Bosses")
  taulatemp = "c:\temporal.mdb"
  ratoli "espera"
  'Me.Caption = "Processant... "
  If Not existeix(taulatemp) Then DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  Set db = OpenDatabase(cami)
  Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
  On Error Resume Next
  Set db2 = OpenDatabase(taulatemp)
  db2.Execute ("drop table llistatlots")
  db2.Execute ("create table llistatlots (comanda double,pantone string,codiclient double,nomclient string)")
  On Error GoTo 0
  Set rsttmp2 = db2.OpenRecordset("llistatlots")
  were = "(trim(comandabosses1)='" + numerodelot + "')"
  were = were + " or (trim(comandabosses2)='" + numerodelot + "')"
  Set rsttmp = dbbaixes.OpenRecordset("select * from rebobinadorestot where " + were)
  While Not rsttmp.EOF
    Set rstclient = db.OpenRecordset("select client from comandes where comanda=" + atrim(cadbl(rsttmp!comanda)))
    If Not rstclient.EOF Then
      Set rstclient = db.OpenRecordset("select codi,nom from clients where codi=" + atrim(cadbl(rstclient!client)))
      If Not rstclient.EOF Then
        rsttmp2.AddNew
        rsttmp2!comanda = rsttmp!comanda
        rsttmp2!codiclient = rstclient!codi
        rsttmp2!nomclient = rstclient!nom
        rsttmp2.Update
      End If
    End If
    rsttmp.MoveNext
  Wend
  r = "Comandes afectades pel lot de bosses Nº: " + numerodelot
  llistat.DataFiles(0) = taulatemp
  llistat.WindowState = crptMaximized
  llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "comandesxlot.rpt"
  llistat.Formulas(0) = "nomdelllistat=" + "'" + r + "'"
  llistat.Action = 1
  
  
  ratoli "normal"
  Set db = Nothing
  Set db2 = Nothing
 ' Set dbbaixes = Nothing
  Set rsttmp = Nothing
  Set rsttmp2 = Nothing
  Set rstclient = Nothing
End Sub

Private Sub clientseguiment_Click()
   formseguimentclients.Show 1
End Sub

Private Sub comandesacabases_Click()
   Dim rstcom As Recordset
   ratoli "espera"
   Set dbtmp = OpenDatabase(cami)
   Set rstcom = dbtmp.OpenRecordset("select proximaseccio,linkcomanda1,linkcomanda2 from comandes where proximaseccio='T' And producte <> 'PC' And producte <> 'PC2'")
   rstcom.MoveLast
   MsgBox rstcom.RecordCount
   While Not rstcom.EOF
    If cadbl(rstcom!linkcomanda1) > 0 Then dbtmp.Execute "update comandes set proximaseccio='T' where comanda=" + atrim(rstcom!linkcomanda1)
    If cadbl(rstcom!linkcomanda2) > 0 Then dbtmp.Execute "update comandes set proximaseccio='T' where comanda=" + atrim(rstcom!linkcomanda2)
    DoEvents
     rstcom.MoveNext
   Wend
   ratoli "normal"
End Sub

Private Sub Command1_Click()
 reindexant.Visible = False
End Sub

Private Sub accessoris_Click()
  Load formaccessoris
  formaccessoris.Caption = "Manteniment d'Accessoris"
  formaccessoris.autonum = "accessoris"
  formaccessoris.Data1.DatabaseName = cami
  formaccessoris.Data1.RecordSource = "select * from accessoris"
  formaccessoris.refrescar
  formaccessoris.Width = 11000
  'formaltarep.DBGrid1.Columns(1).Caption = "T/A/C"
  formaccessoris.DBGrid1.Refresh
  formaccessoris.Show
End Sub

Private Sub adhesius_Click()
  fColesLaminadora.Show 1
  Exit Sub
  
  
  
  'anulat
  
  Load formaltarep
  formaltarep.Caption = "Manteniment de Adhesius"
  formaltarep.Tag = "220"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from adhesius"
  formaltarep.Width = formaltarep.Width * 2
  formaltarep.DBGrid1.Width = formaltarep.DBGrid1.Width * 2
  
  DoEvents
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(2).Button = True
  formaltarep.Show
End Sub

Private Sub aditius_Click()

  Load formaltarep
  formaltarep.Caption = "Manteniment de Aditius"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from aditius"
  formaltarep.Width = formaltarep.Width * 2
  formaltarep.DBGrid1.Width = formaltarep.DBGrid1.Width * 2
  DoEvents
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub altaproductes_Click()
 Load formaltaproductes
  formaltaproductes.Caption = "Manteniment de Productes"
  formaltaproductes.Data1.DatabaseName = cami
  formaltaproductes.Data1.RecordSource = "select * from productes"
  
  formaltaproductes.refrescar
  formaltaproductes.DBGrid1.Refresh
  formaltaproductes.Width = 9500
  
  
  formaltaproductes.Show
End Sub

Private Sub altesformpag_Click()
  Load formaltaformapag
  formaltaformapag.Caption = "Manteniment de Formes de Pagament"
  formaltaformapag.Data1.DatabaseName = cami
  formaltaformapag.Data1.RecordSource = "select * from [formes de pagament]"
  formaltaformapag.refrescar
  formaltaformapag.DBGrid1.Refresh
  formaltaformapag.Show

End Sub

Private Sub camises_Click()
 
 Load formaltarep
  formaltarep.Caption = "Camises de Laminadora"
  formaltarep.autonum = "camises"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from camises"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.Show

End Sub

Private Sub clients_Click()
  formclients.Show
  
End Sub

Private Sub colorants_Click()

  Load formaltarep
  formaltarep.colsbloc = "4"
  formaltarep.Caption = "Manteniment de colorants"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from  colorants"
  formaltarep.Width = formaltarep.Width * 2.2
  formaltarep.DBGrid1.Width = formaltarep.DBGrid1.Width * 2.25
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub Command2_Click()
  Dim db As Database
  Dim rsttmp2 As Recordset
  Set db = OpenDatabase(cami)
  Set rsttmp = db.OpenRecordset("select datacomanda,proximaseccio,linkcomanda1,producte from comandes where proximaseccio<>'T' ")
  While Not rsttmp.EOF
    If Not IsNull(rsttmp!datacomanda) Then
     If Year(rsttmp!datacomanda) < 2007 Then
       rsttmp.Edit: rsttmp!proximaseccio = "T": rsttmp.Update
     End If
    End If
    rsttmp.MoveNext
  Wend
  
  

End Sub

Private Sub cridarcomandes_Click()

End Sub

Private Sub Command3_Click()
     Formarrastrarpressupost.Show
End Sub

Private Sub Command4_Click()
   Formarrastrarcomandes.Show
End Sub

Private Sub Command5_Click()
  formdocumentaciopressupostos.Show 1
End Sub

Private Sub Command6_Click()
   Load formescanejaralbaransproveidor
    
   formescanejaralbaransproveidor.ettipusescaneig.Caption = "Esperant Albarans del Proveïdor..."
   formescanejaralbaransproveidor.Tag = "albarans"
   formescanejaralbaransproveidor.Caption = formescanejaralbaransproveidor.ettipusescaneig.Caption
   formescanejaralbaransproveidor.vcarpeta = "c:\temp\escaneigdocumentacio\"
   formescanejaralbaransproveidor.eliminar_fitxersdelacarpetaescaner
   DoEvents
   formescanejaralbaransproveidor.Show 1
End Sub

Private Sub Command7_Click()
Load formescanejaralbaransproveidor
   
   formescanejaralbaransproveidor.ettipusescaneig.Caption = "Esperant CQ de lots..."
   formescanejaralbaransproveidor.Tag = "certificats"
   formescanejaralbaransproveidor.Caption = formescanejaralbaransproveidor.ettipusescaneig.Caption
   If Not existeix("c:\temp\escaneigdocumentacio\") Then MkDir "c:\temp\escaneigdocumentacio\"
   formescanejaralbaransproveidor.vcarpeta = "c:\temp\escaneigdocumentacio\"
   formescanejaralbaransproveidor.eliminar_fitxersdelacarpetaescaner
   DoEvents
   formescanejaralbaransproveidor.Show
   
 
   
   
End Sub

Private Sub Command8_Click()
   Unload formescanejaralbaransproveidor
   Load formescanejaralbaransproveidor
   formescanejaralbaransproveidor.ettipusescaneig.Caption = "Esperant Albarans SAP o CMR segellats..."
   formescanejaralbaransproveidor.Tag = "albaransSAP"
   formescanejaralbaransproveidor.checktotselsproveidors.Visible = False
   formescanejaralbaransproveidor.bcmr.Visible = True
   formescanejaralbaransproveidor.Caption = formescanejaralbaransproveidor.ettipusescaneig.Caption
   DoEvents
   formescanejaralbaransproveidor.Show 1
End Sub

Private Sub Command9_Click()
  Shell "\\serverprodu\Dades\progcomandes\aplicacio\Manteniments de Fàbrica.exe", vbNormalFocus
End Sub

Private Sub escullirsegonaimpresoracomandes_Click()
  Dim x As Printer
  Dim xx As String
  xx = Printer.DeviceName
  With CommonDialog1
   .ShowPrinter
   For Each x In Printers
      If Printer.DeviceName = x.DeviceName Then escriure_ini "General", "segonaimpresoradecomandes", x.DeviceName, fitxerini
   Next x
  End With
  Establecer_Impresora xx
  
End Sub
Private Function Establecer_Impresora(ByVal NamePrinter As String) As Boolean
On Error GoTo errSub
      
    'Variable de referencia
    Dim obj_Impresora As Object
      
    'Creamos la referencia
    Set obj_Impresora = CreateObject("WScript.Network")
        obj_Impresora.setdefaultprinter NamePrinter
      
    Set obj_Impresora = Nothing
          
        'La función devuelve true y se cambió con éxito
        Establecer_Impresora = True
        'MsgBox "La impresora se cambió correctamente", vbInformation
    Exit Function
      
      
'Error al cambiar la impresora
errSub:
If err.Number = 0 Then Exit Function
   Establecer_Impresora = False
   MsgBox "error: " & err.Number & Chr(13) & "Description: " & err.Description
   On Error GoTo 0
End Function
Private Sub execcomandes_Click()
  formcomandes.Show
  formcomandes.SetFocus
  On Error Resume Next
  AppActivate "Manteniment de Comandes"
  Dim vx As Double
  Dim vy As Double
  
  If vprimercop = False Then
    vx = cadbl(llegir_ini("PosicioFormComandes", "Left", "comandes.ini"))
    vy = cadbl(llegir_ini("PosicioFormComandes", "Top", "comandes.ini"))
    If vx > 0 And vy > 0 Then formcomandes.Left = vx: formcomandes.Top = vy
  End If

  
End Sub
Sub exportarcomanda()
'per exportar posso variable a 1
 
   escriure_ini "General", "exportant", "1", "comandes.ini"
  Menu.Hide
  Load formcomandes
  While formcomandes.Data1.BackColor = QBColor(12)
    DoEvents
  Wend
  formcomandes.exportartotalacomanda cadbl(llegir_ini("baixes", "imprimircomanda", "comandes.ini"))
  
  
  'per exportar posso variable a 1
  escriure_ini "General", "exportant", "0", "comandes.ini"
  End
End Sub
Sub imprimircomanda()
  Menu.Hide
  Load formcomandes
  While formcomandes.Data1.BackColor = QBColor(12)
    DoEvents
  Wend
  formcomandes.imprimirperpantalla cadbl(llegir_ini("baixes", "imprimircomanda", "comandes.ini"))
  While cadbl(llegir_ini("baixes", "imprimircomanda", "comandes.ini")) > 0
    DoEvents
  Wend
  End
End Sub
'Command9_Click
Private Sub famaditius_Click()

  
  Load formaltafamilies
  formaltafamilies.Caption = "Manteniment Families Aditius"
  formaltafamilies.Data1.DatabaseName = cami
  formaltafamilies.subfamilies.DatabaseName = cami
  formaltafamilies.Data1.Tag = "subfamiliesaditius"
  formaltafamilies.Data1.RecordSource = "select * from familiesaditius"
  formaltafamilies.refrescar
  formaltafamilies.DBGrid1.Columns(1).Width = 2500
  formaltafamilies.DBGrid1.Refresh
  formaltafamilies.Show
  
End Sub

Private Sub familiescolorants_Click()

  Load formaltafamilies
  formaltafamilies.Caption = "Manteniment Families Colorants"
  formaltafamilies.Data1.DatabaseName = cami
  formaltafamilies.subfamilies.DatabaseName = cami
  formaltafamilies.Data1.Tag = "subfamiliescolorants"
  formaltafamilies.Data1.RecordSource = "select * from familiescolorants"
  formaltafamilies.refrescar
  formaltafamilies.DBGrid1.Columns(1).Width = 2500
  formaltafamilies.DBGrid1.Refresh
  formaltafamilies.Show
  
End Sub

Private Sub familiesmaterials_Click()
Load formaltafamilies
  formaltafamilies.Caption = "Manteniment Families Materials"
  formaltafamilies.Data1.DatabaseName = cami
  formaltafamilies.subfamilies.DatabaseName = cami
  formaltafamilies.datatolerancies.DatabaseName = cami
  formaltafamilies.Data1.Tag = "subfamiliesmaterials"
  formaltafamilies.Data1.RecordSource = "select * from familiesmaterials"
  formaltafamilies.refrescar
  formaltafamilies.DBGrid1.Columns(1).Width = 2500
  formaltafamilies.DBGrid1.Refresh
  formaltafamilies.Show
End Sub



Sub canvimicres()
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim rst3 As Recordset
  ' funcio nomes per calcul de micres a soldadores NO SERVEIX PER RES MES
  Set rst = dbtmp.OpenRecordset("select * from comandesmesextres where mesuraesp=10 and instr(ruta,'S')>0")
  While Not rst.EOF
    Me.Caption = atrim(rst!comanda): DoEvents
    vespessor = cadbl(rst!espessor)
    If rst!linkcomanda1 > 0 Then
        Set rst2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rst!linkcomanda1))
        If Not rst2.EOF Then
            If cadbl(rst2!mesuraesp) <> 10 Then GoTo proxima
            vespessor = vespessor + cadbl(rst2!espessor)
        End If
    End If
    If rst!linkcomanda2 > 0 Then
        Set rst2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rst!linkcomanda2))
        If Not rst2.EOF Then
            If cadbl(rst2!mesuraesp) <> 10 Then GoTo proxima
            vespessor = vespessor + cadbl(rst2!espessor)
        End If
    End If
    rst.Edit
    rst!espessorsol = vespessor
    rst!unitatespsol = 10
    rst.Update
proxima:
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Private Sub Form_Activate()
Set controlcanviat = Nothing
'canvimicres
If Me.Tag = "1" Then
   DoEvents
   'si es la natalia carrego comprovació d'etiquetes
   If nomusuari = "Usr_JM" Then
     If Not existeix(App.path + "\comprovaretreb.exe") Then
        MsgBox "No trobo el programa d'avis d'etiquetes"
       Else
        Shell App.path + "\comprovaretreb.exe " + cami
     End If
     '+ " '"  + cami + "'", vbMinimizedNoFocus
   End If
  'si es la alicia carrego avisos de modificacions (OCTUBRE 2025 L'ALICIA DIU QUE JA NO HO VOL)
   'If nomusuari = "Usr_A" Then
   ' If Not existeix(App.path + "\avisos.exe") Then
   '     MsgBox "No trobo el programa d'avisos"
   '    Else
   '      Shell App.path + "\avisos.exe " + cami
   ' End If
   'End If
   Me.Tag = ""
   'cridarcomandes_Click
  If cridaraniloxos Then maniloxs_Click
  If cridacomandes Then execcomandes_Click
  If imprimircomandes Then
    imprimircomanda
  End If
  If exportarcomandes Then
    exportarcomanda
  End If
End If

  
End Sub
Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  MsgBox "ara"
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then volssortir
End Sub
Sub volssortir()
r = ""
  If MsgBox("Segur que vols tancar?", 64 + 4, "Tancar Programa") = vbYes Then
    
    tancar_formularis
    End
  End If
  r = "no"
End Sub

Private Sub Form_Load()
'  comença_captura

If App.PrevInstance Then MsgBox "El programa ja està obert.", vbCritical, "Atenció": End
'MsgBox Workspaces(0).UserName '= nomordinador
arguments = ObtenerLíneaComando
fitxerini = "comandes.ini"
If Not existeix("c:\temp") Then MkDir "c:\temp"
escriure_ini "General", "exportant", "0", "comandes.ini"
If atrim(arguments(1)) <> "" And atrim(arguments(1)) <> "-" Then fitxerini = atrim(arguments(1))
On Error Resume Next
  Kill "c:\temporal.mdb"
  DBEngine.CreateDatabase "c:\temporal.mdb", dbLangGeneral, DatabaseTypeEnum.dbVersion30
On Error GoTo 0

 If llegir_ini("General", "parar", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini") = "si" Then MsgBox "Ara no es pot entrar al programa s'està actualitzant, espera 5 MINUTS, Gràcies", vbCritical, "Actualització": End
  If Not existeix("c:\ordprog.ini") And Not imprimircomandes Then assignardecimalipunt
  cami = llegir_ini("General", "cami", fitxerini)
  camiclixes = rutadelfitxer(cami) + "clixesnous.mdb"
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  '"c:\misdoc~1\commandes\comandes.mdb"
  If existeix("c:\ordprog.ini") Then cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
  
  'es una modificacio per si algu no te aquesta clau possar-la, daquí un temps es pot eliminar
  If llegir_ini("General", "ruta_stocks", fitxerini) = "{[}]" Then
      escriure_ini "General", "ruta_stocks", "\\serverprodu\dades\progcomandes\dades\palets.mdb", fitxerini
  End If
  
  hora = Now
  centerscreen Me
  'afegeixo la copia de coses a l'inici
  If Not existeix("c:\windows\totsproductes.ini") Then
     Copiar_Fitxer "\\serverprodu\dades\progcomandes\aplicacio\plantilles\totsproductes.ini", "c:\windows"
  End If
  borrar_temporals_llistats
  nomusuari = llegir_ini("General", "usuari", fitxerini)
  If nomusuari = "{[}]" Or nomusuari = "" Then nomusuari = "Std": escriure_ini "General", "usuari", "Std", fitxerini
  Set dbtmp = OpenDatabase(cami)
  borrarmenusocults
  If arguments(2) = "aniloxos" Then cridaraniloxos = True
  If arguments(2) = "comandes" Then cridacomandes = True
  If arguments(2) = "imprimir" Then
    imprimircomandes = True
  End If
  If arguments(2) = "exportar" Then
    exportarcomandes = True
  End If
' If nomordinador = "ORD_PROGRAMACIÓ" Then formproveidorsqualitat.Show 1
End Sub
Sub borrarmenusocults()
   'If nomusuari <> "Usr_A" And nomusuari <> "Usr_JM" And nomusuari <> "Usr_MR" Then mplanificacio.Visible = False
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
Sub borrar_temporals_llistats()
  r = Dir("c:\temp\*.mdb")
  While r <> ""
    On Error Resume Next
    Kill "c:\temp\" + r
    On Error GoTo 0
    r = Dir
  Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
'acava_captura
'volssortir
 
 ' If r = "no" Then Cancel = True
 tancar_formularis
 End
End Sub
Sub tancar_formularis()
Dim frm As Form
On Local Error Resume Next
For Each frm In Forms
   Unload frm
   frm.Hide
   Set frm = Nothing
Next
Set rsttmp = Nothing
    Set dbtmp = Nothing
    Set dbtmpb = Nothing
    Set bdllistat = Nothing
End Sub

Private Sub Frame1_DblClick()
'   Unload Formdisposiciomaterialscomanda
'   Load Formdisposiciomaterialscomanda
   'Formdisposiciomaterialscomanda.etrefinplacsa.Tag = "01C1178I2429"
'   Formdisposiciomaterialscomanda.etrefinplacsa.Tag = "03C6969I7203"

   'Formdisposiciomaterialscomanda.Show 1
   'Unload Formdisposiciomaterialscomanda
End Sub

Private Sub indexarcarpetesclients_Click()
   Dim contduplicats As Integer
   contduplicats = 1
  reindexant.Visible = True
  etiqueta = "Buscant..."
  DoEvents
  directoris.path = ruta_relativa_docs
  directoris.Refresh
  i = 0
  etiqueta = "Borrant index..."
  DoEvents
  Data1.DatabaseName = cami
  Data1.RecordSource = "carpeta_client"
  Data1.Refresh
  While Not Data1.Recordset.EOF
    Data1.Recordset.Delete
    Data1.Recordset.MoveNext
  Wend
  On Error GoTo duplicat
  While i < directoris.ListCount
    
    r = Mid(directoris.List(i), Len(ruta_relativa_docs) + 2)
    'r = fCreateShellLink("c:\prova", r, directoris.List(i))
    If r <> "" Then
       Data1.Recordset.AddNew
       Data1.Recordset!codiclient = cadbl(Mid(r, 1, 7))
       Data1.Recordset!nomcarpeta = r
       Data1.Recordset.Update
    End If
    i = i + 1
    etiqueta = "Actualitzant... " + atrim(i) + "/" + atrim(directoris.ListCount)
    DoEvents
  Wend
  
  Data1.RecordSource = ""
  Data1.Refresh
  etiqueta = "Ok. Acavat."
  Exit Sub
duplicat:
   If Len(r) > 5 Then duplicats.AddItem r: contduplicats = contduplicats + 1
   Resume Next
End Sub

Private Sub llaunesalacomandax_Click()
   Dim v As String
   Dim rst As Recordset
   Dim vi As Integer
   Dim vnumc As String
   Dim db As Database
   Set db = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb", , True)
   
   vnumc = InputBox("Entra la comanda que vols consultar les Llaunes utilitzades.", "Comanda")
   If cadbl(vnumc) = 0 Then Exit Sub
   ratoli "espera"
   Set rst = db.OpenRecordset("select * from impresorespantones where comanda=" + atrim(vnumc))
   If rst.EOF Then MsgBox "Baixa no trobada.", vbCritical, "Error": Exit Sub
   
   v = "Comanda: " + vnumc + Chr(13) + Chr(10) + Chr(13) + Chr(10)
   For vi = 1 To 8
     vlots = atrim(rst.Fields("lot" + atrim(vi)))
     'v = v + "Tinter " + atrim(vi) + ":  " + atrim(rst.Fields("pantone" + atrim(vi))) + Space(40 - Len(vlots)) + Chr(13) + Chr(10)
     vlots = buscarelslots(vlots)
     v = v + "Tinter " + atrim(vi) + ":  " + atrim(rst.Fields("pantone" + atrim(vi))) + Chr(13) + Chr(10) + "        " + vlots + Chr(13) + Chr(10) + Chr(13) + Chr(10)
     DoEvents
   Next vi
   enganxarvariableanotepad v
   ratoli "normal"
   Set db = Nothing
End Sub
Function buscarelslots(ByVal vlots As String) As String
   Dim v As String
   Dim vllistalots(50) As String
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
   If atrim(vlots) = "" Then Exit Function
   vlots = vlots + "+"
   While vlots <> ""
      v = Mid(vlots, 1, InStr(1, vlots, "+") - 1)
      buscarelslots = buscarelslots + v + "(" + numdelotbaseonumlotdellauna(v) + ") "
      vlots = substituir(vlots, v + "+", "")
   Wend
   'buscarelslots = v
   Set dbtintes = Nothing
End Function
Function numdelotbaseonumlotdellauna(vnumlotbase As String, Optional vnumllaunarelacionada As Boolean, Optional vnumllauna As String) As String
  Dim vultimnumlot As String
  Dim vnumlot As String
  vnumlotbase = UCase(vnumlotbase)
  If Mid(vnumlotbase, 1, 1) = "A" And (Len(vnumlotbase) > 4 And Len(vnumlotbase) < 7) Then
    vultimnumlot = buscarlot0delalluna(vnumlotbase)
    If vultimnumlot = "0" Then vultimnumlot = vnumlotbase + " (Sense Lot s´ha de buscar manualment la llauna)"
    While Mid(vultimnumlot, 1, 1) = "A" And (Len(vultimnumlot) > 4 And Len(vultimnumlot) < 7) And vultimnumlot <> vnumlotbase
     
     vnumlot = buscarlot0delalluna(vultimnumlot)
     If vnumlot <> "0" And vnumlot <> "" Then
          vultimnumlot = vnumlot
            Else: vultimnumlot = "Lot d'Inplacsa"
     End If
    Wend
     numdelotbaseonumlotdellauna = vultimnumlot + IIf(vnumllaunarelacionada, "[" + vnumlotbase + "]", "") '+ " (" + atrim(vnumlotbase) + ")"
     vnumllauna = vnumlotbase
       Else: numdelotbaseonumlotdellauna = vnumlotbase
  End If
End Function
Function buscarlot0delalluna(vnumlotbase As String) As String
   Dim rsthistoria As Recordset
   Dim rst As Recordset
   Dim vhistoria As String
   Dim idsdecarga As String
   vhistoria = vnumlotbase
   'Set rsthistoria = dbtintes.OpenRecordset("SELECT llaunes.id,Llaunes.numllauna, historiallauna.data, historiallauna.id,historiallauna.tipusmoviment,historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes RIGHT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) RIGHT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) INNER JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent where (numllauna='" + atrim(nllauna) + "')  ORDER BY historiallauna.data DESC;")
   Set rsthistoria = dbtintes.OpenRecordset("SELECT Llaunes.id, Llaunes.numllauna, historiallauna.data, historiallauna.id, historiallauna.tipusmoviment, historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) LEFT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) LEFT JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent Where (numllauna='" + atrim(vnumlotbase) + "') ORDER BY historiallauna.data DESC;")
   If rsthistoria.EOF Then vhistoria = "0"
   While Not rsthistoria.EOF
       If rsthistoria!tipusmoviment = "C" And rsthistoria!idcomponent = 0 Then vhistoria = atrim(rsthistoria!numlotbase)
       rsthistoria.MoveNext
   Wend
   buscarlot0delalluna = vhistoria
   Set rsthistoria = Nothing
End Function
Sub borrarfitxer(vfile As String)
   On Error GoTo fi
   Kill vfile
   Exit Sub
fi:
   MsgBox "No es pot eliminar el fitxer temporal, revisa que no tinguis el bloc de notes obert.", vbCritical, "Error"
End Sub
Sub enganxarvariableanotepad(x As String)
   Dim vfile As String
   vfile = "c:\temp\~tmpllaunescomanda.txt"
   If existeix(vfile) Then borrarfitxer vfile
   If existeix(vfile) Then Exit Sub
   Open vfile For Output As #1
   Print #1, x
   Close #1
   Shell "notepad.exe " + vfile, vbNormalFocus
End Sub

Private Sub Llistatagrupacioardo_Click()
  Dim rst As Recordset
  Dim vagrupacio As String
  Dim vnomfitxer As String
  vnomfitxer = "c:\temp\Llistat Agrupació Ardo.csv"
  vagrupacio = UCase(atrim(InputBox("Escriu el numero d'agrupació d'ARDO:", "Llistat de l'agrupació d'ARDO")))
  If vagrupacio = "" Then Exit Sub
  ratoli "espera"
  Set rst = dbtmp.OpenRecordset("Select * from comandes order by comanda")
  rst.FindFirst "numpressupost='" + vagrupacio + "'"
  If rst.NoMatch Then GoTo fi
  If existeix(vnomfitxer) Then borrarfitxer vnomfitxer
  If existeix(vnomfitxer) Then GoTo fi
  Open vnomfitxer For Output As #1
  Print #1, "Comanda;PVP"
  While Not rst.NoMatch 'Not rst!numpressupost = vagrupacio
     Print #1, atrim(rst!comanda) + ";" + atrim(rst!pvp)
     rst.FindNext "numpressupost='" + vagrupacio + "'"
     If rst.EOF Then GoTo cont
  Wend
cont:
  Close #1
  If existeix(vnomfitxer) Then obrir_document vnomfitxer
fi:
  Set rst = Nothing
  ratoli "normal"
End Sub

Private Sub llistatconsumsinstorics_Click()
Load Llistatconsums
  Llistatconsums.Show
End Sub

Private Sub llistatdestatcomandescrops_Click()
   Formllistatestatcomandes.Show
End Sub

Function triarclient() As Double
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select * from clients"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   triarclient = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   'cnomdelclient.Caption = atrim(formseleccio.data1.Recordset!nom)
  End If
  Unload formseleccio
End Function
Private Sub llistatestocgeneral_Click()
   Dim sql As String
   Dim elwhere As String
   Dim fitxertmpestats As String
   Dim v As String
   Dim inici As Date
   Dim fi As Date
   Dim agruparper As String
   Dim vcodiclient As String
   Dim rst As Recordset
   Dim vq As Double
   Dim uq As String
   
   fitxertmpestats = "c:\temp\consultarefinp_tmp.mdb"
   'vcodiclient = cadbl(InputBox("Entra el codi de client que vols consultar.", "Codi client", "7231"))
   vcodiclient = triarclient
   If vcodiclient = 0 Then GoTo fi
   
   'v = InputBox("Entra la data d'inici de la consulta.", "Inici consulta")
   'If Not IsDate(v) Then MsgBox "La data no es correcte.": Exit Sub
   'inici = CVDate(v)
   'v = InputBox("Entra la data de fi de la consulta.", "Fi consulta", Format(Now, "dd/mm/yy"))
   'If Not IsDate(v) Then MsgBox "La data no es correcte.": Exit Sub
   'fi = CVDate(v)
   'agruparper = UCase(InputBox("Entra per quin camp vols agrupar el llistat" + Chr(10) + "(R)ReferenciaClient    (T)NºTreball d'impresio", "Agrupar per...", "R"))
   'If agruparper <> "R" And agruparper <> "T" Then MsgBox "Opcio no vàlida": GoTo fi
   crearfitxertemp fitxertmpestats
  Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificacio.mdb")
  Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
  If existeix(fitxertmpestats) Then
     Set dbconsulta = DBEngine.OpenDatabase(fitxertmpestats)
      Else: GoTo fi
  End If
  ratoli "espera"
  Set rstc = dbtmp.OpenRecordset("select * from comandes where (producte<>'PC' AND producte<>'PC2' and producte<>'PCP') and proximaseccio<>'T' and client=" + atrim(vcodiclient))
  If rstc.EOF Then MsgBox "No hi ha comandes pendents d'aquest client", vbCritical, "Atenció": Exit Sub
  Set rstconsulta = dbconsulta.OpenRecordset("generalstock")
  While Not rstc.EOF
    'aquesta liniaper crear la descripcio material falta tota la sel.leccio
    rstconsulta.AddNew
    rstconsulta!comanda = rstc!comanda
    rstconsulta!datacomanda = rstc!datacomanda
    rstconsulta!comandavisual = atrim(rstc!comanda) + IIf(rstc!linkcomanda1 > 0, "/" + Mid(atrim(rstc!linkcomanda1), 5, 2), "") + IIf(rstc!linkcomanda2 > 0, "/" + Mid(atrim(rstc!linkcomanda2), 5, 2), "")
    rstconsulta!refclient = rstc!refclient
    rstconsulta!linia_marca = rstc!marcailinia
    rstconsulta!estat = IIf(rstc!proximaseccio <> "V", "IN PRODUCTION", "READY TO BE DISPATCHED")
    rstconsulta!dataentregaclient = rstc!datamaterial
    rstconsulta!dataentregainplacsa = buscardataentregaplanificacio(rstc!comanda)
    rstconsulta!amplada = rstc!amplereb * 10
    rstconsulta!desarroll = rstc!dessarroll
    rstconsulta!quantitatdemanada = cadbl(rstc!tubbaseext)
    rstconsulta!tintes = buscartinters(cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio))
    vq = 0
    buscarquantitatentregadaiunitatdentrega cadbl(rstc!comanda), vq, uq, cadbl(rstc!dessarroll)
    rstconsulta!quantitatentregada = vq
    rstconsulta!unitatdequantitat = uq
    rstconsulta!adreçaenviament = buscaradreçadenviament(rstc!comanda)
    possarespesorimaterial rstconsulta, rstc!comanda, rstc!linkcomanda1, rstc!linkcomanda2
    rstconsulta.Update
    rstc.MoveNext
  Wend
 
'  actualitzarcampsrestants
  exportar_llistat_generalstock_xls
fi:
  ratoli "normal"
  Set rstconsulta = Nothing
  Set rsct = Nothing
  Set dbclixes = Nothing
  Set dbplanificacio = Nothing
  'SET DBBAIXES = NOTHING
  Set dbconsulta = Nothing
End Sub
Sub exportar_llistat_generalstock_xls()
   Set rst = dbconsulta.OpenRecordset("select * from generalstock")
   If rst.EOF Then MsgBox "No hi ha dades per exportar", vbCritical, "Error": Exit Sub
   Open "c:\temp\consultageneralstockinplacsa.csv" For Output As #1
   Print #1, ";Order date;Inplacsa's Batch Nº;Reference Nº;Brand/Product Name;Status;Desired Delivery Date;Confirmed Dispatching Date;Width;Repeat;NºInks;Film Thickness;Film Structure;Order Quantity;Delivered Quantity;Unit;Delivey Adress"
   
   While Not rst.EOF
    linia = ""
    For i = 0 To rst.Fields.Count - 1
      If rst.Fields(i).Name <> "comanda" Then
        linia = linia + IIf(linia = "", "", ";") + """" + atrim(rst.Fields(i)) + """"
      End If
    Next i
    Print #1, linia
    rst.MoveNext
   Wend
   Close #1
   wait 2
   obrir_document "c:\temp\consultageneralstockinplacsa.csv"
   Set rst = Nothing
End Sub
Function buscaradreçadenviament(vnumc As Double) As String
    Dim rst As Recordset
    Set rst = dbtmp.OpenRecordset("SELECT comandes.comanda, Clients_envios.nome, Clients_envios.domicilie FROM comandes INNER JOIN Clients_envios ON comandes.direnvio = Clients_envios.id where comanda=" + atrim(vnumc))
    If Not rst.EOF Then buscaradreçadenviament = atrim(rst!nome) + "  (" + atrim(rst!domicilie) + ")"
End Function
Sub buscarquantitatentregadaiunitatdentrega(vnumc As Double, ByRef vquantitatentregada As Double, ByRef vunitatdentrega As String, vdesarroll As Double)
    Dim rst As Recordset
    Dim rstc As Recordset
    Set rst = dbbaixes.OpenRecordset("SELECT rebobinadores.comanda, Sum(bobinesreb.kilos) AS SumaDekilos, Sum(bobinesreb.metres) AS SumaDemetres FROM bobinesreb INNER JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id  where comanda=" + atrim(vnumc) + " GROUP BY rebobinadores.comanda")
    Set rstc = dbtmp.OpenRecordset("SELECT comandes.comanda, mesures.unitatinterna FROM comandes INNER JOIN mesures ON comandes.mesurapvp = mesures.codi where comanda=" + atrim(vnumc))
    If Not rstc.EOF Then vunitatdentrega = Mid(rstc!unitatinterna, 3)
    If Not rst.EOF Then
      If vunitatdentrega = "K" Then vquantitatentregada = cadbl(rst!SumaDekilos)
      If vunitatdentrega = "M" Then vquantitatentregada = cadbl(rst!SumaDemetres)
      If vunitatdentrega = "U" Then vquantitatentregada = Redondejar((cadbl(rst!SumaDemetres) / vdesarroll) * 1000, 0)
    End If
   Set rst = Nothing
End Sub
Function buscartinters(vtreball As Double, vordre As Double) As Double
  Dim rsttmp As Recordset
  Set rsttmp = dbclixes.OpenRecordset("select tinters from modificacions where id_treball=" + atrim(vtreball) + " and ordre=" + atrim(vordre))
  If Not rsttmp.EOF Then buscartinters = cadbl(rsttmp!tinters)
  Set rsttmp = Nothing
End Function
Function buscardataentregaplanificacio(vnumc As Double) As Date
  Dim rsttmp As Recordset
  Set rsttmp = dbplanificacio.OpenRecordset("select * from planificaciototes where comanda=" + atrim(cadbl(vnumc)))
  If Not rsttmp.EOF Then
        If Not IsNull(rsttmp!Data2) Then buscardataentregaplanificacio = atrim(rsttmp!Data2)
        '  Else: buscardataentregaplanificacio = Null
  End If
  Set rsttmp = Nothing
End Function
Function traduir(valor As String, Idioma As String) As String
   Dim rst As Recordset
   traduir = atrim(valor)
   Set rst = dbclixes.OpenRecordset("select * from diccionari where idioma='" + atrim(Idioma) + "' and trim(pertraduir)='" + atrim(valor) + "'")
   If Not rst.EOF Then
      traduir = atrim(rst!traduit)
   End If
End Function
Function descripciomaterialconcatenat(rstmat As Recordset) As String
   Dim c As String
   Dim vnomcolor As String
   c = Mid(atrim(rstmat![familiesmaterials.descripcio]) + " ", 1, InStr(1, atrim(rstmat![familiesmaterials.descripcio]) + " ", " "))
   vnomcolor = atrim(Mid(atrim(rstmat![familiescolorants.descripcio]) + " ", 1, InStr(1, atrim(rstmat![familiescolorants.descripcio]) + " ", " ")))
   'If vnomcolor = "TRANSPARENT" Then vnomcolor = ""
   vnomcolor = traduir(vnomcolor, "EN")
   c = vnomcolor + " " + c
   descripciomaterialconcatenat = c
End Function
Sub possarespesorimaterial(rstnou As Recordset, numc1 As Double, numc2 As Double, numc3 As Double)
    Dim rstmat1 As Recordset
  Dim rstmat2 As Recordset
  Dim rstmat3 As Recordset
  Dim espesormat1 As Double
  Dim espesormat2 As Double
  Dim espesormat3 As Double
  Dim descripciomat As String
  Dim tipusfilm As String
  Dim codimat As String
  Dim rstcomandes As Recordset
  Set rstcomandes = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(numc1) + " or comanda=" + atrim(numc2) + " or comanda=" + atrim(numc3), dbOpenSnapshot, dbReadOnly)
  If rstcomandes.EOF Then Exit Sub
  rstcomandes.FindFirst "comanda=" + atrim(numc1)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat1 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));", dbOpenSnapshot, dbReadOnly)
  rstcomandes.FindFirst "comanda=" + atrim(numc2)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat2 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));", dbOpenSnapshot, dbReadOnly)
  rstcomandes.FindFirst "comanda=" + atrim(numc3)
  codimat = IIf(Not rstcomandes.NoMatch, cadbl(rstcomandes!materialex), 0)
  Set rstmat3 = dbtmp.OpenRecordset("SELECT familiesmaterials.descripcio, familiescolorants.descripcio, subfamiliesmaterials.descripcio FROM ((familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) INNER JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi WHERE (((materials.codi)=" + atrim(codimat) + "));", dbOpenSnapshot, dbReadOnly)
  If Not rstmat1.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc1)
     If Not rstcomandes.NoMatch Then
        descripciomat = descripciomaterialconcatenat(rstmat1)  'atrim(rstmat1![familiesmaterials.descripcio]), atrim(rstmat1![familiescolorants.descripcio]))rstmat1![subfamiliesmaterials.descripcio]
        espesormat1 = formcomandes.micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))
     End If
  End If
  If Not rstmat2.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc2)
     If Not rstcomandes.NoMatch Then
        descripciomat = descripciomat + "+" + descripciomaterialconcatenat(rstmat2)
        espesormat2 = formcomandes.micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))
     End If
  End If
  If Not rstmat3.EOF Then
     rstcomandes.FindFirst "comanda=" + atrim(numc3)
     If Not rstcomandes.NoMatch Then
        descripciomat = descripciomat + " // " + descripciomaterialconcatenat(rstmat3)
        espesormat3 = formcomandes.micresmaterial(cadbl(rstcomandes!mesuraesp), cadbl(rstcomandes!espessor), atrim(rstcomandes!tubolam))
     End If
  End If
  rstnou!micres = atrim(espesormat1) + IIf(cadbl(espesormat2) <> 0, "+" + atrim(espesormat2), "") + IIf(cadbl(espesormat3) <> 0, "+" + atrim(espesormat3), "")
  rstnou!descfamiliamat = descripciomat
  Set rstmat1 = Nothing
  Set rstmat2 = Nothing
  Set rstmat3 = Nothing
  Set rstcomandes = Nothing
End Sub
Sub crearfitxertemp(vnomfitxer As String)
   taula_tmp = "generalstock"
   If existeix(vnomfitxer) Then Kill vnomfitxer
   DBEngine.CreateDatabase vnomfitxer, dbLangGeneral, DatabaseTypeEnum.dbVersion30
   Set dbconsulta = DBEngine.OpenDatabase(vnomfitxer)
   dbconsulta.Execute ("create table " + taula_tmp + " (id long)")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column datacomanda date")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column comanda double")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column comandavisual string")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column refclient string")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column linia_marca string")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column estat string")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column dataentregaclient date")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column dataentregainplacsa date")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column amplada double")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column desarroll double")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column tintes double")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column micres string")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column descfamiliamat string")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column quantitatdemanada double")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column quantitatentregada double")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column unitatdequantitat string")
   dbconsulta.Execute ("alter table " + taula_tmp + " add column adreçaenviament string")
  
   
End Sub

Private Sub llistatproduccions_Click()
  Load Llistatproduccio
  Llistatproduccio.Show
End Sub

Private Sub llperfamiliaentredates_Click()
  Dim vdatai As Date
  Dim vdataf As Date
  Dim vresp As String
  Dim vcodifam As Integer
  Dim vcodisubfam As Integer
  Dim vnomfamilies As String
  vresp = InputBox("Entra la data inici del llistat. [dd/mm/yy]", "Data inici")
  If Not IsDate(vresp) Then MsgBox "Data no vàlida", vbCritical, "Error": Exit Sub
  vdatai = CVDate(vresp)
  vresp = InputBox("Entra la data final del llistat. [dd/mm/yy]", "Data fi")
  If Not IsDate(vresp) Then MsgBox "Data no vàlida", vbCritical, "Error": Exit Sub
  vdataf = CVDate(vresp)
  escullir_familiamat vcodifam, vcodisubfam, vnomfamilies
  If vcodifam = 0 Then MsgBox "No s'ha escullit cap familia per filtrar.", vbCritical, "Error": Exit Sub
  ferllistatmaterialgastatentredates vdatai, vdataf, vcodifam, vcodisubfam, vnomfamilies
End Sub
Sub ferllistatmaterialgastatentredates(vdatai As Date, vdataf As Date, vcodifam As Integer, vcodifubfam As Integer, vnomfamilies As String)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.report
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatmaterialgastatentredates.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "palets.mdb"
  'oreport.RecordSelectionFormula = "{aniloxosinformacio.id} in (SELECT First({aniloxos_informacio.id]) From {aniloxos_informacio} GROUP BY {aniloxos_informacio.matricula}"
  oreport.FormulaFields.GetItemByName("vdatai").Text = "'" + atrim(vdatai) + "'"
  oreport.FormulaFields.GetItemByName("vdataf").Text = "'" + atrim(vdataf) + "'"
  oreport.FormulaFields.GetItemByName("vcodifam").Text = "'" + atrim(vcodifam) + "'"
  oreport.FormulaFields.GetItemByName("vcodisubfam").Text = "'" + atrim(cadbl(vcodifubfam)) + "'"
  oreport.FormulaFields.GetItemByName("vnomfamisubfam").Text = "'" + atrim(vnomfamilies) + "'"
  
  If cadbl(vcodifubfam) > 0 Then
     oreport.RecordSelectionFormula = "{materials.familia}=" + atrim(cadbl(vcodifam)) + " and {materials.subfamilia}=" + atrim(cadbl(vcodisubfam)) + " and ({Parcials.data} >=#" + Format(vdatai, "mm/dd/yy 00:00:00") + "# and {Parcials.data}<=#" + Format(vdataf, "mm/dd/yy 23:59:00") + "#)"
       Else: oreport.RecordSelectionFormula = "{materials.familia}=" + atrim(cadbl(vcodifam)) + " and ({Parcials.data} >=#" + Format(vdatai, "mm/dd/yy 00:00:00") + "# and {Parcials.data}<=#" + Format(vdataf, "mm/dd/yy 23:59:00") + "#)"
  End If
'  oreport.DiscardSavedData
 
  ' MsgBox oreport.RecordSelectionFormula
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
End Sub
Sub escullir_familiamat(vcodifam As Integer, vcodifubfam As Integer, vnomfamilies As String)
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select CODI,DESCRIPCIO from FAMILIESMATERIALS where codi>499"
  formseleccio.refrescar
  
  formseleccio.Show 1
  If seleccioret = 1 Then
   vcodifam = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   vnomfamilies = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  If vcodifam = 0 Then GoTo fi
  Unload formseleccio
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select CODI,DESCRIPCIO from subFAMILIESMATERIALS where codifam=" + atrim(vcodifam)
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   vcodifubfam = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   vnomfamilies = vnomfamilies + " - " + atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio

fi:
End Sub

Private Sub lotscanutuscomandes_Click()
   Dim numerodelot As String
  Dim db As Database
  Dim db2 As Database
  Dim were As String
  Dim rsttmp2 As Recordset
  Dim rstclient As Recordset
  Dim taulatemp As String
  numerodelot = InputBox("Entra el numero de lot del canutu que vols buscar:", "Lots de Canutus")
  taulatemp = "c:\temporal.mdb"
  ratoli "espera"
  'Me.Caption = "Processant... "
  If Not existeix(taulatemp) Then DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  Set db = OpenDatabase(cami)
  Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
  On Error Resume Next
  Set db2 = OpenDatabase(taulatemp)
  db2.Execute ("drop table llistatlots")
  db2.Execute ("create table llistatlots (comanda double,pantone string,codiclient double,nomclient string)")
  On Error GoTo 0
  Set rsttmp2 = db2.OpenRecordset("llistatlots")
  were = "(trim(comandacanutus1)='" + numerodelot + "')"
  were = were + " or (trim(comandacanutus2)='" + numerodelot + "')"
  Set rsttmp = dbbaixes.OpenRecordset("select * from rebobinadorestot where " + were)
  While Not rsttmp.EOF
    Set rstclient = db.OpenRecordset("select client from comandes where comanda=" + atrim(cadbl(rsttmp!comanda)))
    If Not rstclient.EOF Then
      Set rstclient = db.OpenRecordset("select codi,nom from clients where codi=" + atrim(cadbl(rstclient!client)))
      If Not rstclient.EOF Then
        rsttmp2.AddNew
        rsttmp2!comanda = rsttmp!comanda
        rsttmp2!codiclient = rstclient!codi
        rsttmp2!nomclient = rstclient!nom
        rsttmp2.Update
      End If
    End If
    rsttmp.MoveNext
  Wend
  r = "Comandes afectades pel lot de canutus Nº: " + numerodelot
  llistat.DataFiles(0) = taulatemp
  llistat.WindowState = crptMaximized
  llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "comandesxlot.rpt"
  llistat.Formulas(0) = "nomdelllistat=" + "'" + r + "'"
  llistat.Action = 1
  
  
  ratoli "normal"
  Set db = Nothing
  Set db2 = Nothing
  'SET DBBAIXES = NOTHING
  Set rsttmp = Nothing
  Set rsttmp2 = Nothing
  Set rstclient = Nothing
End Sub

Private Sub m_palets_Click()
  Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe", vbNormalFocus
End Sub

Private Sub maccessorisSol_Click()
   FormAccessorisSoldadora.Show 1
End Sub

Private Sub mactualitzarsap_Click()
  Dim horaentrada As Date
  Dim sincronitzant As Boolean
  horaentrada = Now
  escriure_ini "General", "sincronitzarsap", "Si", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini"
  escriure_ini "General", "sincronitzarsapusuari", Environ("computername"), llegir_ini("General", "rutallistats", fitxerini) + "parar.ini"
  sincronitzant = True
  While DateDiff("s", horaentrada, Now) < 30 And sincronitzant
      ratoli "espera"
      sincronitzant = IIf(llegir_ini("General", "sincronitzarsap", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini") = "Si", True, False)
  Wend
  ratoli "normal"
  escriure_ini "General", "sincronitzarsapusuari", "", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini"
  If sincronitzant Then
     MsgBox "Hi ha hagut algun error amb el Dimoni de sincronització assegura que al servidor hi ha el programa engegat, Gràcies", vbCritical, "Error"
       Else: MsgBox "Proces acabat"
  End If
  
End Sub

Private Sub madhesius_Click()
   formadhesiusmuntadora.Show
End Sub

Private Sub malbarans_Click()
    Shell "\\serverprodu\dades\progcomandes\aplicacio\vendes.exe", vbNormalFocus
End Sub

Private Sub malcadespalets_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment d'Alçades Palets"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from alcadespalets"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Show
End Sub

Private Sub maniloxs_Click()
   formaniloximpresores.Show 1
End Sub

Private Sub manproveidors_Click()
   proveidorsproduccio.Show
  
End Sub

Private Sub mant_operaris_Click()

  Load formaltamaquines
  formaltamaquines.Caption = "Manteniment d'operaris"
  formaltamaquines.Data1.DatabaseName = cami
  formaltamaquines.Data1.RecordSource = "select maquina,codi,descripcio,actiu from Operaris order by maquina,codi"
  formaltamaquines.DBGrid1.Tag = "select maquina,codi,descripcio,actiu from Operaris order by maquina,codi"
  formaltamaquines.refrescar
  formaltamaquines.bcontrasenya.Visible = True
  formaltamaquines.bextres.Visible = False
  formaltamaquines.DBGrid1.Refresh
  formaltamaquines.Show
End Sub

Private Sub mantenimentsdefabrica_Click()
   Shell "\\serverprodu\dades\progcomandes\aplicacio\manteniments fàbrica.exe", vbNormalFocus
End Sub

Private Sub mantenimentdecalloffs_Click()
   formcalloff.Show 1
End Sub

Private Sub mantfamadhesius_Click()
Load formaltarep
  formaltarep.Caption = "Mantenimet Families Adhesius Muntadora."
'  formaltarep.autonum = "transportistes"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from famadhesiusmunt"

  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Width = formaltarep.Width + 700
  formaltarep.DBGrid1.Width = formaltarep.DBGrid1.Width + 700
  formaltarep.Show
End Sub

Private Sub mantfamendur_Click()
  Load formaltarep
  formaltarep.Caption = "Manteniment subfamilia de coles"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from subfamiliescoles"
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.DBGrid1.Columns(1).Width = 150 * 25
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub mantfamres_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment familia de coles"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from familiescoles"
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.DBGrid1.Columns(1).Width = 150 * 25
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub mantobser_Click(Index As Integer)
Load formaltarep
  formaltarep.Caption = "Manteniment Observacions de Baixes"
  'formaltarep.colsbloc = "2"
  formaltarep.Data1.DatabaseName = llegir_ini("General", "camibaixes", fitxerini)
  formaltarep.Data1.RecordSource = "select observacio from constantsobservacio"
  formaltarep.Width = (formaltarep.Width * 2) + 200
  formaltarep.DBGrid1.Width = (formaltarep.DBGrid1.Width * 2) + 400
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub mantsubfamiliesadhesius_Click()
Load formaltarep
  formaltarep.Caption = "Mantenimet SUBFamilies Adhesius Muntadora."
'  formaltarep.autonum = "transportistes"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from subfamadhesiusmunt"

  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Width = formaltarep.Width + 700
  formaltarep.DBGrid1.Width = formaltarep.DBGrid1.Width + 700
  formaltarep.Show
End Sub

Private Sub maquines_Click()

  Load formaltamaquines
  formaltamaquines.Caption = "Manteniment de Maquines"
  formaltamaquines.Data1.DatabaseName = cami
  formaltamaquines.Data1.RecordSource = "select * from maquines order by maquina"
  formaltamaquines.DBGrid1.Tag = "select * from maquines "
  formaltamaquines.refrescar
  formaltamaquines.DBGrid1.Refresh
  formaltamaquines.Show
End Sub

Private Sub massignartranspaenvio_Click()
   Dim vnumc As String
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim rstc2 As Recordset
   Dim rste As Recordset
   Dim vkgcomandes As Double
   Dim vcomandes As String
   Dim vcalloff As String
   Dim vcrops As Boolean
   Dim vnomclient As String
   Dim vcomandesdelclient As String
   Dim vnomtransport As String
   Dim vvalues As String
   Dim vpalets As Integer
   Dim veuros As Double
   Dim v As String
   Dim vpreusuggerit As Double
   Dim vidtransport As Long
   Dim vmsg As String
   Dim videnvio As Long
   Dim vhihaunparcial As Boolean
   Dim vcomandesafectades As String
   
   
   Set dbvendes = OpenDatabase(rutadelfitxer(cami) + "vendes.mdb")
   Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificacio.mdb")
   
   vnumc = InputBox("Escriu la comanda que vols assignar-hi transport.", "Assignar transport")
   If cadbl(vnumc) = 0 Then Exit Sub
   Set rstc = dbtmp.OpenRecordset("SELECT comandes.comanda, comandes.client,comandes.direnvio,clients_envios.codipostale,clients_envios.pais,clients_envios.nome,Clients_envios.id, trim(clients_envios.domicilie)&' '&trim(clients_envios.codipostale)&' '&Clients_envios.poblacioe+' ['+clients_envios.provinciae+']' AS direccioenviament, transportistes.descripcio FROM (comandes_extres LEFT JOIN transportistes ON comandes_extres.transportista_albara = transportistes.codi) RIGHT JOIN (comandes LEFT JOIN Clients_envios ON comandes.direnvio = Clients_envios.id) ON comandes_extres.comanda = comandes.comanda where comandes.comanda = " + atrim(vnumc))
   If rstc.EOF Then MsgBox "Comanda no trobada.", vbCritical, "Error": GoTo fi
   videnvio = rstc!direnvio
   vnomclient = rstc!nome
   If atrim(rstc!descripcio) <> "" Then   'si ja està assignat el transport
        If MsgBox("Aquesta comanda ja té assignat el transport " + atrim(rstc!descripcio) + vbNewLine + "Vols reassignar-lo ACCEPTAR o CANCELAR (Copiar CALLOFFS)?", vbCritical + vbDefaultButton2 + vbOKCancel, "Transport assignat.") = vbCancel Then
            If rstc!client = 6841 Then
                  vcrops = True
                   Else: GoTo fi
            End If
        End If
   End If
   
   Set rst = dbplanificacio.OpenRecordset("Select * from planificaciototes where comanda=" + atrim(vnumc) + " order by comanda")
   If rst.EOF Then MsgBox "Data expedició no trobada": GoTo fi
   If IsNull(rst!dataexpedicio) Then
      MsgBox "No hi ha la data d'expedició a planificació.", vbCritical, "Error": GoTo fi
    'Set rst = dbplanificacio.OpenRecordset("SELECT planificaciototes.*, comandes.direnvio, comandes.comandaclient,comandes.rebkilos, comandes_extres.totalspesMesTA FROM (planificaciototes LEFT JOIN comandes ON planificaciototes.comanda = comandes.comanda) LEFT JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda where (transportista_albara=0 or transportista_albara=null) and direnvio=" + atrim(rstc!ID) + " and dataexpedicio=#" + atrim(Format(rst!dataexpedicio, "mm/dd/yy")) + "#" + " order by planificaciototes.comanda")
      Else
        Set rst = dbplanificacio.OpenRecordset("SELECT planificaciototes.*, comandes.direnvio, comandes.comandaclient,comandes.rebkilos, comandes_extres.totalspesMesTA FROM (planificaciototes LEFT JOIN comandes ON planificaciototes.comanda = comandes.comanda) LEFT JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda where  direnvio=" + atrim(rstc!ID) + " and dataexpedicio=#" + atrim(Format(rst!dataexpedicio, "mm/dd/yy")) + "#" + " order by planificaciototes.comanda")
      '  Clipboard.Clear
      '  Clipboard.SetText "SELECT planificaciototes.*, comandes.direnvio, comandes.comandaclient,comandes.rebkilos, comandes_extres.totalspesMesTA FROM (planificaciototes LEFT JOIN comandes ON planificaciototes.comanda = comandes.comanda) LEFT JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda where  direnvio=" + atrim(rstc!ID) + " and dataexpedicio=#" + atrim(Format(rst!dataexpedicio, "mm/dd/yy")) + "#" + " order by planificaciototes.comanda"
        If rst.EOF Then MsgBox "No s'ha trobat aquesta comanda a planificació": GoTo fi
   End If
   While Not rst.EOF
      Set rstc2 = dbtmp.OpenRecordset("select * from comandes_extres where comanda=" + atrim(rst!comanda))
      If cadbl(rstc2!transportista_albara) = 0 Then
        If vcrops Then
          v = buscarcalloff(rst!comanda)
          vcalloff = IIf(InStr(1, vcalloff, v) = 0, vcalloff + "[" + v + "]" + vbNewLine, vcalloff)
        End If
        If InStr(1, vcomandesdelclient, atrim(rst!comandaclient)) = 0 Then
             vcomandesdelclient = vcomandesdelclient + IIf(vcomandesdelclient <> "", ", ", "") + atrim(rst!comandaclient)
        End If
        vcomandes = vcomandes + atrim(rst!comanda) + " "
      End If
      rst.MoveNext
   Wend
   Set rste = dbvendes.OpenRecordset("select * from registre_enviaments where comandesrelacionades='" + atrim(vcomandes) + "'")
   If Not rste.EOF Then If MsgBox("Aquesta relació de comandes ja t'he un transport assignat." + vbNewLine + "VOLS SUBSTITUIR-LO?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then GoTo fi
   rst.MoveFirst
   If vcrops And vcalloff <> "" Then passarcalloffaNOTEPAD vcalloff, atrim(rst!dataexpedicio): GoTo fi
   If rst.RecordCount > 1 Then
       If MsgBox("Aquesta comanda te una direcció d'enviament que afecte a varies comandes: " + vbNewLine + atrim(rstc!direccioenviament) + vbNewLine + vcomandes + vbNewLine + "Vols assignar aquest transportista a totes? ", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then
         Set rst = dbplanificacio.OpenRecordset("SELECT planificaciototes.*, comandes.direnvio, comandes.rebkilos, comandes_extres.totalspesMesTA FROM (planificaciototes LEFT JOIN comandes ON planificaciototes.comanda = comandes.comanda) LEFT JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda where comandes.comanda=" + atrim(vnumc) + " order by planificaciototes.comanda")
       End If
   End If
   rst.MoveFirst
   While Not rst.EOF
      If InStr(1, vcomandes, atrim(rst!comanda)) > 0 Then
        vkgcomandes = vkgcomandes + IIf(cadbl(rst!totalspesMesTA) > 0, cadbl(rst!totalspesMesTA), cadbl(rst!rebkilos))
        Set rstpesb = dbvendes.OpenRecordset("SELECT Sum(liniesalbara.kgtotalsbruts) AS Kgentregats, liniesalbara.lotinplacsa FROM capcaleraalbara LEFT JOIN liniesalbara ON capcaleraalbara.numalbara = liniesalbara.numalbara where ((Not (capcaleraalbara.numalbaraSAP) Is Null And (capcaleraalbara.numalbaraSAP) <> 0)) GROUP BY liniesalbara.lotinplacsa HAVING (((liniesalbara.lotinplacsa)=" + atrim(rst!comanda) + "));")
        If cadbl(rstpesb!kgentregats) > 0 Then
             vkgcomandes = vkgcomandes - cadbl(rstpesb!kgentregats)
             vhihaunparcial = True
        End If
      End If
      rst.MoveNext
   Wend
   vpalets = calcularpalets(vkgcomandes)
   vidtransport = 0
   vpreusuggerit = buscar_preutransport(videnvio, vpalets, vkgcomandes, vidtransport, vnomtransport, vmsg)
   If MsgBox("Direcció d'entrega: " + rstc!direccioenviament + vbNewLine + vbNewLine + "Kg Totals: " + atrim(vkgcomandes) + "KG." + IIf(vhihaunparcial, "  [Hi ha un parcial]", "") + vbNewLine + atrim(vpalets) + " Palets" + vbNewLine + "Transport: " + vnomtransport + vbNewLine + atrim(vpreusuggerit) + "€" + vbNewLine + vbNewLine + vmsg, vbInformation + vbDefaultButton1 + vbYesNo, "Informació") = vbNo Then
        vidtransport = escullir_transportista(vnomtransport)
        vpreusuggerit = 0
   End If
   If vidtransport > 0 Then
       rst.MoveFirst
       While Not rst.EOF
          dbtmp.Execute "update comandes_extres set transportista_albara=" + atrim(vidtransport) + " where comanda=" + atrim(rst!comanda)
          vcomandesafectades = vcomandesafectades + IIf(vcomandesafectades <> "", ",", "") + atrim(rst!comanda)
          rst.MoveNext
       Wend
       If MsgBox("Segur que vols assignar aquest transportista a la o les comandes?" + vbNewLine + vcomandesafectades, vbExclamation + vbDefaultButton2 + vbYesNo, "Escullir transportista") = vbNo Then GoTo fi
       rst.MoveFirst
       If atrim(rstc!pais) <> "ES" Then
            possar_correuplantillaalPORTAPAPERS vnomclient, rst!dataexpedicio, rstc!direccioenviament, vkgcomandes, vcomandesdelclient, vpalets
            MsgBox "Assignació de transportista feta. " + vbNewLine + "S'HA COPIAT AL PORTAPAPERS EL CORREU PLANTILLA PEL TRANSPORTISTA." + vbNewLine
             Else: MsgBox "Assignació de transport NACIONAL fet.", vbInformation, "Fet"
       End If
       If vpreusuggerit = 0 Then
            v = InputBox("Entra el preu de tarifa per aquest enviament." + vbNewLine + "Kg: " + atrim(vkgcomandes) + "   Palets:" + atrim(vpalets), "Preu teòric de l'enviament.")
              Else: v = atrim(vpreusuggerit)
       End If
       If cadbl(v) = 0 Then v = "0"
       veuros = substituir(v, ".", ",")
    'registro el moviment a registre_enviaments
       dbvendes.Execute "delete * from registre_enviaments where comandesrelacionades='" + atrim(vcomandes) + "'"
       vvalues = "(" + atrim(vidtransport) + ",'" + treure_apostruf(vnomtransport) + "','" + rstc!direccioenviament + "'," + atrim(rstc!ID) + "," + atrim(passaradecimalpunt(atrim(vkgcomandes))) + "," + atrim(cadbl(vpalets)) + "," + atrim(substituir(atrim(veuros), ",", ".")) + ",'" + atrim(vcomandes) + "')"
       dbvendes.Execute "insert into registre_enviaments (id_transport,nomtransport,desti,id_desti,kgteorics,paletsteorics,eurosports,comandesrelacionades) values" + vvalues
       demanarINCOTERMS cadbl(videnvio), atrim(vcomandesafectades)
   End If
   
fi:
   Set rst = Nothing
   Set rste = Nothing
   Set rstc = Nothing
   
End Sub
Sub demanarINCOTERMS(videnvio As Double, vcomandes As String)
    Dim rst As Recordset
    Dim v As String
    Set rst = dbtmp.OpenRecordset("Select * from clients_envios where id=" + atrim(videnvio))
    If rst.EOF Then v = "DAP"
    v = atrim(rst!INCOTERM)
    If Len(v) > 3 Then
      While vr = "" Or InStr(1, v, vr) = 0
       vr = InputBox("Escull quin INCOTERM vols utilitzar per les comanda/s. " + vbNewLine + vcomandes + vbNewLine + "Valors permesos:  " + v, "INCOTERM")
      Wend
      v = vr
    End If
    dbtmp.Execute "update comandes_EXTRES set incoterm_envio='" + atrim(treure_apostruf(v)) + "' where comanda in (" + atrim(vcomandes) + ")"
    Set rst = Nothing
End Sub
Function buscar_preutransport(vid_envio As Long, vpalets As Integer, vkgs As Double, vidtransport As Long, vnomtransport As String, vmsg As String) As Double
   Dim rst As Recordset
   Dim rstenvio As Recordset
   Dim vcodipostal As String
   Dim vpais As String
   Dim vidtransportFAVORIT  As Long
   Dim vsqltransportFAVORIT  As String
   Dim vmsgp As String
   Dim vpreuperpalet As Double
   Dim vkgs_original As Double
   vkgs_original = cadbl(vkgs)
   If vid_envio = 0 Then GoTo fi
   Set rstenvio = dbvendes.OpenRecordset("select * from clients_envios where id=" + atrim(vid_envio))
   If rstenvio.EOF Then GoTo fi
   vpais = atrim(rstenvio!pais)
   If vpais = "" Then MsgBox "Aquest client no té el país assignat a la direcció d'enviament, primer possa-li.", vbCritical, "Atenció": Exit Function
   vcodipostal = rstenvio!codipostale
   vidtransportFAVORIT = cadbl(rstenvio!id_transportFAVORIT)
   If vidtransport > 0 Then vidtransportFAVORIT = vidtransport
   vsqltransportFAVORIT = IIf(vidtransportFAVORIT > 0, "id_transport=" + atrim(vidtransportFAVORIT), "")
   vnomtransport = ""
   vidtransport = 0
   
    'BUSCO PER PALETS
   Set rst = dbvendes.OpenRecordset("select * FROM Tarifes_ports LEFT JOIN transportistes ON Tarifes_ports.id_transport = transportistes.codi  where tarifaperpaletsokg='P'" + IIf(vsqltransportFAVORIT <> "", " and " + vsqltransportFAVORIT, "") + " and numpalets=" + atrim(vpalets) + " and pais='" + atrim(vpais) + "' and (codipostal='" + atrim(Mid(Trim(vcodipostal) + "  ", 1, 2)) + "' or codipostal='" + atrim(vcodipostal) + "') order by preu asc")
   If rst.EOF Then Set rst = dbvendes.OpenRecordset("select * FROM Tarifes_ports LEFT JOIN transportistes ON Tarifes_ports.id_transport = transportistes.codi where tarifaperpaletsokg='P' " + IIf(vsqltransportFAVORIT <> "", " and " + vsqltransportFAVORIT, "") + " and numpalets=" + atrim(vpalets) + " and pais='" + atrim(vpais) + "' and codipostal='Tots' order by preu asc")
   If rst.EOF Then GoTo buscarperKG
   buscar_preutransport = rst!preu
   vidtransport = rst!id_transport
   vnomtransport = rst!descripcio
   While Not rst.EOF
      vmsg = vmsg + atrim(rst!descripcio) + "  ->  " + atrim(rst!preu) + "€" + vbNewLine
      vpreuperpalet = rst!preu
      rst.MoveNext
   Wend
   If vmsg <> "" Then vmsgp = "Tarifes coincidents per PALETS." + vbNewLine + "=======================" + vbNewLine + vmsg
   
buscarperKG:
   'BUSCO PER KILOS
   vmsg = ""
   vsql = "SELECT FIRST(transportistes.redondeigeurokg) as redondeigeurokgs, first(tarifes_ports.id_transport) as idtransport,First(transportistes.descripcio) AS nomtransport,First(Tarifes_ports.preu) AS PreuKg FROM Tarifes_ports LEFT JOIN transportistes ON Tarifes_ports.id_transport = transportistes.codi "
   vsql = vsql + " where Tarifes_ports.numpalets>" + atrim(Redondejar(cadbl(vkgs), 0)) + " And Tarifes_ports.tarifaperpaletsokg = 'K' " + IIf(vsqltransportFAVORIT <> "", " and " + vsqltransportFAVORIT, "") + " and pais='" + atrim(vpais) + "' and (codipostal='" + atrim(Mid(Trim(vcodipostal) + "  ", 1, 2)) + "' or codipostal='" + atrim(vcodipostal) + "') "
   vsql = vsql + " GROUP BY Tarifes_ports.id_transport, Tarifes_ports.pais, Tarifes_ports.codipostal order by First(Tarifes_ports.preu) asc;"
   
   Set rst = dbvendes.OpenRecordset(vsql)
   If rst.EOF Then
      vsql = "SELECT first(tarifes_ports.id_transport) as idtransport,First(transportistes.descripcio) AS nomtransport,First(Tarifes_ports.preu) AS PreuKg FROM Tarifes_ports LEFT JOIN transportistes ON Tarifes_ports.id_transport = transportistes.codi "
      vsql = vsql + " where Tarifes_ports.numpalets >" + atrim(Redondejar(cadbl(vkgs), 0)) + " And Tarifes_ports.tarifaperpaletsokg = 'K' " + IIf(vsqltransportFAVORIT <> "", " and " + vsqltransportFAVORIT, "") + " and pais='" + atrim(vpais) + "' and codipostal='Tots' "
      vsql = vsql + " GROUP BY Tarifes_ports.id_transport, Tarifes_ports.pais, Tarifes_ports.codipostal  order by First(Tarifes_ports.preu) asc;"
      Set rst = dbvendes.OpenRecordset(vsql)
   End If
   'Clipboard.Clear
   'Clipboard.SetText vsql
   If rst.EOF Then GoTo fi
   If cadbl(rst!redondeigeurokgs) > 0 And rst!PreuKg < 0 Then vkgs = Int((vkgs / rst!redondeigeurokgs) + 0.99) * rst!redondeigeurokgs
   If cadbl(rst!redondeigeurokgs) > 0 Then buscar_preutransport = IIf(rst!PreuKg < 0, (vkgs / cadbl(rst!redondeigeurokgs)) * (rst!PreuKg * -1), rst!PreuKg)
   'vidtransport = rst!idtransport
   'vnomtransport = rst!nomtransport
   While Not rst.EOF
      vkgs = vkgs_original
      vredondeixeuroskg = cadbl(rst!redondeigeurokgs)
      If vredondeixeuroskg = 0 Then vredondeixeuroskg = 1
      If cadbl(rst!redondeigeurokgs) > 0 And rst!PreuKg < 0 Then vkgs = Int((vkgs / vredondeixeuroskg) + 0.99) * rst!redondeigeurokgs
      vmsg = vmsg + atrim(rst!nomtransport) + "  ->  " + atrim(IIf(rst!PreuKg < 0, (vkgs / vredondeixeuroskg) * (rst!PreuKg * -1), rst!PreuKg)) + "€" + vbNewLine
      If IIf(rst!PreuKg < 0, (vkgs / vredondeixeuroskg) * (rst!PreuKg * -1), rst!PreuKg) < buscar_preutransport Then
        buscar_preutransport = IIf(rst!PreuKg < 0, (vkgs / vredondeixeuroskg) * (rst!PreuKg * -1), rst!PreuKg)
        vidtransport = rst!idtransport
        vnomtransport = rst!nomtransport
      End If
      rst.MoveNext
   Wend
   
fi:
   'If vpreuperpalet < buscar_preutransport And vpreuperpalet > 0 Then buscar_preutransport = vpreuperpalet
   If vmsg <> "" Then vmsgp = vmsgp + vbNewLine + vbNewLine + "Tarifes coincidents per Kilos." + vbNewLine + "=======================" + vbNewLine + vmsg
   vmsg = vmsgp
   vkgs = vkgs_original
   Set rst = Nothing
   Set rstenvio = Nothing
End Function
Function diadelasetmana(vdata As Date) As String
   Dim vdiaset As Variant
   vdiaset = Array("Dilluns", "Dimarts", "Dimecres", "Dijous", "Divendres", "Dissabtes", "Diumenges")
   diadelasetmana = vdiaset(Format(vdata, "w", vbMonday) - 1)
End Function
Function calcularpalets(vkg As Double) As Double
    calcularpalets = Fix(vkg / 650)
    If calcularpalets < 1 Then
       calcularpalets = 1
         Else: calcularpalets = calcularpalets + 1
    End If
End Function
Function possar_correuplantillaalPORTAPAPERS(vnomclient As String, vdataexp As Date, vdireccio As String, vkg As Double, vcomandesdelclient As String, vpalets As Integer)
    Dim vmsg As String
    Dim vmsg2 As String
    Dim vmsg3 As String
    If InStr(1, UCase(vnomclient), " VIOLAINES") > 0 Then
         vmsg2 = "**** ULL!!! ANOTACIÓ MÉS AVALL - ES TÉ DE FER RESERVA, PER PODER DESCARREGAR. ****" + vbNewLine + vbNewLine + vbNewLine + vbNewLine
         vmsg3 = "=========================================================================================================" + vbNewLine
         vmsg3 = vmsg3 + "Descàrrega de dilluns a divendres de 8 a 13h" + vbNewLine
         vmsg3 = vmsg3 + "Caldrà fer la reserva per poder fer la descàrrega enviant un correu el dia abans, abans de les 12h" + vbNewLine
         vmsg3 = vmsg3 + "a anita.hanck@ardo.com" + vbNewLine
         vmsg3 = vmsg3 + "Els camions que arribin sense reserva prèvia, no es descarregaran." + vbNewLine
         vmsg3 = vmsg3 + "=========================================================================================================" + vbNewLine + vbNewLine
    End If
    If InStr(1, UCase(vnomclient), "ARDO UK") > 0 Then
         vmsg3 = vbNewLine + "És molt important que contacteu amb aquest client per demanar 'booking' (dia i hora d'entrega). De no ser així, no hi hauria possibilitat de descàrrega de la mercaderia." + vbNewLine
         vmsg3 = vmsg3 + "Cal que us poseu en contacte per email amb el responsable del magatzem - Mr. Marc Sharrad - Marc.Sharrad@ardouk.com (Tel. 01233714714)." + vbNewLine
         vmsg3 = vmsg3 + "En cas que no fos possible contactar amb aquesta persona, ho haurieu de fer amb el seu ajudant - Mr. Marcus Lilly - Marcus.Lilley@ardouk.com (Tel. 01233714714)." + vbNewLine
         vmsg3 = vmsg3 + "Els camions que arribin sense reserva prèvia, no es descarregaran." + vbNewLine
         vmsg3 = vmsg3 + "aquest destinatari té agent propi d'aduanes, contacte: Stephen Ierston <Stephen@merseyforwarding.com> + John Dorrington <John.Dorrington@ardo.com>" + vbNewLine + vbNewLine
    End If
    If InStr(1, vdireccio, "Ardooide") > 0 Then 'dir envio de D'Arta Ardooie
         vmsg2 = "**** ULL!!! ANOTACIÓ MÉS AVALL - ES TÉ DE FER RESERVA, PER PODER DESCARREGAR. ****" + vbNewLine
         vmsg3 = "=========================================================================================================" + vbNewLine
         vmsg3 = vmsg3 + "Descàrrega de dilluns a divendres de 7 a 11h i de 14 a 18 h." + vbNewLine
         vmsg3 = vmsg3 + "Caldrà fer la reserva per poder fer la descàrrega enviant un correu el dia abans a charlotte.vercruysse@darta.com amb còpia a rika.bouckaert@darta.com." + vbNewLine
         vmsg3 = vmsg3 + "Els camions que arribin sense reserva prèvia, no es descarregaran." + vbNewLine
         vmsg3 = vmsg3 + "=========================================================================================================" + vbNewLine
    End If
    If InStr(1, UCase(vnomclient), "BEWITAL") > 0 Then 'client de D'BEWITAL
         vmsg2 = "**** ULL!!! ANOTACIÓ MÉS AVALL - ES TÉ DE FER RESERVA, PER PODER DESCARREGAR. ****" + vbNewLine
         vmsg3 = "=========================================================================================================" + vbNewLine
         vmsg3 = vmsg3 + "Descàrrega de dilluns a dijous de 8 a 17h i divendres de 8 a 13h." + vbNewLine
         vmsg3 = vmsg3 + "Caldrà fer la reserva amb 24 hores d'antelació per poder fer la descàrrega enviant un correu a truck@bewital.de. El telèfon de contacte és: +49 2862-581-740." + vbNewLine
         vmsg3 = vmsg3 + "Els camions que arribin sense reserva prèvia, no es descarregaran." + vbNewLine
         vmsg3 = vmsg3 + "=========================================================================================================" + vbNewLine + vbNewLine
    End If
    
    vmsg = "Bon dia," + vbNewLine + vbNewLine
    vmsg = vmsg + "Per " + diadelasetmana(vdataexp) + ", dia " + atrim(Format(vdataexp, "dd/mm")) + ", tindrem una nova expedició al següent destí :" + vbNewLine + vbNewLine
    vmsg = vmsg + vmsg2 + vbNewLine + vbNewLine
    vmsg = vmsg + " - ( Uns " + atrim(vpalets) + IIf(cadbl(vpalets) > 1, " palets", " palet") + " - " + Format(vkg, "#,##0") + " kgs ) => " + atrim(vnomclient) + "(" + vdireccio + ")" + vbNewLine + vbNewLine
    'If InStr(1, UCase(vnomclient), "CROP´S") > 0 Then vmsg = vmsg + " ( FALTARÀ, EL DOCUMENT SLOT RESERVATION ) L'ENVIAREM, EN UN ALTRE E-MAIL." + vbNewLine + vbNewLine + vbNewLine + vbNewLine
    If InStr(1, UCase(vnomclient), "BEWITAL") > 0 Then vmsg = vmsg + " - Números de comanda del client: " + vcomandesdelclient + vbNewLine
    If InStr(1, UCase(vnomclient), "DUJARDIN") > 0 Then
         vmsg = vmsg + " - Números de comanda del client: " + vcomandesdelclient + vbNewLine
         'If InStr(1, UCase(vdireccio), "KOOLSKAMP") > 0 Then
         vmsg = vmsg + "Per fer la descàrrega de la mercaderia d'aquest client, el xofer ha de passar primer per les oficines logístiques i allà li informaran a quin moll ha d'anar a descarregar." + vbNewLine + vbNewLine
    End If
    vmsg = vmsg + vmsg3 + vbNewLine + vbNewLine
    vmsg = vmsg + "Agrairem procediu a la recollida d'aquesta mercaderia i ens informeu de la data de lliurament. " + vbNewLine + vbNewLine + vbNewLine
    vmsg = vmsg + "Moltes gràcies!" + vbNewLine + vbNewLine
    vmsg = vmsg + "Salutacions cordials,"
    Clipboard.Clear
    Clipboard.SetText vmsg
End Function
Sub passarcalloffaNOTEPAD(vcalloff As String, vdia As String)
   If atrim(vcalloff) = "" Then MsgBox "NO HI HA CALLOFFS": Exit Sub
   Open "c:\temp\calloffsCROPS.txt" For Output As 1
   Print #1, "Llista de calloffs de CROP'S    DATA:" + vdia
   Print #1, "==============================================="
   Print #1, vcalloff
   Close 1
   If existeix("c:\temp\calloffsCROPS.txt") Then obrir_document "c:\temp\calloffsCROPS.txt"
End Sub
Function buscarcalloff(vnumc As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select distinct numcalloff,entregat from bobinesent where (entregat='N' or entregat=null) and (numcalloff<>'' and numcalloff<>null) and comanda=" + atrim(vnumc) + " order by entregat")
   If Not rst.EOF Then
       buscarcalloff = atrim(rst!numcalloff)
   End If
   If buscarcalloff = "" Then
      Set rst = dbtmp.OpenRecordset("select numcalloff from comandes_extres where comanda=" + atrim(vnumc))
      If Not rst.EOF Then buscarcalloff = atrim(rst!numcalloff)
   End If
   Set rst = Nothing
End Function
Function escullir_transportista(vnomtransport As String) As Long
    Load formseleccio
  'formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,descripcio from transportistes where visible=1"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 4000
  'formseleccio.Width = 9000
  'formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           escullir_transportista = formseleccio.DBGrid2.Columns("codi")
           vnomtransport = formseleccio.DBGrid2.Columns("descripcio")
        End If
  End If
End Function
Private Sub materials_Click()
 
 ' Load formaltarep
'  formaltarep.Caption = "Manteniment de Materials"
  'formaltarep.colsbloc = "4"
  'formaltarep.Data1.DatabaseName = cami
  'formaltarep.Data1.RecordSource = "select * from materials"
  'formaltarep.Width = (formaltarep.Width * 2) + 200
  'formaltarep.DBGrid1.Width = (formaltarep.DBGrid1.Width * 2) + 400
  'formaltarep.refrescar
  'formaltarep.DBGrid1.Refresh
  'formaltarep.Show
  fmaterials.Show
End Sub

Private Sub mavariaimpresores_Click()
 Load formaltarep
  formaltarep.Caption = "Avaries impresores"
'  formaltarep.autonum = "transportistes"
  formaltarep.Data1.DatabaseName = rutadelfitxer(cami) + "baixes.mdb"
  formaltarep.Data1.RecordSource = "select tipificacioavaria as [Tipus avaria] from impresores_tipificacionsavaria"

  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  
  formaltarep.Width = formaltarep.Width + 1500
  formaltarep.DBGrid1.Width = formaltarep.DBGrid1.Width + 1500
  formaltarep.DBGrid1.Columns(0).Width = 5000
  formaltarep.Left = Me.Left + 500
  formaltarep.Show
End Sub

Private Sub mavisosseccions_Click()
   Shell "\\serverprodu\dades\progcomandes\aplicacio\avisosseccions.exe", vbNormalFocus
End Sub

Private Sub mbaixescostos_Click()
  baixescostos.Show
End Sub

Private Sub mbaixesmuntadoraentredates_Click()
  Shell """\\serverprodu\dades\progcomandes\aplicacio\baixes muntadora.exe"" LLISTATBAIXESMUNTADORA", vbNormalFocus
End Sub

Private Sub mbobinesembolicades_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment Bobines embolicades"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from bobinesembolicades"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Show
End Sub

Private Sub mborrarcomanda_Click()
   Dim vnumc As Double
   Dim rst As Recordset
   If UCase(InputBoxEx("Entra la CONTRASENYA d'accés.", "Molt perillos", , , , , , SPassword)) <> "INPLACSAELIMINAR" Then Exit Sub
   vnumc = cadbl(InputBox("Entra el numero de comanda que vols eliminar definitivament." + Chr(10) + "AQUESTA ELIMINACIÓ NO REVISA SI HI HA COMPRES, ASSIGNACIONS, BOBINES ETC..." + vbNewLine + " NOMES ES BORRARÀ AQUEST NUMERO DE COMANDA NO ELS COMPLEXES.", "BORRAR COMANDA DEFINITIVAMENT"))
   If cadbl(vnumc) = 0 Then GoTo fi
   Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
   If Not rst.EOF Then Set rst = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(rst!client))
   If rst.EOF Then GoTo fi
   If MsgBox("Segur que vols eliminar la comanda " + atrim(vnumc) + vbNewLine + "Del client: " + atrim(rst!nom), vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then GoTo fi
   dbtmp.Execute "delete * from comandes where comanda=" + atrim(vnumc)
   dbtmp.Execute "delete * from comandes_extres where comanda=" + atrim(vnumc)
   MsgBox "Comanda " + atrim(vnumc) + " eliminada.", vbInformation, "Atenció"
   
fi:
   Set rst = Nothing
End Sub

Private Sub mcertqualitat_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment de Cert. qualitat"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from cert_qualitat"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Show
End Sub

Private Sub mcilindres_Click()
   formcilindresimpresores.Show 1
End Sub

Private Sub mclixes_Click()
  'If nomusuari <> "Usr_Disseny" Then MsgBox "No tens drets per editar clixes." + Chr(10) + Chr(13): Exit Sub
  If Not existeix("c:\windows\system32\MSCOMCT2.OCX") Then
      Copiar_Fitxer "\\serverprodu\dades\progcomandes\aplicacio\instalaciocomandes\mscom*.*", "c:\windows\system32"
  End If
   Shell "\\serverprodu\dades\progcomandes\aplicacio\clixesnous.exe", vbNormalFocus
  
End Sub

Private Sub mclixesnous_Click()
Shell "\\serverprodu\dades\progcomandes\aplicacio\clixes.exe", vbNormalFocus
End Sub

Private Sub mcompres_Click()

   Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "compres.exe", vbNormalFocus
End Sub

Private Sub mcomprovarcomandessensetemperatures_Click()
  Dim rst As Recordset
  Dim entrat As Boolean
  ratoli "espera"
  'Set rst = dbtmp.OpenRecordset("select comanda from comandes where proximaseccio='T' and producte<>'PC' and producte<>'PC2' and producte<>'PCP' and materialex>499 order by comanda DESC")
  Set rst = dbtmp.OpenRecordset("SELECT comandes.comanda, InStr(1,[codi],'I') AS Expr1 FROM comandes INNER JOIN productes ON comandes.producte = productes.codi WHERE (((InStr(1,[codi],'I'))>0)) and comandes.materialex>499 order by comanda DESC;")
  If Not rst.EOF Then rst.MoveLast: rst.MoveFirst
  Open "c:\temp\notemperatures.txt" For Output As 1
  Print #1, "*********************************************************"
  Print #1, "***   LLISTA DE COMANDES SENSE FITXER DE TEMPERATURES ***"
  Print #1, "*********************************************************"
  Print #1, " "
  Print #1, " "
  While Not rst.EOF
     If Not existeix(llegir_ini("ruta", "ruta_documentacio_temperatures", rutadelfitxer(cami) + "valorsprograma.ini") + "\" + atrim(rst!comanda) + ".txt") Then
        Print #1, "  Comanda:  " + atrim(rst!comanda)
        entrat = True
     End If
     Me.Caption = " Revisant les comandes sense fitxer de temperatures: ----> " + atrim(Int((rst.AbsolutePosition * 100) / rst.RecordCount)) + "%"
     DoEvents
     rst.MoveNext
  Wend
  Close 1
  If entrat Then Shell "c:\windows\notepad.exe c:\temp\notemperatures.txt", vbNormalFocus
  ratoli "normal"
  Me.Caption = "Menu de Comandes"
End Sub

Private Sub mconosprotectors_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment de Conos protectors"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from conosprotectors"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Show
End Sub

Private Sub mconsulsreferenciescli_Click()
   Dim sql As String
   Dim elwhere As String
   Dim fitxertmpestats As String
   Dim v As String
   Dim inici As Date
   Dim fi As Date
   Dim agruparper As String
   Dim vcodiclient As String
   fitxertmpestats = "c:\temp\consultarefinp_tmp.mdb"
   vcodiclient = cadbl(InputBox("Entra el codi de client que vols consultar.", "Codi client", "6841"))
   If vcodiclient = 0 Then GoTo fi
   v = InputBox("Entra la data d'inici de la consulta.", "Inici consulta")
   If Not IsDate(v) Then MsgBox "La data no es correcte.": Exit Sub
   inici = CVDate(v)
   v = InputBox("Entra la data de fi de la consulta.", "Fi consulta")
   If Not IsDate(v) Then MsgBox "La data no es correcte.": Exit Sub
   fi = CVDate(v)
   agruparper = UCase(InputBox("Entra per quin camp vols agrupar el llistat" + Chr(10) + "(R)ReferenciaClient    (T)NºTreball d'impresio   (D)Referencia+Detall_Tintes", "Agrupar per...", "R"))
   If agruparper <> "R" And agruparper <> "T" And agruparper <> "D" Then MsgBox "Opcio no vàlida": GoTo fi
   MsgBox "El llistat pot trigar una mica si l'interval de dates es molt gran." + Chr(10) + "PREM ACCEPTAR PER COMENÇAR EL LLISTAT", vbInformation, "Atenció"
   ratoli "espera"
   Set dbtmp = OpenDatabase(cami)
   elwhere = " first(comandes.client)=" + vcodiclient + " and (first(bobinesent.data)>=#" + Format(inici, "mm/dd/yy") + "# and first(bobinesent.data)<=#" + Format(fi, "mm/dd/yy") + "#) "
   sql = "SELECT first(comandes.client),First(comandes_extres.refinplacsa) AS refinplacsa, First(comandes.producte) AS Pr, First(comandes.refclient) AS Ref_, Min(1) AS Q, comandes.comanda AS maxcomanda, First(bobinesent.data) AS maxdata FROM (comandes INNER JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda) INNER JOIN bobinesent ON comandes_extres.comanda = bobinesent.comanda GROUP BY comandes.comanda "
   sql = sql + " having " + elwhere
  Set rstconsulta = dbtmp.OpenRecordset(sql)
  
  If rstconsulta.EOF Then MsgBox "No hi han resultats per aquest client": GoTo fi
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
  r = "nopoblar"
  wait 2
  Load Formconsultarefinplacsa
  Unload Formconsultarefinplacsa
  If existeix(fitxertmpestats) Then
     Set dbconsulta = DBEngine.OpenDatabase(fitxertmpestats)
      Else: GoTo fi
  End If
  If existeixlataula(fitxertmpestats, "consultaestats") Then
    dbconsulta.Execute "alter table consultaestats add column kilos double"
    dbconsulta.Execute "alter table consultaestats add column metres double"
    dbconsulta.Execute "alter table consultaestats add column unitats double"
    dbconsulta.Execute "alter table consultaestats add column treball double"
    dbconsulta.Execute "alter table consultaestats add column totaleuros double"
    dbconsulta.Execute "alter table consultaestats add column nomtinta1 text"
    dbconsulta.Execute "alter table consultaestats add column nomtinta2 text"
    dbconsulta.Execute "alter table consultaestats add column nomtinta3 text"
    dbconsulta.Execute "alter table consultaestats add column nomtinta4 text"
    dbconsulta.Execute "alter table consultaestats add column nomtinta5 text"
    dbconsulta.Execute "alter table consultaestats add column nomtinta6 text"
    dbconsulta.Execute "alter table consultaestats add column nomtinta7 text"
    dbconsulta.Execute "alter table consultaestats add column nomtinta8 text"
    dbconsulta.Execute "alter table consultaestats add column nomtintaREPRINT text"
    actualitzarcampsrestants
    exportar_llistat_xls agruparper
     Else: MsgBox "NO S'HA GENERAT CAP INFORMACIÓ I NO ES POT TREURE UN LLISTAT.", vbInformation: GoTo fi
  End If
fi:
  Set dbclixes = Nothing
  ratoli "normal"
End Sub
Sub exportar_llistat_xls(agruparper As String)
   Dim i As Byte
   Dim rst As Recordset
   Dim linia As String
   Dim sql As String
   'agrupat per referencia client
   If agruparper = "D" Then
    sql = "SELECT First(consultaestats.datacomanda) AS datacomanda, Count(consultaestats.numcomandes) AS CuentaDenumcomandes, "
    sql = sql + " First(consultaestats.refinplacsa) AS refinplacsa, consultaestats.refclient, Max(consultaestats.numcomanda) AS MáxDenumcomanda, First(consultaestats.producte) AS producte, First(consultaestats.texteimpresio) AS texteimpresio, First(consultaestats.ampleext) AS ampleext, First(consultaestats.amplereb) AS amplereb, First(consultaestats.desarrollimp) AS desarrollimp, First(consultaestats.tintes) AS tintes, First(consultaestats.simulteneitatreb) AS simulteneitatreb, First(consultaestats.amplesol) AS amplesol, First(consultaestats.longitud) AS longitud, First(consultaestats.solapa) AS solapa, First(consultaestats.tipussoldadura) AS tipussoldadura, First(consultaestats.micres) AS micres, First(consultaestats.descfamiliamat) AS descfamiliamat, "
    sql = sql + " First(consultaestats.treball) AS treball,fIRST(consultaestats.nummodificacio) as VersióTreball, Sum(consultaestats.kilos) AS SumaDekilos, Sum(consultaestats.metres) AS SumaDemetres, Sum(consultaestats.unitats) AS SumaDeunitats, sum(consultaestats.totaleuros) as SumaDeEuros, first(nomtinta1) as Tinta1,first(nomtinta2) as Tinta2,first(nomtinta3) as Tinta3,first(nomtinta4) as Tinta4,first(nomtinta5) as Tinta5,first(nomtinta6) as Tinta6,first(nomtinta7) as Tinta7,first(nomtinta8) as Tinta8,first(nomtintaREPRINT) as Tinta_REPRINT From consultaestats GROUP BY consultaestats.refclient;"
   End If
   If agruparper = "R" Then
    sql = "SELECT First(consultaestats.datacomanda) AS datacomanda, Count(consultaestats.numcomandes) AS CuentaDenumcomandes, "
    sql = sql + " First(consultaestats.refinplacsa) AS refinplacsa, consultaestats.refclient, Max(consultaestats.numcomanda) AS MáxDenumcomanda, First(consultaestats.producte) AS producte, First(consultaestats.texteimpresio) AS texteimpresio, First(consultaestats.ampleext) AS ampleext, First(consultaestats.amplereb) AS amplereb, First(consultaestats.desarrollimp) AS desarrollimp, First(consultaestats.tintes) AS tintes, First(consultaestats.simulteneitatreb) AS simulteneitatreb, First(consultaestats.amplesol) AS amplesol, First(consultaestats.longitud) AS longitud, First(consultaestats.solapa) AS solapa, First(consultaestats.tipussoldadura) AS tipussoldadura, First(consultaestats.micres) AS micres, First(consultaestats.descfamiliamat) AS descfamiliamat, "
    sql = sql + " First(consultaestats.treball) AS treball, Sum(consultaestats.kilos) AS SumaDekilos, Sum(consultaestats.metres) AS SumaDemetres, Sum(consultaestats.unitats) AS SumaDeunitats, sum(consultaestats.totaleuros) as SumaDeEuros From consultaestats GROUP BY consultaestats.refclient;"
   End If
   If agruparper = "T" Then
         sql = "SELECT First(consultaestats.datacomanda) AS datacomanda, Count(consultaestats.numcomandes) AS CuentaDenumcomandes, "
         sql = sql + " First(consultaestats.refinplacsa) AS refinplacsa, first(consultaestats.refclient) As refclient, Max(consultaestats.numcomanda) AS MáxDenumcomanda, First(consultaestats.producte) AS producte, First(consultaestats.texteimpresio) AS texteimpresio, First(consultaestats.ampleext) AS ampleext, First(consultaestats.amplereb) AS amplereb, First(consultaestats.desarrollimp) AS desarrollimp, First(consultaestats.tintes) AS tintes, First(consultaestats.simulteneitatreb) AS simulteneitatreb, First(consultaestats.amplesol) AS amplesol, First(consultaestats.longitud) AS longitud, First(consultaestats.solapa) AS solapa, First(consultaestats.tipussoldadura) AS tipussoldadura, First(consultaestats.micres) AS micres, First(consultaestats.descfamiliamat) AS descfamiliamat, "
         sql = sql + " consultaestats.treball, Sum(consultaestats.kilos) AS SumaDekilos, Sum(consultaestats.metres) AS SumaDemetres, Sum(consultaestats.unitats) AS SumaDeunitats, sum(consultaestats.totaleuros) as SumaDeEuros From consultaestats GROUP BY consultaestats.treball;"
   End If
   Set rst = dbconsulta.OpenRecordset(sql)
   If rst.EOF Then MsgBox "No hi ha dades per exportar", vbCritical, "Error": Exit Sub
   Open "c:\temp\consultarefinplacsa.csv" For Output As #1
   If Not rst.EOF Then
    For i = 0 To rst.Fields.Count - 1
      linia = linia + IIf(linia = "", "", ";") + atrim(rst.Fields(i).Name)
    Next i
    Print #1, linia
   End If
   While Not rst.EOF
    linia = ""
    For i = 0 To rst.Fields.Count - 1
      linia = linia + IIf(linia = "", "", ";") + """" + IIf(rst.Fields(i).Name = "codibarres", "Nº: ", "") + atrim(rst.Fields(i)) + """"
    Next i
    Print #1, linia
    rst.MoveNext
   Wend
   Close #1
   wait 2
   obrir_document "c:\temp\consultarefinplacsa.csv"
End Sub
Sub actualitzarcampsrestants()
  Dim rst As Recordset
  Dim rstent As Recordset
  Dim rsttintes As Recordset
  Dim rstc As Recordset
  Set dbvendes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "vendes.mdb")
  Set rst = dbconsulta.OpenRecordset("select * from consultaestats")
  While Not rst.EOF
    Set rstc = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(rst!numcomanda))
    Set rstent = dbtmp.OpenRecordset("SELECT bobinesent.comanda, first(bobinesent.numalbara) as tnumalb,Sum(bobinesent.metresisacs) AS tmetres, Sum(bobinesent.kilosiunitats) AS tkilos From bobinesent where comanda=" + atrim(rst!numcomanda) + " GROUP BY bobinesent.comanda;")
    rst.Edit
    If Not rst.EOF And Not rstent.EOF Then
       rst!treball = cadbl(rstc!numtreball)
       rst!metres = cadbl(rstent!tmetres)
       rst!kilos = cadbl(rstent!tkilos)
       rst!totaleuros = buscarpreu(cadbl(rstent!tnumalb), rstent!comanda)
       If cadbl(rst!desarrollimp) > 0 Then rst!unitats = Redondejar(cadbl(rstent!tmetres) / (cadbl(rst!desarrollimp) / 1000), 0)
'       rst.Update
    End If
    Set rsttintes = dbtmp.OpenRecordset("select * from tintes where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and (ordremodificacio=" + atrim(cadbl(rstc!numordremodificacio)) + " or ordremodificacio=" + atrim(cadbl(rstc!numordremodificacio) * -1) + ") order by ordretinter")
    While Not rsttintes.EOF
      If rsttintes!ordremodificacio > 0 Then
          rst.Fields("nomtinta" + atrim(rsttintes!ordretinter)) = atrim(rsttintes!color)
            Else: If atrim(rsttintes!color) <> "" Then rst!nomtintaREPRINT = atrim(rsttintes!color)
      End If
      rsttintes.MoveNext
    Wend
    rst.Update
    rst.MoveNext
  Wend
  Set rstc = Nothing
  Set rstent = Nothing
  Set rst = Nothing
  Set dbvendes = Nothing
End Sub
Function buscarpreu(numalb As Double, numc As Double)
   Dim rst As Recordset
   Set rst = dbvendes.OpenRecordset("select quantitat,preuvenda from liniesalbara where numalbara=" + atrim(numalb) + " and lotinplacsa=" + atrim(numc))
   If Not rst.EOF Then buscarpreu = Redondejar(cadbl(rst!quantitat) * cadbl(rst!preuvenda), 2)
   Set rst = Nothing
End Function

Private Sub mcontrolprl_Click()
   Shell "\\serverprodu\dades\progcomandes\aplicacio\controlPRL.exe", vbNormalFocus
End Sub

Private Sub membanonims_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment de Emb. Anónims"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from embalatgesanonims"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Show
End Sub

Private Sub menviamentipaqueteria_Click()
   Set dbvendes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "vendes.mdb")
   formenviaments.Show 1
End Sub

Private Sub mexpCSVcrèdit_Click()
   fer_CSV_cronoligia_credit_client
End Sub
Sub fer_CSV_cronoligia_credit_client()
    Dim vcodicomptable As Double
    Dim vrisc As TipusVrisc
    Dim vnomfitxerCSV As String
    Dim rst As Recordset
    Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
    vcodicomptable = cadbl(triar_client_codicomptable)
    If vcodicomptable = 0 Then Exit Sub
    vnomfitxerCSV = "c:\temp\llistat CSV Cronologia crèdit client.csv"
    calcular_credit_delclient cadbl(vcodicomptable), vrisc
    dbtmp.Execute "delete * from Credit_client_cronologic where codicomptable=" + atrim(vcodicomptable)
    Set rst = dbtmp.OpenRecordset("select * from Credit_client_cronologic")
    rst.AddNew  'poso el saldo inicial amb el crèdit
    rst!codicomptable = vcodicomptable: rst!Data = Format(0, "dd/mm/yy 00:01"): rst!concepte = "Inici": rst!deure = 0: rst!haver = vrisc.creditsap
    rst.Update
    rst.AddNew
    rst!codicomptable = vcodicomptable: rst!Data = Format(0, "dd/mm/yy 00:02"): rst!concepte = "Facturat": rst!deure = Redondejar(vrisc.creditgastatsap, 0): rst!haver = 0
    rst.Update
    rst.AddNew
    rst!codicomptable = vcodicomptable: rst!Data = Format(0, "dd/mm/yy 00:03"): rst!concepte = "SAP pendent facturar": rst!deure = Redondejar(vrisc.valoralbaranspendentsSAP, 0): rst!haver = 0
    rst.Update
    possar_comandes_perentregar vcodicomptable, rst
    possar_pagaments_SAP vcodicomptable, rst
    fer_calculs_totalitzacio vcodicomptable
    rst.MoveFirst
    exportar_CSV_credit vnomfitxerCSV, vrisc.nomdelclient, vcodicomptable
    Set rst = Nothing
End Sub
Sub exportar_CSV_credit(vnomfitxerCSV As String, vnomclient As String, vcodicomptable As Double)
    Dim rst As Recordset
    Set rst = dbtmp.OpenRecordset("select * from Credit_client_cronologic where codicomptable=" + atrim(vcodicomptable) + " order by data asc")
    Open vnomfitxerCSV For Output As #1
    Print #1, ""
    Print #1, ""
    Print #1, Format(Now, "dddd, d mmm yyyy")
    Print #1, "Client: " + atrim(vcodicomptable) + " - " + atrim(vnomclient)
    Print #1, ""
  Print #1, "DATA;CONCEPTE;DEURE;HAVER;ENTREGA COMANDES PREVISTES;VENCIMENTS;CREDIT DISPONIBLE"
  While Not rst.EOF
     vlinia = atrim(Format(rst!Data, "dd/mm/yy")) + ";" + atrim(rst!concepte) + ";" + atrim(rst!deure) + ";" + atrim(rst!haver) + ";" + atrim(rst!entregaprevista) + ";" + atrim(rst!venciments) + ";" + atrim(rst!creditdisponible)
     Print #1, vlinia
     rst.MoveNext
  Wend
  Set rst = Nothing
  Close #1
  If existeix(vnomfitxerCSV) Then obrir_document vnomfitxerCSV
End Sub
Sub fer_calculs_totalitzacio(vcodicomptable As Double)
    Dim rst As Recordset
    Set rst = dbtmp.OpenRecordset("select * from Credit_client_cronologic where codicomptable=" + atrim(vcodicomptable) + " order by data asc")
    If Not rst.EOF Then rst.Edit: rst!creditdisponible = rst!haver: vtotal = rst!haver: rst.Update: rst.MoveNext
    While Not rst.EOF
      rst.Edit
      If cadbl(rst!deure) > 0 Then vtotal = vtotal - cadbl(rst!deure)
      If cadbl(rst!entregaprevista) > 0 Then vtotal = vtotal - cadbl(rst!entregaprevista)
      If cadbl(rst!venciments) > 0 Then vtotal = vtotal + cadbl(rst!venciments)
      rst!creditdisponible = vtotal: rst.Update
      rst.MoveNext
    Wend
    Set rst = Nothing
End Sub
Sub possar_pagaments_SAP(vcodicomptable As Double, rst As Recordset)
    Dim dbsap As Database
    Dim rstsap As Recordset
    Set dbsap = OpenDatabase(rutadelfitxer(cami) + "connexiosap.mdb")
    Set rstsap = dbsap.OpenRecordset("select * from Importada_RebutspendentsClients_Inplacsa where codicomptableclient='" + atrim(vcodicomptable) + "' order by datavenciment")
    While Not rstsap.EOF
      rst.AddNew
      rst!codicomptable = vcodicomptable: rst!Data = Format(rstsap!datavenciment, "dd/mm/yy"): rst!concepte = "Pago previst Fact: " + atrim(rstsap!numfactura): rst!venciments = Redondejar(rstsap!totalpendent, 0)
      rst.Update
      rstsap.MoveNext
    Wend
    Set rstsap = Nothing
    Set dbsap = Nothing
End Sub
Function calcular_servit(vnumc As Double) As Double
     calcular_servit = 0
     Set rstvendes = dbtmp.OpenRecordset("SELECT Sum(quantitat) AS quantitatentregada, First([dataenvioasap]) AS vdataenvioasap FROM liniesalbara LEFT JOIN capcaleraalbara ON liniesalbara.numalbara = capcaleraalbara.numalbara WHERE liniesalbara.lotinplacsa=" + atrim(vnumc))
     If Not rstvendes.EOF Then
        If Not IsNull(rstvendes!vdataenvioasap) Then calcular_servit = cadbl(rstvendes!quantitatentregada)
     End If
End Function
Sub possar_comandes_perentregar(vcodicomptable As Double, rst As Recordset)
    Dim rstc As Recordset
    Dim vvalorcomanda As Double
    
    Dim vquantitatservida As Double
    Set rstc = dbtmp.OpenRecordset("select datacomanda,comanda,codicomptable,proximaseccio,pvp,tubbaseext,rebkilos,cantitatsol,dataentrega from comandesmesextres where producte<>'PC' and producte<>'PC2' and producte<>'PCP' and proximaseccio<>'T' and codicomptable=" + atrim(vcodicomptable) + " order by proximaseccio")
    While Not rstc.EOF
        vquantitatservida = calcular_servit(rstc!comanda)
        vvalorcomanda = Redondejar(cadbl(rstc!pvp) * (cadbl(rstc!tubbaseext) - vquantitatservida), 2)
        If vvalorcomanda < 0 Then vvalorcomanda = 0
        rst.AddNew
        rst!codicomptable = vcodicomptable: rst!Data = Format(IIf(IsNull(rstc!dataentrega), rstc!datacomanda, rstc!dataentrega), "dd/mm/yy"): rst!concepte = "Comanda " + atrim(rstc!comanda) + " " + atrim(rstc!proximaseccio): rst!entregaprevista = vvalorcomanda
        rst.Update
        rstc.MoveNext
    Wend
    Set rstc = Nothing
End Sub
Function triar_client_codicomptable() As Double
   Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codicomptable,nomclient,predeterminat from clients_codiscomptables order by predeterminat asc"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(0).Width = 1500
  formseleccio.DBGrid2.Columns(1).Width = 3000
  formseleccio.DBGrid2.Columns(2).Width = 0
  'formseleccio.Width = 9000
  'formseleccio.Left = formseleccio.Left - 3000
  If formseleccio.Data1.Recordset.EOF Then MsgBox "Aquest client no te cap CODI COMPTABLE ASSIGNAT.": Exit Function
  formseleccio.Data1.Recordset.MoveLast
  formseleccio.Data1.Recordset.MoveFirst
  If formseleccio.Data1.Recordset.RecordCount > 1 Then
         formseleccio.Show 1
     While formseleccio.Visible
        DoEvents
     Wend
    Else: seleccioret = 1
  End If
  
   If seleccioret = 1 Then If Not formseleccio.Data1.Recordset.EOF Then triar_client_codicomptable = cadbl(formseleccio.DBGrid2.Columns("codicomptable"))
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Function


Private Sub mguardarmostres_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment de Guardar mostres"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from guardarmostres"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Show
End Sub

Private Sub mimpostenvasos_Click()
   Shell "\\serverprodu\dades\progcomandes\aplicacio\Manteniment Impost Envasos.exe", vbNormalFocus
End Sub

Private Sub mllenvxrfacturaCSV_Click()
  Dim vsql As String
  Dim vnumfactura As String
  Dim rst As Recordset
  Dim vlinia As String
  
  vnumfactura = InputBox("Escriu el numero de factura que vols detallar la informació de l'Impost de envasos.", "Factura")
  If cadbl(vnumfactura) = 0 Then Exit Sub
  vsql = "SELECT capcaleraalbara.numfacturaSAP, liniesalbara.numcomandacli, liniesalbara.refclient, liniesalbara.marcailinia, ([Kgimpostenvasos])*[eurokg_impost] AS EurosKg "
  vsql = vsql + " FROM liniesalbara INNER JOIN capcaleraalbara ON liniesalbara.numalbara = capcaleraalbara.numalbara "
  vsql = vsql + " WHERE (((capcaleraalbara.numfacturaSAP)=" + vnumfactura + "));"
  Set rst = dbtmp.OpenRecordset(vsql)
  If rst.EOF Then MsgBox "No hi ha informació d'aquesta factura.", vbCritical, "Error": Exit Sub
  On Error GoTo err
  Open "c:\temp\llistat ImpostEnv per factura.csv" For Output As #1
  Print #1, "Nº FACTURA;Nº PEDIDO;CÓDIGO ARTÍCULO;DESCRIPCIÓN ARTÍCULO;IMPORTE IMPUESTO PLÁSTICO"
  While Not rst.EOF
     vlinia = atrim(rst!numfacturaSAP) + ";" + atrim(rst!numcomandacli) + ";" + atrim(rst!refclient) + ";" + atrim(rst!marcailinia) + ";" + atrim(Redondejar(cadbl(rst!EurosKg), 2))
     Print #1, vlinia
     rst.MoveNext
  Wend
  Set rst = Nothing
  Close #1
  If existeix("c:\temp\llistat ImpostEnv per factura.csv") Then obrir_document "c:\temp\llistat ImpostEnv per factura.csv"
  Exit Sub
err:
   MsgBox "No es pot escriure en el fitxer CSV. Assegura que o estigui obert.", vbCritical, "Error"
End Sub

Private Sub mllistamuntadorapendent_Click()
  Dim rst As Recordset
  Dim rsti As Recordset
  Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
  Open "c:\llistat.txt" For Output As #1
  Set rst = dbbaixes.OpenRecordset("SELECT * from muntadora_ordremuntatge order by ordre")
  rst.MoveLast
  rst.MoveFirst
  While Not rst.EOF
         Set rsti = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rst!comanda))
         Print #1, " Comanda: " + atrim(rst!comanda) + "  ->  " + atrim(rsti!producte)
         rst.MoveNext
  Wend
  Close 1
  Set rst = Nothing
  'SET DBBAIXES = NOTHING
  Shell "notepad.exe c:\llistat.txt", vbNormalFocus
End Sub

Sub generar_llistat_credit()
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim vnomusuari As String
   Dim vsql As String
   Dim oapp As CRAXDDRT.Application
   Dim oreport As CRAXDDRT.report
   Dim vrisc As TipusVrisc
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
'   vsql = "SELECT comandesmesextres.*, Clients_codiscomptables.nomclient"
'   vsql = vsql + " FROM comandesmesextres LEFT JOIN Clients_codiscomptables ON comandesmesextres.codicomptable = Clients_codiscomptables.codicomptable "
   vsql = "SELECT First(comandesmesextres.client) AS PrimeroDeclient, comandesmesextres.codicomptable, First(Clients_codiscomptables.nomclient) AS PrimeroDenomclient "
   vsql = vsql + " FROM comandesmesextres LEFT JOIN Clients_codiscomptables ON comandesmesextres.codicomptable = Clients_codiscomptables.codicomptable "
   

   vnomusuari = nomordinador
   dbtmp.Execute "delete * from tmp_llistatcreditclients where usuari='" + vnomusuari + "'"
   datainici = demanardata("Entra data Inici per mirar clients que consumeixen.")
   If datainici = "" Then ratoli "normal": Exit Sub
   MsgBox "La consulta pot trigar uns minuts, sigueu pacients." + Chr(10) + "PREMEU ACCEPTAR PER COMENÇAR", vbExclamation, "ATENCIÓ"
   ratoli "espera"
   Set rstc = dbtmp.OpenRecordset("select * from tmp_llistatcreditclients")
   Set rst = dbtmp.OpenRecordset(vsql + " where datacomanda>#" + Format(datainici, "mm/dd/yy") + "#" + " GROUP BY comandesmesextres.codicomptable")
   While Not rst.EOF
     If cadbl(rst!codicomptable) > 0 Then
            calcular_credit_delclient cadbl(rst!codicomptable), vrisc
            rstc.AddNew
            rstc!usuari = vnomusuari
            rstc!client = cadbl(rst!codicomptable)
            rstc!nomclient = atrim(rst!PrimeroDenomclient)
            rstc!creditsap = vrisc.creditsap
            rstc!creditgastatsap = vrisc.creditgastatsap
            rstc!valorestoc = vrisc.valorestoc
            rstc!valorproduccio = vrisc.valorproduccio
            rstc!valordelsclixes = vrisc.valordelsclixes
            rstc!valorpendent = vrisc.valorpendent
            rstc!valordiferencial = vrisc.valordiferencial
            rstc.Update
     End If
     rst.MoveNext
     DoEvents
   Wend
  wait 2
  ratoli "normal"
   
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatcreditclients.rpt", 1)
  oreport.Database.Tables.Item(1).Location = cami
  oreport.RecordSelectionFormula = "{tmp_llistatcreditclients.usuari}='" + vnomusuari + "'"
  oreport.FormulaFields.GetItemByName("datafilte").Text = "'" + datainici + "'"
  
  Load veurereport
  veurereport.CRViewer.ReportSource = oreport
  veurereport.CRViewer.DisplayGroupTree = False
  veurereport.CRViewer.ViewReport
  veurereport.WindowState = 2
  veurereport.Show 1
  
   Set dbclixes = Nothing
   Set rst = Nothing
   Set rstc = Nothing
End Sub
Function demanardata(titol As String) As String
  Dim d As String
  d = "."
  While Not IsDate(d) And d <> ""
    d = InputBox(titol, "Entra una data", Date)
    If Not IsDate(d) Then MsgBox "Data Erronea"
  Wend
  demanardata = d
End Function



Private Sub mllistatcreditEXCEL_Click()
   crear_llistat_credit_EXCEL
End Sub
Sub crear_llistat_credit_EXCEL()
   Dim rst As Recordset
   Dim vlinia As String
   Dim vnomfitxer As String
   vnomfitxer = "c:\temp\llistatcreditclients.csv"
   Set rst = dbtmp.OpenRecordset("select * from clients_codisSAP where valordiferencial<>0 order by valordiferencial Desc")
   Open vnomfitxer For Output As #2
   Print #2, "CodiSAP;Nom del Client;Credit SAP;Credit Gastat;Valor estoc;Valor pendent; Valor producció;Valor dels clixes;Valor dels albarans;Valor diferencial"
   While Not rst.EOF
      vlinia = atrim(rst!codisap) + ";" + atrim(rst!nomclient) + ";" + atrim(rst!creditsap) + ";" + atrim(rst!creditgastatsap) + ";" + atrim(rst!valorestoc) + ";" + atrim(rst!valorpendent) + ";" + atrim(rst!valorproduccio) + ";" + atrim(rst!valordelsclixes) + ";" + atrim(rst!valoralbaranspendentsSAP) + ";" + atrim(rst!valordiferencial)
      Print #2, vlinia
      rst.MoveNext
   Wend
   Close #2
   If existeix(vnomfitxer) Then obrir_document vnomfitxer
   Set rst = Nothing
End Sub

Private Sub mllistatcropspartits_Click()
  Dim vsql As String
  Dim vnumalbara As String
  Dim rst As Recordset
  Dim vlinia As String
  Dim vnomfitxer As String
  
  vnomfitxer = "c:\temp\llistat Crops Contractes Partits.csv"
  
  vsql = "SELECT comandesmesextres.datacomanda, comandesmesextres.comandaclient, comandesmesextres.refclient, comandesmesextres.client, comandesmesextres.comanda, comandesmesextres.refilate AS Fulla, comandesmesextres.proximaseccio, comandesmesextres.producte, comandesmesextres.tubbaseext AS Q_Demanada, [comandaclient] & ' - ' & [refclient] AS [Contracte+Referencia], comandesmesextres.marcailinia From comandesmesextres "
  vsql = vsql + " where (((comandesmesextres.datacomanda) > #1/1/2018#) And ((comandesmesextres.client) = 6841) And ((comandesmesextres.producte) <> 'PC' And (comandesmesextres.producte) <> 'PC2' And (comandesmesextres.producte) <> 'PCP')) ORDER BY [comandaclient] & ' - ' & [refclient], comandesmesextres.refilate;"
  
  Set rst = dbtmp.OpenRecordset(vsql)
  If rst.EOF Then MsgBox "No hi ha informació.", vbCritical, "Error": Exit Sub
  On Error GoTo err
  For i = 0 To rst.Fields.Count - 1
    vlinia = vlinia + IIf(vlinia <> "", ";", "") + UCase(atrim(rst.Fields(i).Name))
  Next i
  Open vnomfitxer For Output As #1
  Print #1, vlinia  'escriu la capçalera
  While Not rst.EOF
     vlinia = ""
     For i = 0 To rst.Fields.Count - 1
       vlinia = vlinia + IIf(vlinia <> "", ";", "") + atrim(rst.Fields(i).Value)
     Next i
     Print #1, vlinia
     rst.MoveNext
  Wend
  Set rst = Nothing
  Close #1
  If existeix(vnomfitxer) Then obrir_document vnomfitxer
  Exit Sub
err:
   MsgBox "No es pot escriure en el fitxer CSV. Assegura que o estigui obert.", vbCritical, "Error"
End Sub

Private Sub mllistatimpenvalb_Click()
 Dim vsql As String
  Dim vnumalbara As String
  Dim rst As Recordset
  Dim vlinia As String
  
  vnumalbara = InputBox("Escriu el numero d'ALBARÀ SAP que vols detallar la informació de l'Impost de envasos i preu/unitat(KG)." + vbNewLine + "POTS POSAR ALBARANS SEPARATS PER COMES", "ALBARÀ")
  If StrPtr(vnumalbara) = 0 Then Exit Sub
  If atrim(vnumalbara) = "" Then Exit Sub
  vnumalbara = substituir(vnumalbara, ";", ",")
  vnumalbara = substituir(vnumalbara, ".", ",")
  vsql = "SELECT capcaleraalbara.numalbaraSAP, liniesalbara.numcomandacli, liniesalbara.refclient, liniesalbara.marcailinia, ([Kgimpostenvasos])*[eurokg_impost] AS EurosKgImpost, [quantitat]*[preuvenda] AS Eurosproducte, ([EurosKgImpost]+[Eurosproducte])/[kgtotalsbruts] AS PreuKg "
  vsql = vsql + " FROM liniesalbara INNER JOIN capcaleraalbara ON liniesalbara.numalbara = capcaleraalbara.numalbara "
  vsql = vsql + " WHERE (((capcaleraalbara.numalbaraSAP) IN (" + vnumalbara + ")));"
  'Clipboard.Clear
'  Clipboard.SetText vsql
  Set rst = dbtmp.OpenRecordset(vsql)
  If rst.EOF Then MsgBox "No hi ha informació d'aquest ALBARÀ.", vbCritical, "Error": Exit Sub
  On Error GoTo err
  Open "c:\temp\llistat ImpostEnv_i_TANXKG per albarà.csv" For Output As #1
  Print #1, "Nº ALBARAN;Nº PEDIDO;CÓDIGO ARTÍCULO;DESCRIPCIÓN ARTÍCULO;IMPORTE IMPUESTO PLÁSTICO;IMPORTE PRODUCTO;IMPORTE/KG(IMPUESTO INCLUIDO)"
  While Not rst.EOF
     vlinia = atrim(rst!numalbaraSAP) + ";" + atrim(rst!numcomandacli) + ";" + atrim(rst!refclient) + ";" + atrim(rst!marcailinia) + ";" + atrim(Redondejar(cadbl(rst!EurosKgImpost), 2)) + ";" + atrim(Redondejar(cadbl(rst!Eurosproducte), 2)) + ";" + atrim(Redondejar(cadbl(rst!PreuKg), 2))
     Print #1, vlinia
     rst.MoveNext
  Wend
  Set rst = Nothing
  Close #1
  If existeix("c:\temp\llistat ImpostEnv_i_TANXKG per albarà.csv") Then obrir_document "c:\temp\llistat ImpostEnv_i_TANXKG per albarà.csv"
  Exit Sub
err:
   MsgBox "No es pot escriure en el fitxer CSV. Assegura que o estigui obert.", vbCritical, "Error"
End Sub

Private Sub mllistatrefenestoc_Click()
  
  Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe " + fitxerini + " llistatreferencies", vbNormalFocus
  
End Sub

Private Sub mmuntadora_Click()
  Shell """\\serverprodu\dades\progcomandes\aplicacio\baixes muntadora.exe""", vbNormalFocus
End Sub

Private Sub mnomesunclient_Click()
  Dim vcodicomptable As String
  vcodicomptable = escullir_codicomptable
  If Len(vcodicomptable) > 5 Then
     Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
     informe_credit_unclient vcodicomptable
     Set dbclixes = Nothing
  End If
End Sub
Function escullir_codicomptable() As String
    Load formseleccio
  'formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select distinct codicomptable,nomclient from clients_codiscomptables"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(0).Width = 1500
  formseleccio.DBGrid2.Columns(1).Width = 4000
  'formseleccio.Width = 9000
  'formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           escullir_codicomptable = formseleccio.DBGrid2.Columns("codicomptable")
        End If
   End If
End Function

Private Sub mordreimpresio_Click()
  Shell "\\serverprodu\dades\progcomandes\aplicacio\baixesimpresoramaquina.exe ORDREIMPRESSIO '9' OFICINA", vbNormalFocus

End Sub

Private Sub mpeucomanda_Click()
  missatgespeucomandescompra.Show 1
End Sub

Private Sub mpeuimprenta_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment de peu d'imprenta i data"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from peuimprenta"
  formaltarep.Width = 11000
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.DBGrid1.Columns(1).Width = 3000
  formaltarep.DBGrid1.Columns(2).Width = 6500
  
  formaltarep.Show
End Sub

Private Sub mplanificacio_Click()
 
    Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "planificacio.exe", vbNormalFocus
 
End Sub

Private Sub mrelaciocomandes_Click()
   formllistatreferencies.Show 1
End Sub
Function possargrmm2(vcodimat As Double, vmicres As Double, vnommaterial As String) As Double
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(vcodimat))
   If Not rst.EOF Then
       possargrmm2 = vmicres * rst!grmcm3
       vnommaterial = atrim(rst!descripcio)
   End If
   Set rst = Nothing
End Function

Private Sub mrelaciorefpes_Click()
   Dim vcodi As Double
   Dim vnomclient As String
   Dim rst As Recordset
   Dim rstalb As Recordset
   Dim vlinia As String
   Dim vnomfitxer As String
   Dim vpesbobina As Double
   Dim vpesmetre As Double
   Dim vpespalet As Double
   Dim rstc As Recordset
   Dim vnommaterial As String
   Dim vgrmm2 As Double
   Dim vmicres As Double
   Dim vlink1 As Double
   Dim vlink2 As Double
   
   Set dbtmpb = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
   vnomfitxer = "c:\temp\llistat_relacio_referencies_Kg.csv"
   escullir_client vcodi, vnomclient
   If vcodi = 0 Then Exit Sub
   Open vnomfitxer For Output As #2
   vlinia = "Lot inplacsa;Ref Inplacsa;Ref Client;Espessor;Descripció mides;Marca i linia;Mtrs bob teòric;Ample mm;Kg Metre;Kg Bobina;Kg Palet;Espessor C1;Grm2 C1;Material C1;Espessor C2;Grmm2 C2;Material C2;Espessor C3;Grmm2 C3;Material C3"
   Print #2, vlinia
   Set rstc = dbtmp.OpenRecordset("select * from comandes")
   Set rst = dbtmp.OpenRecordset("select distinct refinplacsa from tarifes_referencies where codiclient='" + atrim(vcodi) + "' and not inactiva")
   While Not rst.EOF
     Set rstalb = dbtmp.OpenRecordset("select * from liniesalbara where codiproducte='" + atrim(rst!refinplacsa) + "' order by id desc")
     rstc.FindFirst "comanda=" + atrim(cadbl(rstalb!lotinplacsa))
     vlinia = ""
     vpesbobina = 0
     vpesmetre = 0
     vpespalet = 0
     'If rst!refinplacsa = "02C6908I5392" Then Stop
     If Not rstalb.EOF Then
       If cadbl(rstalb!numbobs) > 0 Then vpesbobina = buscar_pesmaxbobina(rstalb!lotinplacsa)
           'vpesbobina = (cadbl(rstalb!kgtotalsbruts) - cadbl(rstalb!pespalets)) / cadbl(rstalb!numbobs)
       If cadbl(rstalb!metreslineals) > 0 Then vpesmetre = cadbl(rstalb!kgtotalsbruts) / cadbl(rstalb!metreslineals)
       vpespalet = vpesbobina * buscar_bobines_per_palet(rst!refinplacsa)
       vlinia = atrim(rstalb!lotinplacsa) + "   ;" + atrim(rst!refinplacsa) + ";" + atrim(rstalb!refclient) + ";" + atrim(rstalb!espesor) + " " + atrim(rstalb!mesuraespesor) + ";" + atrim(rstalb!descripciomides) + ";" + atrim(rstalb!marcailinia) + ";" + atrim(rstc!mtrslinbob) + ";" + atrim(rstalb!ampladamaterial) + ";" + atrim(Redondejar(vpesmetre, 4)) + ";" + atrim(Redondejar(vpesbobina, 1)) + ";" + atrim(Redondejar(vpespalet, 2))
       'capa1
       vmicres = micresmaterial(cadbl(rstc!mesuraesp), cadbl(rstc!espessor), atrim(rstc!tubolam))
       vgrmm2 = possargrmm2(rstc!materialex, vmicres, vnommaterial)
       vlinia = vlinia + ";" + atrim(vmicres) + ";" + atrim(vgrmm2) + ";" + atrim(vnommaterial)
       'capa2
       vlink1 = cadbl(rstc!linkcomanda1): vlink2 = cadbl(rstc!linkcomanda2)
       If cadbl(vlink1) > 0 Then
            rstc.FindFirst "comanda=" + atrim(cadbl(rstc!linkcomanda1))
            If Not rstc.NoMatch Then
             vmicres = micresmaterial(cadbl(rstc!mesuraesp), cadbl(rstc!espessor), atrim(rstc!tubolam))
             vgrmm2 = possargrmm2(cadbl(rstc!materialex), vmicres, vnommaterial)
             vlinia = vlinia + ";" + atrim(vmicres) + ";" + atrim(vgrmm2) + ";" + atrim(vnommaterial)
            End If
       End If
       'capa3
       If cadbl(vlink2) > 0 Then
            rstc.FindFirst "comanda=" + atrim(cadbl(rstc!linkcomanda2))
            If Not rstc.NoMatch Then
                vmicres = micresmaterial(cadbl(rstc!mesuraesp), cadbl(rstc!espessor), atrim(rstc!tubolam))
                vgrmm2 = possargrmm2(rstc!materialex, vmicres, vnommaterial)
                vlinia = vlinia + ";" + atrim(vmicres) + ";" + atrim(vgrmm2) + ";" + atrim(vnommaterial)
            End If
       End If
       Print #2, vlinia
     End If
     rst.MoveNext
   Wend
   Close #2
   If existeix(vnomfitxer) Then obrir_document vnomfitxer
   Set rst = Nothing
   Set rstalb = Nothing
   
End Sub
Function micresmaterial(codimesuralineal As Byte, espesor As Double, tubolam As String) As Double
  Dim rstmesural As Recordset
  Set rstmesural = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(codimesuralineal), dbOpenSnapshot, dbReadOnly)
  If rstmesural.EOF Then Exit Function
  r = espesor
  If rstmesural!descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = Format(espesor / 4, "#,##0")
                  Else: r = Format(espesor / 2, "#,##0")
            End If
  End If
  If InStr(1, rstmesural!descripcio, "GR/") > 0 Then
    r = espesor * -1
  End If
  micresmaterial = r
End Function

Function buscar_pesmaxbobina(vnumc As Double) As Double
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select max(kilosiunitats) as Tkg from bobinesent where comanda=" + atrim(cadbl(vnumc)))
   buscar_pesmaxbobina = cadbl(rst!Tkg)
   Set rst = Nothing
End Function
Function buscar_bobines_per_palet(vrefinplacsa As String) As Double
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim vnumpalets As Double
   Dim vbobmesgran As Double
   Dim vcomandabobmesgran As Double
   Set rst2 = dbtmp.OpenRecordset("select comanda from comandes_extres where refinplacsa='" + atrim(vrefinplacsa) + "'")
   vnumpalets = 0
   While vnumpalets < 2 And Not rst2.EOF
      Set rst = dbtmpb.OpenRecordset("select comanda,numalbara,numbob,numpalet from bobinesent where comanda=" + atrim(cadbl(rst2!comanda)) + " order by numpalet desc")
      If Not rst.EOF Then
         vnumpalets = cadbl(rst!numpalet)
         If vnumpalets = 0 Then vnumpalets = 1
         If vbobmesgran < cadbl(rst!numbob) Then vbobmesgran = rst!numbob: vcomandabobmesgran = rst2!comanda
      End If
      rst2.MoveNext
   Wend
   If rst.EOF Then Set rst = dbtmpb.OpenRecordset("select comanda,numalbara,numbob from bobinesent where comanda=" + atrim(vcomandabobmesgran) + " order by numpalet desc")
   'Set rst = dbtmpb.OpenRecordset("select comanda,numalbara from bobinesent where numpalet>1 and comanda in (select comanda from comandes_extres where refinplacsa='" + atrim(vrefinplacsa) + "') order by data desc")
   'If rst.EOF Then
   '   Set rst = dbtmpb.OpenRecordset("select comanda,numalbara from bobinesent where comanda in (select comanda from comandes_extres where refinplacsa='" + atrim(vrefinplacsa) + "') order by data desc,numbob desc")
   'End If
   If Not rst.EOF Then
       Set rst = dbtmp.OpenRecordset("select count(*) as Tbobines from bobinesent where comanda=" + atrim(rst!comanda) + " group by numpalet order by count(*) desc ")
       If Not rst.EOF Then buscar_bobines_per_palet = cadbl(rst!Tbobines)
       If buscar_bobines_per_palet = 0 Then buscar_bobines_per_palet = vbobmesgran
   End If
   Set rst = Nothing
End Function
Sub escullir_client(vcodi As Double, vnomclient As String)
    Load formseleccio
  'formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,nom from clients"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(0).Width = 1500
  formseleccio.DBGrid2.Columns(1).Width = 4000
  'formseleccio.Width = 9000
  'formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           vcodi = formseleccio.DBGrid2.Columns("codi")
           vnomclient = formseleccio.DBGrid2.Columns("nom")
        End If
   End If
End Sub
Private Sub mrepasdeclixes_Click()
   Shell "\\serverprodu\dades\progcomandes\aplicacio\repas de clixes.exe", vbNormalFocus
End Sub

Private Sub mrevcq_Click()
    Shell "\\serverprodu\dades\progcomandes\aplicacio\ControlQualitat.exe", vbNormalFocus
End Sub

Private Sub mrevisarescaneig_Click()
   formrevisatescanejats.Show 1
End Sub

Private Sub msubstancies_Click()

  Load formaltarep
  formaltarep.colsbloc = "4"
  formaltarep.Caption = "Manteniment de substancies"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from  substancies"
  formaltarep.Width = formaltarep.Width * 2.2
  formaltarep.DBGrid1.Width = (formaltarep.DBGrid1.Width * 2.25)
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(0).Width = 800
  formaltarep.DBGrid1.Columns(1).Width = 1500
  formaltarep.DBGrid1.Columns(2).Width = 7500
  formaltarep.DBGrid1.Columns(3).Width = 800
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub mtarifesref_Click()
formtarifesperreferencia.Show 1
End Sub

Private Sub mtintes_Click()
   Shell "\\serverprodu\dades\progcomandes\aplicacio\manteniment tintes.exe", vbNormalFocus
End Sub

Private Sub mtipusetreb_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment de Tipus Etiquetes Rebobinadora"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from tipusetiquetes"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Show
End Sub

Private Sub mtipusiva_Click()
  Dim iva As Double
  iva = cadbl(llegir_ini("General", "iva", fitxerini))
  If InputBox("Aquesta opcio es per canviar el tipus d'IVA de treball de compres i vendes." + Chr(10) + Chr(13) + "Estas segur que vols canviar-lo? (Escriu SI per fer-ho)", "Atenció") = "SI" Then
     iva = cadbl(InputBox("Canviar tipus d'iva de compres i vendes." + Chr(10) + Chr(13) + "OJU AMB EL CANVI D'IVA QUE AFECTARÀ LES COMPRES I VENDES", "Atenció", atrim(iva)))
     If iva > 0 Then
        escriure_ini "General", "iva", atrim(iva), fitxerini
     End If
  End If
End Sub

Private Sub mtipuspalets_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment de Tipus Palets"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from tipuspalets"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Show
End Sub

Private Sub mtipuspaperfrontal_Click()
  Load formaltarep
  formaltarep.Caption = "Manteniment de Tipus Papers frontals"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from tipuspaperfrontal"
  formaltarep.refrescar
  formaltarep.Tag = "tipuspaperfrontal"
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Show
End Sub

Private Sub mtipusproteccions_Click()
Load formaltarep
  formaltarep.Caption = "Manteniment de Tipus Proteccions"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from tipusproteccions"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.Show
End Sub

Private Sub mtm_Click()

  Load formaltarep
  formaltarep.Caption = "Manteniment de Tolerancies de Màquina"
  'formaltarep.colsbloc = "2"
  formaltarep.Data1.DatabaseName = llegir_ini("General", "camibaixes", fitxerini)
  formaltarep.Data1.RecordSource = "select * from toleranciesmaquina"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.DBGrid1.AllowAddNew = False
  formaltarep.DBGrid1.Columns(0).Caption = "Tintes"
  formaltarep.DBGrid1.Columns(0).Width = 900
  formaltarep.DBGrid1.Columns(0).Alignment = 2
  formaltarep.DBGrid1.Columns(0).Locked = True
  
  formaltarep.DBGrid1.Columns(1).Caption = "Metres"
  formaltarep.DBGrid1.Columns(1).Width = 900
  formaltarep.DBGrid1.Columns(1).Alignment = 2
  
  formaltarep.DBGrid1.Columns(2).Caption = "Tolerancia"
  formaltarep.DBGrid1.Columns(2).Width = 1500
  formaltarep.DBGrid1.Columns(2).Alignment = 2
  formaltarep.alta.Enabled = False
  formaltarep.eliminar.Enabled = False
  
  'formaltarep.Width = (formaltarep.Width * 2) + 200
  'formaltarep.DBGrid1.Width = (formaltarep.DBGrid1.Width * 2) + 400
  
  formaltarep.Show
  
End Sub

Private Sub mtotselsclients_Click()
   generar_llistat_credit
End Sub

Private Sub mtractamentcares_Click()
  Load formaltarep
  formaltarep.Caption = "Manteniment tractament cares del material"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from tractamentcares order by descripcio"
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.DBGrid1.Columns(1).Width = 150 * 50
  formaltarep.Width = formaltarep.DBGrid1.Columns(1).Width + 2000
  formaltarep.autonum = "tractamentcares"
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub mvalorimpostenvasos_Click()
   Dim v As String
   v = atrim(cadbl(llegir_ini("General", "PreuImpostEnvasos", rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini")))
   v = InputBox("Escriu el valor €/Kg de l'Impost sobre els Envasos.", "Valor Impost Envasos", v)
   If StrPtr(v) = 0 Or cadbl(v) = 0 Then Exit Sub
   If cadbl(v) > 2 Then MsgBox "Aquest valor es massa gran. " + atrim(cadbl(v)): Exit Sub
   If UCase(InputBox("Segur que vols canviar el valor d'aquest IMPOST?" + vbNewLine + "ESCRIU [ACCEPTO] PER APLICAR EL CANVI.", "CANVI VALOR IMPOST")) = "ACCEPTO" Then
     escriure_ini "General", "PreuImpostEnvasos", atrim(v), rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "valorsprograma.ini"
     MsgBox "Canvi Fet.  " + atrim(cadbl(v)) + " €/Kg"
   End If
End Sub

Private Sub representants_Click()
  
  Load formaltarep
  formaltarep.Caption = "Manteniment de Representats"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from representants"
  formaltarep.Width = 6500
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.Show 1
End Sub

Private Sub sortir_Click()
 volssortir
End Sub

Private Sub sortirs_Click()
  volssortir
End Sub

Private Sub subfaditius_Click()
 Load formaltarep
  formaltarep.Caption = "Manteniment SubFamilies d'Aditius"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select codi,codifam as [Familia],descripcio from subfamiliesaditius"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
  formaltarep.Width = 6100
End Sub

Private Sub subfamcol_Click()
 Load formaltarep
  formaltarep.Caption = "Manteniment SubFamilies Colorants"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select codi,codifam as [Familia],descripcio from subfamiliescolorants"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
  formaltarep.Width = 6100
End Sub

Private Sub subfammaterials_Click()
  Load formaltarep
  formaltarep.Caption = "Manteniment SubFamilies Materials"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select codi,codifam as [Familia],descripcio from subfamiliesmaterials"
  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
  formaltarep.Width = 6100
End Sub

Private Sub Timer1_Timer()
 
'If hora <> Now Then hora = Now
  controldeteclat
  canviarelscolorsdelscontrolsalentrar
  
End Sub
Private Sub Timer2_Timer()
  
mirarsiparar
End Sub
Sub mirarsiparar()
 Static contar
 Dim paraula As String
 paraula = llegir_ini("General", "parar", llegir_ini("General", "rutallistats", fitxerini) + "parar.ini")
  If paraula = "si" Or InStr(1, paraula, "[comandes]") > 0 Then
    contar = contar + 1
     If contar = 1 Then MsgBox2 "El programa es pararà d'aqui a 1 minut. TANCA TOT I ESPERA CINC MINUTS.", 5, "Actualització", vbCritical
     If contar = 15 Then MsgBox2 "El programa es pararà d'aqui a 30 segons. TANCA TOT I ESPERA CINC MINUTS.", 5, "Actualització, vbCritical"
     If contar = 27 Then MsgBox2 "El programa es pararà d'aqui a 5 segons. TANCA TOT I ESPERA CINC MINUTS.", 3, "Actualització", vbCritical
     If contar > 30 Then End
   Else: contar = 0
  End If
  If paraula = "ja" Then End
End Sub

Private Sub tintescomandesafectades_Click()
  Dim numerodelot As String
  Dim db As Database
  Dim db2 As Database
  Dim dbtintes As Database
  Dim were As String
  Dim rsttintes As Recordset
  Dim rsttmp2 As Recordset
  Dim rstclient As Recordset
  Dim vnomtinta As String
  Dim taulatemp As String
  numerodelot = InputBox("Entra el numero de lot de Tinta que vols buscar:", "Lot de Tinta")
  If atrim(numerodelot) = "" Then Exit Sub
  taulatemp = "c:\temporal.mdb"
  ratoli "espera"
 ' Me.Caption = "Processant... "
  If Not existeix(taulatemp) Then DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  Set db = OpenDatabase(cami)
  Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  On Error Resume Next
  Set db2 = OpenDatabase(taulatemp)
  db2.Execute ("drop table llistatlots")
  db2.Execute ("create table llistatlots (comanda double,pantone  string,codiclient double,nomclient string)")
  On Error GoTo 0
  Set rsttmp2 = db2.OpenRecordset("llistatlots")
  were = "(lot1 like '*" + numerodelot + "*')"
  For i = 2 To 10
    were = were + " or (lot" + Trim(i) + " like '*" + numerodelot + "*')"
  Next i

  Set rsttmp = dbbaixes.OpenRecordset("select * from impresorespantones where " + were)
  While Not rsttmp.EOF
    Set rstclient = db.OpenRecordset("select client from comandes where comanda=" + atrim(cadbl(rsttmp!comanda)))
    If Not rstclient.EOF Then
      Set rstclient = db.OpenRecordset("select codi,nom from clients where codi=" + atrim(cadbl(rstclient!client)))
      If Not rstclient.EOF Then
        rsttmp2.AddNew
         rsttmp2!comanda = rsttmp!comanda
         rsttmp2!codiclient = rstclient!codi
         rsttmp2!pantone = escullirpantone(rsttmp, numerodelot)
         rsttmp2!nomclient = rstclient!nom

         If rsttmp2!pantone = "" Then
             rsttmp2.CancelUpdate
           Else:
                      rsttmp2!pantone = "": rsttmp2.Update
         End If

      End If
    End If
    rsttmp.MoveNext
  Wend
  Set rsttintes = dbtintes.OpenRecordset("select * from dadesllaunestotes where numllauna='" + atrim(numerodelot) + "'", , ReadOnly)
  If Not rsttintes.EOF Then vnomtinta = atrim(rsttintes!descripcio)
  r = "Comandes afectades pel lot de tinta: " + numerodelot
  llistat.DataFiles(0) = taulatemp
  llistat.WindowState = crptMaximized
  llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "comandesxlot.rpt"
  llistat.Formulas(0) = "nomdelllistat=" + "'" + r + "'"
  llistat.Formulas(1) = "nomlottinta=" + "'" + vnomtinta + "'"
  llistat.Action = 1
  
  
  ratoli "normal"
  Set db = Nothing
  Set db2 = Nothing
  'SET DBBAIXES = NOTHING
  Set dbtintes = Nothing
  Set rsttmp = Nothing
  Set rsttmp2 = Nothing
  Set rstclient = Nothing
  Set rsttintes = Nothing
End Sub
Function escullirpantone(rsttmp As Recordset, numerodelot As String) As String
  Dim i As Byte
  numerodelot = UCase(numerodelot)
  For i = 1 To 8
    If InStr(1, UCase(rsttmp.Fields("lot" + atrim(i)) + "+"), numerodelot + "+") > 0 Then escullirpantone = rsttmp.Fields("pantone" + atrim(i)): GoTo fi
  Next i
fi:
End Function

Private Sub tipusentregues_Click()
 Load formaltarep
  formaltarep.Caption = "Manteniment Tipus Entregues"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from tipusentregues"
  formaltarep.refrescar
  formaltarep.Width = formaltarep.Width * 1.5
  formaltarep.DBGrid1.Width = formaltarep.DBGrid1.Width * 1.5
  formaltarep.DBGrid1.Columns(1).Width = 150 * 35
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub tipussoldadures_Click()
  Load formaltarep
  formaltarep.Caption = "Manteniment Tipus Soldadures"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from tipussoldadura"
  formaltarep.refrescar
  formaltarep.Width = 8000
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub transportistes_Click()
  Load formaltarep
  formaltarep.Caption = "Transportistes"
'  formaltarep.autonum = "transportistes"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from transportistes"

  formaltarep.refrescar
  formaltarep.DBGrid1.Refresh
  
  formaltarep.Width = formaltarep.Width + 14200
  formaltarep.DBGrid1.Width = formaltarep.DBGrid1.Width + 14000
  formaltarep.DBGrid1.Columns(1).Width = 3500
  formaltarep.DBGrid1.Columns(2).Width = 3000
  formaltarep.DBGrid1.Columns(3).Width = 3000
  formaltarep.DBGrid1.Columns(4).Width = 3000
  formaltarep.DBGrid1.Columns(5).Width = 1500
  formaltarep.DBGrid1.Columns(6).Width = 1000
  
  
  formaltarep.Show
  formaltarep.Left = (Screen.Width / 2) - (formaltarep.Width / 2)
End Sub

Private Sub tubbase_Click()
  Load formaltarep
  formaltarep.Caption = "Manteniment Tubs Base"
  formaltarep.autonum = "tubbase"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from tubbase"
  formaltarep.Width = Menu.Width
  formaltarep.refrescar
  
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub unitats_Click()
  Load formaltarep
  formaltarep.Caption = "Manteniment Mesures"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from mesures"
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(1).Width = 150 * 15
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub unitatslineals_Click()
  Load formaltarep
  formaltarep.Caption = "Manteniment Mesures Lineals"
  formaltarep.Data1.DatabaseName = cami
  formaltarep.Data1.RecordSource = "select * from mesureslineals"
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(1).Width = 150 * 15
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub


