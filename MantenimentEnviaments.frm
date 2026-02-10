VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form formenviaments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviaments de material i paqueteria variada."
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   18750
   Icon            =   "MantenimentEnviaments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   18750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Framefactures 
      Caption         =   "Factures transportistes"
      Height          =   6690
      Left            =   930
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   18705
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   13875
         Top             =   1365
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Data datacmrs 
         Caption         =   "datacmrs"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4515
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "transportistes_factures_CMR"
         Top             =   5835
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Data datafactures 
         Caption         =   "datafactures"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "transportistes_factures"
         Top             =   5835
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FDDECE&
         Caption         =   "CMRs Relacionats"
         Height          =   5460
         Left            =   7950
         TabIndex        =   35
         Top             =   360
         Width           =   4185
         Begin VB.CommandButton belimninarc 
            Height          =   450
            Left            =   1095
            Picture         =   "MantenimentEnviaments.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "Eliminacio Registres"
            Top             =   195
            Width           =   465
         End
         Begin VB.CommandButton Command7 
            Height          =   420
            Left            =   660
            Picture         =   "MantenimentEnviaments.frx":0B14
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Afegir factura del proveidor"
            Top             =   225
            Width           =   435
         End
         Begin VB.CommandButton Command5 
            Height          =   420
            Left            =   150
            Picture         =   "MantenimentEnviaments.frx":109E
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   225
            Width           =   435
         End
         Begin MSDBGrid.DBGrid DBGrid3 
            Bindings        =   "MantenimentEnviaments.frx":1628
            Height          =   4665
            Left            =   135
            OleObjectBlob   =   "MantenimentEnviaments.frx":163B
            TabIndex        =   37
            Top             =   675
            Width           =   3765
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Factures Transportista"
         Height          =   5460
         Left            =   195
         TabIndex        =   32
         Top             =   360
         Width           =   6285
         Begin VB.CommandButton Command8 
            Caption         =   "CMRs Pendents"
            Height          =   405
            Left            =   4605
            TabIndex        =   43
            Top             =   165
            Width           =   1575
         End
         Begin VB.CommandButton blinkpdf 
            BackColor       =   &H0025EFAD&
            Caption         =   "Arrastra PDF"
            Height          =   675
            Left            =   4515
            OLEDropMode     =   1  'Manual
            Picture         =   "MantenimentEnviaments.frx":2036
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1125
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.CommandButton beliminarf 
            Height          =   420
            Left            =   1110
            Picture         =   "MantenimentEnviaments.frx":25C0
            Style           =   1  'Graphical
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   "Eliminacio Registres"
            Top             =   225
            Width           =   435
         End
         Begin VB.CommandButton Command6 
            Height          =   420
            Left            =   660
            Picture         =   "MantenimentEnviaments.frx":2B4A
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Afegir factura del proveidor"
            Top             =   225
            Width           =   435
         End
         Begin VB.CommandButton Command4 
            Height          =   420
            Left            =   150
            Picture         =   "MantenimentEnviaments.frx":30D4
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   225
            Width           =   435
         End
         Begin MSDBGrid.DBGrid DBGrid2 
            Bindings        =   "MantenimentEnviaments.frx":365E
            Height          =   4665
            Left            =   165
            OleObjectBlob   =   "MantenimentEnviaments.frx":3675
            TabIndex        =   33
            Top             =   645
            Width           =   5835
         End
      End
      Begin VB.Image Image1 
         Height          =   660
         Left            =   6630
         Picture         =   "MantenimentEnviaments.frx":43CC
         Stretch         =   -1  'True
         Top             =   2655
         Width           =   1065
      End
   End
   Begin VB.Frame frametarifes 
      Caption         =   "Tarifes Transportistes"
      Height          =   6690
      Left            =   345
      TabIndex        =   3
      Top             =   -90
      Visible         =   0   'False
      Width           =   18705
      Begin VB.CommandButton Command3 
         Height          =   330
         Left            =   1365
         Picture         =   "MantenimentEnviaments.frx":4BB6
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Enganxar desde portapapers"
         Top             =   1875
         Width           =   570
      End
      Begin VB.ListBox llistatarifes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5925
         Left            =   6165
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   405
         Width           =   5385
      End
      Begin VB.CommandButton modificar 
         Height          =   330
         Left            =   772
         Picture         =   "MantenimentEnviaments.frx":5140
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Modificació Registres"
         Top             =   1875
         Width           =   570
      End
      Begin VB.CommandButton btreurecomanda 
         Height          =   330
         Left            =   5100
         Picture         =   "MantenimentEnviaments.frx":56CA
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Eliminació Registres"
         Top             =   1875
         Width           =   645
      End
      Begin VB.CommandButton alta 
         Height          =   330
         Left            =   180
         Picture         =   "MantenimentEnviaments.frx":5C54
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   1875
         Width           =   570
      End
      Begin VB.Data dataports 
         Caption         =   "dataports"
         Connect         =   "Access"
         DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\vendes.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   6390
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tarifes_ports"
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin MSDBGrid.DBGrid reixa 
         Bindings        =   "MantenimentEnviaments.frx":61DE
         Height          =   4230
         Left            =   195
         OleObjectBlob   =   "MantenimentEnviaments.frx":61F2
         TabIndex        =   11
         Top             =   2265
         Width           =   5625
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FDDECE&
         Height          =   1665
         Left            =   135
         TabIndex        =   4
         Top             =   195
         Width           =   5910
         Begin VB.Frame Frame3 
            BackColor       =   &H00EEE4D7&
            Height          =   795
            Left            =   1605
            TabIndex        =   20
            Top             =   0
            Width           =   4215
            Begin VB.TextBox credondeigeurokg 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   30
               TabIndex        =   29
               Text            =   "0"
               Top             =   450
               Width           =   465
            End
            Begin VB.TextBox cdatafifuel 
               Height          =   285
               Left            =   3120
               TabIndex        =   27
               Top             =   150
               Width           =   1035
            End
            Begin VB.TextBox cdatainicifuel 
               Height          =   285
               Left            =   1875
               TabIndex        =   25
               Top             =   150
               Width           =   1035
            End
            Begin VB.TextBox cfuel 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   30
               TabIndex        =   21
               Text            =   "0"
               Top             =   135
               Width           =   465
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Redondeig kilos €/Kg"
               Height          =   285
               Left            =   555
               TabIndex        =   30
               Top             =   465
               Width           =   1575
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "a"
               Height          =   195
               Left            =   2970
               TabIndex        =   26
               Top             =   165
               Width           =   270
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "% de fuel vàlid de:"
               Height          =   285
               Left            =   570
               TabIndex        =   22
               Top             =   165
               Width           =   1575
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00EAD9CE&
            Height          =   795
            Left            =   105
            TabIndex        =   19
            Top             =   0
            Width           =   1425
            Begin VB.TextBox cseguro 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   75
               TabIndex        =   23
               Text            =   "0"
               Top             =   150
               Width           =   360
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "% seguro"
               Height          =   285
               Left            =   600
               TabIndex        =   24
               Top             =   165
               Width           =   765
            End
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "MantenimentEnviaments.frx":6BCB
            Left            =   3825
            List            =   "MantenimentEnviaments.frx":6BD2
            TabIndex        =   9
            Top             =   1185
            Width           =   1650
         End
         Begin VB.ComboBox Combopais 
            Height          =   315
            Left            =   1185
            TabIndex        =   7
            Top             =   1200
            Width           =   1650
         End
         Begin VB.ComboBox Combotransportista 
            Height          =   315
            Left            =   1200
            TabIndex        =   5
            Top             =   855
            Width           =   4305
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Comença per. Ex: 88"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   3855
            TabIndex        =   12
            Top             =   1470
            Width           =   1650
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Codi Postal:"
            Height          =   195
            Left            =   2940
            TabIndex        =   10
            Top             =   1245
            Width           =   1050
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Pais d'envio:"
            Height          =   195
            Left            =   150
            TabIndex        =   8
            Top             =   1245
            Width           =   1050
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Transportista:"
            Height          =   195
            Left            =   180
            TabIndex        =   6
            Top             =   900
            Width           =   1050
         End
      End
      Begin VB.Label ettanperkilo 
         BackStyle       =   0  'Transparent
         Caption         =   "Si el preu es negatiu serà per Kg."
         Height          =   225
         Left            =   2115
         TabIndex        =   28
         Top             =   1980
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.Label Label5 
         Caption         =   "Llistat de tarifes entrades"
         Height          =   180
         Left            =   6210
         TabIndex        =   17
         Top             =   210
         Width           =   3345
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   210
      Picture         =   "MantenimentEnviaments.frx":6BDC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exportar a Excel"
      Top             =   15
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   195
      Picture         =   "MantenimentEnviaments.frx":7166
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Filtrar dades de la reixa."
      Top             =   390
      Width           =   540
   End
   Begin VB.Data dataenviaments 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "registre_enviaments"
      Top             =   330
      Visible         =   0   'False
      Width           =   2250
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "MantenimentEnviaments.frx":76F0
      Height          =   5640
      Left            =   150
      OleObjectBlob   =   "MantenimentEnviaments.frx":7709
      TabIndex        =   0
      Top             =   810
      Width           =   18465
   End
   Begin VB.Menu mtarifes 
      Caption         =   "Tarifes Transportista"
   End
   Begin VB.Menu mfacturestransport 
      Caption         =   "Factures transportista"
   End
End
Attribute VB_Name = "formenviaments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vcamp As String
Dim vvalor As String

Private Sub alta_Click()
  Dim v As String
  Dim i As Byte
  Dim vsalt As Double
  Dim vultimvalor As Double
  Dim vPaletsoKilos As String
  Dim vidtransport As Double
  Dim vsql As String
  Dim vdesctarifa As String
  
  If Combotransportista = "" Or Combopais = "" Or Combo1 = "" Then MsgBox "Primer s'ha de carregar les dades de busqueda", vbCritical, "Error": Exit Sub
  vPaletsoKilos = UCase(InputBox("Escriu si vols tarifa de Kilos o de Palets [K] o [P].", "Tarifa de Kilos o Palets", IIf(Not dataports.Recordset.EOF, dataports.Recordset!tarifaperpaletsokg, "K")))
  If vPaletsoKilos <> "P" And vPaletsoKilos <> "K" Then MsgBox "Aquesta lletra no es vàlida", vbCritical, "Error": Exit Sub
 ' vPaletsoKilos = "K"
  v = InputBox("Entra el numero de linies que vols generar tarifa.", "Palets")
  If StrPtr(v) = 0 Then Exit Sub
  If cadbl(v) = 0 Then Exit Sub
  If cadbl(v) > 50 Then MsgBox "No es pot crear una tarifa de mes de 50 palets/Kilos.": Exit Sub
  vsalt = 1
  If vPaletsoKilos = "K" Then
      If v = 1 Then
            vsalt = cadbl(InputBox("Escriu els Kilos que vols possar a la nova linia.", "Assignar kilos a la nova linia", 50))
           Else:
             vsalt = cadbl(InputBox("Escriu de quants Kilos vols que sigui el salt.", "Salt", 50))
      End If
        
  End If
  If vsalt = 0 Then GoTo fi
  If Not dataports.Recordset.EOF Then
       If dataports.Recordset!tarifaperpaletsokg = vPaletsoKilos Then dataports.Recordset.MoveLast: vultimvalor = dataports.Recordset!numpalets
  End If
  vidtransport = cadbl(Combotransportista.ItemData(Combotransportista.ListIndex))
  For i = 1 To cadbl(v)
     If v > 1 Then dataports.Recordset.FindFirst "numpalets=" + atrim((i * vsalt) + vultimvalor)
     If v = 1 Then dataports.Recordset.FindFirst "numpalets=" + atrim(vsalt): vultimvalor = 0
     If dataports.Recordset.NoMatch Then
       dataports.Recordset.AddNew
       dataports.Recordset!id_transport = atrim(vidtransport)
       dataports.Recordset!pais = Combopais
       dataports.Recordset!codipostal = Combo1
       dataports.Recordset!numpalets = (i * vsalt) + vultimvalor
       dataports.Recordset!tarifaperpaletsokg = vPaletsoKilos
       dataports.Recordset.Update
     End If
  Next i
fi:
  vdesctarifa = vPaletsoKilos & "-" & Combotransportista & " [" & Combopais & "] " & Combo1
  'vsql = "select * from tarifes_ports where id_transport=" + atrim(vidtransport) + " and pais='" + atrim(Combopais) + "' and codipostal='" + atrim(Combo1) + "' and tarifaperpaletsokg='" + vPaletsoKilos + "' order by numpalets"
  carregar_llistatarifes vdesctarifa
 ' dataports.RecordSource = vsql
 ' dataports.Refresh
  
End Sub

Private Sub beliminarf_Click()
   If datafactures.Recordset.EOF Then Exit Sub
   If MsgBox("Segur que vols eliminar aquesta factura de la llista?" + vbNewLine + "  AIXÓ TAMBÉ ELIMINARÀ TOTS ELS CMRS VINCULATS.", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
       If Not datacmrs.Recordset.EOF Then
        datacmrs.Recordset.MoveFirst
        While Not datacmrs.Recordset.EOF
            datacmrs.Recordset.Delete
            datacmrs.Recordset.MoveNext
        Wend
       End If
       datafactures.Recordset.Delete
       datafactures.Refresh
   End If
End Sub

Private Sub belimninarc_Click()
   If datacmrs.Recordset.EOF Then Exit Sub
   If MsgBox("Segur que vols eliminar aquest CMR de la llista?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
       datacmrs.Recordset.Delete
       datacmrs.Refresh
   End If
End Sub

Private Sub blinkpdf_Click()
   Dim vnomfitxer As String
   Dim vdata As Date
   vdata = datafactures.Recordset!datafactura
   vnomfitxer = "FraTrans_" + atrim(Format(datafactures.Recordset!ID, "000000")) + " [" + Format(vdata, "dd-mm-yy") + "] " + treuresimbolsnovalidsnomfitxer(datafactures.Recordset!numerofactura) + ".pdf"
   vnomfitxer = "\\ord_copies\AlbaransSAPClients\FacturesTransport\" + vnomfitxer
   If existeix(vnomfitxer) Then
        obrir_document (vnomfitxer)
         Else:
            If datafactures.Recordset!escanejada Then
                  MsgBox "No trobo el PDF relacionat." + vbNewLine + "Si l'acabes de guardar podria ser que tardes una estona a estar disponible." + vbNewLine + "SI VOLS SOBREESCRIURE-LA HAURAS D'ARRASTRAR LA FACTURA SOBRE EL BOTÓ.", vbCritical, "Error"
                   Else:
                      vnomfitxer = escullir_fitxerfactura
                      If existeix(vnomfitxer) Then desar_facturatransport vnomfitxer
            End If
   End If
End Sub
Function escullir_fitxerfactura() As String
  Dim vdirectori As String
  CommonDialog1.CancelError = True
  vdirectori = llegir_ini("General", "directorifacturestransport", "comandes.ini")
  On Error Resume Next
  With CommonDialog1
   .DialogTitle = "Seleccionar fitxer factura transport"
   .flags = cdlOFNExplorer
   .DefaultExt = ".pdf"
   .InitDir = vdirectori
   .ShowOpen
  End With
  If err.Number <> &H7FF3 Then
      escullir_fitxerfactura = CommonDialog1.FileName
      vdirectori = rutadelfitxer(escullir_fitxerfactura)
      escriure_ini "General", "directorifacturestransport", vdirectori, "comandes.ini"
  End If
End Function

Private Sub blinkpdf_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim vnomfitxer As String
    vnomfitxer = Data.Files(1)
    If InStr(1, LCase(vnomfitxer), ".pdf") = 0 Then MsgBox "El fitxer ha de ser PDF.", vbCritical, "Error": Exit Sub
    desar_facturatransport vnomfitxer
End Sub

Sub desar_facturatransport(vfitxer As String)
   Dim vnomfitxerfinal As String
   Dim vnumalb As Double
   Dim vdata As Date
   vdata = datafactures.Recordset!datafactura
   vnomfitxerfinal = "FraTrans_" + atrim(Format(datafactures.Recordset!ID, "000000")) + " [" + Format(vdata, "dd-mm-yy") + "] " + treuresimbolsnovalidsnomfitxer(datafactures.Recordset!numerofactura) + ".pdf"
   FileCopy vfitxer, rutadelfitxer(cami) + "Cache_escanejarexpedicions\AlbaransSAP\" + vnomfitxerfinal
   datafactures.Recordset.Edit
   datafactures.Recordset!escanejada = True
   datafactures.Recordset.Update
   datafactures.Recordset.Move 0
   MsgBox "Factura " + atrim(datafactures.Recordset!numerofactura) + "- PDF guardat.", vbInformation, "INFORMACIÓ"
End Sub

Private Sub btreurecomanda_Click()
   If dataports.Recordset.EOF And dataports.Recordset.BOF Then MsgBox "No hi ha cap registre per elimninar.", vbCritical, "Atenció": Exit Sub
   If MsgBox("Segur que vols eliminar aquest registre?" + vbNewLine + "Vols eliminar-lo?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   'dataports.Recordset.MoveLast
   If Not dataports.Recordset.EOF Then dataports.Recordset.Delete
   'emplenar_combotransportistes_paisos
   dataports.Refresh
   carregar_llistatarifes
   
End Sub

Private Sub cdatafifuel_LostFocus()
actualitzar_valors_transportista cadbl(Combotransportista.ItemData(Combotransportista.ListIndex))
End Sub

Private Sub cdatainicifuel_LostFocus()
actualitzar_valors_transportista cadbl(Combotransportista.ItemData(Combotransportista.ListIndex))
End Sub

Private Sub cfuel_LostFocus()
  actualitzar_valors_transportista cadbl(Combotransportista.ItemData(Combotransportista.ListIndex))
End Sub

Private Sub Combo1_Change()
   actualitzar_reixa
End Sub

Private Sub Combo1_Click()
  actualitzar_reixa
End Sub

Private Sub Combopais_Click()
  actualitzar_reixa
  
End Sub

Private Sub Combopais_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub Combopais_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Combotransportista_Click()
   If Screen.ActiveControl.Name = "Combotransportista" Then actualitzar_reixa
End Sub

Private Sub Combotransportista_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub Combotransportista_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Command1_Click()
  vvalor = InputBox("Entra el valor que vols filtrar del camp " + vcamp, "Filtre", vvalor)
  If vcamp = "data" And vvalor = "" Then vvalor = "*"
  If vcamp = "comandesrelacionades" Then vvalor = "*" + vvalor + "*"
  dataenviaments.RecordSource = "select * from registre_enviaments where trim(" + vcamp + ") like '" + atrim(treure_apostruf(vvalor)) + "' order by data desc"
  dataenviaments.Refresh
End Sub

Private Sub Command2_Click()
   Dim rst As Recordset
   Dim v As String
   Dim vnomfitxer As String
   vnomfitxer = "c:\temp\Registre enviaments.csv"
   Set rst = dataenviaments.Recordset.Clone
   If rst.EOF Then MsgBox "No hi ha dades per exportar", vbCritical, "Error": Exit Sub
   Open vnomfitxer For Output As #1
   For i = 0 To rst.Fields.Count - 1
      If Mid(UCase(rst.Fields(i).Name), 1, 3) <> "ID_" Then v = v + IIf(v = "", "", ";") + UCase(rst.Fields(i).Name)
   Next i
   Print #1, "LLISTAT D'ENVIAMENTS " + atrim(Now) + vbNewLine + vbNewLine
   Print #1, v
   
   While Not rst.EOF
    linia = ""
    For i = 0 To rst.Fields.Count - 1
      If Mid(UCase(rst.Fields(i).Name), 1, 3) <> "ID_" Then
        v = atrim(rst.Fields(i))
        If UCase(rst.Fields(i).Name) = "NUMEROAVIS" Then v = " " + v
        If UCase(rst.Fields(i).Name) = "COMANDESRELACIONADES" Then v = " " + v
        linia = linia + IIf(linia = "", "", ";") + """" + v + """"
      End If
    Next i
    Print #1, linia
    rst.MoveNext
   Wend
   Close #1
   wait 2
   If existeix(vnomfitxer) Then obrir_document vnomfitxer
   Set rst = Nothing
End Sub

Private Sub Command3_Click()
   Dim v As String
   Dim vv As String
   Dim i As Long
   v = Clipboard.GetText
   v = substituirtots(v, " ", "")
   v = substituirtots(v, "€", " ")
   If dataports.Recordset.EOF Then Exit Sub
   dataports.Recordset.MoveFirst
   i = 1
   vv = Mid(v, i, InStr(i, v + " ", " ") - i)
   If Not IsNumeric(vv) Then MsgBox "Aquestes dades del portapapers no semblen apropiades per enganxar." + vbNewLine + "LES DADES HAN DE SER UNA FILA D'EUROS AMB EL SIMBOL € AL FINAL.", vbCritical, "Error": vv = ""
   While Not dataports.Recordset.EOF And vv <> ""
      dataports.Recordset.Edit
      If IsNumeric(vv) Then dataports.Recordset!preu = cadbl(vv)
      dataports.Recordset.Update
      dataports.Recordset.MoveNext
      i = InStr(i + 1, v + " ", " ")
      If i < Len(v) Then
           vv = Mid(v, i, InStr(i + 1, v + " ", " ") - i)
            Else: vv = ""
      End If
      
   Wend
   
End Sub

Private Sub Command6_Click()
    Dim vidtransport As Double
    Dim vnomtransport As String
    Dim vdatafactura As String
    Dim vnumfactura As String
    vidtransport = escullir_transportista(vnomtransport)
    If vidtransport = 0 Then Exit Sub
    vnumfactura = InputBox("Entra el número de la factura de transport.", "Numero de factura")
    If atrim(vnumfactura) = "" Then Exit Sub
    vdatafactura = InputBox("Entra la data de la factura de transport." + vbNewLine + "Ex: " + Format(Now, "dd/mm/yyyy") + "  o   " + Format(Now, "dd/mm/yy"), "Data de la factura")
    If Not IsDate(vdatafactura) Then MsgBox "Aquesta data no es vàlida.", vbCritical, "Error": Exit Sub
    datafactures.Recordset.FindFirst "numerofactura='" + atrim(vnumfactura) + "' and datafactura<>" + atrim(vdatafactura)
    If Not datafactures.Recordset.NoMatch Then MsgBox "Aquesta factura ja està entrada.", vbCritical, "Error": Exit Sub
    datafactures.Recordset.AddNew
    datafactures.Recordset!numerofactura = atrim(vnumfactura)
    datafactures.Recordset!datafactura = atrim(vdatafactura)
    datafactures.Recordset!idtransport = vidtransport
    datafactures.Recordset!nomtransport = atrim(vnomtransport)
    datafactures.Recordset!escanejada = False
    datafactures.Recordset.Update
    datafactures.Refresh
    datafactures.Recordset.FindFirst "numerofactura='" + atrim(vnumfactura) + "' and datafactura<>" + atrim(vdatafactura)
End Sub
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

Private Sub Command7_Click()
  Dim vnumcmr As String
  Dim vpreu As Double
  Dim v As String
   If datafactures.Recordset.EOF Then Exit Sub
  'SELECT Transportistes_avisos.numeroavis
'FROM transportistes LEFT JOIN Transportistes_avisos ON transportistes.codi = Transportistes_avisos.coditransport
'WHERE (((Transportistes_avisos.numeroavis) Not In (select numeroCMR from transportistes_factures_CMR)) AND ((transportistes.codi)=10) AND ((DateDiff("m",[datarecullida],Now()))<4));
  vnumcmr = escullir_cmr(cadbl(datafactures.Recordset!idtransport))
  If vnumcmr = "" Then Exit Sub
  v = InputBox("Entra el preu de l'enviament d'aquest CMR.", "Preu enviament")
  vpreu = cadbl(substituir(v, ".", ","))
  If vpreu = 0 Then Exit Sub
  
  datacmrs.Recordset.AddNew
  datacmrs.Recordset!ID = datafactures.Recordset!ID
  datacmrs.Recordset!numeroCMR = vnumcmr
  datacmrs.Recordset!preuenviament = vpreu
  datacmrs.Recordset.Update
  
End Sub
Function escullir_cmr(vidtransport As Double) As String
   Dim vsql As String
   Unload formseleccio
   Load formseleccio
  'formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
  formseleccio.Data1.RecordSource = "select cmr,datarecullida,pais from llistat_avisos_transportista where codi=" + atrim(datafactures.Recordset!idtransport)
  formseleccio.refrescar
  wait 1
  formseleccio.DBGrid2.Columns(0).Width = 1500
  formseleccio.DBGrid2.Columns(1).Width = 2000
  formseleccio.DBGrid2.Columns(2).Width = 1000
  formseleccio.CommandXLS.Visible = True
  
  
  'formseleccio.Width = 9000
  'formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           escullir_cmr = formseleccio.DBGrid2.Columns("CMR")
        End If
  End If
End Function


Private Sub Command8_Click()
 Dim vsql As String
   Unload formseleccio
   Load formseleccio
  'formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
  formseleccio.Data1.RecordSource = "select cmr,datarecullida,pais,Nom_Transport from llistat_avisos_transportista"
  formseleccio.refrescar
  wait 1
  formseleccio.DBGrid2.Columns(0).Width = 1500
  formseleccio.DBGrid2.Columns(1).Width = 2000
  formseleccio.DBGrid2.Columns(2).Width = 1000
  formseleccio.DBGrid2.Columns(3).Width = 3000
  formseleccio.Width = 10000
  formseleccio.CommandXLS.Visible = True
  
  
  'formseleccio.Width = 9000
  'formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  
End Sub

Sub guardar_pdf_factura(vnomfitxer As String)
   If InStr(1, LCase(vnomfitxer), ".pdf") = 0 Then MsgBox "Ha de ser un fitxer PDF.", vbCritical, "Error": Exit Sub
   
End Sub
Private Sub credondeigeurokg_LostFocus()
   actualitzar_valors_transportista cadbl(Combotransportista.ItemData(Combotransportista.ListIndex))
End Sub

Private Sub cseguro_LostFocus()
  actualitzar_valors_transportista cadbl(Combotransportista.ItemData(Combotransportista.ListIndex))
End Sub
Sub actualitzar_valors_transportista(vidtransport As Long)
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from transportistes where codi=" + atrim(vidtransport))
   If Not rst.EOF Then
      rst.Edit
      rst![%seguro] = cadbl(cseguro)
      rst![%fuel] = cadbl(cfuel)
      rst!redondeigeurokg = cadbl(credondeigeurokg)
      If IsDate(cdatainicifuel) Then rst!datainicifuel = cdatainicifuel
      If IsDate(cdatafifuel) Then rst!datafifuel = cdatafifuel
      rst.Update
   End If
   Set rst = Nothing
   carregar_dades_transportista vidtransport
End Sub

Private Sub datafactures_Reposition()
  If Not datafactures.Recordset.EOF Then
     datacmrs.RecordSource = "select * from transportistes_factures_cmr where id=" + atrim(datafactures.Recordset!ID)
       Else: datacmrs.RecordSource = "select * from transportistes_factures_cmr where id=0"
  End If
  datacmrs.Refresh
End Sub

Private Sub DBGrid1_DblClick()
   If DBGrid1.Columns(DBGrid1.col).DataField = "eurosports" Then
       v = InputBox("Escriu el preu dels ports", "Preu ports")
       If Not IsNumeric(v) Then Exit Sub
       dataenviaments.Recordset.Edit
       dataenviaments.Recordset!eurosports = cadbl(v)
       dataenviaments.Recordset.Update
       dataenviaments.Recordset.Move 0
   End If
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   vcamp = DBGrid1.Columns(DBGrid1.col).DataField
   vvalor = DBGrid1.Text
End Sub
Sub actualitzar_codispostals()
  Dim rst As Recordset
  Dim vwhere As String
  Dim v As String
  Static vsocdins As Boolean
  If vsocdins Then Exit Sub
  vsocdins = True
  vwhere = "id_transport =" + atrim(cadbl(Combotransportista.ItemData(Combotransportista.ListIndex)))
  vwhere = vwhere + " and pais='" + atrim(Combopais) + "'"
  Set rst = dataports.Database.OpenRecordset("select distinct codipostal from tarifes_ports where " + vwhere)
  v = Combo1
  Combo1.Clear
  While Not rst.EOF
    Combo1.AddItem rst!codipostal
    rst.MoveNext
  Wend
  If Combo1.ListCount = 0 Then Combo1.AddItem "Tots"
'  If v <> "" Then Combo1 = v
  Set rst = Nothing
  vsocdins = False
End Sub
Sub actualitzar_reixa()
  Dim vwhere As String
  
  If Combotransportista.ListIndex = -1 Then dataports.RecordSource = "select * from tarifes_ports where id_transport=-1": dataports.Refresh: Exit Sub
  If Screen.ActiveControl.Name = "Combopais" Then actualitzar_codispostals: SendKeys "{TAB}"
  vwhere = " id_transport=" + atrim(cadbl(Combotransportista.ItemData(Combotransportista.ListIndex)))
  vwhere = vwhere + " and pais='" + atrim(Combopais) + "' " + IIf(atrim(Combo1) <> "", " and codipostal ='" + atrim(Combo1) + "'", "")
  If Combotransportista = "" Or Combopais = "" Or Combo1 = "" Then
      dataports.RecordSource = "select * from tarifes_ports where id_transport=-1"
       Else:
          dataports.RecordSource = "select * from tarifes_ports where " + vwhere + " order by numpalets asc"
  End If
  dataports.Refresh
  carregar_dades_transportista cadbl(Combotransportista.ItemData(Combotransportista.ListIndex))
  If Not dataports.Recordset.EOF Then If dataports.Recordset!tarifaperpaletsokg = "P" Then reixa.Columns(0).Caption = "Palets": ettanperkilo.Visible = False Else reixa.Columns(0).Caption = "Kilos": ettanperkilo.Visible = True
  reixa.AllowUpdate = False
End Sub
Sub carregar_dades_transportista(vidtransport As Long)
   Dim rst As Recordset
   cseguro = "": cfuel = "": cdatainicifuel = "": cdatafifuel = ""
   Set rst = dbtmp.OpenRecordset("select * from transportistes where codi=" + atrim(vidtransport))
   If Not rst.EOF Then
      cseguro = atrim(rst![%seguro])
      cfuel = atrim(rst![%fuel])
      credondeigeurokg = atrim(rst!redondeigeurokg)
      cdatainicifuel = atrim(rst!datainicifuel)
      cdatafifuel = atrim(rst!datafifuel)
   End If
   Set rst = Nothing
End Sub

Private Sub DBGrid2_DblClick()
   If DBGrid2.Columns(DBGrid2.col).DataField = "datafactura" Then
        v = InputBox("Entra la data de la factura:", "Data factura", DBGrid2.Text)
        If IsDate(v) Then DBGrid2.Text = v:  datafactures.Recordset.Move 0
   End If
   If DBGrid2.Columns(DBGrid2.col).DataField = "numerofactura" Then
        v = InputBox("Entra el numero de la factura:", "Data factura", DBGrid2.Text)
        If Len(v) > 1 Then DBGrid2.Text = v: datafactures.Recordset.Move 0
   End If
End Sub

Private Sub DBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  colocarBotoPDF
End Sub
Sub colocarBotoPDF()
  blinkpdf.Top = DBGrid2.RowTop(DBGrid2.row) + DBGrid2.Top + DBGrid2.RowHeight
  blinkpdf.Left = DBGrid2.Left + DBGrid2.Columns(2).Left + DBGrid2.Columns(2).Width - blinkpdf.Width
  blinkpdf.Visible = True
End Sub

Private Sub Form_Load()
   
   dataenviaments.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
   dataenviaments.RecordSource = "select * from registre_enviaments order by data desc"
   dataports.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
   dataports.RecordSource = "select * from tarifes_ports where id_transport=-1"
   dataenviaments.Refresh
   datafactures.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
   datafactures.RecordSource = "select * from transportistes_factures order by datafactura desc"
   datafactures.Refresh
   datacmrs.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
   
   dataports.Refresh
   actualitzar_dataexpedicio
   frametarifes.Top = 1: frametarifes.Left = 60
   Framefactures.Top = 1: Framefactures.Left = 60
   frametarifes.Caption = ""
   frametarifes.Visible = False
   carregar_llistatarifes
End Sub
Sub carregar_llistatarifes(Optional vant As String)
   Dim rst As Recordset
   Dim vpos As Long
   Set rst = dataports.Database.OpenRecordset("SELECT DISTINCT tarifes_ports.tarifaperpaletsokg & '-' & transportistes.descripcio & ' [' & Tarifes_ports.pais & '] ' & Tarifes_ports.codipostal AS desca FROM Tarifes_ports LEFT JOIN transportistes ON Tarifes_ports.id_transport = transportistes.codi")
   If llistatarifes.ListIndex <> -1 And vant = "" Then vant = llistatarifes.Text
   llistatarifes.Clear
   While Not rst.EOF
     llistatarifes.AddItem rst!desca
     If rst!desca = vant Then vpos = llistatarifes.NewIndex
     rst.MoveNext
   Wend
   If llistatarifes.ListCount > 0 Then llistatarifes.ListIndex = vpos
   Set rst = Nothing
End Sub
Sub actualitzar_dataexpedicio()
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim rstplan As Recordset
   Dim v As String
   Dim vv As String
   Dim vpreu As Double
   Dim vnumc As Double
   Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificacio.mdb")
   dataenviaments.Database.Execute "UPDATE (Clients_envios LEFT JOIN clients ON Clients_envios.codi = clients.codi) RIGHT JOIN registre_enviaments ON Clients_envios.id = registre_enviaments.id_desti SET registre_enviaments.nomclient = [clients].[nom] WHERE (((registre_enviaments.nomclient) Is Null Or (registre_enviaments.nomclient)=''));"
   possar_numeroavis_enviaments
   Set rst = dataports.Database.OpenRecordset("select * from registre_enviaments where dataexpedicio=null or dataexpedicio=0")
   While Not rst.EOF
    Set rst2 = dataports.Database.OpenRecordset("select * from tots_transportistes_envios where numeroavis='" + atrim(rst!numeroavis) + "'")
    vnumc = cadbl(Mid(rst!comandesrelacionades, 1, InStr(1, rst!comandesrelacionades + " ", " ")))
    Set rstplan = dbplanificacio.OpenRecordset("select * from planificaciototes where comanda=" + atrim(vnumc))
    rst.Edit
    If Not rst2.EOF Then
        rst!kgteorics = cadbl(rst2!kgs)
        rst!paletsteorics = cadbl(rst2!bases)
        rst!m3teorics = cadbl(rst2!metres3)
        rst!dataexpedicio = IIf(Not IsNull(rst2!datarecullida), rst2!datarecullida, Null)
        vpreu = Menu.buscar_preutransport(rst!id_desti, rst!paletsteorics, rst!kgteorics, rst!id_transport, v, vv)
        If vpreu > 0 Then rst!eurosports = vpreu
    End If
    rst!dataexpedicioteorica = IIf(Not rstplan.EOF, IIf(IsNull(rstplan!dataexpedicio), Null, atrim(rstplan!dataexpedicio)), Null)
    rst.Update
    rst.MoveNext
   Wend
   Set rst = Nothing
   Set rst2 = Nothing
   
   dataenviaments.Database.Execute "UPDATE registre_enviaments LEFT JOIN Transportistes_avisos ON registre_enviaments.numeroavis = Transportistes_avisos.numeroavis SET registre_enviaments.dataexpedicio = [Transportistes_avisos].[datarecullida] WHERE (((registre_enviaments.dataexpedicio) Is Null));"
   
End Sub
Sub possar_numeroavis_enviaments()
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim vnumc As Double
   Set rst = dataports.Database.OpenRecordset("select * from registre_enviaments where numeroavis='' or numeroavis =null")
   While Not rst.EOF
      vnumc = cadbl(Mid(rst!comandesrelacionades, 1, InStr(1, rst!comandesrelacionades + " ", " ")))
      Set rst2 = dataports.Database.OpenRecordset("select * from liniesalbara where lotinplacsa=" + atrim(vnumc))
      If Not rst2.EOF Then
          Set rst2 = dataports.Database.OpenRecordset("select * from transportistes_avisos where numalbara=" + atrim(rst2!numalbara))
          If Not rst2.EOF Then rst.Edit: rst!numeroavis = atrim(rst2!numeroavis): rst.Update
      End If
      rst.MoveNext
   Wend
   Set rst = Nothing
   Set rst2 = Nothing
End Sub

Private Sub llistatarifes_Click()
  Dim rst As Recordset
   Set rst = dataports.Database.OpenRecordset("select * FROM Tarifes_ports LEFT JOIN transportistes ON Tarifes_ports.id_transport = transportistes.codi where tarifes_ports.tarifaperpaletsokg & '-' & transportistes.descripcio&' ['& Tarifes_ports.pais&'] '& Tarifes_ports.codipostal='" + llistatarifes + "'")
   If Not rst.EOF Then
       For i = 0 To Combotransportista.ListCount - 1: If Combotransportista.ItemData(i) = atrim(rst!id_transport) Then Combotransportista.ListIndex = i
       Next i
       Combopais = rst!pais
       vwhere = " tarifaperpaletsokg='" + atrim(rst!tarifaperpaletsokg) + "' and id_transport=" + atrim(rst!id_transport)
       vwhere = vwhere + " and pais='" + atrim(rst!pais) + "' and codipostal = '" + atrim(rst!codipostal) + "'"
       Combo1 = rst!codipostal
       dataports.RecordSource = "select * from tarifes_ports where " + vwhere + " order by numpalets"
       dataports.Refresh
       reixa.Columns(0).Caption = IIf(rst!tarifaperpaletsokg = "P", "Palets", "Kilos")
   End If
  Set rst = Nothing
End Sub

Private Sub mfacturestransport_Click()
   Framefactures.Visible = Not Framefactures.Visible
   frametarifes.Visible = False
   
End Sub

Private Sub modificar_Click()
   reixa.AllowUpdate = True
End Sub

Private Sub mtarifes_Click()
   emplenar_combotransportistes_paisos
   frametarifes.Visible = Not frametarifes.Visible
   Framefactures.Visible = False
End Sub
Sub emplenar_combotransportistes_paisos()
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from transportistes where visible=1 order by descripcio")
   Combotransportista.Clear
   While Not rst.EOF
       Combotransportista.AddItem rst!descripcio
       Combotransportista.ItemData(Combotransportista.NewIndex) = cadbl(rst!codi)
       rst.MoveNext
   Wend
   Set rst = Nothing
   
   Set rst = dbtmp.OpenRecordset("select distinct pais from clients_envios where pais<>null and pais<>'' order by pais")
   Combopais.Clear
   While Not rst.EOF
       Combopais.AddItem atrim(rst!pais)
       rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

