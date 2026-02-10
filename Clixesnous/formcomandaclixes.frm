VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formcomandaclixes 
   BackColor       =   &H00FDDECE&
   Caption         =   "Comanda clixes a fotogravador"
   ClientHeight    =   9855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   Icon            =   "formcomandaclixes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   675
      Top             =   240
   End
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   8115
      Picture         =   "formcomandaclixes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Enviar comanda per mail amb pdf"
      Top             =   375
      Width           =   585
   End
   Begin VB.Frame Frame6 
      Caption         =   "Reposicions"
      Height          =   1755
      Left            =   90
      TabIndex        =   27
      Top             =   8085
      Width           =   8580
      Begin VB.TextBox cpostit 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4290
         TabIndex        =   37
         Text            =   "  Dos clics per canviar data, albarà i preu de l'albarà"
         Top             =   0
         Width           =   3960
      End
      Begin VB.Data datareposicions 
         Caption         =   "datareposicions"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5235
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   840
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   8085
         Picture         =   "formcomandaclixes.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Eliminar reposicio"
         Top             =   615
         Width           =   420
      End
      Begin VB.CommandButton alta 
         Height          =   375
         Left            =   8085
         Picture         =   "formcomandaclixes.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Nova reposicio"
         Top             =   210
         Width           =   420
      End
      Begin MSDBGrid.DBGrid reixareposicions 
         Bindings        =   "formcomandaclixes.frx":1628
         Height          =   1485
         Left            =   45
         OleObjectBlob   =   "formcomandaclixes.frx":1642
         TabIndex        =   28
         Top             =   210
         Width           =   8010
      End
   End
   Begin VB.CommandButton Command8 
      Height          =   375
      Left            =   7530
      Picture         =   "formcomandaclixes.frx":270A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Verificar comanda"
      Top             =   375
      Width           =   585
   End
   Begin VB.Frame Frame1 
      Height          =   7410
      Left            =   60
      TabIndex        =   1
      Top             =   690
      Width           =   8670
      Begin VB.ComboBox ctipusproducte 
         Height          =   315
         ItemData        =   "formcomandaclixes.frx":2C94
         Left            =   4335
         List            =   "formcomandaclixes.frx":2C9E
         TabIndex        =   35
         Top             =   405
         Width           =   1515
      End
      Begin VB.Frame fenviant 
         BackColor       =   &H00F3B378&
         BorderStyle     =   0  'None
         Height          =   2850
         Left            =   1185
         TabIndex        =   31
         Top             =   2355
         Visible         =   0   'False
         Width           =   6165
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Enviant el mail al fotogravador"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   330
            Left            =   780
            TabIndex        =   32
            Top             =   2250
            Width           =   4845
         End
         Begin VB.Image Image1 
            Height          =   1950
            Left            =   1575
            Picture         =   "formcomandaclixes.frx":2CB1
            Stretch         =   -1  'True
            Top             =   315
            Width           =   2505
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   2385
            Left            =   375
            Shape           =   4  'Rounded Rectangle
            Top             =   270
            Width           =   5310
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Passar a Clixes Rebuts"
         Height          =   1020
         Left            =   7365
         Picture         =   "formcomandaclixes.frx":8734F
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Passar la comanda del proveidor a rebuda."
         Top             =   645
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox cpreu 
         Height          =   285
         Left            =   5355
         TabIndex        =   20
         Top             =   795
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Frame Frame3 
         Caption         =   "Orientació Muntatge"
         Height          =   4410
         Left            =   60
         TabIndex        =   13
         Top             =   2865
         Width           =   8550
         Begin VB.Frame Frame5 
            Caption         =   "   Imatge 1        Imatge2"
            Height          =   1620
            Left            =   6390
            TabIndex        =   16
            Top             =   1350
            Width           =   2055
            Begin VB.Image r1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1215
               Left            =   75
               Stretch         =   -1  'True
               Top             =   285
               Width           =   900
            End
            Begin VB.Image r2 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   1215
               Left            =   1050
               Stretch         =   -1  'True
               Top             =   285
               Width           =   900
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Imatge 2"
            Height          =   1995
            Left            =   165
            TabIndex        =   15
            Top             =   2310
            Width           =   6180
            Begin VB.CommandButton capformat 
               Height          =   315
               Left            =   30
               Picture         =   "formcomandaclixes.frx":878D9
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   525
               Width           =   285
            End
            Begin VB.Image fosel 
               Height          =   240
               Left            =   75
               Picture         =   "formcomandaclixes.frx":87E63
               Top             =   225
               Width           =   240
            End
            Begin VB.Image fo1 
               Appearance      =   0  'Flat
               Height          =   1770
               Left            =   360
               Picture         =   "formcomandaclixes.frx":883ED
               Top             =   180
               Width           =   1335
            End
            Begin VB.Image fo2 
               Height          =   1770
               Left            =   1815
               Picture         =   "formcomandaclixes.frx":89C8B
               Top             =   180
               Width           =   1335
            End
            Begin VB.Image fo3 
               Height          =   1770
               Left            =   3255
               Picture         =   "formcomandaclixes.frx":8AC14
               Top             =   180
               Width           =   1335
            End
            Begin VB.Image fo4 
               Height          =   1770
               Left            =   4710
               Picture         =   "formcomandaclixes.frx":8BB73
               Top             =   180
               Width           =   1335
            End
         End
         Begin VB.Frame framw 
            Caption         =   "Imatge 1"
            Height          =   1995
            Left            =   165
            TabIndex        =   14
            Top             =   315
            Width           =   6180
            Begin VB.CommandButton capfilm 
               Height          =   315
               Left            =   30
               Picture         =   "formcomandaclixes.frx":8CAA9
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   555
               Width           =   285
            End
            Begin VB.Image fsel 
               Height          =   240
               Left            =   60
               Picture         =   "formcomandaclixes.frx":8D033
               Top             =   225
               Width           =   240
            End
            Begin VB.Image ff4 
               Height          =   1770
               Left            =   4695
               Picture         =   "formcomandaclixes.frx":8D5BD
               Top             =   180
               Width           =   1335
            End
            Begin VB.Image ff3 
               Height          =   1770
               Left            =   3240
               Picture         =   "formcomandaclixes.frx":8E4F3
               Top             =   180
               Width           =   1335
            End
            Begin VB.Image ff2 
               Height          =   1770
               Left            =   1800
               Picture         =   "formcomandaclixes.frx":8F452
               Top             =   180
               Width           =   1335
            End
            Begin VB.Image ff1 
               Appearance      =   0  'Flat
               Height          =   1770
               Left            =   345
               Picture         =   "formcomandaclixes.frx":903DB
               Top             =   180
               Width           =   1335
            End
         End
         Begin VB.Label Label6 
            Caption         =   "Doble clic per sel.leccionar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00F3B378&
            Height          =   360
            Left            =   2160
            TabIndex        =   17
            Top             =   120
            Width           =   3045
         End
      End
      Begin VB.TextBox cobservacions 
         Height          =   810
         Left            =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   2040
         Width           =   8025
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pauta cèl.lula"
         Height          =   930
         Left            =   135
         TabIndex        =   6
         Top             =   795
         Width           =   3870
         Begin VB.TextBox ccolorscelula 
            Height          =   285
            Left            =   1155
            TabIndex        =   9
            Text            =   "TODOS"
            Top             =   585
            Width           =   1545
         End
         Begin VB.TextBox ctamanycelula 
            Height          =   285
            Left            =   1170
            TabIndex        =   7
            Text            =   "PDF/CROMALIN"
            Top             =   270
            Width           =   1545
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Colors:"
            Height          =   345
            Left            =   225
            TabIndex        =   10
            Top             =   585
            Width           =   825
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Tamany:"
            Height          =   345
            Left            =   225
            TabIndex        =   8
            Top             =   300
            Width           =   825
         End
      End
      Begin VB.TextBox cdataestimada 
         Height          =   285
         Left            =   2310
         TabIndex        =   3
         Top             =   405
         Width           =   1260
      End
      Begin VB.TextBox cdata 
         Height          =   285
         Left            =   225
         TabIndex        =   2
         Top             =   405
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Tipus de producte"
         Height          =   180
         Left            =   4410
         TabIndex        =   36
         Top             =   195
         Width           =   1365
      End
      Begin VB.Label etdatarecepcio 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   6675
         TabIndex        =   26
         Top             =   630
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Image imatgeenviat 
         Height          =   450
         Left            =   7815
         Picture         =   "formcomandaclixes.frx":91C79
         Stretch         =   -1  'True
         Top             =   150
         Width           =   435
      End
      Begin VB.Image imatgeacceptat 
         Height          =   450
         Left            =   7245
         Picture         =   "formcomandaclixes.frx":92203
         Stretch         =   -1  'True
         Top             =   165
         Width           =   435
      End
      Begin VB.Label Label8 
         Caption         =   "€"
         Height          =   285
         Left            =   6360
         TabIndex        =   22
         Top             =   795
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label7 
         Caption         =   "Preu dels Clixes:"
         Height          =   345
         Left            =   4050
         TabIndex        =   21
         Top             =   795
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label5 
         Caption         =   "Observacions:"
         Height          =   345
         Left            =   390
         TabIndex        =   12
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Data estimada:"
         Height          =   345
         Left            =   2430
         TabIndex        =   5
         Top             =   195
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Data comanda:"
         Height          =   345
         Left            =   270
         TabIndex        =   4
         Top             =   195
         Width           =   1125
      End
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   8115
      Picture         =   "formcomandaclixes.frx":9278D
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Eliminar tot lo d'aquesta comanda"
      Top             =   0
      Width           =   585
   End
   Begin VB.Label etidioma 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   33
      Top             =   0
      Width           =   330
   End
   Begin VB.Label etdescripciocomanda 
      Alignment       =   2  'Center
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
      Height          =   600
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   8055
   End
End
Attribute VB_Name = "formcomandaclixes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbtintes As Database
Sub possarexemplemuntatge()
   r1.Picture = Nothing: r2.Picture = Nothing
   If cadbl(fsel.tag) = 0 Then Exit Sub
   r1.Picture = formcomandaclixes.Controls("ff" + atrim(fsel.tag)).Picture
   If cadbl(fosel.tag) = 0 Then r2.visible = False: Exit Sub
   r2.visible = True
   r2.Picture = formcomandaclixes.Controls("fo" + atrim(fosel.tag)).Picture
End Sub
Sub colocarselfilm(n As Byte)
   If n > 0 Then
        fsel.Left = formcomandaclixes.Controls("ff" + atrim(n)).Left + 100
      Else: fsel.Left = 60: n = 0
   End If
   fsel.tag = atrim(n)
   possarexemplemuntatge
End Sub
Sub colocarselformat(n As Byte)
   If n > 0 Then
        fosel.Left = formcomandaclixes.Controls("fo" + atrim(n)).Left + 100
      Else: fosel.Left = 60: n = 0
   End If
   fosel.tag = atrim(n)
   possarexemplemuntatge
End Sub

Private Sub alta_Click()
    altareposicio
End Sub
Sub altareposicio()
   Dim clixesdemanats As String
   clixesdemanats = seleccionarcolors
   prepararenviament clixesdemanats
End Sub
Sub prepararenviament(clixesdemanats As String)
   Dim cos As String
   Dim mides As String
   Dim salutacio As String
   Dim vmodificatperinplacsa As Boolean
   If Mid(clixesdemanats + "  ", 1, 2) = "#M" Then
      clixesdemanats = atrim(Mid(clixesdemanats + "    ", 3))
      vmodificatperinplacsa = True
   End If
   Load formenviomails
   formenviomails.destinatari = emailfotogravador(formclixes.modificacions.Recordset!fotograbador)
   formenviomails.asumpte = "CLICHÉ REPOSICIÓN " + formclixes.marcaproducte + " - " + formclixes.liniaproducte
   formenviomails.cosdelmissatge = formenviomails.cosdelmissatge + " Hola," + vbCrLf + "Necesitaríamos un cliché de reposición de la referencia " + formclixes.marcaproducte + " - " + formclixes.liniaproducte + vbCrLf + vbCrLf
   cos = " COLORES:  " + clixesdemanats + vbCrLf + vbCrLf
   mides = possarlesmides
   salutacio = vbCrLf + "Gracias," + vbCrLf + "Eva Llinàs" + vbCrLf + vbCrLf + "Dep. Mk." + vbCrLf + "Tel.: +34 972 460 190" + vbCrLf + "e-mail: ellinas@inplacsa.com" + vbCrLf + "web: www.inplacsa.com"
   formenviomails.cosdelmissatge = formenviomails.cosdelmissatge + cos + mides + salutacio
   Copiar_Fitxer formclixes.rutapdftreball, "c:\temp\clixescomandes\PDFdeltreball_" + atrim(id_treball) + "-" + atrim(ordremodificacio) + ".pdf"
   formenviomails.nomfitxeradjunt = "c:\temp\clixescomandes\PDFdeltreball_" + atrim(id_treball) + "-" + atrim(ordremodificacio) + ".pdf"
   formenviomails.Show 1
   If formenviomails.enviar.tag = "enviar" Then
        If enviaremail2 Then
           
           MsgBox "Missatge enviat correctament.", vbInformation, "Envio Reposicio"
           'crear la reposicio a la base de dades
           datareposicions.Recordset.AddNew
           datareposicions.Recordset!id_treball = id_treball
           datareposicions.Recordset!ordremodificacio = ordremodificacio
           datareposicions.Recordset!dataenviament = Now
           datareposicions.Recordset!datateoricarecepcio = DateAdd("d", 7, Now)
           datareposicions.Recordset!descripcio = "Reposició del colors: " + clixesdemanats
           datareposicions.Recordset!modificatperinplacsa = vmodificatperinplacsa
           datareposicions.Recordset.Update
           datareposicions.Refresh
           MsgBox "Passarè l'estat del clixé a REPOSICIÓ.", vbInformation, "Atenció"
           possar_liniamodificacio_reposicio cadbl(id_treball), cadbl(ordremodificacio), Now, DateAdd("d", 7, Now)
             Else: MsgBox "Error en l'envio del mail", vbCritical, "Error"
        End If
   End If
   If existeix(formenviomails.nomfitxeradjunt) Then Kill formenviomails.nomfitxeradjunt
   Unload formenviomails
End Sub
Sub possar_liniamodificacio_reposicio(id_treball As Double, ordre As Double, Data1 As Date, data2 As Date)
   Dim rst As Recordset
   Dim rstestat As Recordset
   Dim vordrenou As Byte
   Dim vordrevisual As Byte
   Set rstestat = dbclixes.OpenRecordset("select * from clixes_estats where descripcio='REPOSICIÓ DEL CLIXE'")
   If rstestat.EOF Then MsgBox "No he trobat l'estat de clixé REPOSICIÓ DEL CLIXE a la taula d'estats de clixes", vbCritical, "Error": Exit Sub
   vordrenou = 1
   vordrevisual = 1
   Set rst = dbclixes.OpenRecordset("select * from clixes_modifi where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordre) + " order by ordre desc", , ReadOnly)
   If Not rst.EOF Then
      vordrenou = rst!ordre + 1
      vordrevisual = cadbl(Mid(rst!descripcio, 1, 2)) + 1
      rst.AddNew
      rst!id_treball = id_treball
      rst!ordremodificacio = ordre
      rst!ordre = vordrenou
      rst!id_estatclixe = rstestat!id_estat
      rst!descripcioestat = rstestat!descripcio
      rst!descripcio = Format(vordrevisual, "00") + " REPOSICIÓ DEL CLIXE"
      rst!data_inici = Data1
      rst!data_prevista = data2
      rst.Update
   End If
   Set rst = Nothing
   Set rstestat = Nothing
   formclixes.possarestatclixe
End Sub
Function possarlesmides() As String
    Dim rsttintes As Recordset
    Set rsttintes = dbclixes.OpenRecordset("Select * from tintes where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " order by ordretinter")
    possarlesmides = "Medidas-Ancho: " + formclixes.camplelamina + IIf(cadbl(formclixes.Text10) > 1, "(" + formclixes.Text10 + " caídas)", "")
    possarlesmides = possarlesmides + vbCrLf + "Desarrollo: " + formclixes.Text5 + "mm (cilindro " + IIf(rsttintes.EOF, "", atrim(rsttintes!cilindre)) + ")"
    possarlesmides = possarlesmides + vbCrLf + "Espesor polímero: " + formclixes.Text11 + vbCrLf + vbCrLf
    possarlesmides = possarlesmides + "IMPRESIÓN: " + formclixes.comboformaimpresio + vbCrLf
    Set rstintes = Nothing
End Function
Function idiomafotogravador(idfoto As Long) As String
   Dim rst As Recordset
   idiomafotogravador = "ES"
   Set rst = dbclixes.OpenRecordset("Select idioma from fotogravadors where codi=" + atrim(idfoto))
   If Not rst.EOF Then idiomafotogravador = atrim(rst!Idioma)
   Set rst = Nothing
End Function
Function emailfotogravador(idfoto As Long) As String
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("Select email from fotogravadors where codi=" + atrim(idfoto))
   If Not rst.EOF Then emailfotogravador = atrim(rst!email)
   Set rst = Nothing
End Function
Function seleccionarcolors() As String
   Dim resp As String
   Dim vmodxinp As Boolean
   Dim vmodificatperinplacsa As Boolean
   Do
     vmodificatperinplacsa = False
     resp = triarcolorafectat(vmodificatperinplacsa)
     If resp <> "" Then
        If InStr(1, seleccionarcolors, "(" + resp + ")") = 0 Then
            seleccionarcolors = seleccionarcolors + "  (" + resp + ")"
            If vmodificatperinplacsa Then
              While UCase(resp) <> "OK"
                resp = InputBox("Aquest clixé ha estat MODIFICAT PER INPLACSA, tingues-ho en compte al fer la comanda." + Chr(10) + "Escriu OK per continuar.", "Atenció")
              Wend
            End If
           Else: MsgBox "Aquest color ja l'has escullit", vbCritical, "Error"
        End If
          Else: GoTo cont
     End If
     If vmodificatperinplacsa Then vmodxinp = True
    Loop Until MsgBox("Vols afegir mes colors a la sel.lecció?" + Chr(10) + seleccionarcolors, vbYesNo + vbInformation, "Atenció") = vbNo
    If vmodxinp Then seleccionarcolors = "#M" + seleccionarcolors
cont:
   
End Function
Function triarcolorafectat(vmodificatperinplacsa As Boolean)
   Load formseleccio
   formseleccio.Data1.DatabaseName = camiclixes
   formseleccio.Data1.RecordSource = "select color,modificatperinplacsa from tintes where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio)
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar

   'formseleccio.DBGrid2.Columns(0).Width = 0
   formseleccio.DBGrid2.Columns("modificatperinplacsa").visible = False
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           triarcolorafectat = formseleccio.DBGrid2.Columns("color")
           vmodificatperinplacsa = formseleccio.DBGrid2.Columns("modificatperinplacsa")
        End If
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
   
   
End Function
Function enviaremail2() As Boolean
  Dim usuarim As String
  Dim contrasenyam As String
   fenviant.visible = True
   DoEvents
   formcomandaclixes.Enabled = False
   ratoli "espera"
   enviaremail2 = False
   'If Command7.tag <> "enviar" Then Exit Function
   usuarim = llegir_ini("Enviomails", "usuari", "comandes.ini")
   contrasenyam = llegir_ini("Enviomails", "contrasenya", "comandes.ini")
   If usuarim = "{[}]" Or contrasenyam = "{[}]" Then MsgBox "L'usuari o la contrasenya no estan entrades", vbCritical, "Error": Exit Function
   
'creo el fitxer de cos de missatge
   Open "c:\temp\cosmissatge.txt" For Output As #2
   Print #2, formenviomails.cosdelmissatge
   Close #2
   
   
    Set objMessage = CreateObject("CDO.Message")
    objMessage.Subject = formenviomails.asumpte
    objMessage.from = usuarim
    objMessage.To = formenviomails.destinatari
    objMessage.TextBody = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\temp\cosmissatge.txt", 1).ReadAll
    objMessage.AddAttachment formenviomails.nomfitxeradjunt
    
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = usuarim
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = contrasenyam
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 100
    objMessage.Configuration.Fields.Update
    
    
    '==End remote SMTP server configuration section==
    If cadbl(objMessage.Send) = 0 Then enviaremail2 = True
    ratoli "normal"
    fenviant.visible = False
    formcomandaclixes.Enabled = True
End Function

Private Sub capfilm_Click()
   If MsgBox("Segur que vols borrar la sel.lecció?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   colocarselfilm 0
   colocarselformat 0
End Sub

Private Sub capformat_Click()
   If MsgBox("Segur que vols borrar la sel.lecció?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   colocarselformat 0
End Sub

Private Sub cdata_LostFocus()
  If IsDate(cdata) And cdataestimada = "" Then cdataestimada = DateAdd("d", 4, cdata)
End Sub

Private Sub Command1_Click()
   passararebut
   gravarcomanda
   carregarcomanda
End Sub
Sub passararebut()
   Dim datarebut As String
   datarebut = InputBox("Entra la data de recepció dels clixés", "Recepcio clixés", Date)
   If Not IsDate(datarebut) Then MsgBox "Aquesta data no es correcte", vbCritical, "Error": Exit Sub
   etdatarecepcio.tag = datarebut
   etdatarecepcio = "Data Rebuda: " + Format(datarebut, "dd/mm/yy")
   etdatarecepcio.visible = True
   Command1.visible = False
End Sub

Sub generarcomandaimpresa(Optional ferpdf As Boolean)
   Dim nomdbtemporal As String
   Dim dbtemporal As Database
   Dim rsttemporal As Recordset
   Dim vnomCSV As String
   gravarcomanda
   nomdbtemporal = nomfitxertemporal
   If Not existeix(nomdbtemporal) Then DBEngine.CreateDatabase nomdbtemporal, dbLangGeneral, DatabaseTypeEnum.dbVersion30
   borrartaulatemporal dbtemporal
   dbclixes.Execute "select * into comandesproveidor IN '" + nomdbtemporal + "' from comandesfotogravador where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio)
   Set dbtemporal = OpenDatabase(nomdbtemporal)
   
   Set rsttemporal = dbtemporal.OpenRecordset("select * from comandesproveidor")
   If rsttemporal.EOF Then MsgBox "No hi ha res per imprimir": GoTo fi
   Set rsttemporal = Nothing
   afegircamps dbtemporal
   Set rsttemporal = dbtemporal.OpenRecordset("select * from comandesproveidor")
   possardadesaltemporal rsttemporal
   wait 2
   If ferpdf Then
           vnomCSV = "c:\temp\clixescomandes\CSV_clixescomandaproveidor_" + atrim(id_treball) + "-" + atrim(ordremodificacio) + ".csv"
           generarpdfcomanda nomdbtemporal
           If cadbl(formclixes.modificacions.Recordset!fotograbador) = 40 Then generarcsvcomanda dbtemporal, vnomCSV
           Copiar_Fitxer formclixes.rutapdftreball, "c:\temp\clixescomandes\PDFdeltreball_" + atrim(id_treball) + "-" + atrim(ordremodificacio) + ".pdf"
           enviar_fitxers
       Else
         visualitzarcomanda nomdbtemporal
   End If
fi:
   Set dbtemporal = Nothing
End Sub
Sub enviar_fitxers()
  ratoli "espera"
  Shell "c:\windows\system32\cmd.exe /c start mailto:"
  wait 4
  idp = ShellExecute(Me.hWnd, "Open", "c:\windows\explorer.exe", " " + "c:\temp\clixescomandes", "", 1)
  
  ratoli "normal"
End Sub
Sub borrarelstemporals()
   On Error Resume Next
   MkDir "c:\temp\clixespressupostos"
   MkDir "c:\temp\clixescomandes"
   Kill "c:\temp\clixespressupostos\clixespressupost*.*"
   Kill "c:\temp\clixescomandes\*.*"
End Sub
Sub generarcsvcomanda(dbtemporal As Database, vnomCSV As String)
  Dim rst As Recordset
  Dim vcamps As String
  vcamps = "nomtreball, datacomanda, dataprevista, nomclient, id_treball, ordremodificacio, numcomanda, nomproveidor, ample, llarg, material, impressio, tintes, tinta1, tinta2, tinta3, tinta4, tinta5, tinta6, tinta7, tinta8, polimers, espesorpolimer, muntatgebandes, pautacelula, pautasituacio, tamanycelula, colorscelula, cilindre, motius, continu, bandaseguiment, bandasituacio, banda_a, deixarsang, sangsituacio, sang_de, codidebarres, tipusproducte, tintareprint1, tintareprint2, tintareprint3, tintareprint4, tintareprint5, tintareprint6, tintareprint7, tintareprint8, impressioreprint, muntatge1, muntatge2, observacions"
  Set rst = dbtemporal.OpenRecordset("select " + vcamps + " from comandesproveidor")
  If existeix(vnomCSV) Then Kill vnomCSV
  If Not rst.EOF Then
      Open vnomCSV For Output As #1
      For i = 0 To rst.Fields.Count - 1
        Print #1, UCase(rst.Fields(i).Name) + ";" + substituirtot(atrim(rst.Fields(i).Value), ";", ",")
      Next i
      Print #1, "AMPLE TOTAL" + ";" + atrim(cadbl(rst!ample) * cadbl(rst!muntatgebandes))
      Close 1
  End If
  Set rst = Nothing
End Sub
Sub generarpdfcomanda(nomdbtemporal As String)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Dim fitxerpdftemporal As String
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "clixescomandaproveidors.rpt", 1)
  borrarelstemporals
  fitxerpdftemporal = "c:\temp\clixescomandes\clixescomandaproveidor_" + atrim(id_treball) + "-" + atrim(ordremodificacio) + ".pdf"
  oreport.ExportOptions.DiskFileName = fitxerpdftemporal
  oreport.ExportOptions.PDFExportAllPages = True
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  oreport.Database.Tables.Item(1).Location = nomdbtemporal
  oreport.DiscardSavedData
  passaraidioma oreport, etidioma

  oreport.Export False
  ratoli "normal"
End Sub

Sub visualitzarcomanda(nomdbtemporal As String)
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "clixescomandaproveidors.rpt", 1)
  
  oreport.Database.Tables.Item(1).Location = nomdbtemporal
  oreport.DiscardSavedData
  passaraidioma oreport, etidioma
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.ViewReport
   veurereport.width = formclixes.width - 200
   veurereport.Height = formclixes.Height - 300
   
   veurereport.Show 1, Me
  ratoli "normal"
End Sub
Sub passaraidioma(oreport As CRAXDDRT.Report, Idioma As String)
    Dim valorcamp As String
    Dim nomcamp As String
    For i = 1 To oreport.FormulaFields.Count
       nomcamp = oreport.FormulaFields.Item(i).FormulaFieldName
       If Mid(oreport.FormulaFields.GetItemByName(nomcamp).Text, 1, 1) = """" Then
            valorcamp = treure_apostruficometes(oreport.FormulaFields.GetItemByName(nomcamp).Text)
            oreport.FormulaFields.GetItemByName(nomcamp).Text = """" + traduir(valorcamp, Idioma) + """"
       End If
    Next i
End Sub
Function treure_apostruficometes(valor As String) As String
   Dim n As String
   n = valor
   While InStr(n, "'")
     n = Mid(n, 1, InStr(1, n, "'") - 1) + "´" + Mid(n, InStr(1, n, "'") + 1)
   Wend
   While InStr(n, """")
     n = Mid(n, 1, InStr(1, n, """") - 1) + "" + Mid(n, InStr(1, n, """") + 1)
   Wend

   If n = "{[}]" Then n = ""
   treure_apostruficometes = n

End Function
Function traduir(valor As String, Idioma As String) As String
       Dim rst As Recordset
   traduir = atrim(valor)
   Set rst = dbclixes.OpenRecordset("select * from diccionari where idioma='" + atrim(Idioma) + "' and trim(pertraduir)='" + treure_apostruf(atrim(valor)) + "'")
   If Not rst.EOF Then
      traduir = atrim(rst!traduit)
   End If
   
End Function
Function buscarcomandespendents() As String
   Dim i As Byte
   For i = 0 To formclixes.llistadecomandespendents.ListCount
     If buscarcomandespendents = "" Then
       If InStr(1, formclixes.llistadecomandespendents.List(i), "v" + atrim(ordremodificacio)) > 0 Then
           buscarcomandespendents = atrim(Mid(formclixes.llistadecomandespendents.List(i), 1, 6))
       End If
     End If
   Next i
   
End Function
Function buscarcolormaterialpc2(numc As Double, camp As String) As String
   Dim rst As Recordset
   If numc = 0 Then Exit Function
   Set rst = dbcomandes.OpenRecordset("select " + camp + " as elcolor from comandes where comanda=" + atrim(numc))
   If Not rst.EOF Then
    Clipboard.Clear
    Clipboard.SetText "SELECT subfamiliesmaterials.descripcio as Desc_Subfamilia, familiescolorants.descripcio as DescColorant, familiesmaterials.descripcio as DescMaterial FROM (familiesmaterials RIGHT JOIN (comandes INNER JOIN (familiescolorants RIGHT JOIN materials ON familiescolorants.codi = materials.familiacol) ON comandes.materialex = materials.codi) ON familiesmaterials.codi = materials.familia) LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi where comanda=" + atrim(rst!elcolor)
    Set rst = dbcomandes.OpenRecordset("SELECT subfamiliesmaterials.descripcio as Desc_Subfamilia, familiescolorants.descripcio as DescColorant, familiesmaterials.descripcio as DescMaterial FROM (familiesmaterials RIGHT JOIN (comandes INNER JOIN (familiescolorants RIGHT JOIN materials ON familiescolorants.codi = materials.familiacol) ON comandes.materialex = materials.codi) ON familiesmaterials.codi = materials.familia) LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi where comanda=" + atrim(rst!elcolor))
    If Not rst.EOF Then
      buscarcolormaterialpc2 = atrim(rst!DescMaterial) + " " + atrim(rst!DescColorant)
      If InStr(1, " " + rst!Desc_Subfamilia + " ", " MATE ") > 0 Or InStr(1, " " + Desc_Subfamilia + " ", "+MATE ") > 0 Then buscarcolormaterialpc2 = buscarcolormaterialpc2 + " MATE"
      If InStr(1, " " + rst!Desc_Subfamilia + " ", " ALOX ") > 0 Or InStr(1, " " + Desc_Subfamilia + " ", "+ALOX ") > 0 Then buscarcolormaterialpc2 = buscarcolormaterialpc2 + " ALOX"
    End If
   End If
   Set rst = Nothing
End Function
Function traduirlatinta(color As String, coditinta As String, detalltinter As String) As String
  Dim primeraparaula As String
  Dim segonaparaula As String
  Dim colorrestant As String
  Dim rsttinta As Recordset
  If detalltinter <> "TEXTES" And detalltinter <> "TRAMES" Then detalltinter = ""
  If atrim(coditinta) <> "" Then
     Set rsttinta = dbtintes.OpenRecordset("select * from tintes_tot where codi='" + atrim(coditinta) + "'")
     If Not rsttinta.EOF Then color = atrim(rsttinta!descripciofamcol) + " " + atrim(rsttinta!referenciacolor)
     Set rsttinta = Nothing
  End If
  color = color + " "
  primeraparaula = Mid(color, 1, InStr(1, color, " "))
  If primeraparaula <> "" Then
     colorrestant = Trim(Mid(color, InStr(1, color, " ") + 1)) + " "
     segonaparaula = Trim(Mid(colorrestant, 1, InStr(1, colorrestant, " ")))
     If segonaparaula = "PRIMAR" Then primeraparaula = atrim(primeraparaula) + "PRIMAR": segonaparaula = ""
     colorrestant = Trim(Mid(colorrestant, InStr(1, colorrestant, " ") + 1)) + " "
     traduirlatinta = Trim(traduir(primeraparaula, etidioma) + " " + traduir(segonaparaula, etidioma) + " " + atrim(colorrestant)) + " " + traduir(detalltinter, etidioma)
       Else: traduirlatinta = Trim(color)
  End If
  If traduirlatinta = atrim(primeraparaula + " " + segonaparaula) Then traduirlatinta = color
     
End Function
Sub possardadesaltemporal(rsttemp As Recordset)
  Dim rstclixe As Recordset
  Dim rstmodifi As Recordset
  Dim rsttintes As Recordset
  Dim rsttintesreprint As Recordset
  Dim vtintesafectades As Byte
  Set rstclixe = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(id_treball))
  Set rstmodifi = dbclixes.OpenRecordset("Select * from modificacions where id_Treball=" + atrim(id_treball) + " and ordre=" + atrim(ordremodificacio))
  Set rsttintes = dbclixes.OpenRecordset("Select * from tintes where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " order by ordretinter")
  If Not rsttintes.EOF Then
      If cadbl(rsttintes!tinterlinkambid_treball) > 0 Then
         Set rsttintes = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(rsttintes!tinterlinkambid_treball))
      End If
  End If
  If rstclixe.EOF Or rstmodifi.EOF Or rsttintes.EOF Then MsgBox "No s'ha trobat la informació necessaria per generar l'informe", vbCritical, "Error": Exit Sub
  rsttemp.Edit
  With rsttemp
  !nomtreball = atrim(rstclixe!marca) + " - " + atrim(rstclixe!linia)
  !nomclient = atrim(rstclixe!nomclienttemporal)
  !numcomanda = cadbl(buscarcomandespendents)
  !nomproveidor = formclixes.nomproveidor
  !ample = cadbl(rstmodifi!amplelamina)
  !llarg = cadbl(rstmodifi!desarroll) / 10
  !material = atrim(buscarcolormaterialpc2(!numcomanda, "linkcomanda1"))
  !material = traduir(atrim(buscarcolormaterialpc2(!numcomanda, "comanda")), etidioma) + IIf(!material <> "", "+" + traduir(atrim(!material), etidioma), "")
  !impressio = formclixes.comboformaimpresio
  !impressioreprint = IIf(rstmodifi!reprintformaimpres = "T", "Transparent", IIf(rstmodifi!reprintformaimpres = "N", "Normal", ""))
  !cilindre = cadbl(rsttintes!cilindre) / 10
  Set rsttintes = dbclixes.OpenRecordset("Select * from tintes where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " order by ordretinter")
  Set rsttintesreprint = dbclixes.OpenRecordset("Select * from tintes where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio * -1) + " order by ordretinter")
  For i = 1 To 8
    If rsttintes!afectatspelcanvi Then
       vtintesafectades = vtintesafectades + 1
       .Fields("tinta" + atrim(i)) = Mid(traduirlatinta(rsttintes!color, rsttintes!coditinta, rsttintes!detalltinter), 1, 30)
       .Fields("linia" + atrim(i)) = atrim(rsttintes!anilox)
       If rsttintes!continuu Then !continu = True
    End If
    If Not rsttintesreprint.EOF And !impressioreprint <> "" Then
      If rsttintesreprint!afectatspelcanvi Then
       .Fields("tintareprint" + atrim(i)) = traduirlatinta(rsttintesreprint!color, rsttintesreprint!coditinta, rsttintesreprint!detalltinter)
       .Fields("liniareprint" + atrim(i)) = atrim(rsttintesreprint!anilox)
       If rsttintesreprint!continuu Then !continureprint = True
      End If
      rsttintesreprint.MoveNext
    End If
    rsttintes.MoveNext
    If rsttintes.EOF Then i = 8
  Next i
  
  !tintes = vtintesafectades 'cadbl(FormClixes.Text19)
  !polimers = 1
  !espesorpolimer = cadbl(formclixes.Text11)
  !muntatgebandes = cadbl(formclixes.Text10)
  !pautacelula = IIf(formclixes.Combo4 = "", "No", "Si")
  !pautasituacio = atrim(formclixes.Combo4)
  !motius = !cilindre / IIf(!llarg = 0, 1, !llarg)
  !bandaseguiment = IIf(formclixes.Combo5 = "", "No", "Si")
  !bandasituacio = formclixes.Combo5
  !banda_a = atrim(formclixes.Text6) + "_mm"
  !deixarsang = IIf(formclixes.Combo3 = "", "No", "Si")
  !sangsituacio = formclixes.Combo3
  !sang_de = cadbl(formclixes.Text7)
  !codidebarres = atrim(rstclixe!codidebarres)
  If !figurafilm <> 0 Then copiafoto rutadelfitxer(llegir_ini("General", "cami", fitxerini)) + "figuresmuntatge\F" + atrim(!figurafilm) + ".jpg", !muntatge1
  If !figuraformat <> 0 Then copiafoto rutadelfitxer(llegir_ini("General", "cami", fitxerini)) + "figuresmuntatge\F" + atrim(!figuraformat) + ".jpg", !muntatge2
  traduirtotselsvalors rsttemp
  End With
  rsttemp.Update
  Set rstclixe = Nothing
  Set rstmodifi = Nothing
  Set rsttintes = Nothing
End Sub
Sub traduirtotselsvalors(rst As Recordset)
  Dim i As Byte
  For i = 0 To rst.Fields.Count - 1
      If rst.Fields(i).Type = 10 Then
        rst.Fields(i).Value = traduir(atrim(rst.Fields(i).Value), etidioma)
        'MsgBox rst.Fields(i).Name
      End If
  Next i
End Sub
Function copiafoto(foto As String, fldTO As Field)

'This function takes the source field image and copies it
'into the destination field.
'The function first saves the image in the source field to a
'temp file on disc. Then reads this temp file into
'the destination field.
'The temp file is then deleted
'On Error Resume Next

Dim iFieldSize  As Long
Dim varChunk    As Variant
Dim baData()    As Byte
Dim iOffset     As Long
Dim sFName      As String
Dim iFileNum    As Long
Dim cnt         As Long
Dim z()         As Byte

Const CONCHUNKSIZE As Long = 16384

Dim iChunks As Long
Dim iFragmentSize As Long
    
    'Get a unique random filename
    If Not existeix(foto) Then Exit Function
    sFName = foto
    
    Open sFName For Binary Access Read As #1
    ReDim z(FileLen(sFName))
    Get #1, , z()
     fldTO.AppendChunk z
    Close #1
    
    'Delete the file
    'Kill (sFName)
    
End Function

Sub afegircamps(bd As Database)
      
    bd.Execute ("alter table comandesproveidor add column nomtreball text(90)")
    bd.Execute ("alter table comandesproveidor add column nomclient text(50)")
    bd.Execute ("alter table comandesproveidor add column numcomanda double")
    bd.Execute ("alter table comandesproveidor add column nomproveidor text(25)")
    bd.Execute ("alter table comandesproveidor add column ample double")
    bd.Execute ("alter table comandesproveidor add column llarg double")
    bd.Execute ("alter table comandesproveidor add column material text(100)")
    bd.Execute ("alter table comandesproveidor add column impressio text(20)")
    bd.Execute ("alter table comandesproveidor add column impressioreprint text(20)")
    bd.Execute ("alter table comandesproveidor add column tintes byte")
    bd.Execute ("alter table comandesproveidor add column tinta1 text(30)")
    bd.Execute ("alter table comandesproveidor add column tinta2 text(30)")
    bd.Execute ("alter table comandesproveidor add column tinta3 text(30)")
    bd.Execute ("alter table comandesproveidor add column tinta4 text(30)")
    bd.Execute ("alter table comandesproveidor add column tinta5 text(30)")
    bd.Execute ("alter table comandesproveidor add column tinta6 text(30)")
    bd.Execute ("alter table comandesproveidor add column tinta7 text(30)")
    bd.Execute ("alter table comandesproveidor add column tinta8 text(30)")
    bd.Execute ("alter table comandesproveidor add column linia1 double")
    bd.Execute ("alter table comandesproveidor add column linia2 double")
    bd.Execute ("alter table comandesproveidor add column linia3 double")
    bd.Execute ("alter table comandesproveidor add column linia4 double")
    bd.Execute ("alter table comandesproveidor add column linia5 double")
    bd.Execute ("alter table comandesproveidor add column linia6 double")
    bd.Execute ("alter table comandesproveidor add column linia7 double")
    bd.Execute ("alter table comandesproveidor add column linia8 double")
    
    bd.Execute ("alter table comandesproveidor add column tintareprint1 text(30)")
    bd.Execute ("alter table comandesproveidor add column tintareprint2 text(30)")
    bd.Execute ("alter table comandesproveidor add column tintareprint3 text(30)")
    bd.Execute ("alter table comandesproveidor add column tintareprint4 text(30)")
    bd.Execute ("alter table comandesproveidor add column tintareprint5 text(30)")
    bd.Execute ("alter table comandesproveidor add column tintareprint6 text(30)")
    bd.Execute ("alter table comandesproveidor add column tintareprint7 text(30)")
    bd.Execute ("alter table comandesproveidor add column tintareprint8 text(30)")
    bd.Execute ("alter table comandesproveidor add column liniareprint1 double")
    bd.Execute ("alter table comandesproveidor add column liniareprint2 double")
    bd.Execute ("alter table comandesproveidor add column liniareprint3 double")
    bd.Execute ("alter table comandesproveidor add column liniareprint4 double")
    bd.Execute ("alter table comandesproveidor add column liniareprint5 double")
    bd.Execute ("alter table comandesproveidor add column liniareprint6 double")
    bd.Execute ("alter table comandesproveidor add column liniareprint7 double")
    bd.Execute ("alter table comandesproveidor add column liniareprint8 double")
    
    bd.Execute ("alter table comandesproveidor add column polimers byte")
    bd.Execute ("alter table comandesproveidor add column espesorpolimer double")
    bd.Execute ("alter table comandesproveidor add column muntatgebandes double")
    bd.Execute ("alter table comandesproveidor add column pautacelula text(3)")
    bd.Execute ("alter table comandesproveidor add column pautasituacio text(20)")
    bd.Execute ("alter table comandesproveidor add column cilindre double")
    bd.Execute ("alter table comandesproveidor add column motius double")
    bd.Execute ("alter table comandesproveidor add column continu text(3)")
    bd.Execute ("alter table comandesproveidor add column continureprint text(3)")
    bd.Execute ("alter table comandesproveidor add column bandaseguiment text(3)")
    bd.Execute ("alter table comandesproveidor add column bandasituacio text(10)")
    bd.Execute ("alter table comandesproveidor add column banda_a text(60)")
    bd.Execute ("alter table comandesproveidor add column deixarsang text(3)")
    bd.Execute ("alter table comandesproveidor add column sangsituacio text(20)")
    bd.Execute ("alter table comandesproveidor add column sang_de double")
    bd.Execute ("alter table comandesproveidor add column codidebarres text(20)")
    bd.Execute ("alter table comandesproveidor add column muntatge1 longbinary")
    bd.Execute ("alter table comandesproveidor add column muntatge2 longbinary")
   
End Sub
Sub borrartaulatemporal(db As Database)
  On Error Resume Next
   db.Execute "drop table comandesproveidor"
End Sub
Function nomfitxertemporal() As String
     nomfitxertemporal = "c:\temp\~cl" + Format(Now, "ddmmhhnnss") + ".mdb"
     On Error Resume Next
     Kill "c:\temp\~cl*.*"
End Function

Private Sub Command2_Click()
   If Not datareposicions.Recordset.EOF Then
       If InputBox("Segur que vols borrar aquesta reposicio?" + Chr(10) + "Escriu [SEGUR] per borrar-la.", "Borrar reposició") = "SEGUR" Then
          datareposicions.Recordset.Delete
          datareposicions.Refresh
       End If
   End If
End Sub

Private Sub Command3_Click()
   eliminar_totalacomanda
End Sub
Sub eliminar_totalacomanda()
   If UCase(InputBox("Segur que vols eliminar tota la informació d'aquesta comanda al fotogravador?" + Chr(10) + "Escriu [SEGUR] per eliminarla.")) = "SEGUR" Then
        formcomandaclixes.tag = "nogravar"
        dbclixes.Execute "delete * from comandesfotogravador where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio)
        dbclixes.Execute "delete * from reposicionsfotogravador where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio)
        Unload formcomandaclixes
   End If
End Sub

Private Sub Command7_Click()
   Dim fitxerpdftemporal As String
   If MsgBox("Si fas d'acord, no puc controlar si realment has enviat aquesta comanda al fotogravador" + Chr(10) + " passaré la comanda a enviada suposant que l'enviament ha estat correcte.", vbCritical + vbOKCancel, "Atenció") = vbCancel Then Exit Sub
   generarcomandaimpresa True
   
   fitxerpdftemporal = "c:\temp\clixescomandes\clixescomandaproveidor_" + atrim(id_treball) + "-" + atrim(ordremodificacio) + ".pdf"
   copiarelpdfalacarpetadeltreball fitxerpdftemporal, atrim(id_treball) + "-" + atrim(ordremodificacio)
   imatgeenviat.visible = True
   gravarcomanda
   carregarcomanda
   possarmodificacionspolimersoclixes
End Sub
Sub possarmodificacionsclixesrebuts()
  Dim rst As Recordset
  Dim proximnum As Integer
  Dim numordre As Integer
  Dim liniamodificacio As String
  Set rst = dbclixes.OpenRecordset("select * from clixes_modifi where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
  If rst.EOF Then Exit Sub
  rst.MoveLast
  proximnum = cadbl(Mid(rst!descripcio, 1, 3)) + 1
  If rst!id_estatclixe = 20 Then GoTo fi
  liniamodificacio = InputBox("Escriu la linia que vols que s'afegeixi a modificacions." + Chr(10) + "per exemple possar l'estantaria.", "Linia modificacion CLIXES REBUTS", atrim(proximnum) + " CLIXES REBUTS")
  If liniamodificacio = "" Then MsgBox "No s'ha afegit cap linia de modificació", vbInformation, "Atenció": GoTo fi
  numordre = rst!ordre + 1
  rst.Edit
  rst!data_fi = Date
  rst.Update
  rst.AddNew
  rst!id_treball = id_treball
  rst!ordremodificacio = ordremodificacio
  rst!id_estatclixe = 20
  rst!descripcioestat = "CLIXES REBUTS"
  rst!data_inici = Date
  rst!data_prevista = Date
  rst!ordre = numordre
  rst!descripcio = liniamodificacio
  rst.Update
fi:
  Set rst = Nothing
End Sub

Sub possarmodificacionspolimersoclixes()
  Dim rst As Recordset
  Dim proximnum As Integer
  Dim numordre As Integer
  Set rst = dbclixes.OpenRecordset("select * from clixes_modifi where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
  If rst.EOF Then Exit Sub
  rst.MoveLast
  proximnum = cadbl(Mid(rst!descripcio, 1, 3)) + 1
  If rst!id_estatclixe = 15 And atrim(Mid(rst!descripcio, 4, 8)) = "PASSADA" Then GoTo fi
  numordre = rst!ordre + 1
  rst.Edit
  rst!data_fi = Date
  rst.Update
  rst.AddNew
  rst!id_treball = id_treball
  rst!ordremodificacio = ordremodificacio
  rst!id_estatclixe = 15
  rst!descripcioestat = "POLIMERS O CLIXES"
  rst!data_inici = cdata
  rst!data_prevista = cdataestimada
  rst!ordre = numordre
  rst!descripcio = atrim(proximnum) + " PASSADA FULLA A FOTOGRAVADOR PER FER POLIMERS"
  rst.Update
fi:
  Set rst = Nothing
End Sub
Sub copiarelpdfalacarpetadeltreball(fitxerpdf As String, nump As String)
  Dim rutacarpetatreball As String
  Dim fitxerdesti As String
  formclixes.crearruta ruta_documentacio_clixes + "\" + Format(id_treball, "00000")
  rutacarpetatreball = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Arxiu_documentacio_relacionada" + "\v" + atrim(ordremodificacio)
  formclixes.crearruta rutacarpetatreball
  fitxerdesti = rutacarpetatreball + "\clixescomandaproveidor_" + atrim(nump) + ".pdf"
  If existeix(fitxerdesti) Then
     On Error GoTo fi
     Kill fitxerdesti
  End If
  Copiar_Fitxer fitxerpdf, fitxerdesti
fi:
End Sub
Private Sub Command8_Click()
   
   If imatgeacceptat.visible Then
      If MsgBox("Aquesta comanda ja ha estat revisada vols tornar-hi?" + Chr(10) + "Si dius que si també s'haurà de tornar a enviar.", vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   End If
   imatgeacceptat.visible = False
   imatgeenviat.visible = False
   generarcomandaimpresa
   If MsgBox("Si tot es correcte vols donar OK a aquesta comanda?", vbYesNo + vbDefaultButton2 + vbInformation, "Ok comanda") = vbYes Then
       imatgeacceptat.visible = True
      Else: imatgeacceptat.visible = False
   End If
   gravarcomanda
End Sub

Private Sub ctipusproducte_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub ctipusproducte_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub ff1_DblClick()
    colocarselfilm 1
  
End Sub

Private Sub ff2_DblClick()
    colocarselfilm 2
End Sub

Private Sub ff3_DblClick()
    colocarselfilm 3
End Sub

Private Sub ff4_DblClick()
    colocarselfilm 4
End Sub

Private Sub fo1_DblClick()
    colocarselformat 1
End Sub

Private Sub fo2_DblClick()
    colocarselformat 2
End Sub

Private Sub fo3_DblClick()
    colocarselformat 3
End Sub

Private Sub fo4_DblClick()
    colocarselformat 4
End Sub

Private Sub Form_Load()
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
   etdescripciocomanda = formclixes.marcaproducte + " - " + formclixes.liniaproducte + Chr(10) + " Treball: " + atrim(id_treball) + " v" + atrim(ordremodificacio)
   carregarcomanda
   datareposicions.DatabaseName = camiclixes
   datareposicions.RecordSource = "select * from reposicionsfotogravador where id_treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio) + " order by dataenviament desc"
   datareposicions.Refresh
   etidioma = idiomafotogravador(formclixes.modificacions.Recordset!fotograbador)
End Sub
Sub carregarcomanda()
   Dim rst As Recordset
   imatgeenviat.visible = False
   imatgeacceptat.visible = False
   Set rst = dbclixes.OpenRecordset("select * from comandesfotogravador where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
   If Not rst.EOF Then
      cdata = rst!datacomanda
      cdataestimada = rst!dataprevista
      ctipusproducte = atrim(rst!tipusproducte)
      ctamanycelula = rst!tamanycelula
      ccolorscelula = rst!colorscelula
      cobservacions = rst!observacions
      cpreu = rst!preuclixe
      colocarselfilm cadbl(rst!figurafilm)
      colocarselformat cadbl(rst!figuraformat)
      imatgeenviat.visible = rst!okenviat
      imatgeacceptat.visible = rst!okperenviar
      If IsDate(rst!datarecepcio) Then
         etdatarecepcio.tag = atrim(rst!datarecepcio): etdatarecepcio.visible = True
         etdatarecepcio = "Data Rebuda: " + Format(rst!datarecepcio, "dd/mm/yy")
        Else: If rst!okenviat And rst!okperenviar Then Command1.visible = True
      End If
        Else: possarvalorspredeterminats
   End If
   Set rst = Nothing
End Sub
Function afegirdieslaborablesadata(vdata As Date, diesaafegir As Long) As String
   Dim datafinal As Date
   datafinal = vdata
   For i = 1 To diesaafegir
       datafinal = DateAdd("d", 1, datafinal)
       While Format(datafinal, "dddd", vbMonday) = "sábado" Or Format(datafinal, "dddd", vbMonday) = "domingo"
         datafinal = DateAdd("d", 1, datafinal)
       Wend
   Next i
   afegirdieslaborablesadata = Format(datafinal, "dd/mm/yy")
End Function
Sub possarvalorspredeterminats()
    cdata = Date
    cdataestimada = afegirdieslaborablesadata(Date, 3)
    ctipusproducte = "Film"
End Sub
Private Sub Image4_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
  If imatgeacceptat.visible And imatgeenviat.visible = False Then
     MsgBox "Aquesta comanda la tens acceptada però no està enviada.", vbInformation, "Atenció"
  End If
  If imatgeacceptat.visible Then gravarcomanda
  Set dbtintes = Nothing
End Sub
Function cadate(v As Variant) As Date
   If Not IsDate(v) Then
      cadate = 0
    Else: cadate = CVDate(v)
   End If
   
End Function
Sub gravarcomanda()
    Dim rst As Recordset
    If formcomandaclixes.tag = "nogravar" Then Exit Sub
    Set rst = dbclixes.OpenRecordset("select * from comandesfotogravador where id_Treball=" + atrim(id_treball) + " and ordremodificacio=" + atrim(ordremodificacio))
    If rst.EOF Then
      rst.AddNew
       Else: rst.Edit
    End If
    rst!id_treball = id_treball
    rst!ordremodificacio = ordremodificacio
    rst!datacomanda = cadate(cdata)
    rst!dataprevista = cadate(cdataestimada)
    rst!tipusproducte = atrim(ctipusproducte)
    rst!tamanycelula = atrim(ctamanycelula)
    rst!colorscelula = atrim(ccolorscelula)
    rst!observacions = cobservacions
    rst!preuclixe = cadbl(cpreu)
    rst!figurafilm = cadbl(fsel.tag)
    rst!figuraformat = cadbl(fosel.tag)
    rst!okperenviar = imatgeacceptat.visible
    rst!okenviat = imatgeenviat.visible
    If IsDate(etdatarecepcio.tag) Then
       rst!datarecepcio = etdatarecepcio.tag
    End If
    rst.Update
    Set rst = Nothing
End Sub

Private Sub reixareposicions_DblClick()
   Dim vresp As String
   Dim vidalb As String
   If datareposicions.Recordset.EOF Then Exit Sub
   vidalb = cadbl(datareposicions.Recordset!ID)
   If reixareposicions.Columns(reixareposicions.col).DataField = "preu" Then
       vresp = InputBox("Entra el preu de l'albarà del fotogravador." + "Aquest import mai es facturarà al client", "Preu albarà")
       If vresp <> "" Then
           datareposicions.Recordset.Edit
           datareposicions.Recordset!preu = cadbl(vresp)
           datareposicions.Recordset.Update
           datareposicions.Refresh
       End If
   End If
    If reixareposicions.Columns(reixareposicions.col).DataField = "dataalbara" Then
       vresp = InputBox("Entra la data de l'albarà del fotogravador.", "Data albarà")
       If vresp <> "" And IsDate(vresp) Then
           datareposicions.Recordset.Edit
           datareposicions.Recordset!dataalbara = CVDate(vresp)
           datareposicions.Recordset.Update
           datareposicions.Refresh
       End If
   End If
   If reixareposicions.Columns(reixareposicions.col).DataField = "num_alb" Then
       vresp = InputBox("Entra el numero de l'albarà del fotogravador.", "Nº albarà")
       If vresp <> "" Then
           datareposicions.Recordset.Edit
           datareposicions.Recordset!num_alb = atrim(vresp)
           datareposicions.Recordset.Update
           datareposicions.Refresh
       End If
   End If
   datareposicions.Recordset.FindFirst "id=" + atrim(cadbl(vidalb))
End Sub

Private Sub Timer1_Timer()
  cpostit.visible = False
  Timer1.Enabled = False
End Sub
