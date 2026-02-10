VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form subbusqueda 
   Caption         =   "Fer la busqueda sel.leccionada"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12435
   Icon            =   "subbusqueda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   12435
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Busqueda per Resum"
      Height          =   390
      Left            =   9510
      TabIndex        =   40
      Top             =   1335
      Width           =   1905
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modificacio Xls spontex"
      Height          =   270
      Left            =   3270
      TabIndex        =   37
      Top             =   1380
      Visible         =   0   'False
      Width           =   2010
   End
   Begin Crystal.CrystalReport report 
      Left            =   5565
      Top             =   1410
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1815
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -30
      Visible         =   0   'False
      Width           =   2400
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "subbusqueda.frx":0442
      Height          =   8025
      Left            =   45
      OleObjectBlob   =   "subbusqueda.frx":0452
      TabIndex        =   9
      Top             =   1710
      Width           =   9375
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   135
      TabIndex        =   3
      Tag             =   "12105"
      Top             =   15
      Width           =   12105
      Begin VB.TextBox cnumtreball 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6165
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   900
         Width           =   825
      End
      Begin VB.CheckBox metode 
         Caption         =   "Mètode"
         Height          =   210
         Left            =   75
         TabIndex        =   53
         Top             =   -15
         Width           =   900
      End
      Begin VB.CheckBox checkdataentrega 
         Caption         =   "Amb data d'entrega"
         Height          =   195
         Left            =   3690
         TabIndex        =   51
         ToolTipText     =   "Sense agrupar"
         Top             =   885
         Width           =   1815
      End
      Begin VB.CheckBox checkunperun 
         Caption         =   "Comanda x comanda"
         Height          =   195
         Left            =   3690
         TabIndex        =   50
         ToolTipText     =   "Sense agrupar"
         Top             =   1065
         Width           =   1815
      End
      Begin VB.TextBox cdatafi 
         Height          =   285
         Left            =   2550
         TabIndex        =   48
         Top             =   945
         Width           =   1035
      End
      Begin VB.TextBox cdatainici 
         Height          =   285
         Left            =   1305
         TabIndex        =   46
         Top             =   945
         Width           =   1035
      End
      Begin VB.CheckBox totsproductes 
         Caption         =   "Tots"
         Height          =   225
         Left            =   11130
         TabIndex        =   35
         Tag             =   "11130"
         Top             =   465
         Value           =   1  'Checked
         Width           =   840
      End
      Begin VB.CheckBox pendent 
         Caption         =   "Pendent"
         Height          =   225
         Left            =   11130
         TabIndex        =   33
         Tag             =   "11130"
         Top             =   210
         Width           =   930
      End
      Begin VB.Frame bobrebimpresses 
         Caption         =   "Bobines Impresses Reb"
         Height          =   555
         Left            =   9195
         TabIndex        =   29
         Tag             =   "9195"
         Top             =   690
         Width           =   1950
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SIC"
            Height          =   285
            Index           =   24
            Left            =   1215
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   210
            Width           =   360
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "IC"
            Height          =   285
            Index           =   16
            Left            =   825
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   210
            Width           =   360
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "MR"
            Height          =   285
            Index           =   15
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   210
            Width           =   375
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "UR"
            Height          =   285
            Index           =   14
            Left            =   450
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   210
            Width           =   360
         End
      End
      Begin VB.Frame bobinesreb 
         Caption         =   "BobinesReb"
         Height          =   555
         Left            =   5700
         TabIndex        =   26
         Tag             =   "5700"
         Top             =   690
         Width           =   1485
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "LC"
            Height          =   285
            Index           =   17
            Left            =   810
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   210
            Width           =   360
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "LR"
            Height          =   285
            Index           =   13
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   210
            Width           =   375
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TR"
            Height          =   285
            Index           =   12
            Left            =   450
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   210
            Width           =   360
         End
      End
      Begin VB.Frame bobinesimpresses 
         Caption         =   "Bobines Impresses"
         Height          =   555
         Left            =   7425
         TabIndex        =   23
         Tag             =   "7425"
         Top             =   690
         Width           =   1575
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "U"
            Height          =   285
            Index           =   9
            Left            =   795
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   210
            Width           =   360
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "M"
            Height          =   285
            Index           =   8
            Left            =   405
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   210
            Width           =   375
         End
      End
      Begin VB.Frame bobines 
         Caption         =   "Bobines"
         Height          =   555
         Left            =   11160
         TabIndex        =   20
         Tag             =   "11160"
         Top             =   690
         Width           =   870
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "L"
            Height          =   285
            Index           =   11
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   210
            Width           =   375
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "T"
            Height          =   285
            Index           =   10
            Left            =   450
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   210
            Width           =   360
         End
      End
      Begin VB.Frame formatsimpressos 
         Caption         =   "Formats Impressos"
         Height          =   555
         Left            =   8550
         TabIndex        =   12
         Tag             =   "8550"
         Top             =   120
         Width           =   2565
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "OR"
            Height          =   285
            Index           =   25
            Left            =   870
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   210
            Width           =   345
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "FIR"
            Height          =   285
            Index           =   23
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   210
            Width           =   360
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "FI"
            Height          =   285
            Index           =   18
            Left            =   1890
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   210
            Width           =   270
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "IR"
            Height          =   285
            Index           =   19
            Left            =   1605
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   210
            Width           =   270
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "OIR"
            Height          =   285
            Index           =   7
            Left            =   1215
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   210
            Width           =   375
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "OI"
            Height          =   285
            Index           =   6
            Left            =   555
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   210
            Width           =   315
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "O"
            Height          =   285
            Index           =   5
            Left            =   285
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   210
            Width           =   255
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "I"
            Height          =   285
            Index           =   4
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   210
            Width           =   240
         End
      End
      Begin VB.Frame formats 
         Caption         =   "Formats"
         Height          =   555
         Left            =   5445
         TabIndex        =   10
         Tag             =   "5445"
         Top             =   120
         Width           =   3105
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "F"
            Height          =   285
            Index           =   22
            Left            =   2730
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   210
            Width           =   330
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ACNR"
            Height          =   285
            Index           =   21
            Left            =   2130
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   210
            Width           =   570
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "LNR"
            Height          =   285
            Index           =   20
            Left            =   1650
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   210
            Width           =   450
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BCR"
            Height          =   285
            Index           =   3
            Left            =   1170
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   210
            Width           =   450
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BC"
            Height          =   285
            Index           =   2
            Left            =   810
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   210
            Width           =   345
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "B"
            Height          =   285
            Index           =   1
            Left            =   450
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   210
            Width           =   360
         End
         Begin VB.CommandButton productes 
            BackColor       =   &H00FFFFFF&
            Caption         =   "A"
            Height          =   285
            Index           =   0
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   210
            Width           =   375
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   555
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   225
         Width           =   1290
      End
      Begin VB.Label Label6 
         Caption         =   "Treball:"
         Height          =   210
         Left            =   5595
         TabIndex        =   55
         Top             =   945
         Width           =   705
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5235
         TabIndex        =   54
         Top             =   345
         Width           =   5970
      End
      Begin VB.Shape Shape1 
         Height          =   420
         Left            =   90
         Top             =   855
         Width           =   5445
      End
      Begin VB.Label Label4 
         Caption         =   "i"
         Height          =   210
         Left            =   2400
         TabIndex        =   49
         Top             =   990
         Width           =   180
      End
      Begin VB.Label Label3 
         Caption         =   "Entre dates:"
         Height          =   210
         Left            =   135
         TabIndex        =   47
         Top             =   1005
         Width           =   1125
      End
      Begin VB.Label nomproducte 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Top             =   540
         Width           =   5970
      End
      Begin VB.Label nomclient 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2775
         TabIndex        =   7
         Top             =   255
         Width           =   6165
      End
      Begin VB.Label Label2 
         Caption         =   "Producte:"
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   615
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Codi de Client:"
         Height          =   210
         Left            =   165
         TabIndex        =   5
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.CommandButton acceptar 
      Appearance      =   0  'Flat
      Caption         =   "Acceptar La Busqueda"
      Height          =   360
      Left            =   7050
      TabIndex        =   2
      Top             =   1350
      Width           =   2205
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar per Formulari"
      Height          =   375
      Left            =   135
      TabIndex        =   4
      Top             =   1335
      Width           =   2190
   End
   Begin VB.Label etestatusllistat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "           "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   11520
      TabIndex        =   52
      Top             =   1410
      Width           =   660
   End
   Begin VB.Menu llistats 
      Caption         =   "Llistats"
      Begin VB.Menu desarrolls 
         Caption         =   "Desarrolls"
      End
      Begin VB.Menu general 
         Caption         =   "General"
      End
      Begin VB.Menu quantitatentregada 
         Caption         =   "Quantitat Entregada"
      End
   End
End
Attribute VB_Name = "subbusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbclixes As Database
Dim pend As String
Sub crear_llistaproductes()
  r = ""
  If Trim(Text1) <> "" Then r = " client=" + atrim(Text1)
  If Trim(Text2) <> "" Then
     If r = "" Then
        r = " producte='" + atrim(Text2) + "'"
      Else: r = r + " and producte='" + atrim(Text2) + "'"
     End If
  End If
  For j = 0 To productes.Count - 1
    If productes.Item(j).BackColor = QBColor(12) Then
        If produ = "" Then
           produ = " producte='" + productes.Item(j).Caption + "' "
          Else: produ = produ + " or producte='" + productes.Item(j).Caption + "' "
        End If
    End If
  Next j
  If produ = "" Then produ = " producte<>'@' " 'perque faci una busqueda filtrant qualsevol producte
  If Trim(Text2) = "" And totsproductes.Value <> 1 Then r = r + IIf(r <> "", " and ", "") + "(" + produ + ")"
  pend = IIf(pendent.Value = 1, " and proximaseccio<>'T' ", "")
End Sub
Sub emplenarcampsreixa(Optional ByRef campsreixa)
     iniconfigreixa = iniconfigreixa + ".ini"
     campsreixa = llegir_ini("general", "camps1", iniconfigreixa)
     
     ra = llegir_ini("general", "camps2", iniconfigreixa)
     If ra = "{[}]" Then ra = ""
     campsreixa = campsreixa + ra
     ra = llegir_ini("general", "camps3", iniconfigreixa)
     If ra = "{[}]" Then ra = ""
     campsreixa = campsreixa + ra
     
     If InStr(1, campsreixa, "texteimpressio") > 0 Then campsreixa = Mid(campsreixa, 1, InStr(1, campsreixa, "texteimpressio") - 1) + "marcailinia " + Mid(campsreixa, InStr(1, campsreixa, "texteimpressio") + 15)
     If InStr(1, campsreixa, "refclient") > 0 Then campsreixa = Mid(campsreixa, 1, InStr(1, campsreixa, "refclient") - 1) + "refinplacsa as Ref_Inp, coditarifa as Tarifa,numtreball as Treball,numordremodificacio as Versió, " + Mid(campsreixa, InStr(1, campsreixa, "refclient"))
 '    If InStr(1, campsreixa, "palets") > 0 Then campsreixa = campsreixa + "palets,kilospalets"
     
End Sub
Private Sub acceptar_Click()
  Dim campsreixa As String
  i = 0
  crear_llistaproductes
  
  ratoli "espera"
  If r <> "" Then
   iniconfigreixa = plantillabusqueda
   If iniconfigreixa <> "" Then
     emplenarcampsreixa campsreixa
     'MsgBox campsreixa
     Data1.RecordSource = "select " + campsreixa + " from comandesmesextresmestarifes where " + r + pend + " order by comanda DESC"
     AppActivate "Fer la busqueda sel.leccionada"
     acceptar.Tag = campsreixa
     Data1.Tag = r + pend + " order by comanda DESC"
     refrescar_reixa
     AppActivate "Fer la busqueda sel.leccionada"
     'reixa.SetFocus
       Else: MsgBox "No hi ha seleccionat cap producte", vbInformation, "Atenció": Data1.Tag = "": acceptar.Tag = ""
   End If
    Else:
       r = ""
       i = 2
       Unload subbusqueda
      'MsgBox "No hi ha res a buscar...", 64, "Atenció": Text1.SetFocus
  End If
  ratoli "normal"
 reixa.Tag = "b"
End Sub

Private Sub cdatafi_Change()
  If Not IsDate(cdatafi) Then
      cdatafi.BackColor = QBColor(12)
       Else: cdatafi.BackColor = QBColor(15)
  End If
End Sub

Private Sub cdatainici_Change()
  If Not IsDate(cdatainici) Then
      cdatainici.BackColor = QBColor(12)
       Else: cdatainici.BackColor = QBColor(15)
  End If
End Sub

Private Sub cnumtreball_DblClick()
netejar_reixa_busqueda
End Sub

Private Sub Command1_Click()
 i = 1
 Unload subbusqueda
End Sub

Private Sub Command2_Click()
Dim excel_app As Object
Dim excel_sheet As Object
Dim row As Integer
Dim statement As String

    Screen.MousePointer = vbHourglass
    DoEvents

    ' Create the Excel application.
    Set excel_app = CreateObject("Excel.Application")

    ' Uncomment this line to make Excel visible.
'    excel_app.Visible = True

    ' Open the Excel spreadsheet.
    
    excel_app.Workbooks.Open FileName:="c:\spontex\sagunto.xls"

    ' Check for later versions.
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.activeSheet
    Else
        Set excel_sheet = excel_app
    End If

    
    ' Get data from the database and insert
    ' it into the spreadsheet.
    row = 1
    Do While row < 1000
        r = excel_sheet.Cells(row, 9)
        If cadbl(r) > 100000 Then
         Set rsttmp = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(r)))
         If Not rsttmp.EOF Then
              r = cadbl(rsttmp!cilindres)
              If r <= 100 Then r = r * 10
              excel_sheet.Cells(row, 7) = r
              r = cadbl(rsttmp!dessarroll)
              If r <= 100 Then r = r * 10
              excel_sheet.Cells(row, 8) = r
         End If
        End If
        row = row + 1
    Loop


    ' Comment the rest of the lines to keep
    ' Excel running so you can see it.

    ' Close the workbook saving changes.
    excel_app.ActiveWorkbook.Close True

    ' Close Excel.
    excel_app.Quit
    Set excel_sheet = Nothing
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault
End Sub
Sub llistat_general()
Dim taulatemp As String
Dim camps As String
ratoli "espera"
If acceptar.Tag = "" Then MsgBox "No hi ha res per llistar": ratoli "normal": Exit Sub
taulatemp = "c:\temporal.mdb"
If existeix(taulatemp) Then Kill taulatemp
DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
campsreixa = llegir_ini("general", "camps1", "totsproductes.ini")
ra = llegir_ini("general", "camps2", "totsproductes.ini")
If ra = "{[}]" Then ra = ""
campsreixa = campsreixa + ra
ra = llegir_ini("general", "camps3", "totsproductes.ini")
If ra = "{[}]" Then ra = ""
camps = campsreixa + ra
dbtmp.Execute ("select " + camps + " into temporal in '" + taulatemp + "' from comandes where " + Data1.Tag)
report.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistattotsproductes.rpt"
report.DataFiles(0) = taulatemp
report.Destination = crptToWindow
report.Action = 1
ratoli "normal"
End Sub
Sub triarclient()
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = formcomandes.Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from clients"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text1.Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   nomclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
  End If
  Unload formseleccio
  
End Sub
Sub triarproducte()
  Load formseleccio
  formseleccio.Data1.DatabaseName = formcomandes.Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from productes"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text2.Text = atrim(formseleccio.Data1.Recordset!codi)
   nomproducte.Caption = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  
  Unload formseleccio
  
End Sub
Sub refrescar()
 Dim tipusdato As Byte
 Dim grandoto As Integer
 Dim espais As Byte
 Data1.Refresh
 reixa.Refresh
 reixa.ReBind
 reixa.AllowUpdate = False
 On Error GoTo fi
 For j = 0 To Data1.Recordset.Fields.Count
'   tipusdato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Type
   v = llegir_ini("AmplesReixaBusqueda", UCase(reixa.Columns(j).Caption), fitxerini)
   
   If v = "{[}]" Then
     tipusdato = Data1.Recordset.Fields(j).Type
     grandato = Data1.Recordset.Fields(reixa.Columns(j).DataField).Size
     If grandato < 5 Then grandato = 5
     v = grandato * 125
   End If
   v = cadbl(v)
   reixa.Columns(j).Width = v
   reixa.Columns(j).Caption = UCase(reixa.Columns(j).Caption)
 Next j
fi:
End Sub
Sub refrescar_reixa()
 Dim tipusdato As Byte
 Dim grandoto As Integer
 Dim espais As Byte
 DoEvents
 DoEvents
 'MsgBox Data1.RecordSource
 Data1.Refresh
 reixa.Refresh
 reixa.ReBind
 reixa.AllowUpdate = False
 On Error GoTo fi
 For j = 0 To Data1.Recordset.Fields.Count
'   tipusdato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Type
   v = llegir_ini("AmplesReixaBusqueda", UCase(reixa.Columns(j).Caption), iniconfigreixa)
   
   If v = "{[}]" Then
     tipusdato = Data1.Recordset.Fields(j).Type
     grandato = Data1.Recordset.Fields(reixa.Columns(j).DataField).Size
     If grandato < 5 Then grandato = 5
     v = grandato * 125
   End If
   v = cadbl(v)
   reixa.Columns(j).Width = v
   'reixa.Columns(i).Caption = UCase(reixa.Columns(i).Caption)
 Next j
fi:
End Sub

Sub guardar_amples_reixa_busqueda()
If iniconfigreixa <> "" Then
  For j = 0 To Data1.Recordset.Fields.Count - 1
   'MsgBox UCase(reixa.Columns(i).Caption)
   escriure_ini "AmplesReixaBusqueda", UCase(reixa.Columns(j).Caption), atrim(Redondejar(reixa.Columns(j).Width, 0)), iniconfigreixa
 Next j
End If
End Sub
Sub borrar_o_crear_taulatemporal(dbtemp As String)
  On Error Resume Next
  Kill dbtemp
  DBEngine.CreateDatabase dbtemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  
End Sub
Sub possar_datesdentrega()
  Dim rst As Recordset
  Dim rstc As Recordset
  Set rst = dbbaixes.OpenRecordset("select max(data) as maxdata,first(comanda) as numc from bobinesent group by comanda")
  Set rstc = dbtmp.OpenRecordset("select comanda,dataentrega from comandes_extres where dataentrega=null ")
  While Not rstc.EOF
     rst.FindFirst "numc=" + atrim(rstc!comanda)
     If Not rst.NoMatch Then
        If Not IsNull(rst!maxdata) Then dbtmp.Execute "update comandes_Extres set dataentrega=#" + atrim(rst!maxdata) + "# where comanda=" + atrim(rstc!comanda)
     End If
     rstc.MoveNext
  Wend
End Sub
Private Sub Command3_Click()
   ferlabusquedaperresum
End Sub
Sub ferlabusquedaperresum()
Dim dblocal As Database
 Dim subconsulta As String
 Dim rsttmp2 As Recordset
 Dim dbtemp As String
 Dim campsreixa As String
 Dim refinp As String
 Dim elwhere As String
 Static ultimaconsulta As String
 If vtreballbuscatsubbusqueda <> "" And cnumtreball.Tag = "1a" Then cnumtreball.Tag = "": ultimaconsulta = ""
  If etestatusllistat.Tag = "-" Then
    Unload Formconsultarefinplacsa
    etestatusllistat.Tag = ""
  End If
  ratoli "espera"
  'Workspaces(0).BeginTrans
  'dbtemp = "c:\temp\temporal.mdb"
  'taula_tmp = "tmp_consultaref" + Format(Now, "nnss")
  'borrar_o_crear_taulatemporal dbtemp
  'Set dbconsulta = DBEngine.OpenDatabase(dbtemp)
  'If Not crear_taula_tmp_consulta(taula_tmp) Then MsgBox "Error creant la taula": GoTo fi
  etestatusllistat = "Generant la consulta..."
  crear_llistaproductes
  r = r + pend
  If IsDate(cdatainici) And IsDate(cdatafi) And checkdataentrega.Value <> 1 Then vdates = "and (datacomanda>=#" + Format(cdatainici, "mm/dd/yy") + "# and datacomanda<=#" + Format(cdatafi, "mm/dd/yy") + "#)"
  If IsDate(cdatainici) And IsDate(cdatafi) And checkdataentrega.Value = 1 Then
     vdates = "and (dataentrega>=#" + Format(cdatainici, "mm/dd/yy") + "# and dataentrega<=#" + Format(cdatafi, "mm/dd/yy") + "#)"
     'possar_datesdentrega
  End If
  elwhere = " (" + r + ")"
  If vtreballbuscatsubbusqueda <> "" Then
       If InStr(1, vtreballbuscatsubbusqueda, "I") > 0 Then
              elwhere = elwhere + " and refinplacsa like '*" + Mid(vtreballbuscatsubbusqueda, InStr(1, vtreballbuscatsubbusqueda, "I")) + "'"
           Else: elwhere = elwhere + " and refinplacsa='" + atrim(vtreballbuscatsubbusqueda) + "'"
       End If
  End If
  
  If checkunperun.Value = 1 And IsDate(cdatainici) And IsDate(cdatafi) Then
      sql = "SELECT client as fclient,refinplacsa, Producte as Pr,refclient as Ref_, 1 as Q,datacomanda AS maxdata, comanda AS maxcomanda From comandesmesextresmestarifes"
      sql = sql + IIf(r <> "()", " where " + elwhere, "") + "and producte<>'PC' and producte<>'PC2' and producte<>'PCP' and producte<>'PCI3' " + vdates + " ORDER BY datacomanda DESC;"
'      Clipboard.Clear
'      Clipboard.SetText sql
    Else
     sql = "SELECT first(client) as fclient,refinplacsa, first(coditarifa) as Tr,first(producte) as Pr,last(refclient) as Ref_, count(*) as Q,Max(datacomanda) AS maxdata, Max(comanda) AS maxcomanda From comandesmesextresmestarifes"
     sql = sql + IIf(r <> "()", " where " + elwhere, "") + "and producte<>'PC' and producte<>'PC2' and producte<>'PCP' and producte<>'PCI3' " + vdates + " GROUP BY refinplacsa ORDER BY Max(datacomanda) DESC;"
  End If
  subbusqueda.Tag = IIf(r <> "()", " where " + elwhere, "") + "and producte<>'PC' and producte<>'PC2' and producte<>'PCP' and producte<>'PCI3' " + vdates
  If InStr(1, sql, "()") > 0 Or r = "()" Then MsgBox "Has d'escullir algun client i producte.", vbCritical, "Error": GoTo fi
 ' Clipboard.Clear
 ' Clipboard.SetText sql
  ' el comparar consultes per saltar-la
  
  If ultimaconsulta <> sql Then
    ultimaconsulta = sql
    Unload Formconsultarefinplacsa
    Set dbconsulta = Nothing
    If existeix("c:\temp\consultarefinp_tmp.mdb") Then Kill "c:\temp\consultarefinp_tmp.mdb"
      Else:
        GoTo saltarconsulta
  End If
  ' fi saltar consultes
  Set dblocal = dbtmp
'  Clipboard.Clear
'  Clipboard.SetText sql
  'vhora = Now
  wait 2
  Set rstconsulta = dblocal.OpenRecordset(sql)
  If rstconsulta.EOF Then MsgBox "No hi han resultats per aquest client": GoTo fi
  'Set rstconsulta = dblocal.OpenRecordset(sql)
  rstconsulta.MoveLast
  
saltarconsulta:
  ratoli "normal"
  etestatusllistat = "Carregant la reixa...": DoEvents
  'Workspaces(0).CommitTrans
  Formconsultarefinplacsa.Show 1
  'triarregistresubconsulta rsttmp2
  'If etestatusllistat = "Carregant la reixa..." Then etestatusllistat.Tag = "parant": GoTo fi
  If isloaded("Formconsultarefinplacsa") Then
     If Formconsultarefinplacsa.Caption = "Tancar" Then
        Unload Formconsultarefinplacsa: etestatusllistat.Tag = "-"
        Set dbconsulta = Nothing
        If existeix("c:\temp\consultarefinp_tmp.mdb") Then Kill "c:\temp\consultarefinp_tmp.mdb"
        Exit Sub
     End If
      Else: Exit Sub
  End If
  ratoli "espera"
  
  
  'Set rstconsulta = dbconsulta.OpenRecordset("select * from " + taula_tmp)
  'While Not rsttmp2.EOF
  '   Set rsttmp = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rsttmp2!maxcomanda))
  '   copiarregistreactiuaconsulta
  '   rsttmp2.MoveNext
  'Wend
  iniconfigreixa = plantillabusqueda
  If iniconfigreixa = "" Then GoTo fi
  emplenarcampsreixa campsreixa
  If isloaded("Formconsultarefinplacsa") Then
    refinp = atrim(Formconsultarefinplacsa.reixa.TextMatrix(Formconsultarefinplacsa.reixa.row, 2))
  End If
  If refinp = "Sense Referència" Then
     refinp = elwhere + " and (refinplacsa='' or refinplacsa=null)"
       Else: refinp = "refinplacsa='" + refinp + "'"
  End If
  Data1.RecordSource = "select " + campsreixa + " from comandesmesextresmestarifes where " + refinp + " order by comanda DESC"
 ' MsgBox Data1.RecordSource
  AppActivate "Fer la busqueda sel.leccionada"
  acceptar.Tag = campsreixa
  Data1.Tag = refinp + " order by comanda DESC"
  refrescar_reixa
  
  AppActivate "Fer la busqueda sel.leccionada"
  
fi:
  Set rsttmp2 = Nothing
  Set rsttmp = Nothing
  
  etestatusllistat = ""
  'dbconsulta.Close
  ratoli "normal"
  Set dblocal = Nothing
End Sub
Sub triarregistresubconsulta(rsttmp2 As Recordset)
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = formcomandes.Data1.DatabaseName
  Set formseleccio.Data1.Recordset = rsttmp2
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 1200
  formseleccio.DBGrid2.Columns(2).Width = 1200
  formseleccio.DBGrid2.Columns(4).Width = 1100
  formseleccio.Show 1
  If seleccioret = 1 Then
   rsttmp2.FindFirst "refinplacsa='" + atrim(formseleccio.Data1.Recordset!refinplacsa) + "'"
     Else: rsttmp2.MoveLast
  End If
  Unload formseleccio
  

End Sub
Sub copiarregistreactiuaconsulta()
   rstconsulta.AddNew
   rstconsulta!refclient = atrim(rsttmp!refclient)
   rstconsulta!producte = atrim(rsttmp!producte)
   rstconsulta!dessarroll = cadbl(rsttmp!dessarroll)
   rstconsulta!material1 = nommaterial(rsttmp!comanda)
   rstconsulta!material2 = nommaterial(rsttmp!linkcomanda1)
   rstconsulta!material3 = nommaterial(rsttmp!linkcomanda2)
   rstconsulta!texteimpresio = atrim(rsttmp!marcailinia)
   rstconsulta!Data = IIf(Not IsDate(rsttmp!datacomanda), Null, rsttmp!datacomanda)
   rstconsulta!ean = atrim(rsttmp!codibarras)
   rstconsulta!idtreball = cadbl(rsttmp!numtreball)
   rstconsulta.Update
End Sub
Function nommaterial(numcom As Double) As String
  Dim rst2 As Recordset
  Dim rstcom As Recordset
  Dim rstmat As Recordset
  Dim mat As String
  Set rstcom = dbtmp.OpenRecordset("select * from comandes  where comanda=" + atrim(cadbl(numcom)))
  If rstcom.EOF Then Exit Function
'mat 1
  Set rst2 = dbtmp.OpenRecordset("select familia from materials where codi=" + atrim(cadbl(rstcom!materialex)))
  If Not rst2.EOF Then If atrim(rst2!familia) <> "" Then m = atrim(cadbl(rst2!familia))
  Set rst2 = dbtmp.OpenRecordset("select familia from colorants where codi=" + atrim(cadbl(rstcom!colorex)))
  If Not rst2.EOF Then If atrim(rst2!familia) <> "" Then m = m + "/" + atrim(rst2!familia)
  Set rstmat = dbtmp.OpenRecordset("select * from mesureslineals where codi=" + atrim(cadbl(rstcom!mesuraesp)))
  If Not rstmat.EOF Then m = m + "/" + atrim(micresmaterial(atrim(rstmat!descripcio), cadbl(rstcom!espessor), atrim(rstcom!tubolam)))
  Set rst2 = Nothing
  Set rstcom = Nothing
  Set rstmat = Nothing
End Function

Function micresmaterial(descripcio As String, espesor As Double, tubolam As String) As Double
  r = espesor
  If descripcio = "GALGUES" Then
            If tubolam = "T" Then
                 r = Format(espesor / 4, "#,##0")
                  Else: r = Format(espesor / 2, "#,##0")
            End If
  End If
  If InStr(1, descripcio, "GR/") > 0 Then
    micresmaterial = espesor * -1
  End If
  micresmaterial = r
End Function
Function crear_taula_tmp_consulta(taula_tmp As String) As Boolean
  Dim camps(100, 2) As String
  
  On Error Resume Next
   dbconsulta.Execute "drop table " + taula_tmp
  On Error GoTo error
  i = 1
  camps(i, 1) = "refclient": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "producte": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "ample": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "dessarroll": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "material1": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "material2": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "material3": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "texteimpresio": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "comandes": camps(i, 2) = "double": i = i + 1
  camps(i, 1) = "data": camps(i, 2) = "date": i = i + 1
  camps(i, 1) = "ean": camps(i, 2) = "string": i = i + 1
  camps(i, 1) = "idtreball": camps(i, 2) = "double": i = i + 1
 
  dbconsulta.Execute ("create table " + taula_tmp + " (id long)")
  For i = 1 To 100
    If camps(i, 1) <> "" Then
       dbconsulta.Execute ("alter table " + taula_tmp + " add column " + camps(i, 1) + " " + camps(i, 2))
        Else: i = 1000
    End If
  Next i
  crear_taula_tmp_consulta = True
  Exit Function
error:
   crear_taula_tmp_consulta = False
  'ample double,plegat double,solapa double,espessor double,metres double,kilos double)"
  'dbtmpb.Execute ("create table tmp_lam_empalmes (" + camps + camps2 + camps3 + camps4) + ")"
End Function

Sub llistat_desarrolls()
  Dim taulatemp As String
  Dim camps As String
  Dim rs As Recordset
  ratoli "espera"
  If acceptar.Tag = "" Then MsgBox "No hi ha res per llistar": ratoli "normal": Exit Sub
  taulatemp = "c:\temporal.mdb"
  If existeix(taulatemp) Then Kill taulatemp
  DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  camps = " comanda, datacomanda, producte,numerotintes, dessarroll,cilindres,texteimpressio,obsimp1,amplereb "
'  dbtmp.Execute ("select " + camps + " into temporal in '" + taulatemp + "' from comandes where " + Data1.Tag)
  MsgBox "PENSA QUE ELS VALORS DE L'AMPLADA S'ESCRIUEN AMB , NO AMB . EX: 45,1"
  Set rs = dbtmp.OpenRecordset("select " + camps + " from comandes where " + Data1.Tag)
  While Not rs.EOF
    If rs!producte = "M" Then
      rs.Edit: rs!producte = "MR"
    End If
    If cadbl(rs!amplereb) = 0 Then
      If rs.EditMode = 0 Then rs.Edit
      rs!amplereb = cadbl(InputBox(atrim(rs!obsimp1), "Entra l'ample    Data: " + atrim(rs!datacomanda)))
      If cadbl(rs!amplereb) = 0 Then rs.MoveLast
    End If
   If rs.EditMode > 0 Then rs.Update
    rs.MoveNext
   
  Wend
  dbtmp.Execute ("select " + camps + " into temporal in '" + taulatemp + "' from comandes where " + Data1.Tag)
  report.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatdesarrolls.rpt"
  report.DataFiles(0) = taulatemp
  report.Destination = crptToWindow
  report.Action = 1
  ratoli "normal"
  Set rs = Nothing
End Sub
Function sumar_totals_entregats(numcomanda As String, producte As String) As String
  Dim ventregatm As Double
  Dim vpendentm As Double
  Dim ventregatk As Double
  Dim vpendentk As Double
  Dim rsttmpt As Recordset
  Set rsttmpt = dbbaixes.OpenRecordset("select metresisacs,data,kilosiunitats from bobinesent where comanda=" + numcomanda)
 ' rsttmpt.MoveLast
  While Not rsttmpt.EOF
    
     If rsttmpt!Data <> "" Then
       ventregatm = ventregatm + cadbl(rsttmpt!metresisacs)
       ventregatk = ventregatk + cadbl(rsttmpt!kilosiunitats)
      Else:
         vpendentm = vpendentm + cadbl(rsttmpt!metresisacs)
         vpendentk = vpendentk + cadbl(rsttmpt!kilosiunitats)
    End If
    rsttmpt.MoveNext
  Wend
  If producte = "L" Or producte = "T" Or producte = "LR" Or producte = "TR" Or producte = "LC" Or producte = "M" Or producte = "U" Or producte = "FI" Or producte = "MR" Or producte = "UR" Or producte = "IC" Then
    s = Format(ventregatm, "#,##0 Mtrs/Unit")
     Else: s = Format(ventregatk, "#,##0.00 Kg")
  End If
  
  sumar_totals_entregats = s
  'entregatm = Format(ventregatm, "#,##0.00")
  'pendentm = Format(vpendentm, "#,##0.00")
  'entregatk = Format(ventregatk, "#,##0.00")
  'pendentk = Format(vpendentk, "#,##0.00")

  Set rsttmpt = Nothing
End Function
Function demanardata(titol As String) As String
  Dim d As String
  d = "."
  While Not IsDate(d) And d <> ""
    d = InputBox(titol, "Entra una data", Date)
    If Not IsDate(d) Then MsgBox "Data Erronea"
  Wend
  demanardata = d
End Function
Sub llistat_quantitatentregada()
  Dim taulatemp As String
  Dim camps As String
  Dim rs As Recordset
  Dim dbt As Database
  Dim datainici As String
  Dim datafi As String
  Dim nomclient As String
  ratoli "espera"
  If acceptar.Tag = "" Then MsgBox "No hi ha res per llistar": ratoli "normal": Exit Sub
  datainici = demanardata("Entra data Inici")
  If datainici = "" Then ratoli "normal": Exit Sub
  datafi = demanardata("Entra data Fi")
  If datafi = "" Then ratoli "normal": Exit Sub
  taulatemp = "c:\temporal.mdb"
  If existeix(taulatemp) Then Kill taulatemp
  Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
  
  DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  camps = " client,comanda, datacomanda,dataentrega,refclient, producte,numerotintes, dessarroll,cilindres,amplereb,cantitatex,mesuracantex,'               ' as quantitat,'             ' as quantitatent ,obsimp1,texteimpressio"
'  dbtmp.Execute ("select " + camps + " into temporal in '" + taulatemp + "' from comandes where " + Data1.Tag)
  dbtmp.Execute ("select " + camps + " into temporal in '" + taulatemp + "' from comandes where " + Data1.Tag)
  Set dbt = DBEngine.OpenDatabase(taulatemp)
  Set rs = dbt.OpenRecordset("temporal")
  
  While Not rs.EOF
   If Not (rs!datacomanda >= CVDate(datainici) And rs!datacomanda <= CVDate(datafi)) Then rs.Delete: GoTo proxim
    If nomclient = "" Then
        Set rsttmp = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(rs!client)))
        If Not rsttmp.EOF Then nomclient = rsttmp!nom
    End If
    If Not rs.EOF Then Set rsttmp = dbtmp.OpenRecordset("select * from mesureslineals where codi=" + atrim(cadbl(rs!mesuracantex)))
    rs.Edit
    If Not rsttmp.EOF Then rs!quantitat = rsttmp!descripcio
    rs!quantitatent = sumar_totals_entregats(cadbl(rs!comanda), rs!producte)
    If cadbl(rs!cilindres) < 101 Then rs!cilindres = cadbl(rs!cilindres) * 10
    If cadbl(rs!dessarroll) < 101 Then rs!dessarroll = cadbl(rs!dessarroll) * 10
    If atrim(rs!texteimpressio) = "" Then
        s = rs!obsimp1
        If InStr(1, s, "=") > 0 Then
            s = Mid(s, InStr(1, s, "="))
            r = InStr(InStr(1, s, "=") + 1, s, "=") - InStr(1, s, "=")
            s = Mid(s, InStr(1, s, "="), r + 1)
            rs!texteimpressio = s
        End If
    End If
    rs.Update
proxim:
    rs.MoveNext
    ratoli "espera"
  Wend
  
  report.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatquantitatentregada.rpt"
  report.Formulas(0) = "nomclient='" + nomclient + "'"
  report.DataFiles(0) = taulatemp
  report.Destination = crptToWindow
  report.Action = 1
  ratoli "normal"
  Set rs = Nothing
  Set dbt = Nothing
  
  'SET DBBAIXES = NOTHING
End Sub

Private Sub desarrolls_Click()
llistat_desarrolls
End Sub

Private Sub etestatusllistat_DblClick()
'   subbusqueda.etestatusllistat.Tag = "parant"
End Sub

Private Sub Form_Activate()
Text1.SetFocus
Data1.DatabaseName = formcomandes.Data1.DatabaseName

 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 Then acceptar_Click
  If KeyCode = 13 Then SendKeys ("{TAB}")
  
  If KeyCode = 27 Then r = "sortir": Unload Me
End Sub

Private Sub Form_Load()
'subbusqueda.Width = Screen.Width - 2000
'subbusqueda.Height = Screen.Height - 2000
Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb", , True)
End Sub

Private Sub Form_Resize()
On Error Resume Next
reixa.Width = subbusqueda.Width - 200
Frame1.Width = subbusqueda.Width - 400
reixa.Height = subbusqueda.Height - reixa.Top - 700

colocarframesdinsdeframe1

End Sub
Sub colocarframesdinsdeframe1()
Dim ampleframe As Double
ampleframe = Frame1.Width - cadbl(Frame1.Tag)
formats.Left = cadbl(formats.Tag) + ampleframe
bobinesreb.Left = cadbl(bobinesreb.Tag) + ampleframe
bobinesimpresses.Left = cadbl(bobinesimpresses.Tag) + ampleframe
formatsimpressos.Left = cadbl(formatsimpressos.Tag) + ampleframe
bobrebimpresses.Left = cadbl(bobrebimpresses.Tag) + ampleframe
bobines.Left = cadbl(bobines.Tag) + ampleframe
pendent.Left = cadbl(pendent.Tag) + ampleframe
totsproductes.Left = cadbl(totsproductes.Tag) + ampleframe
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'If subbusqueda.etestatusllistat <> "" Then
 '  subbusqueda.etestatusllistat = "Cancelant la consulta..."
 '  r = "Cancelant la consulta..."
 '  DoEvents
 '  Unload subbusqueda
 '  wait 1
 'End If
 dbclixes.Close
 
  If i = 0 Then
    r = "sortir"
  End If
End Sub

Private Sub general_Click()
llistat_general
End Sub

Private Sub productes_Click(Index As Integer)
   If reixa.Tag <> "" Then netejar_reixa_busqueda
 ' For f = 0 To 17
 '  If productes.Item(f).BackColor = QBColor(12) Then
 '     If productes.Item(f).Container <> productes.Item(Index).Container Then
 '         For j = 0 To 17
            'If productes.Item(j).Enabled Then productes.Item(j).BackColor = QBColor(15)
 '         Next j
 '     End If
 '  End If
 ' Next f
  
  If productes.Item(Index).BackColor <> QBColor(15) Then
      productes.Item(Index).BackColor = QBColor(15)
    Else: productes.Item(Index).BackColor = QBColor(12)
  End If
  If controlcanviat.Name = "productes" Then colorcanviat = productes.Item(Index).BackColor
End Sub
Function plantillabusqueda() As String
 Dim valp As String
 plantillabusqueda = ""
 For f = 0 To productes.Count - 1
   If productes.Item(f).BackColor = QBColor(12) Then
      If plantillabusqueda = "" Then
          plantillabusqueda = productes.Item(f).Container
        Else: valp = "totsproductes"
      End If
   End If
   If Text2 <> "" Then If productes.Item(f).Caption = Trim(Text2) Then valp = productes.Item(f).Container
  Next f

 If totsproductes.Value = 1 Then valp = "totsproductes"
 If valp <> "" Then plantillabusqueda = valp
End Function

Private Sub quantitatentregada_Click()
llistat_quantitatentregada
End Sub

Private Sub reixa_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
guardar_amples_reixa_busqueda
End Sub

Private Sub reixa_DblClick()
  acceptarregistre
End Sub
Function acceptarregistre()
  On Error Resume Next
    r = " comanda=" + atrim(cadbl(Data1.Recordset!lot))
   subbusqueda.Visible = False
   'Unload subbusqueda
End Function

Private Sub reixa_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then acceptarregistre
End Sub

Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim color As Long
 Dim rstsisimp As Recordset
 Dim rstcom As Recordset
  Dim colorclassic As Long
  Dim colorkodak As Long
  Dim coloroffset As Long
  Dim nordre As Integer
  Dim tipus As String
  colorclassic = &H8000000F
  colorkodak = &HC0FFFF
  coloroffset = &HFDD7FD
  DoEvents
  If reixa.Columns.Count < 3 Then Exit Sub
   r = " comanda=" + atrim(cadbl(reixa.Columns("Lot")))
   
    Set rstcom = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where " + r)
    If rstcom.EOF Then Exit Sub
    
    nordre = cadbl(rstcom!numordremodificacio)
    If nordre = 0 Then nordre = 1
   
   Set rstsisimp = dbclixes.OpenRecordset("select sistemadimpresio,bandes from modificacions where id_treball=" + atrim(cadbl(rstcom!numtreball)) + " and ordre=" + atrim(nordre))
   If Not rstsisimp.EOF Then
   
  ' Set rstsisimp = dbclixes.OpenRecordset("select sistemadimpresio,bandes from clixes where id_treball=" + atrim(cadbl(rstcom!numtreball)))
   'If Not rstsisimp.EOF Then
     tipus = atrim(rstsisimp!sistemadimpresio)
     
   End If
   If tipus = "Flexo Std" Then color = colorclassic
     If tipus = "Flexo Kodak" Then color = colorkodak
     If tipus = "Offset" Then color = coloroffset
     If color = 0 Then color = colorclassic
     reixa.BackColor = color
End Sub

Private Sub Text1_Change()
 If reixa.Tag <> "" Then netejar_reixa_busqueda
End Sub
Sub netejar_reixa_busqueda()
  Data1.RecordSource = "select comanda as [_] from comandes where comanda=99999999"
  Data1.Refresh
  reixa.Refresh
  reixa.ClearFields
  vtreballbuscatsubbusqueda = ""
  cnumtreball.Tag = ""
  cnumtreball = ""
  reixa.Tag = ""
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then
    triarclient
    netejar_reixa_busqueda
  End If
   
End Sub

Private Sub Text1_LostFocus()
  Dim rstc As Recordset
  Dim dblocal As Database
  Dim temps As Date
 ' temps = Now
  If reixa.Tag <> "" Then Exit Sub
  formcomandes.lookupde "clients", Text1, nomclient, "nom", "clientsextres"
  If Text2 <> "" Or nomclient = "" Then Exit Sub
  'Set dblocal = OpenDatabase(cami)
  If totsproductes.Value = 1 Then Exit Sub
  Set dblocal = formcomandes.Data1.Database 'dbtmp
  nomclient = "Buscant productes d'aquest client..."
  DoEvents
  pend = IIf(pendent.Value = 1, " where proximaseccio<>'T' ", "")
  'Set rstc = dblocal.OpenRecordset("select distinct producte from comandes where client=" + atrim(cadbl(Text1.Text)) + pend)
  Workspaces(0).BeginTrans
  If metode.Value = 1 Then
       dbtmp.Execute "drop table tmp_consultaproducte "
       dbtmp.Execute "SELECT comandes.producte, comandes.client, comandes.proximaseccio INTO tmp_consultaproducte From comandes WHERE (((comandes.client)=6841));"
       Set rstc = dbtmp.OpenRecordset("SELECT distinct producte From tmp_consultaproducte where " + IIf(pend <> "", pend + " and ", "") + " client=" + atrim(cadbl(Text1.Text)), dbOpenSnapshot, dbReadOnly)
      Else: Set rstc = dblocal.OpenRecordset("SELECT comandes1.producte From comandes1 " + pend + " GROUP BY comandes1.producte,comandes1.client HAVING (((comandes1.client)=" + atrim(cadbl(Text1.Text)) + "))", dbOpenSnapshot, dbReadOnly)
  End If

  For j = 0 To productes.Count - 1
    productes.Item(j).Enabled = False
    productes.Item(j).BackColor = QBColor(0)
  Next j
  While Not rstc.EOF
    For j = 0 To productes.Count - 1
      If productes.Item(j).Caption = atrim(rstc!producte) Then
        productes.Item(j).Enabled = True
        productes.Item(j).BackColor = QBColor(15)
      End If
    Next j
    rstc.MoveNext
  Wend
  Workspaces(0).CommitTrans dbForceOSFlush
  formcomandes.lookupde "clients", Text1, nomclient, "nom", "clientsextres"
  reixa.Tag = "B"
  
  Set dblocal = Nothing
  'MsgBox atrim(DateDiff("s", temps, Now))
End Sub

Private Sub Text2_Change()
 If reixa.Tag <> "" Then netejar_reixa_busqueda
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarproducte
End Sub

Private Sub Text2_LostFocus()
'LOOKUP DE producte
  Set rsttmp = dbtmp.OpenRecordset("select descripcio,ruta from productes where codi='" + atrim((Text2.Text)) + "'")
  If Not rsttmp.EOF Then
     nomproducte.Caption = atrim(rsttmp!descripcio)
    Else: nomproducte.Caption = ""
  End If
  If Text2 <> "" Then
  For j = 0 To productes.Count - 1
    productes.Item(j).Enabled = True
    productes.Item(j).BackColor = QBColor(15)
  Next j
  End If
End Sub

