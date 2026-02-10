VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fColesLaminadora 
   Caption         =   "Manteniment de coles de laminadora"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8460
   HelpContextID   =   100
   Icon            =   "ColesLaminadora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   8310
      Begin VB.Timer Timer1 
         Interval        =   900
         Left            =   6255
         Top             =   210
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   7845
         Picture         =   "ColesLaminadora.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   390
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   520
         Picture         =   "ColesLaminadora.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Consulta Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   965
         Picture         =   "ColesLaminadora.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "ColesLaminadora.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton consultar 
         Height          =   360
         Left            =   7395
         Picture         =   "ColesLaminadora.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda de Registres"
         Top             =   180
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Data datacoles 
         Caption         =   "Coles"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   360
         Left            =   3570
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from adhesius "
         Top             =   180
         Width           =   2430
      End
      Begin VB.CommandButton Command3 
         Height          =   360
         Left            =   1410
         Picture         =   "ColesLaminadora.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton Command9 
         Height          =   360
         Index           =   1
         Left            =   6885
         Picture         =   "ColesLaminadora.frx":26C6
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir llistat de materials >500"
         Top             =   195
         Visible         =   0   'False
         Width           =   420
      End
      Begin Crystal.CrystalReport llistat 
         Left            =   2625
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label estattaula 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1995
         TabIndex        =   24
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.Frame framedades 
      Enabled         =   0   'False
      Height          =   3810
      Left            =   45
      TabIndex        =   25
      Top             =   630
      Width           =   8370
      Begin VB.CheckBox Check2 
         Caption         =   "No visible a Lam."
         DataField       =   "novisiblealaminadora"
         DataSource      =   "datacoles"
         Height          =   315
         Left            =   6435
         TabIndex        =   69
         Top             =   390
         Width           =   1650
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Aportació de cola diversos materials:"
         Height          =   1500
         Left            =   5370
         TabIndex        =   60
         Top             =   1575
         Width           =   2925
         Begin VB.TextBox camps 
            DataField       =   "aportcola_tricapa_impres"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   19
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   67
            Top             =   1155
            Width           =   1500
         End
         Begin VB.TextBox camps 
            DataField       =   "aportcola_impres"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   18
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   65
            Top             =   855
            Width           =   1500
         End
         Begin VB.TextBox camps 
            DataField       =   "aportcola_anonim"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   17
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   63
            Top             =   570
            Width           =   1500
         End
         Begin VB.TextBox camps 
            DataField       =   "aportcola_EVOH"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   15
            Left            =   1260
            MaxLength       =   15
            TabIndex        =   61
            Top             =   270
            Width           =   1500
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   "Tricapa Imprès:"
            Height          =   240
            Index           =   24
            Left            =   60
            TabIndex        =   68
            Top             =   1200
            Width           =   1170
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   "Imprès:"
            Height          =   240
            Index           =   23
            Left            =   120
            TabIndex        =   66
            Top             =   900
            Width           =   840
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   "Anónim:"
            Height          =   240
            Index           =   22
            Left            =   120
            TabIndex        =   64
            Top             =   615
            Width           =   840
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   "Tot EVOH:"
            Height          =   240
            Index           =   17
            Left            =   120
            TabIndex        =   62
            Top             =   315
            Width           =   840
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Temperatura "
         Height          =   765
         Left            =   120
         TabIndex        =   53
         Top             =   2430
         Width           =   5220
         Begin VB.TextBox camps 
            DataField       =   "tempprensa"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   20
            Left            =   3660
            TabIndex        =   70
            Top             =   390
            Width           =   990
         End
         Begin VB.TextBox camps 
            DataField       =   "temptubo"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   14
            Left            =   2475
            TabIndex        =   58
            Top             =   390
            Width           =   945
         End
         Begin VB.TextBox camps 
            DataField       =   "tempenduridor"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   12
            Left            =   1305
            TabIndex        =   56
            Top             =   390
            Width           =   885
         End
         Begin VB.TextBox camps 
            DataField       =   "tempresina"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   13
            Left            =   165
            TabIndex        =   54
            Top             =   390
            Width           =   885
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   "Prensa:"
            Height          =   240
            Index           =   25
            Left            =   3870
            TabIndex        =   71
            Top             =   195
            Width           =   705
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   "Tubo:"
            Height          =   240
            Index           =   16
            Left            =   2685
            TabIndex        =   59
            Top             =   195
            Width           =   705
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   "Enduridor:"
            Height          =   240
            Index           =   14
            Left            =   1395
            TabIndex        =   57
            Top             =   195
            Width           =   840
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   " Resina:"
            Height          =   240
            Index           =   15
            Left            =   270
            TabIndex        =   55
            Top             =   180
            Width           =   690
         End
      End
      Begin VB.TextBox codisubfamilia 
         DataField       =   "idsubfamilia"
         DataSource      =   "datacoles"
         Height          =   285
         Left            =   4665
         TabIndex        =   50
         Top             =   3360
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox codifamilia 
         DataField       =   "idfamilia"
         DataSource      =   "datacoles"
         Height          =   285
         Left            =   630
         TabIndex        =   49
         Top             =   3345
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox codienduridor 
         DataField       =   "codienduridor"
         DataSource      =   "datacoles"
         Height          =   285
         Left            =   600
         TabIndex        =   48
         Top             =   1095
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox codiresina 
         DataField       =   "codiresina"
         DataSource      =   "datacoles"
         Height          =   285
         Left            =   600
         TabIndex        =   47
         Top             =   795
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.TextBox cpredeterminat 
         DataField       =   "predeterminada"
         DataSource      =   "datacoles"
         Height          =   285
         Left            =   5775
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   225
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Predeterminat"
         Height          =   315
         Left            =   6435
         TabIndex        =   45
         Top             =   135
         Width           =   1815
      End
      Begin VB.ComboBox opcions 
         DataField       =   "color"
         DataSource      =   "datacoles"
         Height          =   315
         ItemData        =   "ColesLaminadora.frx":2C50
         Left            =   3285
         List            =   "ColesLaminadora.frx":2C66
         TabIndex        =   1
         Top             =   240
         Width           =   1605
      End
      Begin VB.ComboBox combosubfamilia 
         Height          =   315
         Left            =   4875
         TabIndex        =   16
         Top             =   3315
         Width           =   2460
      End
      Begin VB.ComboBox combofamilia 
         Height          =   315
         Left            =   825
         TabIndex        =   15
         Top             =   3330
         Width           =   2460
      End
      Begin VB.ComboBox comboenduridor 
         DataField       =   "enduridor"
         DataSource      =   "datacoles"
         Height          =   315
         Left            =   810
         TabIndex        =   7
         Top             =   1095
         Width           =   2460
      End
      Begin VB.ComboBox comboresina 
         DataField       =   "resina"
         DataSource      =   "datacoles"
         Height          =   315
         Left            =   810
         TabIndex        =   2
         Top             =   765
         Width           =   2460
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Temperatura aigua"
         Height          =   750
         Left            =   105
         TabIndex        =   36
         Top             =   1590
         Width           =   4380
         Begin VB.TextBox camps 
            DataField       =   "tempaigua4"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   16
            Left            =   3270
            TabIndex        =   51
            Top             =   345
            Width           =   915
         End
         Begin VB.TextBox camps 
            DataField       =   "tempPreHeating"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   11
            Left            =   2265
            TabIndex        =   14
            Top             =   360
            Width           =   915
         End
         Begin VB.TextBox camps 
            DataField       =   "tempaigua2"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   10
            Left            =   1170
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox camps 
            DataField       =   "tempaigua1"
            DataSource      =   "datacoles"
            Height          =   285
            Index           =   9
            Left            =   150
            TabIndex        =   12
            Top             =   360
            Width           =   915
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   "Temp: 4"
            Height          =   240
            Index           =   21
            Left            =   3390
            TabIndex        =   52
            Top             =   135
            Width           =   645
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   "Pre-Heating"
            Height          =   240
            Index           =   11
            Left            =   2265
            TabIndex        =   39
            Top             =   150
            Width           =   885
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   "Temp: 2"
            Height          =   240
            Index           =   10
            Left            =   1380
            TabIndex        =   38
            Top             =   165
            Width           =   645
         End
         Begin VB.Label etcamps 
            BackStyle       =   0  'Transparent
            Caption         =   "Temp: 1"
            Height          =   240
            Index           =   9
            Left            =   345
            TabIndex        =   37
            Top             =   165
            Width           =   645
         End
      End
      Begin VB.TextBox camps 
         DataField       =   "euroskgenduridor"
         DataSource      =   "datacoles"
         Height          =   285
         Index           =   8
         Left            =   7740
         TabIndex        =   11
         Top             =   1050
         Width           =   480
      End
      Begin VB.TextBox camps 
         DataField       =   "euroskgresina"
         DataSource      =   "datacoles"
         Height          =   285
         Index           =   7
         Left            =   7740
         TabIndex        =   6
         Top             =   690
         Width           =   480
      End
      Begin VB.TextBox camps 
         DataField       =   "grmcm3_enduridor"
         DataSource      =   "datacoles"
         Height          =   285
         Index           =   6
         Left            =   6615
         TabIndex        =   10
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox camps 
         DataField       =   "grmcm3_resina"
         DataSource      =   "datacoles"
         Height          =   285
         Index           =   5
         Left            =   6615
         TabIndex        =   5
         Top             =   720
         Width           =   555
      End
      Begin VB.TextBox camps 
         DataField       =   "grausenduridor"
         DataSource      =   "datacoles"
         Height          =   285
         Index           =   4
         Left            =   5295
         TabIndex        =   9
         Top             =   1110
         Width           =   555
      End
      Begin VB.TextBox camps 
         DataField       =   "grausresina"
         DataSource      =   "datacoles"
         Height          =   285
         Index           =   3
         Left            =   5295
         TabIndex        =   4
         Top             =   750
         Width           =   555
      End
      Begin VB.TextBox camps 
         DataField       =   "%enduridor"
         DataSource      =   "datacoles"
         Height          =   285
         Index           =   2
         Left            =   4110
         TabIndex        =   8
         Top             =   1125
         Width           =   555
      End
      Begin VB.TextBox camps 
         DataField       =   "%resina"
         DataSource      =   "datacoles"
         Height          =   285
         Index           =   1
         Left            =   4110
         TabIndex        =   3
         Top             =   765
         Width           =   555
      End
      Begin VB.TextBox camps 
         BackColor       =   &H00E0E0E0&
         DataField       =   "codi"
         DataSource      =   "datacoles"
         Height          =   285
         Index           =   0
         Left            =   885
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   255
         Width           =   630
      End
      Begin VB.Label etcamps 
         Caption         =   "Color del missatge:"
         Height          =   240
         Index           =   20
         Left            =   1875
         TabIndex        =   44
         Top             =   270
         Width           =   1440
      End
      Begin VB.Label etcamps 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Familia:"
         Height          =   240
         Index           =   19
         Left            =   3915
         TabIndex        =   43
         Top             =   3375
         Width           =   990
      End
      Begin VB.Label etcamps 
         BackStyle       =   0  'Transparent
         Caption         =   "Familia "
         Height          =   240
         Index           =   18
         Left            =   210
         TabIndex        =   42
         Top             =   3375
         Width           =   1320
      End
      Begin VB.Label etcamps 
         BackStyle       =   0  'Transparent
         Caption         =   "Enduridor:"
         Height          =   240
         Index           =   13
         Left            =   45
         TabIndex        =   41
         Top             =   1140
         Width           =   990
      End
      Begin VB.Label etcamps 
         BackStyle       =   0  'Transparent
         Caption         =   "Resina:"
         Height          =   240
         Index           =   12
         Left            =   60
         TabIndex        =   40
         Top             =   795
         Width           =   675
      End
      Begin VB.Label etcamps 
         Caption         =   "€/Kg:"
         Height          =   240
         Index           =   8
         Left            =   7230
         TabIndex        =   35
         Top             =   1110
         Width           =   510
      End
      Begin VB.Label etcamps 
         Caption         =   "€/Kg:"
         Height          =   240
         Index           =   7
         Left            =   7230
         TabIndex        =   34
         Top             =   750
         Width           =   630
      End
      Begin VB.Label etcamps 
         Caption         =   "Grm/cm3:"
         Height          =   240
         Index           =   6
         Left            =   5865
         TabIndex        =   33
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label etcamps 
         Caption         =   "Grm/cm3:"
         Height          =   240
         Index           =   5
         Left            =   5865
         TabIndex        =   32
         Top             =   780
         Width           =   915
      End
      Begin VB.Label etcamps 
         Caption         =   "Graus:"
         Height          =   240
         Index           =   4
         Left            =   4755
         TabIndex        =   31
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label etcamps 
         Caption         =   "Graus:"
         Height          =   240
         Index           =   3
         Left            =   4755
         TabIndex        =   30
         Top             =   810
         Width           =   915
      End
      Begin VB.Label etcamps 
         Caption         =   "%Endur.:"
         Height          =   240
         Index           =   2
         Left            =   3330
         TabIndex        =   29
         Top             =   1185
         Width           =   915
      End
      Begin VB.Label etcamps 
         Caption         =   "% Resina:"
         Height          =   240
         Index           =   1
         Left            =   3330
         TabIndex        =   28
         Top             =   825
         Width           =   915
      End
      Begin VB.Label etcamps 
         Caption         =   "Codi cola:"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   300
         Width           =   885
      End
   End
End
Attribute VB_Name = "fColesLaminadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub colsbloc_Change()

End Sub

Private Sub alta_Click()
Dim gran As Long
Dim rst As Recordset
framedades.Enabled = True
Set rst = dbtmp.OpenRecordset("select max(codi) as elgran from adhesius ")
If rst.EOF Then
   gran = 0
     Else: gran = cadbl(rst!elgran)
 End If
'materials.RecordSource = "select * from materials order by codi"
'materials.Recordset.MoveLast
'If Not materials.Recordset.EOF Then gran = materials.Recordset!codi
gran = gran + 1
datacoles.Recordset.AddNew
framedades.Enabled = True
datacoles.Recordset!codi = gran
opcions.SetFocus
camps(0) = gran
'Text1(2).SetFocus
Set rst = Nothing
End Sub

Private Sub Command2_Click()
  If Not existeixrang Then
      espesors.Recordset.AddNew
       espesors.Recordset!codi = materials.Recordset!codi
       espesors.Recordset!micres = cadbl(micres)
       espesors.Recordset!grmsm2 = cadbl(grmm2)
      espesors.Recordset.Update
     Else: MsgBox "Aquest espesor ja existeix."
  End If
End Sub
Function existeixrang() As Boolean
  existeixrang = False
  If cadbl(micres) > 0 And espesors.Recordset.RecordCount > 0 Then
   espesors.Recordset.FindFirst "micres=" + atrim(cadbl(micres))
   If Not espesors.Recordset.NoMatch Then
      existeixrang = True
   End If
  End If
  If (cadbl(micres) = 0 And cadbl(grmm2) > 0) And espesors.Recordset.RecordCount > 0 Then
   espesors.Recordset.FindFirst "grmsm2=" + atrim(cadbl(grmm2))
   If Not espesors.Recordset.NoMatch Then
      existeixrang = True
   End If
  End If
  
End Function

Private Sub Command1_Click()

End Sub

Sub carregar_comboresinaenduridor(vcombo As Control, vcolaoenduridor As String)
  Dim vsql As String
  Dim rst As Recordset
  Dim dbtintes As Database
  Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb", , True)
  vsql = "SELECT tintes_tot.codi,tintes_tot.descripcio, tintes_tot.descripciofam From tintes_tot "
  vsql = vsql + " WHERE (((tintes_tot.descripciofam)='" + vcolaoenduridor + "'));"
  vcombo.Clear
  Set rst = dbtintes.OpenRecordset(vsql)
  While Not rst.EOF
     vcombo.AddItem UCase(atrim(rst!descripcio))
     vcombo.ItemData(vcombo.NewIndex) = cadbl(rst!codi)
     rst.MoveNext
  Wend
  Set dbtintes = Nothing
End Sub

Private Sub Check1_Click()
  If Screen.ActiveControl.Name = "Check1" Then
    datacoles.Database.Execute "update adhesius set predeterminada=' '"
    If Check1.Value = 1 Then datacoles.Database.Execute "update adhesius set predeterminada='S' where codi=" + atrim(datacoles.Recordset!codi)
    gravar_registre
  End If
End Sub

Private Sub Check2_Click()
' If Screen.ActiveControl.Name = "Check2" Then
'    If Check1.Value = 1 Then
'       datacoles.Database.Execute "update adhesius set novisiblealaminadora=true where codi=" + atrim(datacoles.Recordset!codi)
'         Else: datacoles.Database.Execute "update adhesius set novisiblealaminadora=false where codi=" + atrim(datacoles.Recordset!codi)
'    End If
'    gravar_registre
'  End If
End Sub

Private Sub comboenduridor_Click()
  codienduridor = comboenduridor.ItemData(comboenduridor.ListIndex)
End Sub

Private Sub combofamilia_Click()

   If combofamilia.ListIndex <> -1 Then codifamilia = combofamilia.ItemData(combofamilia.ListIndex)
End Sub

Private Sub comboresina_Click()
  codiresina = comboresina.ItemData(comboresina.ListIndex)
End Sub

Private Sub combosubfamilia_Click()
   If combosubfamilia.ListIndex <> -1 Then codisubfamilia = combosubfamilia.ItemData(combosubfamilia.ListIndex)
End Sub

Private Sub Command3_Click()
  gravar_registre
End Sub

Private Sub Command9_Click(Index As Integer)
  If MsgBox("Vols imprimir agrupant per proveidor?", vbInformation + vbYesNo, "Atenció") = vbYes Then
     llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatdematerials.rpt"
    Else
      llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatdematerialssensegrups.rpt"
  End If
  llistat.DataFiles(0) = cami
  llistat.Destination = crptToWindow
  llistat.Action = 1
End Sub

Private Sub consultar_Click()
   Dim b As String
   b = InputBox("Entra la Descripcio/RefProducte a buscar o el Codi" + Chr(10) + " No escriguis res per treure els filtres", "Busqueda")
   b = treure_apostruf(b)
   If cadbl(b) > 0 Then
     materials.RecordSource = "select * from materials where codi>499 and codi=" + atrim(cadbl(b)) + ""
     materials.Refresh
     b = ""
      Else
       If b <> "" Then
        materials.RecordSource = "select * from materials where codi>499 and descripcio like '*" + b + "*' or refproducte like '*" + b + "*'"
        materials.Refresh
          Else
             materials.RecordSource = "select * from materials where codi>499 "
             materials.Refresh
       End If
   End If
End Sub

Private Sub datacoles_Reposition()
  posarcolorcomboopcions
  Check1.Value = IIf(cpredeterminat = "S", 1, 0)
  posarfamiliescoles
End Sub
Sub posarfamiliescoles()
   Dim i As Byte
   combofamilia.ListIndex = -1
   For i = 0 To combofamilia.ListCount - 1
     If combofamilia.ItemData(i) = cadbl(codifamilia) Then combofamilia.ListIndex = i
   Next i
   
   combosubfamilia.ListIndex = -1
   For i = 0 To combosubfamilia.ListCount - 1
     If combosubfamilia.ItemData(i) = cadbl(codisubfamilia) Then combosubfamilia.ListIndex = i
   Next i
End Sub
Sub posarcolorcomboopcions()
  Dim codicolor As Double
  Select Case opcions
    Case "VERD"
       codicolor = QBColor(10)
    Case "TARONJA"
       codicolor = &H62B1F2
    Case "BLAU"
       codicolor = QBColor(9)
    Case "ROSA"
       codicolor = &HC78DFA
    Case "GROC"
       codicolor = QBColor(6)
    Case "VERMELL"
       codicolor = QBColor(12)
    Case "BLANC"
       codicolor = QBColor(15)
    Case Else
       codicolor = &H8000000F
  End Select
  opcions.BackColor = codicolor
  Set controlcanviat = Nothing
End Sub
Private Sub eliminar_Click()
  If MsgBox("Segur que vols borrar aquesta cola?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
     If InputBox("Escriu la paraula [ELIMINAR] per fer efectiu l'eliminació", "Control de seguretat") = "ELIMINAR" Then
         dbtmp.Execute ("delete * from adhesius where codi=" + atrim(cadbl(datacoles.Recordset!codi)))
         datacoles.Recordset.Delete
         datacoles.Refresh
     End If
  End If
  
End Sub

Private Sub Form_Activate()
posarcolorcomboopcions
posarfamiliescoles
End Sub

Private Sub Form_Click()
   'While Not materials.Recordset.BOF
   '  modificar_Click
   '  DoEvents
   '  gravar_registre
   '  materials.Recordset.MovePrevious
   'Wend
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then gravar_registre
End Sub
Sub gravar_registre()
   If datacoles.Recordset.EditMode > 0 Then datacoles.Recordset.Update
   framedades.Enabled = False
   
'  Dim i As Byte
'  Dim nogravar As Boolean
'  nogravar = False
'  If materials.Recordset.EditMode > 0 Then
'    If atrim(mesespcompra) = "" Then MsgBox "No hi ha mesura del producte entrada", vbCritical, "Error": Exit Sub
'    If mesespcompra = "Grm/m2" And cadbl(Text1(12)) = 0 Then MsgBox "Has d'entrar l'equivalencia en micres al camp d'Espessor quan es Grm/m2", vbInformation, "Atenció": Exit Sub
'    If mesespcompra <> "Unitats" Then
'        For i = 4 To 9
'           If cadbl(Text1(i)) = 0 Then nogravar = True
'        Next i
'        If nogravar Then MsgBox "Falta possar alguna familia o subfamilia abans de guardar les dades": Exit Sub
'        If cadbl(Text1(4)) < 500 Or cadbl(Text1(5).Text) < 500 Or cadbl(Text1(6).Text) < 500 Then MsgBox "El codi de les families ha de ser mes gran de 500": Exit Sub
'    End If
'    If Not comprovar_families(cadbl(Text1(4)), cadbl(Text1(7)), cadbl(Text1(5)), cadbl(Text1(8)), cadbl(Text1(6)), cadbl(Text1(9))) Then
 '       MsgBox "Falta possar alguna familia o subfamilia abans de guardar les dades": Exit Sub
  '  End If
  '  generar_id_families cadbl(Text1(4)), cadbl(Text1(7)), cadbl(Text1(5)), cadbl(Text1(8)), cadbl(Text1(6)), cadbl(Text1(9))
'    materials.Recordset.Update
''    framematerials.Enabled = False
''    materials.Recordset.Bookmark = materials.Recordset.LastModified
 ' End If
End Sub
Function comprovar_families(fam As Double, subfam As Double, famcol As Double, subfamcol As Double, famad As Double, subfamad As Double) As Boolean
   comprovar_families = True
   If fam = 0 Or subfam = 0 Or famcol = 0 Or subfamcol = 0 Or famad = 0 Or subfamad = 0 Then comprovar_families = False
End Function
Sub generar_id_families(fam As Double, subfam As Double, famcol As Double, subfamcol As Double, famad As Double, subfamad As Double)
   Dim rstf As Recordset
   Set rstf = dbtmp.OpenRecordset("select id_familia from materials where id_familia<>null and familia=" + atrim(fam) + " and subfamilia=" + atrim(subfam) + " and familiacol=" + atrim(famcol) + " and subfamiliacol=" + atrim(subfamcol) + " and familiaad=" + atrim(famad) + " and subfamiliaad=" + atrim(subfamad))
   'MsgBox "select id_familia from materials where familia=" + atrim(fam) + " and subfamilia=" + atrim(subfam) + " and familiacol=" + atrim(famcol) + " and subfamiliacol=" + atrim(subfamcol) + " and familiaad=" + atrim(famad) + " and subfamiliaad=" + atrim(subfamad)
   If Not rstf.EOF Then
    If cadbl(rstf!id_familia) > 0 Then
       materials.Recordset!id_familia = rstf!id_familia
       Exit Sub
    End If
   End If
   Set rstf = dbtmp.OpenRecordset("select max(id_familia) as gran from materials")
   materials.Recordset!id_familia = 1
   If Not rstf.EOF Then materials.Recordset!id_familia = cadbl(rstf!gran) + 1
End Sub
Sub cancelar_registre()

 If datacoles.EditMode > 0 Then datacoles.Recordset.CancelUpdate
 framedades.Enabled = False
End Sub

Private Sub Form_Load()
  datacoles.DatabaseName = cami
'  espesors.DatabaseName = cami
  carregar_comboresinaenduridor comboresina, "resina"
  carregar_comboresinaenduridor comboenduridor, "enduridor"
  carregar_combofamilies combofamilia, "familiescoles"
  carregar_combofamilies combosubfamilia, "subfamiliescoles"
End Sub
Sub carregar_combofamilies(vcombo As Control, vfam As String)
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from " + vfam + " order by descripcio")
  vcombo.Clear
  While Not rst.EOF
    vcombo.AddItem UCase(atrim(rst!descripcio))
    vcombo.ItemData(vcombo.NewIndex) = cadbl(rst!codi)
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Private Sub Form_Resize()
  'reixa.Width = fmaterials.Width - reixa.Left - 200
  
End Sub

Private Sub materials_Reposition()
  framematerials.Enabled = False
  carregar_camps
  If Not materials.Recordset.EOF Then materials.Caption = "Mat. " + atrim(1 + cadbl(materials.Recordset.AbsolutePosition)) + "/" + atrim(cadbl(materials.Recordset.RecordCount))
End Sub
Sub carregar_camps()
 ' Dim rstp As Recordset
 ' nomproveidor = ""
 ' Label3(0) = ""
 ' Label3(1) = ""
 ' Label3(2) = ""
 ' If materials.Recordset.EOF Then
 '   espesors.RecordSource = ""
 '   espesors.Refresh
 '   Exit Sub
 ' End If
 ' If cadbl(materials.Recordset!codi) > 0 Then
 '  espesors.RecordSource = "select * from materials_espesors where codi=" + atrim(cadbl(materials.Recordset!codi))
 '    Else: espesors.RecordSource = ""
 ' End If
 ' espesors.Refresh
 ' Set rstp = materials.Database.OpenRecordset("select * from proveidors where codi=" + atrim(cadbl(materials.Recordset!proveidor)))
 ' If Not rstp.EOF And cadbl(materials.Recordset!proveidor) > 0 Then
 '    nomproveidor = atrim(rstp!nom)
 '      Else: nomproveidor = ""
 ' End If
 ' descfamilies
  
End Sub
Sub descfamilies()
  'Dim rstp As Recordset
  'Dim rstp2 As Recordset
  'Dim l As String
  ''families materials
  'Set rstp = materials.Database.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(Text1(4))))
  'Set rstp2 = materials.Database.OpenRecordset("select descripcio from subfamiliesmaterials where codi=" + atrim(cadbl(Text1(7))))
  'If Not rstp.EOF Then
  '   l = atrim(rstp!descripcio)
  '   If Not rstp2.EOF Then l = l + " - " + atrim(rstp2!descripcio)
  '   Label3(0) = l
  'End If
  
  ''families colorants
  'Set rstp = materials.Database.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(Text1(5))))
  'Set rstp2 = materials.Database.OpenRecordset("select descripcio from subfamiliescolorants where codi=" + atrim(cadbl(Text1(8))))
  'If Not rstp.EOF Then
  '   l = atrim(rstp!descripcio)
  '   If Not rstp2.EOF Then l = l + " - " + atrim(rstp2!descripcio)
  '   Label3(1) = l
  'End If
 '
 ' 'families aditius
 ' Set rstp = materials.Database.OpenRecordset("select descripcio from familiesaditius where codi=" + atrim(cadbl(Text1(6))))
  'Set rstp2 = materials.Database.OpenRecordset("select descripcio from subfamiliesaditius where codi=" + atrim(cadbl(Text1(9))))
  'If Not rstp.EOF Then
  '   l = atrim(rstp!descripcio)
'     If Not rstp2.EOF Then l = l + " - " + atrim(rstp2!descripcio)
'     Label3(2) = l
'  End If
  
 ' Set rstp = Nothing
 ' Set rstp2 = Nothing
End Sub

Private Sub mesespcompra_LostFocus()
   'If mesespcompra = "Grm/m2" And cadbl(Text1(12)) = 0 Then MsgBox "Pensa a entrar l'equivalencia en micres al camp d'Espessor", vbInformation, "Atenció"
End Sub

Private Sub modificar_Click()
   If Not datacoles.Recordset.EOF Then
     datacoles.Recordset.Edit
     framedades.Enabled = True
     comboresina.SetFocus
   End If
End Sub

Private Sub opcions_Click()
    posarcolorcomboopcions
End Sub

Private Sub opcions_GotFocus()
  Set controlcanviat = Nothing
End Sub

Private Sub sortir_Click()
 Unload fColesLaminadora
End Sub

Private Sub Timer1_Timer()
  estattaula = IIf(datacoles.Recordset.EditMode = 0, "", "Editant")
End Sub
Sub triar_proveidor()
 ' Load formseleccio
 ' formseleccio.Data1.DatabaseName = materials.DatabaseName
 ' formseleccio.Data1.RecordSource = "select * from proveidors"
 ' formseleccio.refrescar
'  formseleccio.Show 1
'  If seleccioret = 1 Then
'   Text1(1).Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
'   materials.Recordset!proveidor = Text1(1).Text
'   nomproveidor.Caption = atrim(formseleccio.Data1.Recordset!nom)
'  End If
'  Unload formseleccio
  
End Sub



Private Sub Text1_Change(Index As Integer)
   'If Index = 12 And materials.Recordset.EditMode > 0 Then If cadbl(Text1(12)) = 0 Then Text1(12) = "0"
End Sub

Private Sub Text1_GotFocus(Index As Integer)
  'If materials.Recordset.Fields(Text1(Index).DataField).Type = 10 Then
'     Text1(Index).MaxLength = materials.Recordset.Fields(Text1(Index).DataField).Size
'  End If
'  If Index = 0 Then
'    If espesors.RecordSource = "" Then Exit Sub
'    If Not espesors.Recordset.EOF Then
'       Text1(1).SetFocus
'       Text1(0).Locked = True
'       MsgBox "No pots editar aquest camp si hi ha micres asignades"
'      Else: Text1(0).Locked = False
'    End If
'  End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 ' If KeyCode = 113 Then
 '     Select Case Index
 '       Case 1
 '          triar_proveidor
 '       Case 4
 '         Text1(Index) = triar_familia("select * from familiesmaterials where codi>499")
 '       Case 5
 '         Text1(Index) = triar_familia("select * from familiescolorants where codi>499")
 '       Case 6
'          Text1(Index) = triar_familia("select * from familiesaditius where codi>499")
'        Case 7
'           Text1(Index) = triar_familia("select * from subfamiliesmaterials where codifam=" + atrim(cadbl(Text1(4))))
'        Case 8
'           Text1(Index) = triar_familia("select * from subfamiliescolorants where codifam=" + atrim(cadbl(Text1(5))))
'        Case 9
'           Text1(Index) = triar_familia("select * from subfamiliesaditius where codifam=" + atrim(cadbl(Text1(6))))
'      End Select
 '     descfamilies
 ' End If
 '
End Sub
Function triar_familia(seleccio As String) As String
   Load formseleccio
   formseleccio.Caption = "Triar Familia o Subfamilia"
   formseleccio.Data1.DatabaseName = materials.DatabaseName
   formseleccio.Data1.RecordSource = seleccio
   formseleccio.refrescar
   formseleccio.Show 1
   If seleccioret = 1 Then
     triar_familia = atrim(cadbl(formseleccio.Data1.Recordset!codi))
      Else: triar_familia = "0"
   End If
  Unload formseleccio
End Function
Sub comprovarmesde500(Index As Integer)
 '  If cadbl(Text1(Index)) = 0 Then Exit Sub
 '  Select Case Index
 '       Case 4
 '         If cadbl(Text1(Index)) < 500 Then Text1(Index) = "0": MsgBox "La familia ha de ser superior a 500"
 '       Case 5
 '         If cadbl(Text1(Index)) < 500 Then Text1(Index) = "0": MsgBox "La familia ha de ser superior a 500"
 '       Case 6
 '         If cadbl(Text1(Index)) < 500 Then Text1(Index) = "0": MsgBox "La familia ha de ser superior a 500"
 '  End Select
   
   
End Sub
Private Sub Text1_LostFocus(Index As Integer)

  'comprovarmesde500 Index

  'If Index = 4 Or Index = 5 Or Index = 6 Or Index = 7 Or Index = 8 Or Index = 9 Then
'      descfamilies
'  End If
'  If Index = 11 Then
'      If cadbl(Text1(11)) > 0 And cadbl(Text1(10)) > 0 Then
''         If MsgBox("Hi ha un valor entrat a Grm/cm3 que no pot coexistir amb els Grm/m2." + Chr(10) + Chr(13) + "Vols eliminar els Grm/cm3?", vbYesNo, "Atenció") = vbYes Then
'           Text1(10) = 0
'             Else: Text1(11) = 0
'         End If
'      End If
'  End If
'  If Index = 10 Then
'      If cadbl(Text1(10)) > 0 And cadbl(Text1(11)) > 0 Then
'        If MsgBox("Hi ha un valor entrat a Grm/m2 que no pot coexistir amb els Grm/cm3." + Chr(10) + Chr(13) + "Vols eliminar els Grm/m2?", vbYesNo, "Atenció") = vbYes Then
'           Text1(11) = 0
'             Else: Text1(10) = 0
'         End If
'      End If
'  End If
End Sub
