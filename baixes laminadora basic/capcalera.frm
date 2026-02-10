VERSION 5.00
Begin VB.Form capcalera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capçalera Laminadora"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11805
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   60
      Picture         =   "capcalera.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3585
      Width           =   11670
   End
   Begin VB.Data capcalera 
      Caption         =   "capcalera"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4605
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   3525
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   11730
      Begin VB.CommandButton Command1 
         Height          =   330
         Left            =   -30
         Picture         =   "capcalera.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Intercambiar Desbobinador"
         Top             =   1410
         Width           =   270
      End
      Begin VB.Frame Frame4 
         Height          =   2910
         Left            =   5490
         TabIndex        =   21
         Top             =   210
         Width           =   6150
         Begin VB.TextBox Text16 
            DataField       =   "observacions"
            DataSource      =   "capcalera"
            Height          =   1515
            Left            =   135
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   1245
            Width           =   5880
         End
         Begin VB.TextBox Text11 
            DataField       =   "tensioreb"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   5265
            TabIndex        =   30
            Top             =   480
            Width           =   600
         End
         Begin VB.TextBox Text10 
            DataField       =   "impresvisual"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   4170
            TabIndex        =   28
            Top             =   495
            Width           =   600
         End
         Begin VB.TextBox Text9 
            DataField       =   "adhesiu"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   2790
            TabIndex        =   26
            Top             =   495
            Width           =   600
         End
         Begin VB.TextBox Text8 
            DataField       =   "cilincola"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   1500
            TabIndex        =   24
            Top             =   495
            Width           =   600
         End
         Begin VB.TextBox Text7 
            DataField       =   "camisa"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   300
            TabIndex        =   22
            Top             =   495
            Width           =   600
         End
         Begin VB.Label Label16 
            Caption         =   "Observacions"
            Height          =   255
            Left            =   135
            TabIndex        =   39
            Top             =   960
            Width           =   1650
         End
         Begin VB.Label Label13 
            Caption         =   "Tensió Reb."
            Height          =   315
            Left            =   5175
            TabIndex        =   31
            Top             =   225
            Width           =   1020
         End
         Begin VB.Label Label12 
            Caption         =   "Impresió Visual %"
            Height          =   315
            Left            =   3855
            TabIndex        =   29
            Top             =   225
            Width           =   1290
         End
         Begin VB.Label Label11 
            Caption         =   "Adhesiu (gr/m2)"
            Height          =   315
            Left            =   2520
            TabIndex        =   27
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label10 
            Caption         =   "Cilin. Cola (valor)"
            Height          =   315
            Left            =   1200
            TabIndex        =   25
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label Label9 
            Caption         =   "Camisa (cm)"
            Height          =   315
            Left            =   180
            TabIndex        =   23
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame desb2 
         Caption         =   "Desb 2"
         Height          =   1530
         Left            =   105
         TabIndex        =   3
         Top             =   1560
         Width           =   5250
         Begin VB.TextBox lotdesb2 
            DataField       =   "comdesb2"
            DataSource      =   "capcalera"
            Height          =   285
            Left            =   900
            TabIndex        =   41
            Text            =   "Text18"
            Top             =   240
            Visible         =   0   'False
            Width           =   165
         End
         Begin VB.TextBox Text15 
            DataField       =   "obsmat2"
            DataSource      =   "capcalera"
            Height          =   330
            Left            =   4080
            TabIndex        =   36
            ToolTipText     =   "Observació visual del material"
            Top             =   885
            Width           =   1095
         End
         Begin VB.TextBox Text13 
            DataField       =   "matdesb2"
            DataSource      =   "capcalera"
            Height          =   285
            Left            =   75
            TabIndex        =   33
            Top             =   915
            Width           =   3270
         End
         Begin VB.TextBox Text6 
            DataField       =   "tensio2"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   2925
            TabIndex        =   19
            Top             =   435
            Width           =   645
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Tractat a dalt"
            DataField       =   "tractatadalt"
            DataSource      =   "capcalera"
            Height          =   345
            Left            =   3810
            TabIndex        =   16
            Top             =   390
            Width           =   1335
         End
         Begin VB.TextBox Text4 
            DataField       =   "amplada2"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   2070
            TabIndex        =   13
            Top             =   465
            Width           =   600
         End
         Begin VB.TextBox Text2 
            DataField       =   "micres2"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   1425
            TabIndex        =   9
            Top             =   465
            Width           =   600
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Imprès"
            DataField       =   "matimpres2"
            DataSource      =   "capcalera"
            Height          =   225
            Left            =   3795
            TabIndex        =   6
            Top             =   165
            Width           =   825
         End
         Begin VB.Label ettoleranciadesb2 
            BackColor       =   &H00000000&
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
            ForeColor       =   &H00ED823A&
            Height          =   240
            Left            =   210
            TabIndex        =   47
            Top             =   1260
            Width           =   4950
         End
         Begin VB.Label Label15 
            Caption         =   "Obs Mat:"
            Height          =   180
            Left            =   3390
            TabIndex        =   37
            Top             =   945
            Width           =   765
         End
         Begin VB.Label Label8 
            Caption         =   "Tensió (Kg)"
            Height          =   315
            Left            =   2805
            TabIndex        =   20
            Top             =   180
            Width           =   930
         End
         Begin VB.Label Label6 
            Caption         =   "Mat. (cm)"
            Height          =   315
            Left            =   2040
            TabIndex        =   14
            Top             =   165
            Width           =   690
         End
         Begin VB.Label Label4 
            Caption         =   "Micres"
            Height          =   315
            Left            =   1470
            TabIndex        =   10
            Top             =   210
            Width           =   510
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   720
            TabIndex        =   4
            Top             =   210
            Width           =   360
         End
         Begin VB.Image Image2 
            Height          =   570
            Left            =   90
            Picture         =   "capcalera.frx":0B14
            Top             =   210
            Width           =   750
         End
      End
      Begin VB.Frame desb1 
         Caption         =   "Desb 1"
         Height          =   1425
         Left            =   90
         TabIndex        =   1
         Top             =   105
         Width           =   5250
         Begin VB.TextBox lotdesb1 
            DataField       =   "comdesb1"
            DataSource      =   "capcalera"
            Height          =   285
            Left            =   885
            TabIndex        =   40
            Text            =   "Text17"
            Top             =   150
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox Text14 
            DataField       =   "obsmat1"
            DataSource      =   "capcalera"
            Height          =   330
            Left            =   4110
            TabIndex        =   34
            ToolTipText     =   "Observació visual del material"
            Top             =   810
            Width           =   1095
         End
         Begin VB.TextBox Text12 
            DataField       =   "matdesb1"
            DataSource      =   "capcalera"
            Height          =   285
            Left            =   75
            TabIndex        =   32
            Top             =   870
            Width           =   3285
         End
         Begin VB.TextBox Text5 
            DataField       =   "tensio1"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   2985
            TabIndex        =   17
            Top             =   405
            Width           =   600
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Tractat a baix"
            DataField       =   "tractatabaix"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   3780
            TabIndex        =   15
            Top             =   375
            Width           =   1320
         End
         Begin VB.TextBox Text3 
            DataField       =   "amplada1"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   2085
            TabIndex        =   11
            Top             =   420
            Width           =   600
         End
         Begin VB.TextBox Text1 
            DataField       =   "micres1"
            DataSource      =   "capcalera"
            Height          =   360
            Left            =   1365
            TabIndex        =   7
            Top             =   420
            Width           =   600
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprès"
            DataField       =   "matimpres1"
            DataSource      =   "capcalera"
            Height          =   225
            Left            =   3780
            TabIndex        =   5
            Top             =   165
            Width           =   765
         End
         Begin VB.Label ettoleranciadesb1 
            BackColor       =   &H00000000&
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
            ForeColor       =   &H00ED823A&
            Height          =   240
            Left            =   150
            TabIndex        =   46
            Top             =   1170
            Width           =   4950
         End
         Begin VB.Label Label14 
            Caption         =   "Obs Mat:"
            Height          =   180
            Left            =   3420
            TabIndex        =   35
            Top             =   900
            Width           =   765
         End
         Begin VB.Label Label7 
            Caption         =   "Tensió (Kg)"
            Height          =   315
            Left            =   2850
            TabIndex        =   18
            Top             =   180
            Width           =   885
         End
         Begin VB.Label Label5 
            Caption         =   "Mat. (cm)"
            Height          =   315
            Left            =   2055
            TabIndex        =   12
            Top             =   165
            Width           =   690
         End
         Begin VB.Label Label3 
            Caption         =   "Micres"
            Height          =   315
            Left            =   1470
            TabIndex        =   8
            Top             =   150
            Width           =   510
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   690
            TabIndex        =   2
            Top             =   195
            Width           =   360
         End
         Begin VB.Image Image1 
            Height          =   570
            Left            =   75
            Picture         =   "capcalera.frx":21E6
            Top             =   210
            Width           =   750
         End
      End
      Begin VB.Label ettoleranciatotal 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H005C31DD&
         Height          =   300
         Left            =   135
         TabIndex        =   45
         Top             =   3120
         Width           =   5370
      End
      Begin VB.Label Label17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005C31DD&
         Height          =   300
         Left            =   0
         TabIndex        =   44
         Top             =   345
         Width           =   4770
      End
   End
End
Attribute VB_Name = "capcalera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function descripciomaterial(rstmat As Recordset) As String
  Dim desc As String
  Dim rstfam As Recordset
  Dim dbtmpb As Database
  Set dbtmpb = dbtmp
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from familiesmaterials where codi=" + atrim(cadbl(rstmat!familia)))
  If Not rstfam.EOF Then desc = desc + atrim(rstfam!descripcio)
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliesmaterials where codi=" + atrim(cadbl(rstmat!subfamilia)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from familiescolorants where codi=" + atrim(cadbl(rstmat!familiacol)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliescolorants where codi=" + atrim(cadbl(rstmat!subfamiliacol)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from familiesaditius where codi=" + atrim(cadbl(rstmat!familiaad)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  Set rstfam = dbtmpb.OpenRecordset("select descripcio from subfamiliesaditius where codi=" + atrim(cadbl(rstmat!subfamiliaad)))
  If Not rstfam.EOF Then desc = desc + af(rstfam!descripcio)
  descripciomaterial = desc
End Function

Function af(v As Variant) As String
  v = atrim(v)
  If Len(v) > 1 Then
     v = "+" + v
    Else: v = ""
  End If
  af = v
End Function


Private Sub capcalera_Reposition()
 desb1.Caption = "Desb1-" + lotdesb1.Text
 desb2.Caption = "Desb2-" + lotdesb2.Text
End Sub

Private Sub Command1_Click()
 Dim tmp As String
 Set rsttmp = dbtmp.OpenRecordset("select impressora, lotmatdesb1,lotmatdesb2,espessor,tubolam,mesuraesp from comandes where comanda=" + atrim(cadbl(form1.comanda)))
 If Not rsttmp.EOF Then
  tmp = lotdesb1.Text
  lotdesb1.Text = lotdesb2.Text
  lotdesb2.Text = tmp
  emplenar_dades_capcalera
 End If
End Sub

'Sub obrestocks(Optional noobrirbd As Boolean)
'camistocks = llegir_ini("General", "ruta_stocksmdb", "comandes.ini")
'If camistocks = "{[}]" Then camistocks = "\\Ser2\documentos\Stock Reclamaciones\Estoc inplacsa.mdb"
'If Not existeix(camistocks) Then camistocks = "\\serverprodu\dades\progcomandes\dades\copiaestocinplacsa.mdb"
'If Not noobrirbd Then Set dbstocks = OpenDatabase(camistocks)
  
'End Sub

Private Sub Command2_Click()
  If capcalera.Recordset.EditMode > 0 Then capcalera.Recordset.Update
  
  Me.Hide
End Sub

Private Sub Form_Activate()
  Dim lotsnocoincideixen As Boolean
Set rsttmp = dbtmp.OpenRecordset("select impressora,lotmatdesb1,lotmatdesb2,espessor,tubolam,mesuraesp from comandes where comanda=" + atrim(cadbl(form1.comanda)))
  If Not rsttmp.EOF Then
  If cadbl(lotdesb1.Text) <> ncomanda And cadbl(lotdesb1.Text) <> ncomanda2 Then lotsnocoincideixen = True
  If cadbl(lotdesb2.Text) <> ncomanda And cadbl(lotdesb2.Text) <> ncomanda2 Then lotsnocoincideixen = True
   If cadbl(lotdesb1) = 0 Or lotsnocoincideixen Then
      lotdesb1.Text = form1.comanda 'cadbl(rsttmp!lotmatdesb1)
      lotdesb2.Text = form1.linkcomanda 'cadbl(rsttmp!lotmatdesb2)
      emplenar_dades_capcalera
   End If
  End If
   Set rsttmp = Nothing
   Me.Top = form1.Top + 2015
   Me.Left = form1.Left
   micrescomanda = cadbl(Text1) + cadbl(Text2)
   posar_tolerancies_espesor
End Sub
Sub obrestocks(Optional noobrirbd As Boolean)
 Dim camistocks As String
' Set ws = DBEngine.CreateWorkspace("", "admin", "")
 ' If estaobertstocks Then dbtemp.Execute "delete * from selecciobobentrada": Exit Sub
camistocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
'If camistocks = "{[}]" Then camistocks = "\\Ser2\documentos\Stock Reclamaciones\Estoc inplacsa.mdb"
'If Not existeix(camistocks) Then
'    MsgBox "Error obrint la la base de dades de Estocs (Palets) intentarem obrir la BD per defecte", vbCritical, "Error"
'    camistocks = "\\serverprodu\dades\progcomandes\dades\palets.mdb"
'End If

If camistocks = "{[}]" Then escriure_ini "General", "ruta_stocks", rutadelfitxer(cami) + "palets.mdb", "comandes.ini"
camistocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
If Not noobrirbd Then
   Set dbstocks = OpenDatabase(camistocks)
 '  dbtemp.Execute "delete * from selecciobobentrada"
End If
  
End Sub

Sub emplenar_dades_capcalera()
 Dim rs2 As Recordset
 Dim lot1 As Double, lot2 As Double
 
 obrestocks
 ' empleno el desbobinador 1
 
 If form1.proces = "PC2" Then If cadbl(lotdesb1.Text) < cadbl(lotdesb2.Text) Then lot2 = form1.comanda.Tag  'aqui hi ha el petit dels tres
 lot1 = cadbl(lotdesb1.Text)
 r = ""
 capcalera.Recordset!matdesb1 = Mid(possar_desc_lot2(cadbl(lotdesb1.Text), lot1, lot2), 1, 50)
 capcalera.Recordset!amplada1 = cadbl(r)
 capcalera.Recordset!tractatabaix = True
 r = material_comanda(lotdesb1.Text)
  Check1.Value = IIf(rsttmp!impressora <> 0, 1, 0)
  r = material_comanda(atrim(lot1))
 Set rs2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp!mesuraesp)))
 If Not rs2.EOF Then
    capcalera.Recordset!micres1 = micresmaterial(rs2!descripcio, cadbl(rsttmp!espessor))
 End If
 If lot2 > 0 Then
  r = material_comanda(atrim(lot2))
  Check1.Value = IIf(rsttmp!impressora <> 0, 1, 0)
  Set rs2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp!mesuraesp)))
  If Not rs2.EOF Then
    capcalera.Recordset!micres1 = capcalera.Recordset!micres1 + micresmaterial(rs2!descripcio, cadbl(rsttmp!espessor))
    capcalera.Recordset!matdesb1 = Mid(capcalera.Recordset!matdesb1 + " + " + possar_desc_lot2(cadbl(lot2), lot1, lot2), 1, 50)
  End If
 End If
 'empleno el desbobinador 2
 lot2 = 0
 If form1.proces = "PC2" Then If cadbl(lotdesb2.Text) < cadbl(lotdesb1.Text) Then lot2 = form1.comanda.Tag  'aqui hi ha el petit dels tres
 capcalera.Recordset!matdesb2 = Mid(possar_desc_lot2(cadbl(lotdesb2.Text), lot1, lot2), 1, 50)
 capcalera.Recordset!amplada2 = cadbl(r)
 capcalera.Recordset!tractatadalt = True
 
  r = material_comanda(lotdesb2.Text)
  lot1 = cadbl(lotdesb2.Text)
 Check2.Value = IIf(rsttmp!impressora <> 0, 1, 0)
 r = material_comanda(atrim(lot1))
 Set rs2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp!mesuraesp)))
 If Not rs2.EOF Then
    capcalera.Recordset!micres2 = micresmaterial(rs2!descripcio, cadbl(rsttmp!espessor))
 End If
 If lot2 > 0 Then
  r = material_comanda(atrim(lot2))
  Check2.Value = IIf(rsttmp!impressora <> 0, 1, 0)
  Set rs2 = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rsttmp!mesuraesp)))
  If Not rs2.EOF Then
     capcalera.Recordset!micres2 = capcalera.Recordset!micres2 + micresmaterial(rs2!descripcio, cadbl(rsttmp!espessor))
     capcalera.Recordset!matdesb2 = Mid(capcalera.Recordset!matdesb2 + " + " + possar_desc_lot2(cadbl(lot2), lot1, lot2), 1, 50)
  End If
 End If
 If InStr(1, capcalera.Recordset!matdesb1, "PEBD") Then Text5 = "3"
 If InStr(1, capcalera.Recordset!matdesb1, "OPP") Then Text5 = "8"
 If InStr(1, capcalera.Recordset!matdesb2, "PEBD") Then Text6 = "3"
 If InStr(1, capcalera.Recordset!matdesb2, "OPP") Then
    Text6 = "8"
    If InStr(1, capcalera.Recordset!matdesb1, "OPP") Then Text6 = "6"
 End If
 Text11 = atrim(cadbl(Text5) + cadbl(Text6) + 1)
 'actualitzo les dades desb1 i desb2
 capcalera.Recordset.Update
 capcalera.UpdateControls
 
 capcalera.Recordset.Edit
End Sub
 Sub posar_tolerancies_espesor()
    Dim vtotal As Double
    Dim vesptolerancia As Double
    Dim vesptinta As Double
    Dim vespcola As Double
    Dim vesp As Double
    Dim vtanx100 As Double
    ettoleranciadesb1 = ""
    ettoleranciadesb2 = ""
    ettoleranciatotal = ""
    vesp = form1.calcular_espesorteorica_material(lotdesb1, vesptinta, vespcola, vesptolerancia)
    vesp = capcalera.Recordset!micres1
    vtanx100 = (vesp * (vesptolerancia / 100))
    ettoleranciadesb1 = "Tolerancia espesor D1: " + "Min: " + atrim(Redondejar(vesp - vtanx100, 0)) + "    Max: " + atrim(Redondejar(vesp + vtanx100, 0))
    
    vesp = form1.calcular_espesorteorica_material(lotdesb2, vesptinta, vespcola, vesptolerancia)
    vesp = capcalera.Recordset!micres2
    vtanx100 = (vesp * (vesptolerancia / 100))
    ettoleranciadesb2 = "Tolerancia espesor D2: " + "Min: " + atrim(Redondejar(vesp - vtanx100, 0)) + "    Max: " + atrim(Redondejar(vesp + vtanx100, 0))
    
    vtotal = capcalera.Recordset!micres1 + capcalera.Recordset!micres2
    vesp = vtotal
    vtanx100 = (vesp * (vesptolerancia / 100))
    ettoleranciatotal = "Tolerancia Total espesor: " + "Min: " + atrim(Redondejar(vesp - vtanx100, 0)) + "    Max: " + atrim(Redondejar(vesp + vtanx100, 0))
            
 End Sub
 Function material_comanda(comanda As String)
  Dim rs2 As Recordset
  Dim rr As String
  Set rsttmp = dbtmp.OpenRecordset("select impressora,lotmatdesb1,lotmatdesb2,espessor,tubolam,mesuraesp from comandes where comanda=" + atrim(comanda))
'  Set rststocks = dbstocks.OpenRecordset("select Idpalet from bobines where numcom='" + atrim(comanda) + "'")
'  If Not rststocks.EOF Then
'   Set rststocks = dbstocks.OpenRecordset("select * from palets where idpalet=" + atrim(cadbl(rststocks!idpalet)))
'   Set rs2 = dbstocks.OpenRecordset("select [Idfam] from productes where [Idprod]=" + atrim(cadbl(rststocks!Idprod)))
'   If Not rs2.EOF Then Set rs2 = dbstocks.OpenRecordset("select [Nomfam] from families where [Idfam]=" + atrim(cadbl(rs2!idfam)))
   rr = ""
'   If Not rs2.EOF Then rr = rs2!Nomfam
     'Else: MsgBox "No hi ha bobines asignades a la comanda " + comanda: Exit Function
  'End If
 'Set rs2 = Nothing
 material_comanda = rr
 End Function
 Function possar_desc_lot(numlot As String, ByRef lot1 As Double, ByRef lot2 As Double) As String
  Dim desctmp As String
  Dim rsttmp2 As Recordset
  Dim rsttmp3 As Recordset
  Dim proces As Byte
  Dim ample As Double
  proces = 1
  desctmp = ""
  desclotx = desctmp
  If cadbl(numlot) < 1 Then Exit Function
  Set rsttmp3 = dbtmp.OpenRecordset("select producte,lotmatdesb1,lotmatdesb2,laminadora,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(cadbl(numlot)))
  If rsttmp3.EOF Then possar_desc_lot = "": Exit Function
  If cadbl(rsttmp3!laminadora) > 0 Then
    If rsttmp3!producte = "PC2" Then
      lot1 = numlot: lot2 = 0: proces = 2
       Else
        lot1 = cadbl(rsttmp3!lotmatdesb1)
        lot2 = cadbl(rsttmp3!lotmatdesb2)
     End If
     If rsttmp3!producte <> "PC" And rsttmp3!producte <> "PC2" Then lot2 = 0
     Else:
       lot1 = cadbl(numlot)
       If rsttmp3!producte = "PC" Then lot2 = lot1: lot1 = cadbl(rsttmp3!linkcomanda1)
  End If
  Set rsttmp3 = dbtmp.OpenRecordset("select materialex,colorex,espessor,mesuraesp from comandes where comanda=" + atrim(lot1))
  If Not rsttmp3.EOF Then
    Set rsttmp2 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rsttmp3!materialex)))
    If Not rsttmp2.EOF Then
     possar_desc_lot = descripciomaterial(rsttmp2)
    End If
  End If
  Set rsttmp2 = dbstocks.OpenRecordset("SELECT Max(Palets.Ample) AS elmesample, Parcials.comanda FROM Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + numlot + "'));")
  If Not rsttmp2.EOF Then ample = rsttmp2!elmesample
  'si hi ha dos materials al lot els suma
  If lot2 > 0 And proces = 2 Then
    Set rsttmp3 = dbtmp.OpenRecordset("select linkcomanda1,materialex,colorex,espessor,mesuraesp from comandes where comanda=" + atrim(lot2))
    If Not rsttmp3.EOF Then
     Set rsttmp2 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rsttmp3!materialex)))
     If Not rsttmp2.EOF Then
       possar_desc_lot = atrim(possar_desc_lot) + " + " + descripciomaterial(rsttmp2)
     End If
    End If
    Set rsttmp2 = dbstocks.OpenRecordset("SELECT Max(Palets.Ample) AS elmesample, Parcials.comanda FROM Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + numlot + "'));")
    If Not rsttmp2.EOF Then r = rsttmp2!elmesample
    If cadbl(r) > ample Then ample = cadbl(r)
  End If
  r = ample
  Set rsttmp2 = Nothing
  Set rsttmp3 = Nothing
End Function
Function possar_desc_lot2(numlot As String, ByRef lot1 As Double, ByRef lot2 As Double) As String
  Dim desctmp As String
  Dim rsttmp2 As Recordset
  Dim rsttmp3 As Recordset
  Dim proces As Byte
  r = "0"
  proces = 1
  desctmp = ""
  desclotx = desctmp
  If cadbl(numlot) < 1 Then Exit Function
  Set rsttmp3 = dbtmp.OpenRecordset("select producte,lotmatdesb1,lotmatdesb2,laminadora,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(cadbl(numlot)))
  If rsttmp3.EOF Then possar_desc_lot2 = "": Exit Function
    Set rsttmp3 = dbtmp.OpenRecordset("select materialex,colorex,espessor,mesuraesp from comandes where comanda=" + atrim(numlot))
    If Not rsttmp3.EOF Then
     Set rsttmp2 = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rsttmp3!materialex)))
     possar_desc_lot2 = descripciomaterial(rsttmp2)
    End If
  Set rsttmp2 = dbstocks.OpenRecordset("SELECT Max(Palets.Ample) AS elmesample, Parcials.comanda FROM Parcials INNER JOIN Palets ON Parcials.idpalet = Palets.Idpalet GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + numlot + "'));")
  If Not rsttmp2.EOF Then
     r = rsttmp2!elmesample
  End If
  If amplematerialestoc(cadbl(numlot)) > 0 Then r = atrim(amplematerialestoc(cadbl(numlot)))
  'r = material_comanda(atrim(numlot))
  On Error Resume Next
  'r = rststocks!ample
  Set rsttmp2 = Nothing
  Set rsttmp3 = Nothing
End Function
Function amplematerialestoc(numlot As Double) As Double
   Dim rstm As Recordset
   Set rstm = dbstocks.OpenRecordset("SELECT opcionsdajust.comanda, grupsdepalets.ample FROM opcionsdajust LEFT JOIN grupsdepalets ON opcionsdajust.grupdestoc = grupsdepalets.numerogrup where comanda=" + atrim(cadbl(numlot)))
   If Not rstm.EOF Then
       amplematerialestoc = cadbl(rstm!ample)
         Else: amplematerialestoc = 0
   End If
   Set rstm = Nothing

End Function


Function micresmaterial(descripcio As String, espesor As Double) As Double
  r = espesor
  If descripcio = "GALGUES" Then
            If rsttmp!tubolam = "T" Then
                 r = Format(espesor / 4, "#,##0")
                  Else: r = Format(espesor / 2, "#,##0")
            End If
  End If
  
  micresmaterial = r
End Function

Private Sub Text18_Change()

End Sub

Private Sub Form_Click()
  'MsgBox Me.Top
End Sub

