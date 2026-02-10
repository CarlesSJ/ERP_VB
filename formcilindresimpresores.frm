VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formcilindresimpresores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment de Cilindres d'impresores"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10860
   Icon            =   "formcilindresimpresores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Saber desarroll"
      Height          =   375
      Left            =   9225
      TabIndex        =   31
      Top             =   30
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recalcular Desarrolls"
      Height          =   330
      Left            =   6075
      TabIndex        =   30
      Top             =   45
      Width           =   1995
   End
   Begin VB.CommandButton alta 
      Height          =   450
      Left            =   75
      Picture         =   "formcilindresimpresores.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Alta  Registres"
      Top             =   465
      Width           =   435
   End
   Begin VB.CommandButton eliminar 
      Height          =   450
      Left            =   945
      Picture         =   "formcilindresimpresores.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   465
      Width           =   435
   End
   Begin VB.CommandButton modificar 
      Height          =   450
      Left            =   510
      Picture         =   "formcilindresimpresores.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Modificar Registres"
      Top             =   465
      Width           =   435
   End
   Begin VB.CommandButton guardar 
      Height          =   450
      Left            =   1380
      Picture         =   "formcilindresimpresores.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Acceptar els canvis (F1)."
      Top             =   465
      Width           =   435
   End
   Begin VB.CommandButton triar 
      Caption         =   "Triar"
      Height          =   300
      Left            =   3420
      TabIndex        =   3
      Top             =   75
      Width           =   1155
   End
   Begin VB.ComboBox seleccioimpresora 
      Height          =   315
      Left            =   285
      TabIndex        =   2
      Top             =   75
      Width           =   3105
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   765
      Left            =   45
      TabIndex        =   1
      Top             =   840
      Width           =   10650
      Begin VB.TextBox maqprincipal 
         DataField       =   "nummaquinaprincipal"
         DataSource      =   "datacilindres"
         Height          =   300
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox maqalt3 
         DataField       =   "nummaquinaalternativa3"
         DataSource      =   "datacilindres"
         Height          =   300
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox maqalt2 
         DataField       =   "nummaquinaalternativa2"
         DataSource      =   "datacilindres"
         Height          =   300
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox maqalt1 
         DataField       =   "nummaquinaalternativa1"
         DataSource      =   "datacilindres"
         Height          =   300
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.ComboBox maquinaalt3 
         Height          =   315
         Left            =   9150
         TabIndex        =   12
         Top             =   375
         Width           =   1080
      End
      Begin VB.ComboBox maquinaalt2 
         Height          =   315
         Left            =   8013
         TabIndex        =   11
         Top             =   375
         Width           =   1080
      End
      Begin VB.ComboBox maquinaalt1 
         Height          =   315
         Left            =   6879
         TabIndex        =   10
         Top             =   375
         Width           =   1080
      End
      Begin VB.TextBox gruixpolimer 
         DataField       =   "gruixpolimer"
         DataSource      =   "datacilindres"
         Height          =   300
         Left            =   6315
         TabIndex        =   9
         Top             =   375
         Width           =   510
      End
      Begin VB.TextBox desenvolupament 
         DataField       =   "desenvolupamentcamisa"
         DataSource      =   "datacilindres"
         Height          =   300
         Left            =   5616
         TabIndex        =   8
         Top             =   375
         Width           =   645
      End
      Begin VB.TextBox amplecamisa 
         DataField       =   "amplecamisa"
         DataSource      =   "datacilindres"
         Height          =   300
         Left            =   4872
         TabIndex        =   7
         Top             =   375
         Width           =   690
      End
      Begin VB.TextBox colorcamisa 
         DataField       =   "colorcamisa"
         DataSource      =   "datacilindres"
         Height          =   300
         Left            =   3153
         TabIndex        =   6
         Top             =   375
         Width           =   1665
      End
      Begin VB.TextBox camises 
         DataField       =   "numcamises"
         DataSource      =   "datacilindres"
         Height          =   300
         Left            =   2589
         TabIndex        =   5
         Top             =   375
         Width           =   510
      End
      Begin VB.ComboBox tipuscilindre 
         DataField       =   "portaclixeosleeve"
         DataSource      =   "datacilindres"
         Height          =   315
         ItemData        =   "formcilindresimpresores.frx":1BB2
         Left            =   135
         List            =   "formcilindresimpresores.frx":1BBC
         TabIndex        =   4
         Top             =   375
         Width           =   2400
      End
      Begin VB.Label etiquetes 
         BackStyle       =   0  'Transparent
         Caption         =   "Màquina Alt.3"
         Height          =   195
         Index           =   8
         Left            =   9180
         TabIndex        =   21
         Top             =   150
         Width           =   1170
      End
      Begin VB.Label etiquetes 
         BackStyle       =   0  'Transparent
         Caption         =   "Màquina Alt.2"
         Height          =   195
         Index           =   7
         Left            =   8010
         TabIndex        =   20
         Top             =   150
         Width           =   1170
      End
      Begin VB.Label etiquetes 
         BackStyle       =   0  'Transparent
         Caption         =   "Màquina Alt.1"
         Height          =   195
         Index           =   6
         Left            =   6885
         TabIndex        =   19
         Top             =   150
         Width           =   1170
      End
      Begin VB.Label etiquetes 
         BackStyle       =   0  'Transparent
         Caption         =   "Gruix"
         Height          =   195
         Index           =   5
         Left            =   6345
         TabIndex        =   18
         Top             =   150
         Width           =   780
      End
      Begin VB.Label etiquetes 
         BackStyle       =   0  'Transparent
         Caption         =   "Desenvol."
         Height          =   195
         Index           =   4
         Left            =   5550
         TabIndex        =   17
         Top             =   150
         Width           =   795
      End
      Begin VB.Label etiquetes 
         BackStyle       =   0  'Transparent
         Caption         =   "Ample"
         Height          =   195
         Index           =   3
         Left            =   4965
         TabIndex        =   16
         Top             =   150
         Width           =   705
      End
      Begin VB.Label etiquetes 
         BackStyle       =   0  'Transparent
         Caption         =   "Color de la camisa"
         Height          =   195
         Index           =   2
         Left            =   3390
         TabIndex        =   15
         Top             =   150
         Width           =   1395
      End
      Begin VB.Label etiquetes 
         BackStyle       =   0  'Transparent
         Caption         =   "NºCamises"
         Height          =   195
         Index           =   1
         Left            =   2505
         TabIndex        =   14
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label etiquetes 
         BackStyle       =   0  'Transparent
         Caption         =   "Portaclixé o Sleeve"
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   13
         Top             =   150
         Width           =   1830
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "formcilindresimpresores.frx":1BD4
      Height          =   5895
      Left            =   30
      OleObjectBlob   =   "formcilindresimpresores.frx":1BEC
      TabIndex        =   0
      Top             =   1740
      Width           =   10695
   End
   Begin VB.Data datacilindres 
      Caption         =   "datacilindres"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   6165
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   390
      Visible         =   0   'False
      Width           =   2880
   End
End
Attribute VB_Name = "formcilindresimpresores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
  If datacilindres.Recordset.EditMode > 0 Then Exit Sub
  Frame1.Enabled = True
  datacilindres.Recordset.AddNew
  datacilindres.Recordset!nummaquinaprincipal = seleccioimpresora.ItemData(seleccioimpresora.ListIndex)
  tipuscilindre.SetFocus
End Sub

Private Sub Command1_Click()
   Dim rst As Recordset
   Dim divisor As Double
   Dim cilindre As Double
   Dim id_cilindre As Double
   Set rst = datacilindres.Database.OpenRecordset("select * from cilindres")
   datacilindres.Database.Execute "delete * from desarrolls"
   ratoli "espera"
   While Not rst.EOF
     divisor = 1
     cilindre = cadbl(rst!desenvolupamentcamisa)
     id_cilindre = rst!id_cilindre
     While (cilindre / divisor) > 45
       datacilindres.Database.Execute "insert into desarrolls (id_cilindre,cilindre,divisor,desarroll) values (" + atrim(id_cilindre) + "," + atrim(cilindre) + "," + atrim(divisor) + "," + passaradecimalpunt(Redondejar(cilindre / divisor, 2)) + ")"
       divisor = divisor + 1
     Wend
     rst.MoveNext
   Wend
   
   ratoli "normal"
   MsgBox "Proces Acabat"
End Sub

Private Sub Command2_Click()
  Dim des As Double
  Dim sql As String
  Dim caigudes As Double
  des = cadbl(InputBox("Quin desarroll vols utilitzar? en Cm", "Desarroll"))
  If des = 0 Then Exit Sub
  des = des * 10
  caigudes = cadbl(InputBox("Quantes caigudes vols?", "Caigudes"))
  If caigudes = 0 Then Exit Sub
  sql = "SELECT maquines.descripcio, desarrolls.cilindre, desarrolls.desarroll FROM maquines INNER JOIN (Cilindres INNER JOIN desarrolls ON Cilindres.id_cilindre = desarrolls.id_cilindre) ON maquines.codi = Cilindres.nummaquinaprincipal WHERE maquines.maquina='I'"

  Load formseleccio
  formseleccio.Data1.DatabaseName = datacilindres.DatabaseName
  formseleccio.Data1.RecordSource = sql + "and (desarroll>=" + atrim(des) + "-2 and desarroll<=" + atrim(des + 2) + ") and divisor=" + atrim(caigudes) + " order by cilindre"
  formseleccio.refrescar
  formseleccio.Show 1
  'If seleccioret = 1 Then
 '  Text2.Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
 '  Data1.Recordset!client = Text2.Text
 '  nomclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
  
 ' End If
  Unload formseleccio
  
End Sub

Private Sub eliminar_Click()
   If datacilindres.Recordset.EditMode > 0 Then MsgBox "Estas editant...": Exit Sub
   If MsgBox("Segur que vols eliminar aquest cilindre?", vbCritical + vbYesNo, "Atenció") = vbNo Then Exit Sub
   datacilindres.Database.Execute "delete * from desarrolls where id_cilindre=" + atrim(datacilindres.Recordset!id_cilindre)
   datacilindres.Recordset.Delete
   datacilindres.Refresh
End Sub

Private Sub Form_Load()
   possarnommaquines
   datacilindres.DatabaseName = cami
   datacilindres.RecordSource = "select * from cilindres where id_cilindre=0"
   datacilindres.Refresh
   If seleccioimpresora.ListCount > 1 Then seleccioimpresora.ListIndex = 0
  
   
End Sub
Sub possarnommaquines()
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from maquines where maquina='I' and donadadebaixa=null")
  While Not rst.EOF
    maquinaalt1.AddItem atrim(rst!descripcio)
    maquinaalt1.ItemData(maquinaalt1.NewIndex) = cadbl(rst!codi)
    maquinaalt2.AddItem atrim(rst!descripcio)
    maquinaalt2.ItemData(maquinaalt2.NewIndex) = cadbl(rst!codi)
    maquinaalt3.AddItem atrim(rst!descripcio)
    maquinaalt3.ItemData(maquinaalt3.NewIndex) = cadbl(rst!codi)
    seleccioimpresora.AddItem atrim(rst!descripcio)
    seleccioimpresora.ItemData(seleccioimpresora.NewIndex) = cadbl(rst!codi)
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MsgBox "Si has fet algun canvi de tamany de desarrolls o has afegit algun cilindre pensa a fer el recalcular desarrols per actualitzar-los", vbCritical, "Atenció"
End Sub

Private Sub guardar_Click()
 If datacilindres.Recordset.EditMode = 0 Then Exit Sub
  Frame1.Enabled = False
  datacilindres.Recordset.Update
End Sub

Private Sub maqalt1_Change()
Dim i As Byte
   For i = 0 To maquinaalt1.ListCount - 1
     If maquinaalt1.ItemData(i) = cadbl(maqalt1) Then maquinaalt1.ListIndex = i
   Next i
End Sub

Private Sub maqalt2_Change()
Dim i As Byte
   For i = 0 To maquinaalt2.ListCount - 1
     If maquinaalt2.ItemData(i) = cadbl(maqalt2) Then maquinaalt2.ListIndex = i
   Next i
End Sub

Private Sub maqalt3_Change()
   Dim i As Byte
   For i = 0 To maquinaalt3.ListCount - 1
     If maquinaalt3.ItemData(i) = cadbl(maqalt3) Then maquinaalt3.ListIndex = i
   Next i
End Sub

Private Sub maquinaalt1_Click()
  maqalt1 = maquinaalt1.ItemData(maquinaalt1.ListIndex)
End Sub

Private Sub maquinaalt2_Click()
  maqalt2 = maquinaalt2.ItemData(maquinaalt2.ListIndex)
End Sub

Private Sub maquinaalt3_Click()
maqalt3 = maquinaalt3.ItemData(maquinaalt3.ListIndex)
End Sub

Private Sub modificar_Click()
  If datacilindres.Recordset.EditMode > 0 Then Exit Sub
  Frame1.Enabled = True
  datacilindres.Recordset.Edit
End Sub

Private Sub triar_Click()
   If seleccioimpresora.ListIndex = -1 Then Exit Sub
   datacilindres.RecordSource = "select * from cilindres where nummaquinaprincipal=" + atrim(seleccioimpresora.ItemData(seleccioimpresora.ListIndex))
   datacilindres.Refresh
End Sub
