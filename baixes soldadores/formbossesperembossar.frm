VERSION 5.00
Begin VB.Form formbossesperembossar 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bosses i Canutus"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3795
   Icon            =   "formbossesperembossar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framealtreslots 
      Caption         =   "Altres Lots"
      Height          =   1395
      Left            =   195
      TabIndex        =   26
      Top             =   3525
      Width           =   3480
      Begin VB.CommandButton Command12 
         Height          =   390
         Left            =   2670
         Picture         =   "formbossesperembossar.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Canvi codi de LOT"
         Top             =   690
         Width           =   750
      End
      Begin VB.CommandButton Command11 
         Height          =   390
         Left            =   2670
         Picture         =   "formbossesperembossar.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Canvi codi de LOT"
         Top             =   225
         Width           =   750
      End
      Begin VB.TextBox clotcinta 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   705
         Width           =   1620
      End
      Begin VB.TextBox clotzipper 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   255
         Width           =   1620
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Cinta Adhesiva:"
         Height          =   390
         Left            =   135
         TabIndex        =   30
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Zipper:"
         Height          =   285
         Left            =   150
         TabIndex        =   28
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.Frame framecanvilots 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Possar lots compra de canutus"
      Height          =   1050
      Left            =   495
      TabIndex        =   24
      Top             =   5385
      Visible         =   0   'False
      Width           =   3690
      Begin VB.CommandButton Command10 
         BackColor       =   &H0076B5E9&
         Caption         =   "Canviar Lots de Canutus"
         Height          =   435
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   375
         Width           =   2355
      End
      Begin VB.Image Image3 
         Height          =   795
         Left            =   2805
         Picture         =   "formbossesperembossar.frx":0E9E
         Stretch         =   -1  'True
         Top             =   210
         Width           =   795
      End
   End
   Begin VB.Frame frameactdes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Activar/Desactivar LOTS de BOSSES"
      Height          =   1395
      Left            =   900
      TabIndex        =   16
      Top             =   5010
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Command9 
         Height          =   390
         Left            =   2895
         Picture         =   "formbossesperembossar.frx":BB00
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Acceptar"
         Top             =   315
         Width           =   390
      End
      Begin VB.CommandButton Command8 
         Height          =   420
         Left            =   2025
         Picture         =   "formbossesperembossar.frx":C08A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Cancelar LOT"
         Top             =   915
         Width           =   1065
      End
      Begin VB.CommandButton Command7 
         Height          =   420
         Left            =   540
         Picture         =   "formbossesperembossar.frx":C614
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Acceptar LOT."
         Top             =   915
         Width           =   1125
      End
      Begin VB.TextBox lotperactivar 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1380
         TabIndex        =   17
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "DESActivar LOT"
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   2085
         TabIndex        =   22
         Top             =   735
         Width           =   1200
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Activar LOT"
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   765
         TabIndex        =   21
         Top             =   735
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de LOT:"
         Height          =   285
         Left            =   465
         TabIndex        =   18
         Top             =   405
         Width           =   975
      End
   End
   Begin VB.CommandButton Command6 
      Height          =   390
      Left            =   45
      Picture         =   "formbossesperembossar.frx":CB9E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Comandes visibles i acavades."
      Top             =   5175
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Caixes"
      Height          =   1665
      Left            =   210
      TabIndex        =   8
      Top             =   1815
      Width           =   3375
      Begin VB.TextBox canutus1 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   345
         Width           =   1350
      End
      Begin VB.TextBox canutus2 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1065
         TabIndex        =   11
         Top             =   810
         Width           =   1350
      End
      Begin VB.CommandButton Command5 
         Height          =   330
         Left            =   2415
         Picture         =   "formbossesperembossar.frx":D128
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Acceptar"
         Top             =   315
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton Command2 
         Height          =   330
         Left            =   2430
         Picture         =   "formbossesperembossar.frx":D6B2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Acceptar"
         Top             =   765
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   2490
         Picture         =   "formbossesperembossar.frx":DC3C
         Stretch         =   -1  'True
         Top             =   915
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Lot:"
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Segon  Lot:"
         Height          =   285
         Left            =   165
         TabIndex        =   13
         Top             =   810
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   390
      Left            =   2655
      Picture         =   "formbossesperembossar.frx":11573E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Acceptar"
      Top             =   5190
      Width           =   870
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bosses"
      Height          =   1665
      Left            =   210
      TabIndex        =   0
      Top             =   75
      Width           =   3375
      Begin VB.CommandButton Command4 
         Height          =   390
         Left            =   2415
         Picture         =   "formbossesperembossar.frx":115CC8
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Acceptar"
         Top             =   720
         Width           =   390
      End
      Begin VB.CommandButton Command3 
         Height          =   390
         Left            =   2415
         Picture         =   "formbossesperembossar.frx":116252
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Acceptar"
         Top             =   315
         Width           =   390
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   810
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   345
         Width           =   1350
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   2730
         Picture         =   "formbossesperembossar.frx":1167DC
         Stretch         =   -1  'True
         Top             =   945
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot exterior:"
         Height          =   285
         Left            =   165
         TabIndex        =   5
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot interior:"
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   345
         Width           =   975
      End
   End
End
Attribute VB_Name = "formbossesperembossar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select comandabosses1,comandabosses2,comandacaixes1,comandacaixes2 from soldadorestot where comanda=" + atrim(cadbl(Form1.comanda)))
  If rst.EOF Then Unload Me
  rst.Edit
  If Text1 = "Sense Bossa" Then Text1 = "1"
  If Text2 = "Sense Bossa" Then Text2 = "1"
  rst!comandabosses1 = cadbl(Text1)
  rst!comandabosses2 = cadbl(Text2)
  rst!comandacaixes1 = atrim(canutus1)
  rst!comandacaixes2 = atrim(canutus2)
  rst.Update
  Set rst = Nothing
  Unload Me
End Sub

Private Sub Command10_Click()
  Dim resp As String
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "SELECT cm_int as Mida,lotcompra1 ,lotcompra2 ,lotcompra3  from tubbase"
  formseleccio.caption = "Selecció del canutu"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 1000
  formseleccio.DBGrid2.Columns(1).width = 2100
  formseleccio.DBGrid2.Columns(2).width = 2100
  formseleccio.DBGrid2.Columns(3).width = 2100
  formseleccio.DBGrid2.MarqueeStyle = 2
  'formseleccio..caption = "Per canviar un valor fer doble clic a sobre."
  formseleccio.Show 1
  If seleccioret = 1 Then
    If InStr(1, formseleccio.DBGrid2.Columns(formseleccio.DBGrid2.col).DataField, "lotcompra") Then
         resp = InputBox("Entra el nou LOT de COMPRA:", "LOT DE COMPRA CANUTUS", formseleccio.DBGrid2.Text)
         If resp <> "" Then
            formseleccio.Data1.Recordset.Edit
            formseleccio.Data1.Recordset.Fields(formseleccio.DBGrid2.Columns(formseleccio.DBGrid2.col).DataField) = resp
            formseleccio.Data1.Recordset.Update
         End If
    End If
  End If
  Unload formseleccio
End Sub

Private Sub Command11_Click()
   Dim v As String
   v = InputBox("Entra el numero de LOT del Zipper:" + vbNewLine + "Escaneja'l o escriulo.", "Canvi de Lot de zipper")
   If StrPtr(v) = 0 Or atrim(v) = "" Then Exit Sub
   clotzipper = v
   escriure_ini "Baixes", "LotZipper", v, "comandes.ini"
   
End Sub

Private Sub Command12_Click()
Dim v As String
   v = InputBox("Entra el numero de LOT del cinta adhesiva:" + vbNewLine + "Escaneja'l o escriulo.", "Canvi de Lot de la cinta adhesiva")
   If StrPtr(v) = 0 Or atrim(v) = "" Then Exit Sub
   clotcinta = v
   escriure_ini "Baixes", "LotCinta", v, "comandes.ini"
End Sub

Private Sub Command2_Click()
  Load formseleccio
  formseleccio.Data1.DatabaseName = possartubsatemporal
  formseleccio.Data1.RecordSource = "SELECT * from tubsbase"
  formseleccio.caption = "Selecció del canutu"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 1000
  formseleccio.DBGrid2.Columns(1).width = 4000
  formseleccio.Show 1
  If seleccioret = 1 Then
   canutus2 = atrim(formseleccio.Data1.Recordset!comanda)
  End If
  Unload formseleccio
End Sub

Private Sub Command3_Click()

   Load formseleccio
  formseleccio2.Data1.DatabaseName = cami
  formseleccio2.Data1.RecordSource = "SELECT comandes.comanda, Str([amplesol])+'/'+Str([ampleplegsol])+'X'+Str([longitudsol]) & ' ' & IIf([familiescolorants].[descripcio]<>'TRANSPARENT',[familiescolorants].[descripcio],'') AS Mides FROM (familiescolorants INNER JOIN materials ON familiescolorants.codi = materials.familiacol) INNER JOIN comandes ON materials.codi = comandes.materialex WHERE (((comandes.client)=7) AND ((comandes.proximaseccio)='V') and comandes.comanda in (select comanda from lotsdebosses where activada=true));"
'  formseleccio2.Data1.RecordSource = "SELECT comandes.comanda, Str([amplesol])+'/'+Str([ampleplegsol])+'X'+Str([longitudsol]) AS Mides From comandes WHERE (((comandes.client)=7) AND ((comandes.proximaseccio)='V')) and comandes.comanda in (select comanda from lotsdebosses where activada=true)"
  formseleccio2.caption = "Selecció comanda bosses"
  formseleccio2.refrescar
 ' formseleccio.bsensebossa.visible = True
  formseleccio2.DBGrid2.Columns(0).width = 2500
  formseleccio2.DBGrid2.Columns(1).width = 7000
  formseleccio2.width = 12000
  formseleccio2.DBGrid2.width = 10500
  
  formseleccio2.Show 1
  If seleccioret = 1 Then
   Text1 = formseleccio2.Data1.Recordset!comanda 'IIf(formseleccio.bsensebossa.tag = "1", "1", cadbl(formseleccio.Data1.Recordset!comanda))
   If Text1 = "1" Then Text1 = "Sense Bossa"
   If Text2 = "1" Then Text2 = "Sense Bossa"
  End If
  Unload formseleccio
  


End Sub

Private Sub Command4_Click()
 Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "SELECT comandes.comanda, Str([amplesol])+'/'+Str([ampleplegsol])+'X'+Str([longitudsol]) AS Mides From comandes WHERE (((comandes.client)=7) AND ((comandes.proximaseccio)='V'));"
  formseleccio.caption = "Selecció comanda bosses"
  formseleccio.refrescar
   formseleccio.DBGrid2.Columns(0).width = 2500
  formseleccio.DBGrid2.Columns(1).width = 4500
  formseleccio.Show 1
  If seleccioret = 1 Then
   'Text2 = IIf(formseleccio.bsensebossa.tag = "1", "1", cadbl(formseleccio.Data1.Recordset!comanda))
   If Text1 = "1" Then Text1 = "Sense Bossa"
   If Text2 = "1" Then Text2 = "Sense Bossa"
  End If
  Unload formseleccio
  
End Sub
Sub escullirisortir(diametrecanutu As Double)
   Dim valor As Double
   Me.visible = False
   valor = cadbl(Text1)
   If valor = 0 Then
     Command3_Click
'     Command5_Click
   End If
   If diametrecanutu > 0 And atrim(canutus1) = "" Then
      valor = cadbl(InputBox("Entra el diametre del canutu que utilitzaràs." + Chr(10) + " Hauria de ser " + atrim(diametrecanutu), "Diametre Canutu"))
     If valor = diametrecanutu Then
         canutus1 = UCase(InputBox("Entra la referencia del canutu que utilitzes." + Chr(10) + "Ex: S010115", "Referència canutu"))
        Else: MsgBox "Aquest diametre que has entrat no es el que hi ha apuntat a la comanda." + Chr(10) + " Assegura que sigui correcte.", vbCritical, "Atenció"
     End If
   End If
   Command1_Click
   

End Sub
Function possartubsatemporal() As String
  Dim rst As Recordset
  Dim rsttb As Recordset
  possartubsatemporal = crear_taulatemp_tubsbase
  Set rst = dbtmp.OpenRecordset("tubbase")
  Set rsttb = dbtemp.OpenRecordset("select * from tubsbase")
  While Not rst.EOF
     If atrim(rst!lotcompra1) <> "" Then
       rsttb.AddNew
       rsttb!mida = cadbl(rst!cm_int)
       rsttb!comanda = UCase(rst!lotcompra1)
       rsttb.Update
     End If
     If atrim(rst!lotcompra2) <> "" Then
       rsttb.AddNew
       rsttb!mida = cadbl(rst!cm_int)
       rsttb!comanda = UCase(rst!lotcompra2)
       rsttb.Update
     End If
     If atrim(rst!lotcompra3) <> "" Then
       rsttb.AddNew
       rsttb!mida = cadbl(rst!cm_int)
       rsttb!comanda = UCase(rst!lotcompra3)
       rsttb.Update
     End If
     rst.MoveNext
  Wend
  Set rst = Nothing
  Set rsttb = Nothing
End Function
Function crear_taulatemp_tubsbase() As String
  Dim nomfitxertemporal As String
  nomfitxertemporal = "c:\temp\~tubbase" + Format(Now, "ddmmhhnnss") + ".mdb"
  On Error Resume Next
   MkDir "c:\temp"
   Kill "c:\temp\~tubbase*.*"
   DBEngine.CreateDatabase nomfitxertemporal, dbLangGeneral, dbVersion10
   Set dbtemp = OpenDatabase(nomfitxertemporal)
   'dbtemp.Execute "drop table tmp_imp_empalmes"
  On Error GoTo 0
  camps = "mida double,comanda string(15)"
  dbtemp.Execute ("create table tubsbase (" + camps) + ")"
  crear_taulatemp_tubsbase = nomfitxertemporal
End Function
Private Sub Command5_Click()
  
  Load formseleccio
  formseleccio.Data1.DatabaseName = possartubsatemporal
  formseleccio.Data1.RecordSource = "SELECT * from tubsbase"
  formseleccio.caption = "Selecció del canutu"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 1000
  formseleccio.DBGrid2.Columns(1).width = 4000
  formseleccio.Show 1
  If seleccioret = 1 Then
   canutus1 = atrim(formseleccio.Data1.Recordset!comanda)
  End If
  Unload formseleccio
End Sub

Private Sub Command6_Click()
    If formbossesperembossar.Height < 6500 Then
         If UCase(InputBox("Entra el codi d'encarregat.", "Atenció")) = "INPLACSA" Then
             'formbossesperembossar.Height = 4700 + frameactdes.Height '+ framecanvilots.Height
             frameactdes.Top = 3700
             frameactdes.Left = 30
             frameactdes.visible = True
             
         End If
           Else: formbossesperembossar.Height = 4400
    End If
End Sub

Private Sub Command7_Click()
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(lotperactivar)))
   If rst.EOF Then MsgBox "Aquesta comanda no existeix.", vbCritical, "ERROR": Exit Sub
   If rst!client <> 7 Then MsgBox "Aquesta comanda no es interna de REBOBINADORES no es pot activar per bosses.", vbCritical, "ERROR": Exit Sub
   If rst!proximaseccio <> "V" And rst!proximaseccio <> "T" And rst!proximaseccio <> "P" Then MsgBox "Aquesta comanda encara no està acavada de produïr no es pot activar.", vbCritical, "Atenció": Exit Sub
   If rst!proximaseccio = "T" Then
     If MsgBox("Aquesta comanda ja està marcada com a utilitzada, vols activar-la igualment?", vbInformation, "Atenció") = vbNo Then Exit Sub
   End If
   dbtmp.Execute "update comandes set proximaseccio='V' where comanda=" + atrim(cadbl(lotperactivar))
   dbtmpb.Execute "delete * from lotsdebosses where comanda=" + atrim(cadbl(lotperactivar))
   dbtmpb.Execute "insert into lotsdebosses (comanda,activada) values (" + atrim(cadbl(lotperactivar)) + ",true)"
   MsgBox "La comanda: " + atrim(lotperactivar) + " ACTIVADA."
      
End Sub

Private Sub Command8_Click()
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(lotperactivar)))
   If rst.EOF Then MsgBox "Aquesta comanda no existeix.", vbCritical, "ERROR": Exit Sub
   If rst!client <> 7 Then MsgBox "Aquesta comanda no es interna de REBOBINADORES no es pot activar per bosses.", vbCritical, "ERROR": Exit Sub
   If rst!proximaseccio <> "V" And rst!proximaseccio <> "T" And rst!proximaseccio <> "P" Then MsgBox "Aquesta comanda encara no està acavada de produïr no es pot DESACTIVAR.", vbCritical, "Atenció": Exit Sub
   dbtmp.Execute "update comandes set proximaseccio='T' where comanda=" + atrim(cadbl(lotperactivar))
   dbtmpb.Execute "delete * from lotsdebosses where comanda=" + atrim(cadbl(lotperactivar))
   dbtmpb.Execute "insert into lotsdebosses (comanda,activada) values (" + atrim(cadbl(lotperactivar)) + ",false)"
   MsgBox "La comanda: " + atrim(lotperactivar) + " DESACTIVADA."
End Sub

Private Sub Command9_Click()

   Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "SELECT comandes.comanda, Str([amplesol])+'/'+Str([ampleplegsol])+'X'+Str([longitudsol]) AS Mides From comandes WHERE (((comandes.client)=7) AND ((comandes.proximaseccio)='V')) and comandes.comanda "
  formseleccio.caption = "Selecció comanda bosses"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 2500
  formseleccio.DBGrid2.Columns(1).width = 4500
  formseleccio.Show 1
  If seleccioret = 1 Then
   lotperactivar = atrim(cadbl(formseleccio.Data1.Recordset!comanda))
  End If
  Unload formseleccio
End Sub

Private Sub Form_Load()
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select comandabosses1,comandabosses2,comandacaixes1,comandacaixes2 from soldadorestot where comanda=" + atrim(cadbl(Form1.comanda)))
  If rst.EOF Then Unload Me: Exit Sub
  canutus1 = atrim(rst!comandacaixes1)
  canutus2 = atrim(rst!comandacaixes2)
  Text1 = cadbl(rst!comandabosses1)
  Text2 = cadbl(rst!comandabosses2)
  If Text1 = "1" Then Text1 = "Sense Bossa"
  If Text2 = "1" Then Text2 = "Sense Bossa"
  Set rst = Nothing
  formbossesperembossar.Height = 6000
  clotcinta = llegir_ini("Baixes", "LotCinta", "comandes.ini")
  clotzipper = llegir_ini("Baixes", "LotZipper", "comandes.ini")
End Sub

