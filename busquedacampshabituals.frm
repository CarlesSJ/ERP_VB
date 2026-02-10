VERSION 5.00
Begin VB.Form formbusquedahabitual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda pels camps habituals"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5940
   Icon            =   "busquedacampshabituals.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bproximabusqueda 
      Height          =   450
      Left            =   4620
      Picture         =   "busquedacampshabituals.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Anar a búsqueda avançada"
      Top             =   3690
      Width           =   1065
   End
   Begin VB.CommandButton bconsultar 
      Height          =   450
      Left            =   3360
      Picture         =   "busquedacampshabituals.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Buscar comandes amb aquests criteris"
      Top             =   3690
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   5625
      Begin VB.CommandButton Command1 
         Caption         =   "Disposició materials"
         Height          =   345
         Left            =   3945
         TabIndex        =   26
         ToolTipText     =   "Comprovar la disposició de materials a comanda."
         Top             =   1635
         Width           =   1635
      End
      Begin VB.TextBox crefinplacsa 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1710
         TabIndex        =   23
         Top             =   1620
         Width           =   2115
      End
      Begin VB.TextBox crefalternativa 
         Height          =   315
         Left            =   1710
         TabIndex        =   20
         Top             =   1260
         Width           =   2130
      End
      Begin VB.TextBox ccomandaclient 
         Height          =   315
         Left            =   1710
         TabIndex        =   4
         Top             =   1950
         Width           =   2130
      End
      Begin VB.TextBox ccodiclient 
         Height          =   315
         Left            =   1710
         TabIndex        =   2
         Top             =   570
         Width           =   675
      End
      Begin VB.TextBox ctexteimpresio 
         Height          =   315
         Left            =   1710
         TabIndex        =   7
         Top             =   2985
         Width           =   3615
      End
      Begin VB.TextBox cnumtreball 
         Height          =   315
         Left            =   1710
         TabIndex        =   6
         Top             =   2640
         Width           =   810
      End
      Begin VB.TextBox ccodidebarres 
         Height          =   315
         Left            =   1710
         TabIndex        =   5
         Top             =   2295
         Width           =   2130
      End
      Begin VB.TextBox ccomanda 
         Height          =   315
         Left            =   1710
         TabIndex        =   1
         Top             =   225
         Width           =   1020
      End
      Begin VB.TextBox creferencia 
         Height          =   315
         Left            =   1710
         TabIndex        =   3
         Top             =   930
         Width           =   2115
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. Inplacsa:"
         Height          =   300
         Left            =   75
         TabIndex        =   25
         Top             =   1650
         Width           =   1425
      End
      Begin VB.Label etlikerefinplacsa 
         BackStyle       =   0  'Transparent
         Caption         =   "(utilitza * per filtrar ""like"")"
         ForeColor       =   &H00FF8080&
         Height          =   315
         Left            =   3840
         TabIndex        =   24
         Top             =   1365
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label etlikecodibarres 
         BackStyle       =   0  'Transparent
         Caption         =   "(utilitza * per filtrar ""like"")"
         ForeColor       =   &H00FF8080&
         Height          =   315
         Left            =   3855
         TabIndex        =   22
         Top             =   2340
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref Client+Alternativa:"
         Height          =   300
         Left            =   75
         TabIndex        =   21
         Top             =   1290
         Width           =   2595
      End
      Begin VB.Label etfiltrarref 
         BackStyle       =   0  'Transparent
         Caption         =   "(utilitza * per filtrar ""like"")"
         ForeColor       =   &H00FF8080&
         Height          =   315
         Left            =   3840
         TabIndex        =   19
         Top             =   945
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Comanda Client:"
         Height          =   300
         Left            =   75
         TabIndex        =   18
         Top             =   1995
         Width           =   1425
      End
      Begin VB.Label cnomdelclient 
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2400
         TabIndex        =   14
         Top             =   600
         Width           =   3120
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi del Client:"
         Height          =   300
         Left            =   75
         TabIndex        =   13
         Top             =   615
         Width           =   1425
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Texte Impresió:"
         Height          =   300
         Left            =   75
         TabIndex        =   12
         Top             =   3030
         Width           =   1425
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Treball:"
         Height          =   300
         Left            =   75
         TabIndex        =   11
         Top             =   2685
         Width           =   1425
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi de Barres:"
         Height          =   300
         Left            =   75
         TabIndex        =   10
         Top             =   2340
         Width           =   1425
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comanda Inplacsa:"
         Height          =   300
         Left            =   75
         TabIndex        =   9
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Referència Client:"
         Height          =   300
         Left            =   75
         TabIndex        =   8
         Top             =   960
         Width           =   1425
      End
   End
   Begin VB.Label etstatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscant resultats..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   225
      TabIndex        =   17
      Top             =   3765
      Visible         =   0   'False
      Width           =   3030
   End
End
Attribute VB_Name = "formbusquedahabitual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vprimercop As Boolean
Private Sub bconsultar_Click()
   Dim vsql As String
   Dim vquerydef As QueryDef
   etstatus.Visible = True
   DoEvents
   crear_consulta vsql
   If atrim(vsql) <> "" Then
     
     If InStr(1, vsql, "refinplacsa ") > 0 Then
        'vsql = "select * from comandes where comanda in (select comanda from comandesmesextres where " + vsql + ") order by comanda desc"
        vsql = "select * from comandesmesextres where " + vsql + " order by comanda desc"
         Else: vsql = "select * from comandes where " + vsql + " order by comanda desc"
     End If
     
     'Set vquerydef = formcomandes.data1.Database.CreateQueryDef("", vsql)
     'Set formcomandes.data1.Recordset = vquerydef.OpenRecordset
    formcomandes.Data1.RecordSource = vsql '"select * from comandes where " + vsql + " order by comanda desc"
'    formcomandes.Data1.Recordset.Requery
    formcomandes.Data1.Refresh
       ' EM COL.LOCO EN L'ULTIMA COMANDA PRINCIPAL SALTANT LES FULLES 2 I 3
    With formcomandes
    If Not .Data1.Recordset.EOF And Not .Data1.Recordset.BOF Then
      .Data1.Recordset.MoveLast: .Data1.Recordset.MoveFirst
      If .Data1.Recordset!producte = "PC" Or .Data1.Recordset!producte = "PC2" Or .Data1.Recordset!producte = "PCP" Then
            While .Data1.Recordset.RecordCount - 1 > .Data1.Recordset.AbsolutePosition
                .Data1.Recordset.MoveNext
                If .Data1.Recordset!producte <> "PC" And .Data1.Recordset!producte <> "PC2" And .Data1.Recordset!producte <> "PCP" Then GoTo fi
            Wend
fi:
      End If
    End If
    'If Not formcomandes.Data1.Recordset.EOF Then
    '  formcomandes.Data1.Recordset.MoveLast
    '  formcomandes.Data1.Recordset.MoveFirst
    'End If
    End With
   End If
   etstatus.Visible = False
   formbusquedahabitual.Hide
End Sub
Sub crear_consulta(vsql As String)
  creferencia = substituir(creferencia, "[", "")
  If cadbl(ccomanda) > 0 Then vsql = IIf(vsql <> "", vsql + " and (", "(") + "comanda=" + atrim(cadbl(ccomanda)) + ")"
  If cadbl(ccodiclient) > 0 Then vsql = IIf(vsql <> "", vsql + " and (", "(") + "client=" + atrim(ccodiclient) + ")"
  'If atrim(creferencia) <> "" Then vsql = IIf(vsql <> "", vsql + " and (", "(") + "refclient like '*" + treure_apostruf(atrim(creferencia)) + "*' or refclialt like '*" + treure_apostruf(atrim(creferencia)) + "*')"
  If atrim(creferencia) <> "" Then vsql = IIf(vsql <> "", vsql + " and (", "(") + "refclient like '" + treure_apostruf(atrim(creferencia)) + "') " ' or refclialt like '*" + treure_apostruf(atrim(creferencia)) + "*')"
  If atrim(creferencia) = "" And atrim(crefalternativa) <> "" Then
     vsql = IIf(vsql <> "", vsql + " and (", "(") + "refclient like '*" + treure_apostruf(atrim(crefalternativa)) + "*' or refclialt like '*" + treure_apostruf(atrim(crefalternativa)) + "*')"
  End If
  If atrim(crefinplacsa) <> "" Then vsql = IIf(vsql <> "", vsql + " and (", "(") + "refinplacsa like '" + treure_apostruf(atrim(crefinplacsa)) + "') "
  If atrim(ccodidebarres) <> "" Then
     If InStr(1, ccodidebarres, "*") Then
        vsql = IIf(vsql <> "", vsql + " and (", "(") + "codibarras like '" + treure_apostruf(atrim(ccodidebarres)) + "')"
          Else: vsql = IIf(vsql <> "", vsql + " and ", "(") + "codibarras = '" + treure_apostruf(atrim(ccodidebarres)) + "')"
     End If
  End If
  If atrim(ctexteimpresio) <> "" Then vsql = IIf(vsql <> "", vsql + " and (", "(") + "texteimpressio like '*" + atrim(ctexteimpresio) + "*' or texteimpressio like '*" + atrim(ctexteimpresio) + "*' or obsimp1 like '*" + atrim(ctexteimpresio) + "*' or marcailinia like '*" + atrim(ctexteimpresio) + "*')"
  If cadbl(cnumtreball) > 0 Then vsql = IIf(vsql <> "", vsql + " and (", "(") + "numtreball=" + atrim(cnumtreball) + " and (producte in (SELECT productes.codi From productes WHERE (((InStr(1,[productes].[ruta],'I'))>0)))))"
  If atrim(ccomandaclient) <> "" Then vsql = IIf(vsql <> "", vsql + " and (", "(") + "comandaclient like '*" + atrim(ccomandaclient) + "*')"
End Sub

Private Sub bproximabusqueda_Click()
   bproximabusqueda.Tag = "1"
   vtreballbuscatsubbusqueda = ""
   formbusquedahabitual.Hide
End Sub

Private Sub ccodiclient_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 113 Then triarclient
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 
End Sub
Sub triarclient()
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select * from clients"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   ccodiclient = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   cnomdelclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
  End If
  Unload formseleccio
End Sub

Private Sub vtexteimpresio_Change()

End Sub

Private Sub ccodiclient_LostFocus()
   Dim rstcli As Recordset
   Set rstcli = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(ccodiclient)))
   If rstcli.EOF Then
        cnomdelclient = ""
       Else: cnomdelclient = atrim(rstcli!nom)
   End If
   Set rstcli = Nothing
End Sub
Function valorcampactual() As String
   If TypeOf Screen.ActiveControl Is TextBox Then valorcampactual = atrim(Screen.ActiveControl.Text)
End Function

Private Sub ccodidebarres_GotFocus()
    etlikecodibarres.Visible = True
End Sub

Private Sub ccodidebarres_LostFocus()
    etlikecodibarres.Visible = False
End Sub

Private Sub Command1_Click()
  Dim rst As Recordset
  Dim rstd As Recordset
inici:
  Set rstd = dbtmp.OpenRecordset("select * from referencies_disposiciomaterials where nomcreador<>'" + nomordinador + "' and nomverificador='' or nomverificador=null")
  If Not rstd.EOF Then
       Set rst = dbtmp.OpenRecordset("select * from comandes_extres where refinplacsa='" + atrim(rstd!refinplacsa) + "' order by comanda desc")
       If Not rst.EOF Then
          Set rst = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rst!comanda))
          vnumc = 0
          If Not rst.EOF Then
              vnumc = rst!comanda
              If cadbl(rst!linkcomanda1) < vnumc And cadbl(rst!linkcomanda1) > 0 Then vnumc = cadbl(rst!linkcomanda1)
              If cadbl(rst!linkcomanda2) < vnumc And cadbl(rst!linkcomanda2) > 0 Then vnumc = cadbl(rst!linkcomanda2)
          End If
             Else: rstd.Delete: GoTo inici
       End If
       If vnumc > 0 Then
             formcomandes.Data1.RecordSource = "select * from comandes where comanda=" + atrim(vnumc)
             formcomandes.Data1.Refresh
             Unload formbusquedahabitual
       End If
         Else: MsgBox "No hi ha cap referencia pendent de revisar.", vbCritical, "Revisar disposició de materials"
  End If
  Set rst = Nothing
End Sub

Private Sub creferencia_GotFocus()
   etfiltrarref.Visible = True
End Sub

Private Sub creferencia_LostFocus()
   etfiltrarref.Visible = False
End Sub

Private Sub crefinplacsa_GotFocus()
  etlikerefinplacsa.Visible = True
End Sub

Private Sub crefinplacsa_LostFocus()
etlikerefinplacsa.Visible = False
End Sub

Private Sub Form_Activate()
 Dim vx As Double
  Dim vy As Double
  
  If vprimercop = False Then
    vx = cadbl(llegir_ini("PosicioFormBusquedahabitual", "Left", "comandes.ini"))
    vy = cadbl(llegir_ini("PosicioFormBusquedahabitual", "Top", "comandes.ini"))
    If vx > 0 And vy > 0 Then formbusquedahabitual.Left = vx: formbusquedahabitual.Top = vy
    vprimercop = True
'    formfirmes.Visible = True
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 112 Then bconsultar_Click
  If KeyCode = 13 Then
    If valorcampactual <> "" Then
       bconsultar_Click
      Else: SendKeys "{TAB}"
    End If
  End If
  If KeyCode = 117 Then formbusquedahabitual.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If cadbl(formbusquedahabitual.Left) > 0 Then
  escriure_ini "PosicioFormBusquedahabitual", "Left", atrim(formbusquedahabitual.Left), "comandes.ini"
  escriure_ini "PosicioFormBusquedahabitual", "Top", atrim(formbusquedahabitual.Top), "comandes.ini"
  vprimercop = False
 End If
End Sub
