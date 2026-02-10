VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form avisosxrseccio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avisos de Manteniments"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8925
   Icon            =   "avisosperseccio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data datamanteniments 
      Caption         =   "datamanteniments"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3330
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Consulta1"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1965
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "avisosperseccio.frx":058A
      Height          =   3900
      Left            =   120
      OleObjectBlob   =   "avisosperseccio.frx":05A5
      TabIndex        =   8
      Top             =   1155
      Width           =   8670
   End
   Begin VB.Frame Frame1 
      Height          =   930
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   8700
      Begin VB.CommandButton Command8 
         Height          =   375
         Left            =   7590
         Picture         =   "avisosperseccio.frx":0F9E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir o editar l'avís"
         Top             =   345
         Width           =   720
      End
      Begin VB.TextBox datafi 
         Height          =   285
         Left            =   1425
         TabIndex        =   10
         Top             =   390
         Width           =   1035
      End
      Begin VB.TextBox datainici 
         Height          =   285
         Left            =   300
         TabIndex        =   9
         Top             =   390
         Width           =   1035
      End
      Begin VB.CommandButton buscar 
         Height          =   360
         Left            =   6780
         Picture         =   "avisosperseccio.frx":1528
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   690
      End
      Begin VB.ComboBox seccio 
         Height          =   315
         ItemData        =   "avisosperseccio.frx":1AB2
         Left            =   2805
         List            =   "avisosperseccio.frx":1ACE
         TabIndex        =   2
         Top             =   375
         Width           =   1320
      End
      Begin VB.ComboBox nommaquina 
         Height          =   315
         ItemData        =   "avisosperseccio.frx":1B2B
         Left            =   4185
         List            =   "avisosperseccio.frx":1B2D
         TabIndex        =   1
         Top             =   375
         Width           =   2445
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inici"
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   6
         Top             =   165
         Width           =   1290
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Secció"
         Height          =   255
         Index           =   6
         Left            =   2910
         TabIndex        =   5
         Top             =   135
         Width           =   570
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Màquina"
         Height          =   255
         Index           =   7
         Left            =   4455
         TabIndex        =   4
         Top             =   150
         Width           =   1695
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Fi"
         Height          =   255
         Index           =   8
         Left            =   1635
         TabIndex        =   3
         Top             =   180
         Width           =   840
      End
   End
End
Attribute VB_Name = "avisosxrseccio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub buscar_Click()
   buscaravisos
End Sub
Sub buscaravisos()
   datamanteniments.RecordSource = generarconsulta
   datamanteniments.Refresh
End Sub

Private Sub Command8_Click()
   If Not datamanteniments.Recordset.EOF Then
      Shell "\\serverprodu\dades\progcomandes\aplicacio\Manteniments Fàbrica.exe " + atrim(datamanteniments.Recordset!idmanteniment)
   End If
End Sub

Private Sub Form_Activate()
  ' buscaravisos
End Sub
Function generarconsulta() As String
  generarconsulta = "SELECT manteniments.descripcio, horarismanteniments.* FROM manteniments INNER JOIN horarismanteniments ON manteniments.id = horarismanteniments.idmanteniment "
  generarconsulta = generarconsulta + " where (data>=#" + Format(datainici, "mm/dd/yy") + "# and data<=#" + Format(datafi, "mm/dd/yy") + "#)"
  If seccio <> "" And seccio <> "Totes" Then generarconsulta = generarconsulta + " and seccio='" + atrim(seccio) + "' "
  If nommaquina <> "" Then generarconsulta = generarconsulta + " and maquina=" + atrim(cadbl(nommaquina.Tag))
  generarconsulta = generarconsulta + " and (horarismanteniments.nomoperari=null or horarismanteniments.nomoperari='' or horarismanteniments.nomoperari='_') and not inactiu order by data asc"
End Function
Private Sub Form_Load()

   datamanteniments.DatabaseName = rutadelfitxer(cami) + "mantenimentsfabrica.mdb"
   datamanteniments.RecordSource = "SELECT manteniments.descripcio, horarismanteniments.* FROM manteniments INNER JOIN horarismanteniments ON manteniments.id = horarismanteniments.idmanteniment;"
   datainici = Format(primerdiasetmana, "dd/mm/yy")
   datafi = Format(ultimdiasetmana, "dd/mm/yy")
End Sub
Function primerdiasetmana() As Date
    Dim avui As Byte
    avui = Format(Now, "w")
    If avui = 1 Then avui = 8
    avui = avui - 2
    primerdiasetmana = DateAdd("d", avui * -1, Now)
End Function
Function ultimdiasetmana() As Date
   
    ultimdiasetmana = DateAdd("d", 6, primerdiasetmana)
    
End Function
Private Sub nommaquina_DropDown()
 Load formseleccio
    nommaquina.Tag = ""
   formseleccio.Data1.DatabaseName = camicomandes
   formseleccio.Data1.RecordSource = "select codi,descripcio from maquines where maquina='" + Mid(seccio, 1, 1) + "' and donadadebaixa=null "
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           nommaquina = formseleccio.DBGrid2.Columns("descripcio")
           nommaquina.Tag = formseleccio.DBGrid2.Columns("codi")
        End If
   End If
    If seleccioret = 9 Then
        nommaquina = ""
        nommaquina.Tag = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub reixa_DblClick()
    Command8_Click
End Sub

Private Sub seccio_Click()
   If seccio = "Totes" Then
      nommaquina = ""
      nommaquina.Tag = ""
   End If
End Sub
