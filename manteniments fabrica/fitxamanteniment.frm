VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form fitxamanteniment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fitxa del Manteniment"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9345
   ControlBox      =   0   'False
   Icon            =   "fitxamanteniment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data dataavisos 
      Caption         =   "avisos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "horarismanteniments"
      Top             =   2505
      Visible         =   0   'False
      Width           =   1545
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "fitxamanteniment.frx":058A
      Height          =   3135
      Left            =   105
      OleObjectBlob   =   "fitxamanteniment.frx":059F
      TabIndex        =   13
      Top             =   3180
      Width           =   9135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informació del Manteniment"
      Height          =   2835
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   9210
      Begin VB.CommandButton Command2 
         Height          =   390
         Left            =   8295
         Picture         =   "fitxamanteniment.frx":12F8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Passar avís a fet."
         Top             =   2280
         Width           =   765
      End
      Begin VB.TextBox nomoperari 
         Height          =   285
         Left            =   2370
         TabIndex        =   10
         Top             =   2370
         Width           =   5820
      End
      Begin VB.TextBox dataexecucio 
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   2355
         Width           =   1185
      End
      Begin VB.TextBox observacio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2385
         TabIndex        =   6
         Top             =   1455
         Width           =   6690
      End
      Begin VB.ComboBox dataavis 
         Height          =   315
         Left            =   375
         TabIndex        =   4
         Top             =   1440
         Width           =   1320
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   7680
         Picture         =   "fitxamanteniment.frx":1882
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Visualitzar/Imprimir"
         Top             =   165
         Width           =   930
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   8745
         Picture         =   "fitxamanteniment.frx":1E0C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Sortir"
         Top             =   150
         Width           =   390
      End
      Begin VB.Image imatgenofet 
         Height          =   735
         Left            =   6765
         Picture         =   "fitxamanteniment.frx":2396
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Image imatgefet 
         Height          =   735
         Left            =   6765
         Picture         =   "fitxamanteniment.frx":79AB
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "(*) Nom Operari"
         Height          =   195
         Left            =   2460
         TabIndex        =   11
         Top             =   2175
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "(*) Data Execució"
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Observació"
         Height          =   195
         Left            =   2565
         TabIndex        =   7
         Top             =   1230
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "(*) Data Avís"
         Height          =   195
         Left            =   570
         TabIndex        =   5
         Top             =   1215
         Width           =   1155
      End
      Begin VB.Label descripciomanteniment 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   165
         TabIndex        =   3
         Top             =   195
         Width           =   6585
      End
   End
End
Attribute VB_Name = "fitxamanteniment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  manteniments.imprimirfitxa cadbl(nummanteniment)
End Sub

Private Sub Command2_Click()
  If Not comprovarcamps Then
     MsgBox "Falta emplenar camps basics. (*) " + Chr(10) + "Potser hi ha error en alguna data.", vbExclamation + vbOKOnly, "Atenció"
    Exit Sub
  End If
  
  guardar_avis nummanteniment, dataavis
  carregar_manteniment nummanteniment
  dataavisos.Refresh
End Sub
Function comprovarcamps() As Boolean
   comprovarcamps = True
   If Not IsDate(dataavis) Then comprovarcamps = False
   If Not IsDate(dataexecucio) Then
      comprovarcamps = False
       Else
          If DateDiff("d", dataexecucio, Now) > 30 Then comprovarcamps = False
   End If
   If atrim(nomoperari) = "" Then comprovarcamps = False
End Function
Sub guardar_avis(numm As Long, dataa As Date)
   Dim rstm As Recordset
   Set rstm = dbmanteniments.OpenRecordset("select * from horarismanteniments where idmanteniment=" + atrim(numm) + " and data=#" + Format(dataa, "mm/dd/yy") + "# order by data asc")
   If Not rstm.EOF Then
        rstm.Edit
        rstm!data = Format(dataavis, "dd/mm/yy")
        rstm!observacio = atrim(observacio)
        rstm!dataexecucio = Format(dataexecucio, "dd/mm/yy")
        rstm!nomoperari = atrim(nomoperari)
        rstm.Update
      Else: MsgBox "Error guardant l'avís", vbCritical, "Error"
   End If
   Set rstm = Nothing
End Sub

Private Sub Command3_Click()

End Sub

Private Sub dataavis_Click()
    carregar_manteniment nummanteniment, dataavis.Text
End Sub

Private Sub Form_Load()
  dataavisos.DatabaseName = manteniments.datamanteniments.DatabaseName
  dataavisos.RecordSource = "select * from horarismanteniments where idmanteniment=" + atrim(nummanteniment) + " and (nomoperari<>null and nomoperari<>'_') order by data Desc"
  carregar_manteniment nummanteniment
  If dataavis.ListCount > 0 Then
     dataavis.ListIndex = 0
     carregar_manteniment nummanteniment, dataavis.Text
  End If
End Sub

Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If Screen.ActiveControl.Name = "reixa" Then
        nummanteniment = dataavisos.Recordset!idmanteniment
        carregar_manteniment nummanteniment, reixa.Columns("Data")
   End If
End Sub

Private Sub sortir_Click()
   If manteniments.Tag = "extern" Then
         End
      Else: Unload Me
   End If
End Sub
Sub carregar_manteniment(numm As Long, Optional dataacarregar As Date)
    Dim rstm As Recordset
    Set rstm = dbmanteniments.OpenRecordset("select descripcio from manteniments where id=" + atrim(numm))
    If rstm.EOF Then Exit Sub
    descripciomanteniment = UCase$(rstm!descripcio)
    
    Set rstm = dbmanteniments.OpenRecordset("select * from horarismanteniments where idmanteniment=" + atrim(numm) + " order by data asc")
    
    
    estatavis "CAP"
    observacio = ""
    dataexecucio = ""
    nomoperari = ""
    If rstm.EOF Then Exit Sub
    If dataacarregar <> "0:00:00" Then GoTo norefrescar
    dataavis.Clear
    While Not rstm.EOF
        If Not IsDate(rstm!dataexecucio) Then
           If IsDate(rstm!data) Then dataavis.AddItem Format(rstm!data, "dd/mm/yy")
        End If
        rstm.MoveNext
    Wend
    rstm.MoveFirst
norefrescar:
    If dataacarregar <> "0:00:00" Then
        rstm.FindFirst "data=#" + Format(dataacarregar, "mm/dd/yy") + "# and year(dataexecucio)>2014"
        dataavis.Text = Format(dataacarregar, "dd/mm/yy")
        If rstm.NoMatch Then GoTo sortir
    End If
    dataavis = Format(rstm!data, "dd/mm/yy")
    observacio = atrim(rstm!observacio)
    dataexecucio = Format(rstm!dataexecucio, "dd/mm/yy")
    nomoperari = atrim(rstm!nomoperari)
    
sortir:
    If dataexecucio = "" Then
       dataexecucio = Format(Now, "dd/mm/yy")
       estatavis "PENDENT"
        Else: estatavis "FET"
    End If
    
End Sub
Sub estatavis(estat As String)
   imatgefet.Visible = False
   imatgenofet.Visible = False
   If estat = "PENDENT" Then imatgenofet.Visible = True
   If estat = "FET" Then imatgefet.Visible = True
End Sub
