VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form proveidorsproduccio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment Proveïdors Producció"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   12945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   30
      TabIndex        =   1
      Top             =   -45
      Width           =   12810
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C78DFA&
         Caption         =   "Qualitat"
         Height          =   435
         Left            =   9195
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Proveïdor relacionats amb aquest proveïdor de producció."
         Top             =   120
         Width           =   1125
      End
      Begin VB.CommandButton consultar 
         Height          =   360
         Left            =   1425
         Picture         =   "proveidors produccio.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda de Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   495
         Picture         =   "proveidors produccio.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Modificació Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   945
         Picture         =   "proveidors produccio.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   45
         Picture         =   "proveidors produccio.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "Proveidors Comercial"
         Height          =   435
         Left            =   10350
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Proveïdor relacionats amb aquest proveïdor de producció."
         Top             =   120
         Width           =   1905
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   12315
         Picture         =   "proveidors produccio.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   135
         Width           =   390
      End
      Begin VB.Label etmissatge 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2115
         TabIndex        =   7
         Top             =   240
         Width           =   3210
      End
   End
   Begin VB.Data proveidorsp 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2130
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1545
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "proveidors produccio.frx":1BB2
      Height          =   4470
      Left            =   45
      OleObjectBlob   =   "proveidors produccio.frx":1BC8
      TabIndex        =   0
      Top             =   585
      Width           =   12780
   End
End
Attribute VB_Name = "proveidorsproduccio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
  Dim nom As String
  Dim codi As Double
  nom = InputBox("Entra el nom del proveidor que vols afegir", "Entrada proveïdor")
  If atrim(nom) = "" Then Exit Sub
  If proveidorsp.Recordset.EditMode = 0 Then
      proveidorsp.Recordset.AddNew
      DoEvents
        'busco el mes gran i el poso a codi +1
        Set rsttmp = proveidorsp.Database.OpenRecordset("select max(codi) as [grancodi] from proveidors")
        If Not rsttmp.EOF Then
          codi = cadbl(rsttmp!grancodi) + 1
              Else: codi = 1
        End If
        proveidorsp.Recordset!codi = codi
        proveidorsp.Recordset!nom = nom
        proveidorsp.Recordset.Update
        proveidorsp.Refresh
        proveidorsp.Recordset.FindFirst "CODI=" + atrim(codi)
        formatreixa
     Else: MsgBox "No pots afegir si estàs editant...", vbCritical, "Atenció"
 End If
End Sub

Private Sub Command1_Click()
  If proveidorsp.Recordset.EOF Then MsgBox "Selecciona primer un proveidor", vbCritical + vbOKOnly, "Atenció": Exit Sub
  Unload proveidors
  proveidorsp.Tag = atrim(proveidorsp.Recordset!codi)
  proveidors.Show 1
  Unload proveidors
  proveidorsp.Tag = ""
End Sub

Private Sub Command2_Click()
   
   formproveidorsqualitat.Show 1
   
   
End Sub

Private Sub consultar_Click()
  Dim b As String
   b = InputBox("Entra la Descripcio a buscar o el Codi", "Busqueda")
   If cadbl(b) > 0 Then
     proveidorsp.RecordSource = "select codi,nom,tipuscq ,datacq ,aliastintes,tipusproveidorIMPOST,dataalta,databaixa,alta_desde_sap from proveidors where codi=" + atrim(cadbl(b)) + ""
     proveidorsp.Refresh
     b = ""
      Else
       If b <> "" Then
        proveidorsp.RecordSource = "select codi,nom,tipuscq ,datacq ,aliastintes,tipusproveidorIMPOST,dataalta,databaixa,alta_desde_sap from proveidors where nom like '*" + b + "*'"
        proveidorsp.Refresh
       End If
   End If
   formatreixa
End Sub

Private Sub eliminar_Click()
  Dim rstp As Recordset
  Set rstp = proveidorsp.Database.OpenRecordset("select * from proveidors_comercial where codiproduccio=" + atrim(proveidorsp.Recordset!codi))
  If Not rstp.EOF Then
      MsgBox "No pots eliminar aquest proveïdor ja que te proveidors de comercials assignats. Primer esborra els de comercial.", vbCritical + vbOKOnly, "Atenció"
      Exit Sub
    Else
       If UCase(InputBox("Segur que vols eliminar aquest proveidor." + Chr(10) + Chr(13) + "ESCRIU ELIMINAR PER FER-HO.", "Eliminar proveidor")) = "ELIMINAR" Then
           proveidorsp.Recordset.Delete
           proveidorsp.Refresh
       End If
  End If
End Sub

Private Sub Form_Load()
proveidorsp.DatabaseName = cami
proveidorsp.RecordSource = "select codi,nom,tipuscq ,datacq ,aliastintes,tipusproveidorIMPOST,dataalta,databaixa,alta_desde_sap from proveidors"
proveidorsp.Refresh
formatreixa
End Sub
Sub formatreixa()

reixa.Columns(0).Caption = "Codi"
reixa.Columns(1).Caption = "Nom Proveïdor"
reixa.Columns(2).Caption = "Tipus_CQ(L,C)"
reixa.Columns(3).Caption = "Data_Cad_CQ"
reixa.Columns(4).Caption = "Alias Tintes"
reixa.Columns(5).Caption = "ImpostEnv tipus"
reixa.Columns(6).Caption = "Data Alta"
reixa.Columns(7).Caption = "Data Baixa"
reixa.Columns(8).Caption = "Alta desde SAP"

reixa.Columns(0).Width = 600
reixa.Columns(1).Width = 2500
reixa.Columns(2).Width = 1200
reixa.Columns(3).Width = 1200
reixa.Columns(4).Width = 1000
reixa.Columns(5).Width = 1500
reixa.Columns(6).Width = 1500
reixa.Columns(7).Width = 1500
reixa.Columns(8).Width = 0
reixa.Refresh
End Sub

Private Sub modificar_Click()
Dim nom As String
  Dim codi As Double
  nom = InputBox("Entra el nom del proveidor que vols modificar", "Modificar proveïdor", proveidorsp.Recordset!nom)
  If atrim(nom) = "" Then Exit Sub
  If proveidorsp.Recordset.EditMode = 0 Then
      proveidorsp.Recordset.Edit
      codi = proveidorsp.Recordset!codi
      proveidorsp.Recordset!nom = nom
      proveidorsp.Recordset.Update
      proveidorsp.Refresh
      proveidorsp.Recordset.FindFirst "codi=" + atrim(codi)
      'reixa.Columns(0).Caption = "Codi"
'reixa.Columns(1).Caption = "Nom Proveïdor"
'reixa.Columns(2).Caption = "Alias Tintes"
'reixa.Columns(1).Width = 2500
'reixa.Columns(2).Width = 1000
     formatreixa
  End If
End Sub

Private Sub reixa_Click()
  etmissatge.Caption = "Doble clic per editar."
  etmissatge.Visible = True
End Sub
Sub escullir_tipusImpost()
   Dim v As String
   Dim v2 As String
   Dim codi As Long
   v = UCase(InputBox("Escriu el tipus de proveïdor per l'Impost d'envasos." + vbNewLine + "[No]-No aplica " + vbNewLine + "[ESP]-Estat espanyol" + vbNewLine + "[INTRA]-Intracomunitari" + vbNewLine + "[IMP]-Importació,(ADUANA)", "TIPUS DE PROVEÏDOR"))
   If v = "NO" Then v2 = "No aplica"
   If v = "ESP" Then v2 = "Espanyol"
   If v = "INTRA" Then v2 = "Intracomunitari"
   If v = "IMP" Then v2 = "Importació"
   If v2 <> "" Then
      codi = proveidorsp.Recordset!codi
      proveidorsp.Recordset.Edit
      proveidorsp.Recordset!tipusproveidorIMPOST = v2
      proveidorsp.Recordset.Update
      proveidorsp.Refresh
      proveidorsp.Recordset.FindFirst "codi=" + atrim(codi)
      formatreixa
   End If
End Sub
Sub datadebaixa()
  Dim resp As String
  Dim codi As Double
    resp = UCase(InputBox("Escriu la data de baixa del proveidor." + vbNewLine + "Escriu [Eliminar] per treure la data de baixa.", "Baixa del proveidor", Format(Now, "dd/mm/yy")))
    If resp = "" Then Exit Sub
    codi = proveidorsp.Recordset!codi
    If UCase(resp) = "ELIMINAR" Then proveidorsp.Recordset.Edit: proveidorsp.Recordset!databaixa = Null: proveidorsp.Recordset.Update: GoTo fi
    If DateDiff("d", Now, resp) >= 0 Then
        proveidorsp.Recordset.Edit
        proveidorsp.Recordset!databaixa = resp
        proveidorsp.Recordset.Update
           Else: MsgBox "Aquesta data no es vàlida com a data de baixa.", vbCritical, "Error"
    End If
fi:
    proveidorsp.Refresh
    proveidorsp.Recordset.FindFirst "codi=" + atrim(codi)
    formatreixa
End Sub
Private Sub reixa_DblClick()
   Dim codi As Long
   Dim resp As String
   If reixa.col = 1 Then modificar_Click
   If reixa.col = 5 Then escullir_tipusImpost
   If reixa.col = 7 Then datadebaixa
   If reixa.col = 4 Then
        resp = UCase(InputBox("Entra l'alias que vols fer servir a tintes només dues lletres." + Chr(10) + "Ex: IN, SK,IP...", "Modificar alias", atrim(proveidorsp.Recordset!aliastintes)))
        If resp = "" Then Exit Sub
        If Len(resp) > 2 Then MsgBox "L'alias nomes pot tenir dues lletres.", vbCritical, "Error": Exit Sub
        
        proveidorsp.Recordset.Edit
        codi = proveidorsp.Recordset!codi
        proveidorsp.Recordset!aliastintes = resp
        proveidorsp.Recordset.Update
        proveidorsp.Refresh
        proveidorsp.Recordset.FindFirst "codi=" + atrim(codi)
        'reixa.Columns(0).Caption = "Codi"
        'reixa.Columns(1).Caption = "Nom Proveïdor"
        'reixa.Columns(2).Caption = "Alias Tintes"
        'reixa.Columns(1).Width = 2500
        'reixa.Columns(2).Width = 1000
        formatreixa
   End If
   If reixa.col = 2 Then
        resp = UCase(InputBox("Entra si cal Certificat de Qualitat per cada lot (L) o bé Certificat Concertat (C)." + vbNewLine + "UNA (L) o UNA (C) O [ESPAI]", "Modificar CERTIFICAT DE QUALTIAT", atrim(proveidorsp.Recordset!tipusCQ)))
        
        If resp = "" Or (resp <> "L" And resp <> "C" And resp <> " ") Then MsgBox "Nomes pot tenir 1 lletres.", vbCritical, "Error": Exit Sub
        proveidorsp.Recordset.Edit
        codi = proveidorsp.Recordset!codi
        proveidorsp.Recordset!tipusCQ = resp
        If resp <> "C" Then proveidorsp.Recordset!dataCQ = ""
        proveidorsp.Recordset.Update
        proveidorsp.Refresh
        proveidorsp.Recordset.FindFirst "codi=" + atrim(codi)
        formatreixa
        If Not proveidorsp.Recordset.EOF Then If proveidorsp.Recordset!tipusCQ = "C" Then GoTo posardata
   End If
   If reixa.col = 3 Then
posardata:
        If proveidorsp.Recordset!tipusCQ <> "C" Then MsgBox "SI EL TIPUS DE CERTIFICAT NO ES CONCERTAT (C) NO CAL POSAR DATA", vbCritical, "ERROR": Exit Sub
        resp = UCase(InputBox("Entra la data de caducitat del CERTIFICAT.", "Modificar DATA CERTIFICAT DE QUALTIAT", atrim(proveidorsp.Recordset!dataCQ)))
        If resp = "" Then Exit Sub
        If Not IsDate(resp) And resp <> " " Then MsgBox "Aquesta data no es valida.", vbCritical, "Error": Exit Sub
        proveidorsp.Recordset.Edit
        codi = proveidorsp.Recordset!codi
        proveidorsp.Recordset!dataCQ = Format(resp, "dd/mm/yy")
        proveidorsp.Recordset.Update
        proveidorsp.Refresh
        proveidorsp.Recordset.FindFirst "codi=" + atrim(codi)
        formatreixa
   End If

End Sub

Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If Not proveidorsp.Recordset.EOF Then
      If proveidorsp.Recordset!alta_desde_sap Then
          etmissatge = "Alta del SAP"
            'Else: etmissatge = ""
      End If
   End If
End Sub

Private Sub sortir_Click()
  Unload Me
End Sub
