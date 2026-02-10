VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Llistatconsums 
   Caption         =   "Llistat de consums"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\temporalconsums.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6270
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"Llistathistoricconsums.frx":0000
      Top             =   3060
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Llistathistoricconsums.frx":0138
      Height          =   2055
      Left            =   180
      OleObjectBlob   =   "Llistathistoricconsums.frx":0148
      TabIndex        =   12
      Top             =   3705
      Width           =   7830
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acceptar"
      Height          =   465
      Left            =   5265
      TabIndex        =   5
      Tag             =   "0"
      Top             =   1680
      Width           =   1050
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   5415
      Top             =   2250
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      PrintFileType   =   19
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dates"
      Height          =   1245
      Left            =   510
      TabIndex        =   0
      Top             =   225
      Width           =   5805
      Begin VB.TextBox clients 
         Height          =   330
         Left            =   975
         TabIndex        =   11
         ToolTipText     =   "Entra els codis de clients separats per comes. ex: 2345,4532,286"
         Top             =   825
         Width           =   4665
      End
      Begin MSMask.MaskEdBox inici 
         Height          =   330
         Left            =   885
         TabIndex        =   3
         Top             =   195
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   582
         _Version        =   327680
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox fi 
         Height          =   330
         Left            =   2295
         TabIndex        =   4
         Top             =   195
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   582
         _Version        =   327680
         Format          =   "dd/mm/yy"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Clients:"
         Height          =   240
         Left            =   285
         TabIndex        =   7
         Top             =   855
         Width           =   555
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fi:"
         Height          =   240
         Left            =   2070
         TabIndex        =   2
         Top             =   285
         Width           =   330
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inici:"
         Height          =   240
         Left            =   300
         TabIndex        =   1
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.Frame Escullir 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Escullir"
      Height          =   2145
      Left            =   495
      TabIndex        =   8
      Top             =   1560
      Width           =   4725
      Begin VB.ListBox triar 
         Height          =   1425
         Left            =   135
         MultiSelect     =   1  'Simple
         TabIndex        =   9
         Top             =   600
         Width           =   4290
      End
      Begin VB.Label titoltriar 
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
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   210
         Width           =   4380
      End
   End
   Begin VB.Label maquines 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "Llistatconsums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub acceptar_Click()

End Sub

Sub prepararllistaclients()
 s = clients
 r = ""
 While InStr(1, s, ",") <> 0
    If r = "" Then
       r = "client="
         Else: r = r + " or client="
    End If
    r = r + atrim(cadbl(Mid(s, 1, InStr(1, s, ",") - 1)))
    s = Mid(s, InStr(1, s, ",") + 1, Len(s))
    If InStr(1, s, ",") = 0 Then r = r + " or client=" + atrim(cadbl(s))
 Wend
 If s <> "" Then r = " client=" + atrim(cadbl(s))
 If r <> "" Then r = "(" + r + ")"
End Sub

Private Sub Command1_Click()
  Dim db As Database
  Dim db2 As Database
  Dim taulatemp As String
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim rst3 As Recordset
  Dim rst4 As Recordset
  taulatemp = "c:\temporalconsums.mdb"
  Frame1.Enabled = True
  'ratoli "espera"
  If Command1.Tag = "0" Then
   If existeix(taulatemp) Then Kill taulatemp
   If Not existeix(taulatemp) Then DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
 '  Set dbtmp = DBEngine.OpenDatabase(taulatemp)
   Set db = DBEngine.OpenDatabase(cami)
   r = ""
   prepararllistaclients
   If r = "" Then r = " client>0 "
   db.Execute "select *,' ' as nomclient,0 as metroscomanda,0 as kiloscomanda,' ' as mesuraespes ,' ' as familiamat ,' ' as descmat, ' ' as desccol, ' ' as descaditiu, ' ' as familiacol, ' ' as descmesespesor into consums in '" + taulatemp + "' from comandes where " + r + " and datacomanda>#" + Str(CVDate(inici)) + "# and datacomanda<#" + Str(CVDate(fi)) + "# "
   Set rst = db.OpenRecordset("select * from consums in '" + taulatemp + "'")
   'actualitzo les families de les comandes triades
   While Not rst.EOF
      Set rst2 = db.OpenRecordset("select * from materials inner join familiesmaterials on  materials.familia=familiesmaterials.codi")
      Set rst3 = db.OpenRecordset("select * from colorants inner join familiescolorants on  colorants.familia=familiescolorants.codi")
      Set rst4 = db.OpenRecordset("select * from mesureslineals")
      rst.Edit
      rst2.FindFirst "materials.codi=" + atrim(cadbl(rst!materialex))
      rst3.FindFirst "colorants.codi=" + atrim(cadbl(rst!colorex))
      rst4.FindFirst "codi=" + atrim(cadbl(rst!mesuraesp))
      If Not rst2.NoMatch Then rst!familiamat = rst2![familiesmaterials.descripcio]: rst!descmat = rst2![materials.descripcio]
      If Not rst3.NoMatch Then rst!familiacol = rst3![familiescolorants.descripcio]: rst!desccol = rst3![colorants.descripcio]
      If Not rst4.NoMatch Then rst!descmesespesor = atrim(cadbl(rst!espessor)) + " " + atrim(rst4!descripcio)
      rst.Update
      rst.MoveNext
   Wend
   triar.Clear
   Escullir.Visible = True
   titoltriar.Caption = "Escullir Familia Material"
   Set rst = db.OpenRecordset("select distinct familiamat from consums in '" + taulatemp + "' ")
   While Not rst.EOF
     triar.AddItem atrim(rst!familiamat)
     rst.MoveNext
   Wend
   Command1.Tag = "1"
   Frame1.Enabled = False
   Exit Sub
  End If
  If Command1.Tag = "1" Then
   'trio les families de material
   i = 0
   r = ""
   s = ""
   Set db2 = DBEngine.OpenDatabase(taulatemp)
   While i < triar.ListCount
     If triar.Selected(i) Then
        If r = "" Then
          r = " familiamat<>'" + atrim(triar.List(i)) + "'"
         Else: r = r + " and familiamat<>'" + atrim(triar.List(i)) + "'"
        End If
     End If
     i = i + 1
   Wend
   If r <> "" Then db2.Execute "DELETE * FROM consums where " + r ' un cop borrats els triats ensenyo els colorants
   triar.Clear
   Escullir.Visible = True
   titoltriar.Caption = "Escullir Familia Colorants"
   Set db = DBEngine.OpenDatabase(cami)
   Set rst = db.OpenRecordset("select distinct familiacol from consums in '" + taulatemp + "' ")
   While Not rst.EOF
     triar.AddItem atrim(rst!familiacol)
     rst.MoveNext
   Wend
   
   
   Command1.Tag = "2"
   Exit Sub
  End If
  If Command1.Tag = "2" Then
    'trio les families de colorant
    'trio les families de material
   i = 0
   r = ""
   s = ""
   Set db2 = DBEngine.OpenDatabase(taulatemp)
   While i < triar.ListCount
     If triar.Selected(i) Then
        If r = "" Then
          r = " familiacol<>'" + atrim(triar.List(i)) + "'"
         Else: r = r + " and familiacol<>'" + atrim(triar.List(i)) + "'"
        End If
     End If
     i = i + 1
   Wend
   If r <> "" Then db2.Execute "DELETE * FROM consums where " + r
   
   'carrego els espesors
   triar.Clear
   Escullir.Visible = True
   titoltriar.Caption = "Escullir els espessors desitjats."
   Set db = DBEngine.OpenDatabase(cami)
   Set rst = db.OpenRecordset("select distinct descmesespesor from consums in '" + taulatemp + "' order by descmesespesor")
   While Not rst.EOF
     triar.AddItem atrim(rst!descmesespesor)
     rst.MoveNext
   Wend
   
    Command1.Tag = "3"
    Exit Sub
  End If
  
  If Command1.Tag = "3" Then
    'trio els espessors
   i = 0
   r = ""
   s = ""
   Set db2 = DBEngine.OpenDatabase(taulatemp)
   While i < triar.ListCount
     If triar.Selected(i) Then
        If r = "" Then
          r = " descmesespesor<>'" + atrim(triar.List(i)) + "'"
         Else: r = r + " and descmesespesor<>'" + atrim(triar.List(i)) + "'"
        End If
     End If
     i = i + 1
   Wend
   If r <> "" Then db2.Execute "DELETE * FROM consums where " + r
   Command1.Tag = "4"
  End If
  
  'recorro la taula per possar els valors que em faltes
  Set db = DBEngine.OpenDatabase(cami)
  Set rst = db.OpenRecordset("select * from consums in '" + taulatemp + "'")
   'actualitzo les families de les comandes triades
   Set rst2 = db.OpenRecordset("select * from mesureslineals")
   Set rst3 = db.OpenRecordset("select * from aditius")
   Set rst4 = db.OpenRecordset("select codi,nom from clients")
   While Not rst.EOF
      rst.Edit
      rst2.FindFirst "codi=" + atrim(cadbl(rst!mesuracantex))
      rst3.FindFirst "codi=" + atrim(cadbl(rst!aditiuex))
      rst4.FindFirst "codi=" + atrim(cadbl(rst!client))
      If Not rst2.NoMatch Then rst!mesuraespes = atrim(rst2!descripcio)
      If Not rst3.NoMatch Then rst!descaditiu = atrim(rst3!descripcio)
      If Not rst4.NoMatch Then rst!nomclient = atrim(rst4!nom)
      On Error Resume Next
      If InStr(1, rst!mesuraespes, "KGRS") > 0 Then
          rst!metroscomanda = (cadbl(rst!cantitatex) / cadbl(rts!pes1000mtrs)) * 1000
          rst!kiloscomanda = cadbl(rst!cantitatex)
        Else
            rst!kiloscomanda = (cadbl(rst!cantitatex) / 1000) * cadbl(rst!pes1000mtrs)
            rst!metroscomanda = cadbl(rst!cantitatex)
      End If
      On Error GoTo 0
      rst.Update
      rst.MoveNext
   Wend
   
 Data1.RecordSource = Data1.Tag
 Data1.DatabaseName = "c:\temporalconsums.mdb"
 Data1.Refresh
 DBGrid1.Refresh
 DBGrid1.Visible = True
End Sub

Sub imprimirinforme()
r = "LListat entre dates: " + atrim(inici) + " i " + atrim(fi) + "       Clients: " + atrim(clients)
  llistat.DataFiles(0) = "c:\tmp.mdb"
  llistat.WindowState = crptMaximized
  llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "llistatconsums.rpt"
  llistat.Formulas(0) = "titolinforme=" + "'" + r + "'"
  llistat.Action = 1
End Sub

Private Sub Command2_Click()

End Sub

Private Sub DBGrid1_DblClick()
Dim db As Database
 Dim taulatemp As String
 taulatemp = "c:\tmp.mdb"
 If existeix(taulatemp) Then Kill taulatemp
 If Not existeix(taulatemp) Then DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
 Set db = DBEngine.OpenDatabase(Data1.DatabaseName)
 r = "select * into consums in '" + taulatemp + "'   from consums where consums.familiamat='" + (Data1.Recordset![familiamat]) + "' and consums.familiacol='" + (Data1.Recordset![familiacol]) + "'  and descmesespesor='" + ((Data1.Recordset!descmesespesor)) + "' and cdbl(consums.ampleesq)='" + atrim(cadbl(Data1.Recordset![ampleesq])) + "'"
 db.Execute r
 imprimirinforme
End Sub

Private Sub Form_Click()
' Data1.RecordSource = Data1.Tag
' Data1.DatabaseName = "c:\temporalconsums.mdb"
' Data1.Refresh
' DBGrid1.Refresh
 'DBGrid1.Visible = True
End Sub

Private Sub Form_Load()
 Data1.Tag = Data1.RecordSource
 Data1.RecordSource = "SELECT consums.familiamat, consums.familiacol, consums.descmesespesor, consums.ampleesq, Sum(consums.metroscomanda) AS Mtrs, Sum(consums.kiloscomanda) AS Kg, Count(*) AS Comandes, [Kg]/[Comandes] AS ratio FROM consums GROUP BY consums.familiamat, consums.familiacol, consums.descmesespesor, consums.ampleesq;"
 Data1.DatabaseName = ""
 Data1.Refresh
 DBGrid1.Visible = False
 DBGrid1.Refresh
 inici = "01/01/" + Format(DateAdd("yyyy", -1, Now), "yy")
 fi = "31/12/" + Format(DateAdd("yyyy", -1, Now), "yy")
 'clients = "6657,4811,6721,1903,2326,6425,6670"
 Escullir.Visible = False
 
End Sub


Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

