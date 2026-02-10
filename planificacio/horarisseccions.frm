VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form horarismaquines 
   Caption         =   "Programació del horaris de les màquines"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   Icon            =   "horarisseccions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7755
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   8400
      Begin VB.Data dataescalatsreb 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5550
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "escalatvelocitatsreb"
         Top             =   570
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data datacanvisreb 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "canvimaquinesreb"
         Top             =   990
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame tempscanvisreb 
         Caption         =   "Temps canvis rebobinadora"
         Height          =   6390
         Left            =   2295
         TabIndex        =   22
         Top             =   1275
         Visible         =   0   'False
         Width           =   2910
         Begin MSDBGrid.DBGrid reixacanvisreb 
            Bindings        =   "horarisseccions.frx":058A
            Height          =   2820
            Left            =   105
            OleObjectBlob   =   "horarisseccions.frx":05A2
            TabIndex        =   23
            Top             =   210
            Width           =   2685
         End
         Begin MSDBGrid.DBGrid reixaescalatsreb 
            Bindings        =   "horarisseccions.frx":113C
            Height          =   3270
            Left            =   45
            OleObjectBlob   =   "horarisseccions.frx":1156
            TabIndex        =   24
            Top             =   3060
            Width           =   2805
         End
      End
      Begin VB.Data datacanvis 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   5520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   930
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame tempscanvisbasic 
         Caption         =   "Temps de canvis"
         Height          =   3615
         Left            =   5385
         TabIndex        =   19
         Top             =   1230
         Width           =   2895
         Begin MSDBGrid.DBGrid reixacanvis 
            Bindings        =   "horarisseccions.frx":1CEE
            Height          =   3195
            Left            =   60
            OleObjectBlob   =   "horarisseccions.frx":1D03
            TabIndex        =   20
            Top             =   315
            Width           =   2685
         End
      End
      Begin VB.ComboBox seccio 
         Height          =   315
         ItemData        =   "horarisseccions.frx":2A4E
         Left            =   195
         List            =   "horarisseccions.frx":2A5E
         TabIndex        =   16
         Top             =   450
         Width           =   1695
      End
      Begin VB.ComboBox nommaquina 
         Height          =   315
         Left            =   2175
         TabIndex        =   15
         Top             =   450
         Width           =   3000
      End
      Begin VB.TextBox diaescullit 
         Height          =   330
         Left            =   2430
         TabIndex        =   14
         Top             =   960
         Width           =   1275
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Copiar a tot l'any"
         Height          =   510
         Left            =   6975
         Picture         =   "horarisseccions.frx":2A96
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   705
         Width           =   1305
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   510
         Left            =   6990
         Picture         =   "horarisseccions.frx":3020
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   195
         Width           =   1290
      End
      Begin VB.Frame Frame4 
         Caption         =   "Data"
         Height          =   4275
         Left            =   135
         TabIndex        =   3
         Top             =   1245
         Width           =   3600
         Begin VB.TextBox tany 
            Height          =   315
            Left            =   1650
            TabIndex        =   8
            Top             =   405
            Width           =   870
         End
         Begin VB.ComboBox cmesos 
            Height          =   315
            ItemData        =   "horarisseccions.frx":35AA
            Left            =   120
            List            =   "horarisseccions.frx":35D2
            TabIndex        =   7
            Top             =   405
            Width           =   1365
         End
         Begin VB.Frame fmes 
            Height          =   3510
            Left            =   90
            TabIndex        =   4
            Top             =   675
            Width           =   3405
            Begin VB.Label dies 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   660
               TabIndex        =   6
               Top             =   615
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label Label1 
               BackColor       =   &H00D29F7D&
               Caption         =   "  Dl    Dm    Dx    Dj    Dv   Ds   Dm"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   90
               TabIndex        =   5
               Top             =   195
               Width           =   3195
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'Transparent
               FillColor       =   &H00D29F7D&
               FillStyle       =   0  'Solid
               Height          =   2775
               Left            =   2400
               Top             =   450
               Width           =   900
            End
         End
         Begin VB.Label Label2 
            Caption         =   "Mes"
            Height          =   165
            Left            =   585
            TabIndex        =   10
            Top             =   210
            Width           =   780
         End
         Begin VB.Label Label3 
            Caption         =   "Any"
            Height          =   225
            Left            =   1905
            TabIndex        =   9
            Top             =   195
            Width           =   780
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Hores"
         Height          =   6405
         Left            =   3810
         TabIndex        =   1
         Top             =   1230
         Width           =   1440
         Begin VB.CheckBox Check1 
            Caption         =   "Totes"
            Height          =   195
            Left            =   90
            TabIndex        =   21
            Top             =   225
            Width           =   1080
         End
         Begin VB.ListBox llistadhores 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5820
            ItemData        =   "horarisseccions.frx":3634
            Left            =   60
            List            =   "horarisseccions.frx":363E
            Style           =   1  'Checkbox
            TabIndex        =   2
            Top             =   435
            Width           =   1305
         End
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Secció"
         Height          =   270
         Left            =   240
         TabIndex        =   18
         Top             =   195
         Width           =   1470
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom Màquina"
         Height          =   270
         Left            =   2535
         TabIndex        =   17
         Top             =   210
         Width           =   1470
      End
      Begin VB.Label Label6 
         Caption         =   "Dia:"
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   1005
         Width           =   375
      End
   End
End
Attribute VB_Name = "horarismaquines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
  Dim i As Byte
  Dim valor As Boolean
  valor = Check1.Value
  For i = 0 To llistadhores.ListCount - 1
     llistadhores.Selected(i) = valor
  Next i
End Sub

Private Sub cmesos_Click()
carregar_mes cmesos.ListIndex + 1, cadbl(tany)
End Sub



Private Sub Command1_Click()
  gravar_horari
End Sub
Sub gravar_horari()
  Dim datasql As String
  If seccio = "" Or nommaquina = "" Then MsgBox "Escull seccio i nom de la màquina.": Exit Sub
  datasql = " (year(dataihora)=" + atrim(Year(diaescullit)) + " and month(dataihora)=" + atrim(Month(diaescullit)) + " and day(dataihora)=" + atrim(Day(diaescullit)) + ")"
  dbplanificacio.Execute "delete * from horaris" + seccio + " where maquina=" + atrim(nommaquina.ItemData(nommaquina.ListIndex)) + " and " + datasql
  
  For i = 0 To 23
    If llistadhores.Selected(i) Then
      dbplanificacio.Execute "insert into horaris" + seccio + " (maquina,dataihora) values (" + atrim(nommaquina.ItemData(nommaquina.ListIndex)) + ",#" & Format(diaescullit, "mm/dd/yyyy " + Format(i, "00") + ":00:00") & "#) "
    End If
  Next i
End Sub

Private Sub Command2_Click()
  If seccio = "" Or nommaquina = "" Then MsgBox "Primer escull la seccio i nom de la maquina": Exit Sub
  If MsgBox("Aquesta funció copiarà la programació de la setmana actual a LA RESTA de l'any SOBREESCRIVINT tota la programació que ja hi hagi.", vbCritical + vbYesNo + vbDefaultButton2, "MOLTA ATENCIÓ") = vbYes Then
       ratoli "espera"
       horarismaquines.Enabled = False
       copiarsetmanaactualalrestadelany diaescullit
       ratoli "normal"
       horarismaquines.Enabled = True
  End If
End Sub
Sub copiarsetmanaactualalrestadelany(ByVal dia As Date)
    Dim i As Byte
    Dim datasql As String
    Dim rst As Recordset
    Dim diadesti As Date
    Dim diesset As Variant
    Dim caption As String
    diesset = Array("Dilluns", "Dimarts", "Dimecres", "Dijous", "Divendres", "Dissabtes", "Diumenges")
    While DatePart("w", dia, vbMonday) > 1
       dia = DateAdd("d", -1, dia)
    Wend
    diadesti = DateAdd("d", 7, dia)
    datasql = " (year(dataihora)=" + atrim(Year(diadesti)) + " and dataihora>=#" + Format(diadesti, "mm/dd/yyyy") + "#)"
    dbplanificacio.Execute "delete * from horaris" + seccio + " where maquina=" + atrim(nommaquina.ItemData(nommaquina.ListIndex)) + " and " + datasql
    caption = Me.caption
    For i = 1 To 7
       Me.caption = "Programant tots els " + diesset(i - 1)
       DoEvents
       copiardiarestadany (dia)
       dia = DateAdd("d", 1, dia)
    Next i
    Me.caption = caption
End Sub
Sub copiardiarestadany(ByVal dia As Date)
     Dim datasql As String
     Dim rst As Recordset
     Dim diadesti As Date
    datasql = " (year(dataihora)=" + atrim(Year(dia)) + " and month(dataihora)=" + atrim(Month(dia)) + " and day(dataihora)=" + atrim(Day(dia)) + ")"
    Set rst = dbplanificacio.OpenRecordset("select * from horaris" + seccio + " where maquina=" + atrim(nommaquina.ItemData(nommaquina.ListIndex)) + " and " + datasql + " order by dataihora asc")
    diadesti = DateAdd("d", 7, dia)
    If Not rst.EOF Then
       While Year(dia) = Year(diadesti)
         rst.MoveFirst
         clonarprogramaciodeldia dia, diadesti, rst
         diadesti = DateAdd("d", 7, diadesti)
         
         
       Wend
    End If
End Sub
Sub clonarprogramaciodeldia(dia As Date, diadesti As Date, rst As Recordset)
    While Not rst.EOF
       dbplanificacio.Execute "insert into horaris" + seccio + " (maquina,dataihora) values (" + atrim(rst!maquina) + ",#" & Format(diadesti, "mm/dd/yyyy " + Format(Format(rst!dataihora, "h"), "00") + ":00:00") & "#) "
       rst.MoveNext
    Wend
    
End Sub

Private Sub datacanvisreb_Reposition()
  If datacanvisreb.Recordset.EOF Then Exit Sub
  dataescalatsreb.RecordSource = "select * from escalatvelocitatsreb where idcanvireb=" + atrim(datacanvisreb.Recordset!id)
  dataescalatsreb.Refresh
  If dataescalatsreb.Recordset.EOF Then crearescalat
End Sub
Sub crearescalat()
   Dim id As Long
   id = datacanvisreb.Recordset!id
   With dataescalatsreb.Recordset
   .AddNew: !idcanvireb = id: !metres = 200: .Update
   .AddNew: !idcanvireb = id: !metres = 300: .Update
   .AddNew: !idcanvireb = id: !metres = 400: .Update
   .AddNew: !idcanvireb = id: !metres = 500: .Update
   .AddNew: !idcanvireb = id: !metres = 750: .Update
   .AddNew: !idcanvireb = id: !metres = 1000: .Update
   .AddNew: !idcanvireb = id: !metres = 1500: .Update
   .AddNew: !idcanvireb = id: !metres = 2000: .Update
   .AddNew: !idcanvireb = id: !metres = 2500: .Update
   .AddNew: !idcanvireb = id: !metres = 3000: .Update
   .AddNew: !idcanvireb = id: !metres = 4000: .Update
   .AddNew: !idcanvireb = id: !metres = 6000: .Update
   .AddNew: !idcanvireb = id: !metres = 10000: .Update
   End With
End Sub
Private Sub diaescullit_Change()
  If IsDate(diaescullit) And diaescullit.Tag = "" Then carregarhorari (diaescullit)
End Sub
Sub carregarreixacanvis(seccio As String, nummaquina As Integer)
     
End Sub
Sub carregarhorari(dia As Date)
  Dim rst As Recordset
  Dim datasql As String
  'carregar_reixacanvis seccio, nommaquina.ItemData(nommaquina.ListIndex)
  carregarcanvis
  For i = 0 To 23
    llistadhores.Selected(i) = False
  Next i
  If seccio <> "" And nommaquina <> "" Then
    datasql = " (year(dataihora)=" + atrim(Year(dia)) + " and month(dataihora)=" + atrim(Month(dia)) + " and day(dataihora)=" + atrim(Day(dia)) + ")"
    Set rst = dbplanificacioalicia.OpenRecordset("select * from horaris" + seccio + " where maquina=" + atrim(nommaquina.ItemData(nommaquina.ListIndex)) + " and " + datasql + " order by dataihora asc")
    
    While Not rst.EOF
      llistadhores.Selected((Format(rst!dataihora, "h"))) = True
      rst.MoveNext
    Wend
    marcardiesambprogramacio dia
   Else: MsgBox "Selecciona seccio i maquina primer"
  End If
End Sub
Sub marcardiesambprogramacio(dia As Date)
   Dim datasql As String
   Dim rst As Recordset
   If seccio <> "" And nommaquina <> "" Then
    For i = 0 To 31
       dies(i).BorderStyle = 0
     Next i
    datasql = " (year(dataihora)=" + atrim(Year(dia)) + " and month(dataihora)=" + atrim(Month(dia)) + ")"
    Set rst = dbplanificacio.OpenRecordset("select distinct day(dataihora) as dia from horaris" + seccio + " where maquina=" + atrim(nommaquina.ItemData(nommaquina.ListIndex)) + " and " + datasql)
    While Not rst.EOF
       dies(rst!dia).BorderStyle = 1
       rst.MoveNext
    Wend
   End If
End Sub
Private Sub dies_Click(Index As Integer)
   Dim i As Byte
   For i = 0 To dies.Count - 1
     If dies(i).BackColor = &HD17843 Then
      If dies(i).Tag = "" Then
         dies(i).BackColor = &H80000005
        Else: dies(i).BackColor = dies(i).Tag
      End If
      If dies(i).Tag = "" Then dies(i).ForeColor = QBColor(0)
     End If
   Next i
   dies(Index).BackColor = &HD17843
   dies(Index).ForeColor = QBColor(15)
   diaescullit = Format(Index, "00") + "/" + Format(cmesos.ListIndex + 1, "00") + "/" + tany
End Sub

Sub carregar_mes(mes As Byte, anyy As Long)
   Dim nommes As Variant
   Dim dia As String
   Dim setmana As Byte
   Dim margetop As Integer
   Dim margeleft As Integer
   Dim margeentresetmanes As Integer
   Dim margeentredies As Integer
   Dim primerasetmanadelmes As Integer
   Dim diadelasetmana As Integer
   If Not IsDate("01/" + atrim(mes) + "/" + atrim(anyy)) Then Exit Sub
   
   'nommes = Array("Gener", "Febrer", "Març", "Abril", "Maig", "Juny", "Juliol", "Agost", "Setembre", "Octubre", "Novembre", "Desembre")
   margetop = 700
   margeleft = 200
   margeentresetmanes = 500
   margeentredies = 460
   primerasetmanadelmes = DatePart("ww", "01/" + atrim(mes) + "/" + atrim(anyy))
   
   amagartotselsdies
   For i = 1 To 31
      dia = atrim(i) + "/" + atrim(mes) + "/" + atrim(anyy)
      If IsDate(dia) Then
        setmana = DatePart("ww", dia, vbMonday)
        diadelasetmana = DatePart("w", dia, vbMonday)
        dies(i).Visible = True
        dies(i).Top = ((setmana - primerasetmanadelmes) * margeentresetmanes) + margetop
        dies(i).Left = ((diadelasetmana - 1) * margeentredies) + margeleft
        If diadelasetmana > 5 Then
           'dies(i).BackStyle = 0
           dies(i).ForeColor = QBColor(15)
           dies(i).BackColor = &HD29F7D
           dies(i).Tag = &HD29F7D
            Else: dies(i).BackStyle = 1
        End If
        dies(i) = Day(dia)
      End If
   Next i
   marcardiesambprogramacio "01/" + atrim(mes) + "/" + atrim(anyy)
   Shape1.ZOrder 1
End Sub
Sub amagartotselsdies()
  For i = 0 To 31
    dies(i).ForeColor = QBColor(0)
    dies(i).BackColor = QBColor(15)
    dies(i).BorderStyle = 0
    dies(i).Visible = False
    dies(i).Tag = ""
  Next i
End Sub

Private Sub Form_Load()

  
  datacanvis.DatabaseName = rutadelfitxer(llegir_ini("General", "cami", "comandes.ini")) + "planificacio.mdb"
  datacanvisreb.DatabaseName = datacanvis.DatabaseName
  dataescalatsreb.DatabaseName = datacanvis.DatabaseName
  For i = 1 To 31
    Load dies(i)
  Next i
  diaescullit.Tag = "1"
  cmesos.ListIndex = Month(Now) - 1
  tany = Year(Now)
  carregar_mes cmesos.ListIndex + 1, cadbl(tany)
  dies_Click Day(Now)
  carregar_llista_hores
  diaescullit.Tag = ""
  
End Sub
Sub carregar_llista_hores()
  llistadhores.Clear
  For i = 0 To 23
    llistadhores.AddItem Format(i, "00") + ":00 - " + Format(i + 1, "00") + ":00"
  Next i
End Sub
Function setmanadelany(data As Date) As Double
   setmanadelany = DatePart("ww", data)
End Function

Private Sub HScroll1_Change()

End Sub

Private Sub nommaquina_Click()
 'carregarcanvis
 diaescullit_Change
 possartempscanvis
End Sub
Sub possartempscanvis()
   If seccio = "Rebobinadores" Then
      tempscanvisreb.Visible = True
      tempscanvisreb.Left = tempscanvisbasic.Left
      tempscanvisreb.Top = tempscanvisbasic.Top
      datacanvisreb.RecordSource = "select * from canvimaquinesreb where nummaquina=" + atrim(nommaquina.ItemData(nommaquina.ListIndex))
      datacanvisreb.Refresh
      If datacanvisreb.Recordset.EOF Then crear8bandesreb
        Else: tempscanvisreb.Visible = False
   End If
End Sub
Sub crear8bandesreb()
  Dim i As Byte
  For i = 1 To 8
    datacanvisreb.Recordset.AddNew
    datacanvisreb.Recordset!nummaquina = nommaquina.ItemData(nommaquina.ListIndex)
    datacanvisreb.Recordset!bandes = i
    datacanvisreb.Recordset.Update
  Next i
  datacanvisreb.Refresh
End Sub
Sub carregarcanvis()
  Dim i As Byte
  If seccio = "" Or nommaquina = "" Then Exit Sub
  datacanvis.RecordSource = "select * from canvismaquines where seccio='" + Mid(seccio, 1, 1) + "' and nummaquina=" + atrim(nommaquina.ItemData(nommaquina.ListIndex)) + " order by tintes asc"
  datacanvis.Refresh
  If datacanvis.Recordset.EOF Then
    For i = 1 To 8
      datacanvis.Recordset.AddNew
      datacanvis.Recordset!tintes = i
      datacanvis.Recordset!nummaquina = nommaquina.ItemData(nommaquina.ListIndex)
      datacanvis.Recordset!seccio = Mid(seccio, 1, 1)
      datacanvis.Recordset.Update
    Next i
  End If
  datacanvis.Refresh
End Sub
Private Sub seccio_Click()
  carregar_nommaquines Mid(seccio, 1, 1)
  nommaquina.SetFocus
  SendKeys "%{down}"
  
End Sub
Sub carregar_nommaquines(seccio As String)
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select * from maquines where maquina='" + seccio + "' and isnull(donadadebaixa)")
  nommaquina.Clear
  While Not rst.EOF
     nommaquina.AddItem atrim(rst!codi) + " - " + atrim(rst!descripcio)
     nommaquina.ItemData(nommaquina.NewIndex) = rst!codi
     rst.MoveNext
  Wend
End Sub
Private Sub tany_Change()
  carregar_mes cmesos.ListIndex + 1, cadbl(tany)
End Sub

