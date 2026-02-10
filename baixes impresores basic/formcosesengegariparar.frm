VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formcosesengegariparar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Coses a fer per Engegar i Parar Impresores"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   16695
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   16695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   750
      Left            =   165
      Picture         =   "formcosesengegariparar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   900
   End
   Begin VB.CommandButton afegir 
      Height          =   330
      Left            =   1950
      Picture         =   "formcosesengegariparar.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Afegir una feina"
      Top             =   420
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton eliminar 
      Height          =   330
      Left            =   2310
      Picture         =   "formcosesengegariparar.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Eliminar feina escullida."
      Top             =   420
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton desmarcarTOTS 
      Caption         =   "Tots"
      Height          =   765
      Left            =   1035
      Picture         =   "formcosesengegariparar.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   -15
      Width           =   900
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sortir"
      Height          =   750
      Left            =   14835
      Picture         =   "formcosesengegariparar.frx":17E8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Height          =   7260
      Left            =   165
      TabIndex        =   2
      Top             =   690
      Width           =   16305
      Begin MSFlexGridLib.MSFlexGrid reixa 
         Height          =   7065
         Left            =   120
         TabIndex        =   3
         Top             =   135
         Width           =   16005
         _ExtentX        =   28231
         _ExtentY        =   12462
         _Version        =   393216
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton bparar 
      Caption         =   "Parar"
      Height          =   750
      Left            =   10650
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   -90
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.CommandButton bengegar 
      BackColor       =   &H0017D062&
      Caption         =   "Engegar"
      Height          =   750
      Left            =   8325
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   -90
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Label ettipus 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2010
      TabIndex        =   8
      Top             =   30
      Width           =   3210
   End
   Begin VB.Menu mop 
      Caption         =   "Opcions"
      Begin VB.Menu mafeborrarfeines 
         Caption         =   "Afegir/borrar feines."
      End
   End
End
Attribute VB_Name = "formcosesengegariparar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vfitxerini_engegariparar As String

Private Sub afegir_Click()
   Dim vdescripcio As String
   Dim vTipus As String
   If bengegar.BackColor = &H17D062 Then vTipus = "Engegar" Else vTipus = "Parar"
   vdescripcio = atrim(InputBox("Escriu la descripció de la feina que vols que es faci.", "Afegir"))
   escriure_ini vTipus, "Descripcio" + atrim(reixa.Rows), vdescripcio, vfitxerini_engegariparar
   carregar_feines
End Sub

Private Sub Command1_Click()
   Dim rst As Recordset
   Dim i As Byte
   Dim oapp As CRAXDDRT.Application
   Dim oreport As CRAXDDRT.Report
   Dim vTipus As String
   Dim vnommaquina As String
   vnommaquina = IIf(nummaq = 9, "F2", IIf(nummaq = 7, "FW", "XX"))
   dbtmpb.Execute "delete * from llistat_engegar_parar where nummaquina=" + atrim(nummaq)
   Set rst = dbtmpb.OpenRecordset("select * from llistat_engegar_parar where nummaquina=" + atrim(nummaq))
   i = 1
   While i < reixa.Rows
      rst.AddNew
      rst!descripcio_feina = reixa.TextMatrix(i, 1)
      rst!fet = reixa.TextMatrix(i, 0)
      rst!nummaquina = nummaq
      rst.Update
      i = i + 1
   Wend
   wait 1
   'IMPRIMIR
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistat feines parar i engegar impresores.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "baixes.mdb"
  oreport.RecordSelectionFormula = "{llistat_engegar_parar.nummaquina}=" + atrim(nummaq)
  vTipus = IIf(bengegar.BackColor = &H17D062, "Engegar", "Parar")
  oreport.FormulaFields.GetItemByName("titol").text = "'" + vTipus + " - FEINES ABANS DE " + UCase(vTipus) + " LA IMPRESORA " + vnommaquina + "'"
  oreport.FormulaFields.GetItemByName("nomdelamaquina").text = "'" + vnommaquina + "'"
  oreport.FormulaFields.GetItemByName("nomoperari").text = "'" + form1.nomoperari + "'"
  oreport.PrintOut False
  ' Load veurereport
  '      veurereport.CRViewer.ReportSource = oreport
  '      veurereport.CRViewer.DisplayGroupTree = False
  '      veurereport.CRViewer.ViewReport
  '      veurereport.WindowState = 2
  '      veurereport.Show 1
  Set rst = Nothing
  desmarcarTOTS.tag = ""  'per saber que ja ha imprès
End Sub
Public Sub bparar_Click()
   alternarbotons
End Sub
Public Sub bengegar_Click()
   alternarbotons
End Sub
Sub alternarbotons()
   If bengegar.BackColor = &H17D062 Then bengegar.BackColor = &H8000000F Else bengegar.BackColor = &H17D062
   If bparar.BackColor = &H17D062 Then bparar.BackColor = &H8000000F Else bparar.BackColor = &H17D062
   carregar_feines
End Sub
Sub carregar_feines(Optional vnummaq As Double, Optional vTipus)
  Dim i As Byte
  Dim v As String
  Dim v2 As String
  If cadbl(vnummaq) > 0 Then assignar_fitxerini vnummaq
  If bengegar.BackColor = &H17D062 Then vTipus = "Engegar" Else vTipus = "Parar"
  'ettipus = vTipus + "  " + IIf(vnummaq = 7, "FW", "F2")
  
  '  abans hi havia dues llistes engegar i parar pero ara nomes en volen una i utilitzo la d'engegar
  ettipus = "Manteniment  " + IIf(vnummaq = 7, "FW", "F2")
  i = 1
  reixa.Rows = 1
  While v <> "{[}]"
     v = llegir_ini(vTipus, "Descripcio" + atrim(i), vfitxerini_engegariparar)
     If atrim(v) = "" Then v = "{[}]"
     If v <> "{[}]" Then
       reixa.Rows = reixa.Rows + 1: reixa.TextMatrix(i, 1) = v
       v2 = llegir_ini(vTipus, "Fet" + atrim(i), vfitxerini_engegariparar)
       If atrim(v2) = "" Then v2 = "{[}]"
       reixa.TextMatrix(i, 0) = "q"
       If v2 <> "{[}]" Then reixa.TextMatrix(i, 0) = v2
       
       v2 = llegir_ini(vTipus, "Operari" + atrim(i), vfitxerini_engegariparar)
       If atrim(v2) = "" Then v2 = "{[}]"
       If v2 <> "{[}]" Then reixa.TextMatrix(i, 1) = IIf(cadbl(v2) > 0, "[" + v2 + "] ", "") + reixa.TextMatrix(i, 1)
       v2 = ""
     End If
     i = i + 1
  Wend
  possar_format_reixa
End Sub
Sub possar_format_reixa()
reixa.ColWidth(1) = 14000
   reixa.TextMatrix(0, 0) = "Feta"
   reixa.TextMatrix(0, 1) = "Descripció de la feina."
    With reixa
        .ColWidth(0) = 600
        .RowHeightMin = 300
        If .Rows > 1 Then .row = 1 Else GoTo fi
        .col = 0
        .RowSel = .Rows - 1
        .FillStyle = flexFillRepeat
        .CellFontName = "Wingdings"
        .CellFontSize = .CellFontSize + 6
        .CellAlignment = flexAlignCenterCenter
       ' .text = IIf(.text = "", "q", .text)
        .FillStyle = flexFillSingle
        .row = 0
        .col = 1
        .RowSel = .Rows - 1
        .CellAlignment = flexAlignLeftCenter
        .row = 0
    End With
fi:
End Sub
Private Sub Command3_Click()
  If desmarcarTOTS.tag = "1" Then
      MsgBox "Abans de sortir has d'imprimir el full per l'encarregat.", vbCritical, "Atenció": Exit Sub
  End If
  Unload Me
End Sub

Private Sub Command4_Click()
 
End Sub

Private Sub desmarcarTOTS_Click()
  If MsgBox("Estàs segur que vols passar totes les feines a NO FETES?", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then Exit Sub
  passartotesaUNCHECK
End Sub

Private Sub eliminar_Click()
   Dim i As Byte
   Dim v As String
   Dim vTipus As String
   If reixa.row = 0 Then Exit Sub
   If bengegar.BackColor = &H17D062 Then vTipus = "Engegar" Else vTipus = "Parar"
   If MsgBox("Segur que vols eliminar aquesta feina?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   For i = reixa.row To reixa.Rows - 1
      v = llegir_ini(vTipus, "Descripcio" + atrim(i + 1), vfitxerini_engegariparar)
      If v = "{[}]" Then v = ""
      escriure_ini vTipus, "Descripcio" + atrim(i), v, vfitxerini_engegariparar
      reixa.TextMatrix(i, 1) = v
   Next i
   escriure_ini vTipus, "Descripcio" + atrim(reixa.Rows), "", vfitxerini_engegariparar
   carregar_feines
End Sub

Private Sub Form_Load()
    assignar_fitxerini
    carregar_feines
End Sub
Sub assignar_fitxerini(Optional vnummaq As Double)
  If cadbl(vnummaq) = 0 Then vnummaq = nummaq
  vfitxerini_engegariparar = rutadelfitxer(cami) + "cosesafer_engegar_parar_" + atrim(vnummaq) + ".ini"
End Sub

Private Sub mafeborrarfeines_Click()
  If UCase(InputBoxEx("Escriu la contrasenya per editar les feines.", "Contrasenya", , , , , , SPassword)) <> "INPLACSA" Then Exit Sub
  eliminar.visible = True
  afegir.visible = True
End Sub

Private Sub reixa_Click()
   Dim vTipus As String
   Dim vnumop As Double
   vnumop = numop
   If reixa.col <> 0 Then Exit Sub
   If bengegar.BackColor = &H17D062 Then vTipus = "Engegar" Else vTipus = "Parar"
    With reixa
        If .col = 0 Then
            If .TextMatrix(.row, 0) = "q" Then
                .TextMatrix(.row, 0) = "þ"
                  Else
                   .TextMatrix(.row, 0) = "q"
                   vnumop = 0
            End If
        End If
   escriure_ini vTipus, "Fet" + atrim(.row), .TextMatrix(.row, 0), vfitxerini_engegariparar
   escriure_ini vTipus, "Operari" + atrim(.row), atrim(vnumop), vfitxerini_engegariparar
   reixa.CellAlignment = flexAlignLeftCenter
   End With
   carregar_feines
End Sub
Sub passartotesaUNCHECK()
   Dim vTipus As String
   Dim i As Byte
   If bengegar.BackColor = &H17D062 Then vTipus = "Engegar" Else vTipus = "Parar"
   For i = 1 To reixa.Rows - 1
     escriure_ini vTipus, "Fet" + atrim(i), "q", vfitxerini_engegariparar
     escriure_ini vTipus, "Operari" + atrim(i), "0", vfitxerini_engegariparar
   Next i
   carregar_feines
End Sub
