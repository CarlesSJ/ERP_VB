VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form manteniments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniments de la Fàbrica"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12030
   ControlBox      =   0   'False
   Icon            =   "manteniments fabrica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   12030
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command9 
      Height          =   510
      Left            =   11445
      Picture         =   "manteniments fabrica.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Fitxa del manteniment"
      Top             =   765
      Width           =   465
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "manteniments fabrica.frx":0B14
      Height          =   3765
      Left            =   30
      OleObjectBlob   =   "manteniments fabrica.frx":0B2F
      TabIndex        =   9
      Top             =   2415
      Width           =   11955
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descripció del Manteniment"
      Enabled         =   0   'False
      Height          =   1770
      Left            =   45
      TabIndex        =   8
      Top             =   600
      Width           =   11940
      Begin VB.CheckBox inactiu 
         Caption         =   "Inactiu"
         DataField       =   "inactiu"
         DataSource      =   "datamanteniments"
         Height          =   225
         Left            =   4890
         TabIndex        =   40
         Top             =   135
         Width           =   885
      End
      Begin VB.CommandButton Command7 
         Height          =   285
         Left            =   6285
         Picture         =   "manteniments fabrica.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Afegir tipificacio de manteniment."
         Top             =   1140
         Width           =   300
      End
      Begin VB.CommandButton Command6 
         Height          =   285
         Left            =   6585
         Picture         =   "manteniments fabrica.frx":2E9E
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Editar o esborrar aquesta tipificació del manteniment."
         Top             =   1140
         Width           =   300
      End
      Begin VB.ComboBox descripcioproveidor 
         Height          =   315
         Left            =   4020
         TabIndex        =   34
         Top             =   1155
         Width           =   2220
      End
      Begin VB.ComboBox descripciomanteniment 
         DataField       =   "descripcio"
         DataSource      =   "datamanteniments"
         Height          =   315
         Left            =   1095
         TabIndex        =   32
         Top             =   450
         Width           =   5145
      End
      Begin VB.CommandButton Command5 
         Height          =   285
         Left            =   6570
         Picture         =   "manteniments fabrica.frx":3428
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Editar o esborrar aquesta tipificació del manteniment."
         Top             =   435
         Width           =   300
      End
      Begin VB.CommandButton Command3 
         Height          =   285
         Left            =   6270
         Picture         =   "manteniments fabrica.frx":39B2
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Afegir tipificacio de manteniment."
         Top             =   435
         Width           =   300
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Left            =   8085
         Picture         =   "manteniments fabrica.frx":3F3C
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Borrar la data de Baixa del Treball"
         Top             =   1170
         Width           =   270
      End
      Begin VB.CommandButton Command4 
         Height          =   285
         Left            =   8085
         Picture         =   "manteniments fabrica.frx":44C6
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Borrar la data de Baixa del Treball"
         Top             =   450
         Width           =   270
      End
      Begin VB.TextBox camps 
         DataField       =   "datafi"
         DataSource      =   "datamanteniments"
         Height          =   300
         Index           =   6
         Left            =   6960
         TabIndex        =   25
         Top             =   1170
         Width           =   1125
      End
      Begin VB.ComboBox nommaquina 
         DataField       =   "nommaquina"
         DataSource      =   "datamanteniments"
         Height          =   315
         ItemData        =   "manteniments fabrica.frx":4A50
         Left            =   1515
         List            =   "manteniments fabrica.frx":4A5A
         TabIndex        =   24
         Top             =   1170
         Width           =   2445
      End
      Begin VB.ComboBox seccio 
         DataField       =   "seccio"
         DataSource      =   "datamanteniments"
         Height          =   315
         ItemData        =   "manteniments fabrica.frx":4A68
         Left            =   135
         List            =   "manteniments fabrica.frx":4A84
         TabIndex        =   21
         Top             =   1170
         Width           =   1320
      End
      Begin VB.TextBox camps 
         DataField       =   "cadaxanys"
         DataSource      =   "datamanteniments"
         Height          =   300
         Index           =   5
         Left            =   10560
         TabIndex        =   20
         Top             =   450
         Width           =   525
      End
      Begin VB.TextBox camps 
         DataField       =   "cadaxmesos"
         DataSource      =   "datamanteniments"
         Height          =   300
         Index           =   4
         Left            =   9450
         TabIndex        =   18
         Top             =   450
         Width           =   525
      End
      Begin VB.TextBox camps 
         DataField       =   "cadaxdies"
         DataSource      =   "datamanteniments"
         Height          =   300
         Index           =   3
         Left            =   8415
         TabIndex        =   16
         Top             =   450
         Width           =   525
      End
      Begin VB.TextBox camps 
         DataField       =   "datainici"
         DataSource      =   "datamanteniments"
         Height          =   300
         Index           =   2
         Left            =   6960
         TabIndex        =   14
         Top             =   450
         Width           =   1125
      End
      Begin VB.TextBox camps 
         DataField       =   "id"
         DataSource      =   "datamanteniments"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   405
         TabIndex        =   12
         Top             =   450
         Width           =   525
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveïdor que fa la reparació (F2)"
         Height          =   330
         Index           =   9
         Left            =   4020
         TabIndex        =   37
         Top             =   915
         Width           =   2475
      End
      Begin VB.Label descripciooriginal 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1110
         TabIndex        =   33
         Top             =   765
         Width           =   5055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Foto1            Foto2         Foto3         Foto4"
         Height          =   195
         Left            =   8520
         TabIndex        =   29
         Top             =   765
         Width           =   2970
      End
      Begin VB.Image foto4 
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Left            =   10845
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Top             =   960
         Width           =   810
      End
      Begin VB.Image foto3 
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Left            =   10035
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Top             =   960
         Width           =   810
      End
      Begin VB.Image foto2 
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Left            =   9225
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Top             =   945
         Width           =   810
      End
      Begin VB.Image foto1 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "manteniments fabrica.frx":4AE1
         Height          =   675
         Left            =   8415
         OLEDropMode     =   1  'Manual
         Stretch         =   -1  'True
         Top             =   960
         Width           =   810
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Fi"
         Height          =   255
         Index           =   8
         Left            =   7365
         TabIndex        =   26
         Top             =   960
         Width           =   840
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Màquina"
         Height          =   255
         Index           =   7
         Left            =   1785
         TabIndex        =   23
         Top             =   945
         Width           =   1695
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Secció"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   22
         Top             =   930
         Width           =   570
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cada X Anys"
         Height          =   255
         Index           =   5
         Left            =   10395
         TabIndex        =   19
         Top             =   225
         Width           =   960
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cada X Mesos"
         Height          =   255
         Index           =   4
         Left            =   9285
         TabIndex        =   17
         Top             =   225
         Width           =   1125
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cada X Dies"
         Height          =   255
         Index           =   3
         Left            =   8250
         TabIndex        =   15
         Top             =   225
         Width           =   960
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inici"
         Height          =   255
         Index           =   2
         Left            =   7290
         TabIndex        =   13
         Top             =   225
         Width           =   1290
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripció del manteniment (F2)"
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   11
         Top             =   225
         Width           =   4455
      End
      Begin VB.Label etiqueta 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº"
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   10
         Top             =   225
         Width           =   285
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11985
      Begin VB.CommandButton Command8 
         Height          =   330
         Left            =   10275
         Picture         =   "manteniments fabrica.frx":506B
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   180
         Width           =   1035
      End
      Begin VB.Data datamanteniments 
         Caption         =   "Manteniments"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   315
         Left            =   4470
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "manteniments"
         Top             =   195
         Width           =   2640
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   90
         Picture         =   "manteniments fabrica.frx":55F5
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   965
         Picture         =   "manteniments fabrica.frx":5B7F
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   520
         Picture         =   "manteniments fabrica.frx":6109
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Edicio del  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   11505
         Picture         =   "manteniments fabrica.frx":6693
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Sortir"
         Top             =   150
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   1410
         Picture         =   "manteniments fabrica.frx":6C1D
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Acceptar canvis"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton consultar 
         Height          =   360
         Left            =   1845
         Picture         =   "manteniments fabrica.frx":71A7
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Buscar Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.Label estatedicio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2490
         TabIndex        =   7
         Top             =   150
         Width           =   2025
      End
   End
End
Attribute VB_Name = "manteniments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
  If datamanteniments.Recordset.EditMode > 0 Then MsgBox "Estas editant primer finalitza la operació i despres afegeix.", vbCritical, "Atenció": Exit Sub
   datamanteniments.Recordset.AddNew
   Frame1.Enabled = True
   descripciomanteniment.SetFocus
End Sub
  
Private Sub Command1_Click()
  gravar_canvis True
End Sub
Sub gravar_canvis(Optional generarhoraris As Boolean)
  Dim idactual As Long
  On Error GoTo errror
  Frame1.Enabled = False
  If datamanteniments.Recordset.EditMode = 0 Then Exit Sub
  If Not IsNumeric(camps(5)) Then camps(5) = "0"
  If Not IsNumeric(camps(4)) Then camps(4) = "0"
  If Not IsNumeric(camps(3)) Then camps(3) = "0"
  idactual = datamanteniments.Recordset!id
  datamanteniments.Recordset.Update
  datamanteniments.Recordset.FindFirst "id=" + atrim(idactual)
  If generarhoraris Then generarhorarismanteniments datamanteniments.Recordset!id
 Exit Sub
errror:
   MsgBox err.Description
End Sub
Sub generarhorarismanteniments(Optional nid As Long)
   Dim rstm As Recordset
   Dim datageneradaincial As Date
   Dim datagenerada As Date
   Dim datafi As Date
   
   Set rstm = dbmanteniments.OpenRecordset("select * from manteniments " + IIf(nid > 0, " where id=" + atrim(nid), ""))
   While Not rstm.EOF
     dbmanteniments.Execute "delete * from horarismanteniments where idmanteniment=" + atrim(cadbl(rstm!id)) + " and data>#" + Format(Now, "mm/dd/yy") + "# and (nomoperari=null or nomoperari='' or nomoperari='_')"
     datafi = DateAdd("yyyy", 1, Now)
     datageneradainicial = Now
     If IsDate(rstm!datainici) Then datageneradainicial = IIf(DateDiff("d", Now, rstm!datainici) < 0, rstm!datafi, rstm!datainici)
     If IsDate(rstm!datafi) Then datafi = rstm!datafi
     
     datagenerada = IIf(datageneradainicial = Null, Now, datageneradainicial)
     If cadbl(rstm!cadaxdies) > 0 Then
       While DateDiff("d", datagenerada, datafi) > 0
         guardarhorari nid, DateAdd("d", rstm!cadaxdies, datagenerada), datagenerada
       Wend
     End If
     
     datagenerada = datageneradainicial
     If cadbl(rstm!cadaxmesos) > 0 Then
       While DateDiff("d", datagenerada, datafi) > 0
         guardarhorari nid, DateAdd("m", rstm!cadaxmesos, datagenerada), datagenerada
       Wend
     End If
     
     datagenerada = datageneradainicial
     If cadbl(rstm!cadaxanys) > 0 Then
       While DateDiff("d", datagenerada, datafi) > 0
         guardarhorari nid, DateAdd("yyyy", rstm!cadaxanys, datagenerada), datagenerada
       Wend
     End If
     rstm.MoveNext
   Wend
End Sub
Sub guardarhorari(nid As Long, data As Date, datagenerada As Date)
   dbmanteniments.Execute "insert into horarismanteniments (idmanteniment,data) values (" + atrim(nid) + ",#" + Format(data, "mm/dd/yy") + "#)"
   datagenerada = data
End Sub
Sub borrardataclixe(camp As String)
  gravar_canvis
  dbmanteniments.Execute "update manteniments set " + camp + "=null where id=" + atrim(datamanteniments.Recordset!id)
  datamanteniments.UpdateControls
  modificar_Click
  
End Sub

Private Sub Command2_Click()
   borrardataclixe "datafi"
End Sub

Private Sub Command3_Click()
  Dim resp As String
  resp = InputBox("Entra la tipificació de manteniment que vols utilitzar.", "Nova tipificació")
  If atrim(resp) <> "" Then
      dbmanteniments.Execute "insert into tipusdemanteniments (descripcio) values ('" + treure_apostruf(resp) + "')"
  End If
End Sub

Private Sub Command4_Click()
  borrardataclixe "datainici"
End Sub

Private Sub Command5_Click()
  Dim resp As String
  Dim rstt As Recordset
  gravar_canvis
  modificar_Click
  If cadbl(datamanteniments.Recordset!idtipusmanteniment) = 0 Then MsgBox "No hi ha cap tipificació sel.leccionada.", vbCritical, "Error": Exit Sub
  resp = InputBox("Modifica el tipus de modificació." + Chr(10) + "Escriu [Eliminar] per eliminar la descripció.", "Modificar o Eliminar", descripciooriginal.Tag)
  If atrim(resp) <> "" Then
     If resp <> "Eliminar" Then
         dbmanteniments.Execute "update tipusdemanteniments set descripcio='" + treure_apostruf(resp) + "' where id=" + atrim(datamanteniments.Recordset!idtipusmanteniment)
         If descripciooriginal = "" Then
            descripciomanteniment = treure_apostruf(resp)
           Else: descripciooriginal = treure_apostruf(resp): descripciooriginal.Tag = treure_apostruf(resp)
         End If
        Else
           Set rstt = dbmanteniments.OpenRecordset("select * from manteniments where idtipusmanteniment=" + atrim(datamanteniments.Recordset!idtipusmanteniment) + " and id<>" + atrim(datamanteniments.Recordset!id))
           If rstt.EOF Then
              dbmanteniments.Execute "delete * from tipusdemanteniments where id=" + atrim(datamanteniments.Recordset!idtipusmanteniment)
              datamanteniments.Recordset!idtipusmanteniment = Null
              datamanteniments.Recordset!descripcio = " "
                Else: MsgBox "No es pot eliminar aquesta tipificació perquè hi ha manteniments que l'utilitzen.", vbCritical, "Atenció": Exit Sub
           End If
     End If
     gravar_canvis
     modificar_Click
  End If
End Sub

Private Sub Command6_Click()
 Dim resp As String
  Dim rstt As Recordset
  gravar_canvis
  modificar_Click
  If atrim(descripcioproveidor) = "" Then Exit Sub
  resp = InputBox("Modifica el proveïdor." + Chr(10) + "Escriu [Eliminar] per eliminar-lo.", "Modificar o Eliminar", descripcioproveidor)
  If atrim(resp) <> "" Then
     If resp <> "Eliminar" Then
         dbmanteniments.Execute "update proveidorsreparacions set descripcio='" + treure_apostruf(resp) + "' where id=" + atrim(datamanteniments.Recordset!idproveidor)
         descripcioproveidor = treure_apostruf(resp)
        Else
           Set rstt = dbmanteniments.OpenRecordset("select * from manteniments where idproveidor=" + atrim(datamanteniments.Recordset!idproveidor) + " and id<>" + atrim(datamanteniments.Recordset!id))
           If rstt.EOF Then
              dbmanteniments.Execute "delete * from proveidorsreparacions where id=" + atrim(datamanteniments.Recordset!idproveidor)
              datamanteniments.Recordset!idproveidor = Null
              'datamanteniments.Recordset!descripcioproveidor = " "
                Else: MsgBox "No es pot eliminar aquest proveidor perquè hi ha manteniments que l'utilitzen.", vbCritical, "Atenció": Exit Sub
           End If
     End If
     gravar_canvis
     modificar_Click
  End If
End Sub

Private Sub Command7_Click()
  Dim resp As String
  resp = InputBox("Entra el nom del proveïdor nou.", "Nou proveïdor")
  If atrim(resp) <> "" Then
      dbmanteniments.Execute "insert into proveidorsreparacions (descripcio) values ('" + treure_apostruf(resp) + "')"
  End If
End Sub

Private Sub Command8_Click()
  imprimirfitxa cadbl(datamanteniments.Recordset!id)
End Sub
Sub comprovarquehihaguialgunavisfet(numm As Double)
   Dim rstt As Recordset
   Set rstt = dbmanteniments.OpenRecordset("select * from horarismanteniments where (nomoperari<>'' or nomoperari<>null) and idmanteniment=" + atrim(numm))
   If rstt.EOF Then
      Set rstt = dbmanteniments.OpenRecordset("select * from horarismanteniments where (nomoperari=null or nomoperari='') and idmanteniment=" + atrim(numm) + " order by data")
      If Not rstt.EOF Then
         rstt.Edit
           rstt!nomoperari = "_"
         rstt.Update
         
      End If
   End If
End Sub
Sub imprimirfitxa(numm As Double)
  
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "avismanteniment.rpt", 1)
  comprovarquehihaguialgunavisfet numm
  'oreport.Database.SetDataSource (dbconsulta)
  carregarimatgesdelafitxa numm
  oreport.Database.Tables.Item(1).Location = datamanteniments.DatabaseName
  oreport.RecordSelectionFormula = "{manteniments.id}=" + atrim(cadbl(numm)) + "  and {horarismanteniments.nomoperari}<>''"
  oreport.DiscardSavedData
  'MsgBox oreport.SQLQueryString
  wait 2
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.Show 1, Me
    Else
      oreport.PrintOut False, 1
  End If
  
  

End Sub
Sub carregarimatgesdelafitxa(numm As Double)
   Dim rstm As Recordset
   Dim rstt As Recordset
   dbmanteniments.Execute "delete * from llistat_fitxa"
   Set rstm = dbmanteniments.OpenRecordset("select * from manteniments where id=" + atrim(numm))
   Set rstt = dbmanteniments.OpenRecordset("llistat_fitxa")
   If Not rstm.EOF Then
     rstt.AddNew
     rstt!idmanteniment = numm
     copiafoto atrim(rstm!linkfoto1), rstt!foto1
     copiafoto atrim(rstm!linkfoto2), rstt!foto2
     copiafoto atrim(rstm!linkfoto3), rstt!foto3
     copiafoto atrim(rstm!linkfoto4), rstt!foto4
     rstt.Update
   End If
End Sub

Private Sub Command9_Click()
  nummanteniment = cadbl(datamanteniments.Recordset!id)
  fitxamanteniment.Show 1
End Sub

Private Sub consultar_Click()
   Dim resp As String
   resp = InputBox("Escriu la descripcio que vols buscar o el codi de manteniment.", "Buscar/Filtrar")
   If atrim(resp) <> "" Then
       If cadbl(resp) > 0 Then
           datamanteniments.RecordSource = "select * from manteniments where id=" + atrim(resp)
           datamanteniments.Refresh
          Else
             datamanteniments.RecordSource = "select * from manteniments where descripcio like'*" + treure_apostruf(resp) + "*'"
             datamanteniments.Refresh
       End If
   End If
   
End Sub

Private Sub datamanteniments_Reposition()
   carregar_relacions
   carregar_fotos
End Sub
Sub carregar_relacions()
    Dim rstt As Recordset
    Set rstt = dbmanteniments.OpenRecordset("select * from tipusdemanteniments where id=" + atrim(cadbl(datamanteniments.Recordset!idtipusmanteniment)))
    descripciooriginal = ""
    descripciooriginal.Tag = ""
    descripcioproveidor = ""
    If datamanteniments.Recordset.EOF Then Exit Sub
    If datamanteniments.Recordset!inactiu Then
        Frame1.BackColor = QBColor(8)
       Else: Frame1.BackColor = Command1.BackColor
    End If
    If Not rstt.EOF Then
        descripciooriginal.Tag = atrim(rstt!descripcio)
        If atrim(rstt!descripcio) <> atrim(datamanteniments.Recordset!descripcio) Then descripciooriginal = atrim(rstt!descripcio)
    End If
    Set rstt = dbmanteniments.OpenRecordset("select * from proveidorsreparacions where id=" + atrim(cadbl(datamanteniments.Recordset!idproveidor)))
  
    If Not rstt.EOF Then
         descripcioproveidor = atrim(rstt!descripcio)
    End If
End Sub
Sub carregar_fotos()
  foto1.Picture = Nothing
  foto2.Picture = Nothing
  foto3.Picture = Nothing
  foto4.Picture = Nothing

  If existeix(atrim(datamanteniments.Recordset!linkfoto1)) Then foto1.Picture = LoadPicture(atrim(datamanteniments.Recordset!linkfoto1))
  If existeix(atrim(datamanteniments.Recordset!linkfoto2)) Then foto2.Picture = LoadPicture(atrim(datamanteniments.Recordset!linkfoto2))
  If existeix(atrim(datamanteniments.Recordset!linkfoto3)) Then foto3.Picture = LoadPicture(atrim(datamanteniments.Recordset!linkfoto3))
  If existeix(atrim(datamanteniments.Recordset!linkfoto4)) Then foto4.Picture = LoadPicture(atrim(datamanteniments.Recordset!linkfoto4))
End Sub

Private Sub descripciomanteniment_DropDown()
Load formseleccio
   formseleccio.Data1.DatabaseName = dbmantenimentsName
   formseleccio.Data1.RecordSource = "select id,descripcio from tipusdemanteniments order by descripcio"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns(0).Visible = False
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           descripciomanteniment = formseleccio.DBGrid2.Columns("descripcio")
           datamanteniments.Recordset!idtipusmanteniment = formseleccio.DBGrid2.Columns("id")
        End If
   End If
    If seleccioret = 9 Then
        descripciomanteniment = ""
        datamanteniments.Recordset!idtipusmanteniment = Null
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub descripciomanteniment_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then
     descripciomanteniment_DropDown
  End If
End Sub

Private Sub descripcioproveidor_DropDown()
Load formseleccio
   formseleccio.Data1.DatabaseName = datamanteniments.DatabaseName
   formseleccio.Data1.RecordSource = "select id,descripcio from proveidorsreparacions order by descripcio"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns(0).Visible = False
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           descripcioproveidor = formseleccio.DBGrid2.Columns("descripcio")
           datamanteniments.Recordset!idproveidor = formseleccio.DBGrid2.Columns("id")
        End If
   End If
    If seleccioret = 9 Then
        descripcioproveidor = ""
        datamanteniments.Recordset!idproveidor = Null
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub descripcioproveidor_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then
     descripcioproveidor_DropDown
       Else: KeyCode = 0
   End If
End Sub

Private Sub descripcioproveidor_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub eliminar_Click()
   Dim rstt As Recordset
   Set rstt = dbmanteniments.OpenRecordset("select * from horarismanteniments where idmanteniment=" + atrim(datamanteniments.Recordset!id) + " and dataexecucio<>null")
   If Not rstt.EOF Then
       MsgBox "No es pot eliminar aquest manteniment perquè s'han fet avisos i es perdria l'historic", vbCritical, "Atenció"
       Exit Sub
   End If
   If MsgBox("Segur que vols borrar aquesta programació?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   dbmanteniments.Execute "delete * from horarismanteniments where idmanteniment=" + atrim(datamanteniments.Recordset!id)
   datamanteniments.Recordset.Delete
   datamanteniments.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 Then
       gravar_canvis
   End If
End Sub
Sub generaravisosfuturs()
   Dim rsta As Recordset
   Set rsta = dbmanteniments.OpenRecordset("select * from generals")
   If DateDiff("d", rsta!ultimadataactualitzacio, Now) > 30 Then
      rsta.Edit
       rsta!ultimadataactualitzacio = Now
      rsta.Update
      Set rsta = dbmanteniments.OpenRecordset("select id from manteniments where not inactiu")
      While Not rsta.EOF
         generarhorarismanteniments rsta!id
         rsta.MoveNext
      Wend
      dbmanteniments.Execute "delete * from horarismanteniments where data<#" + Format(DateAdd("yyyy", -1, Now), "mm/dd/yy") + "# and (nomoperari='' or nomoperari=null or nomoperari='_')"
   End If
End Sub
Private Sub Form_Load()
  Dim arguments As Variant
  If App.PrevInstance Then MsgBox "El programa ja està obert.", vbCritical, "Atenció": End
  arguments = ObtenerLíneaComando
  If cadbl(arguments(1)) > 0 Then nummanteniment = cadbl(arguments(1))
  fitxerini = "comandes.ini"
  cami = llegir_ini("General", "cami", fitxerini)
  camicomandes = cami
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  '"c:\misdoc~1\commandes\comandes.mdb"
  If existeix("c:\ordprog.ini") Then cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
  centerscreen Me
  datamanteniments.DatabaseName = rutadelfitxer(cami) + "mantenimentsfabrica.mdb"
  Set dbmanteniments = OpenDatabase(datamanteniments.DatabaseName)
  generaravisosfuturs
  If nummanteniment = 0 Then
     datamanteniments.RecordSource = "manteniments"
     datamanteniments.Refresh
       Else:
         Me.Tag = "extern"
         Me.Visible = False
         datamanteniments.RecordSource = "select * from manteniments where id=" + atrim(nummanteniment)
         datamanteniments.Refresh
         fitxamanteniment.Show 1
  End If
  
 
  
End Sub

Private Sub foto1_DblClick()
  If Not existeix(atrim(datamanteniments.Recordset!linkfoto1)) Then Exit Sub
   Shell "c:\windows\system32\cmd.exe /c """ + datamanteniments.Recordset!linkfoto1 + """"
End Sub

Private Sub foto1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo err
   foto1.Picture = LoadPicture(data.Files(1))
   datamanteniments.Recordset!linkfoto1 = data.Files(1)
   Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub foto2_DblClick()
If Not existeix(atrim(datamanteniments.Recordset!linkfoto2)) Then Exit Sub
Shell "c:\windows\system32\cmd.exe /c """ + datamanteniments.Recordset!linkfoto2 + """"
End Sub

Private Sub foto2_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo err
   foto2.Picture = LoadPicture(data.Files(1))
   datamanteniments.Recordset!linkfoto2 = data.Files(1)
   Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub foto3_DblClick()
If Not existeix(atrim(datamanteniments.Recordset!linkfoto3)) Then Exit Sub
Shell "c:\windows\system32\cmd.exe /c """ + datamanteniments.Recordset!linkfoto3 + """"
End Sub

Private Sub foto3_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo err
   foto3.Picture = LoadPicture(data.Files(1))
   datamanteniments.Recordset!linkfoto3 = data.Files(1)
   Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub foto4_DblClick()
If Not existeix(atrim(datamanteniments.Recordset!linkfoto4)) Then Exit Sub
Shell "c:\windows\system32\cmd.exe /c """ + datamanteniments.Recordset!linkfoto4 + """"
End Sub

Private Sub foto4_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo err
   foto4.Picture = LoadPicture(data.Files(1))
   datamanteniments.Recordset!linkfoto4 = data.Files(1)
   Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub modificar_Click()
 If datamanteniments.Recordset.EditMode > 0 Then MsgBox "Estas editant primer finalitza la operació i despres afegeix.", vbCritical, "Atenció": Exit Sub
   datamanteniments.Recordset.Edit
   Frame1.Enabled = True
   descripciomanteniment.SetFocus
End Sub

Private Sub nommaquina_DropDown()
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select codi,descripcio from maquines where maquina='" + Mid(seccio, 1, 1) + "' and donadadebaixa=null "
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           nommaquina = formseleccio.DBGrid2.Columns("descripcio")
           datamanteniments.Recordset!maquina = formseleccio.DBGrid2.Columns("codi")
        End If
   End If
    If seleccioret = 9 Then
        nommaquina = ""
        datamanteniments.Recordset!maquina = Null
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub sortir_Click()
  'generarhorarismanteniments
  End
  
End Sub

