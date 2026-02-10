VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form formestocseguretat 
   Caption         =   "Estoc de seguretat"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15720
   ClipControls    =   0   'False
   Icon            =   "formestocseguretat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   15720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Height          =   480
      Left            =   675
      Picture         =   "formestocseguretat.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Alta  Registres"
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton beliminar 
      Height          =   480
      Left            =   1545
      Picture         =   "formestocseguretat.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton botoafegir 
      Height          =   480
      Left            =   30
      Picture         =   "formestocseguretat.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Alta  Registres"
      Top             =   60
      Width           =   615
   End
   Begin VB.Data datadetall 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6390
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5205
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CommandButton treurefiltre 
      Height          =   270
      Left            =   30
      Picture         =   "formestocseguretat.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Eliminar totes les linies"
      Top             =   555
      Width           =   240
   End
   Begin VB.TextBox filtre 
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   0
      Left            =   315
      TabIndex        =   34
      ToolTipText     =   "Pots buscar valors separats per comes i a client pots posar el codi de client."
      Top             =   555
      Width           =   630
   End
   Begin VB.Frame frameseleccio 
      Caption         =   "Escullir material"
      Height          =   3600
      Left            =   4200
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   6660
      Begin VB.ListBox List1 
         Height          =   2535
         ItemData        =   "formestocseguretat.frx":1BB2
         Left            =   3675
         List            =   "formestocseguretat.frx":1BB4
         Style           =   1  'Checkbox
         TabIndex        =   50
         Top             =   525
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.CommandButton bvariesfamilies 
         Height          =   330
         Index           =   5
         Left            =   6135
         Picture         =   "formestocseguretat.frx":1BB6
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Agrupar per varies families "
         Top             =   1200
         Width           =   345
      End
      Begin VB.CommandButton bvariesfamilies 
         Height          =   330
         Index           =   4
         Left            =   3285
         Picture         =   "formestocseguretat.frx":2140
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Agrupar per varies families "
         Top             =   1185
         Width           =   345
      End
      Begin VB.CommandButton bvariesfamilies 
         Height          =   330
         Index           =   3
         Left            =   6135
         Picture         =   "formestocseguretat.frx":26CA
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Agrupar per varies families "
         Top             =   855
         Width           =   345
      End
      Begin VB.CommandButton bvariesfamilies 
         Height          =   330
         Index           =   2
         Left            =   3285
         Picture         =   "formestocseguretat.frx":2C54
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Agrupar per varies families "
         Top             =   840
         Width           =   345
      End
      Begin VB.CommandButton bvariesfamilies 
         Height          =   330
         Index           =   1
         Left            =   6135
         Picture         =   "formestocseguretat.frx":31DE
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Agrupar per varies families "
         Top             =   510
         Width           =   345
      End
      Begin VB.CommandButton bvariesfamilies 
         Height          =   330
         Index           =   0
         Left            =   3285
         Picture         =   "formestocseguretat.frx":3768
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Agrupar per varies families "
         Top             =   510
         Width           =   345
      End
      Begin VB.ComboBox famad 
         Height          =   315
         Left            =   705
         TabIndex        =   3
         Top             =   1170
         Width           =   2580
      End
      Begin VB.ComboBox subfamad 
         Height          =   315
         Left            =   3645
         TabIndex        =   4
         Tag             =   "famad"
         Top             =   1155
         Width           =   2490
      End
      Begin VB.ComboBox subfammat 
         Height          =   315
         Left            =   3645
         TabIndex        =   7
         Tag             =   "fammat"
         Top             =   510
         Width           =   2490
      End
      Begin VB.ComboBox subfamcol 
         Height          =   315
         Left            =   3645
         TabIndex        =   6
         Tag             =   "famcol"
         Top             =   840
         Width           =   2490
      End
      Begin VB.ComboBox fammat 
         Height          =   315
         Left            =   705
         TabIndex        =   8
         Top             =   510
         Width           =   2580
      End
      Begin VB.ComboBox famcol 
         Height          =   315
         Left            =   705
         TabIndex        =   5
         Top             =   840
         Width           =   2580
      End
      Begin VB.TextBox kgestoc 
         DataField       =   "micres"
         DataSource      =   "palets"
         Height          =   285
         Left            =   3585
         TabIndex        =   32
         Top             =   2385
         Width           =   660
      End
      Begin VB.CommandButton Command2 
         Height          =   360
         Left            =   5670
         Picture         =   "formestocseguretat.frx":3CF2
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Cancelar"
         Top             =   3150
         Width           =   840
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   4830
         Picture         =   "formestocseguretat.frx":427C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Acceptar canvis"
         Top             =   3150
         Width           =   840
      End
      Begin VB.TextBox iamplemax 
         DataField       =   "Ample"
         DataSource      =   "palets"
         Height          =   285
         Left            =   5025
         TabIndex        =   28
         Top             =   2220
         Width           =   585
      End
      Begin VB.CheckBox imicrop 
         Caption         =   "Microperforat"
         DataField       =   "microperforat"
         DataSource      =   "palets"
         Height          =   285
         Left            =   285
         TabIndex        =   19
         Top             =   2400
         Width           =   1365
      End
      Begin VB.ComboBox iobert 
         DataField       =   "obert"
         DataSource      =   "palets"
         Height          =   315
         ItemData        =   "formestocseguretat.frx":4806
         Left            =   1695
         List            =   "formestocseguretat.frx":4813
         TabIndex        =   18
         Top             =   1965
         Width           =   615
      End
      Begin VB.ComboBox icares 
         DataField       =   "carestractat"
         DataSource      =   "palets"
         Height          =   315
         ItemData        =   "formestocseguretat.frx":4820
         Left            =   885
         List            =   "formestocseguretat.frx":482D
         TabIndex        =   17
         Top             =   1965
         Width           =   615
      End
      Begin VB.ComboBox itl 
         DataField       =   "semielaborat"
         DataSource      =   "palets"
         Height          =   315
         ItemData        =   "formestocseguretat.frx":483A
         Left            =   255
         List            =   "formestocseguretat.frx":4844
         TabIndex        =   16
         Top             =   1965
         Width           =   615
      End
      Begin VB.TextBox iespesor 
         DataField       =   "micres"
         DataSource      =   "palets"
         Height          =   285
         Left            =   3585
         TabIndex        =   15
         Top             =   1980
         Width           =   660
      End
      Begin VB.TextBox iplegat 
         DataField       =   "Plegat"
         DataSource      =   "palets"
         Height          =   285
         Left            =   2400
         TabIndex        =   14
         Top             =   1980
         Width           =   555
      End
      Begin VB.TextBox iamplemin 
         DataField       =   "Ample"
         DataSource      =   "palets"
         Height          =   285
         Left            =   5025
         TabIndex        =   13
         Top             =   1950
         Width           =   585
      End
      Begin VB.TextBox isolapa 
         DataField       =   "solapa"
         DataSource      =   "palets"
         Height          =   285
         Left            =   2985
         TabIndex        =   12
         Top             =   1980
         Width           =   555
      End
      Begin VB.Label etcanviskgestoc 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg d'estoc de seguretat modificar-los desde la reixa."
         ForeColor       =   &H00ED823A&
         Height          =   210
         Left            =   1830
         TabIndex        =   43
         Top             =   2565
         Width           =   3885
      End
      Begin VB.Label lblLabels 
         Caption         =   "Kg d'estoc de seguretat:"
         Height          =   255
         Index           =   5
         Left            =   1815
         TabIndex        =   33
         Top             =   2415
         Width           =   1950
      End
      Begin VB.Label lblLabels 
         Caption         =   "Max:"
         Height          =   255
         Index           =   4
         Left            =   4680
         TabIndex        =   29
         Top             =   2265
         Width           =   630
      End
      Begin VB.Label lblLabels 
         Caption         =   "Min:"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   27
         Top             =   1995
         Width           =   630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Obert"
         Height          =   300
         Index           =   1
         Left            =   1785
         TabIndex        =   26
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cares Tractat"
         Height          =   300
         Index           =   2
         Left            =   735
         TabIndex        =   25
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "T/L"
         Height          =   300
         Index           =   3
         Left            =   375
         TabIndex        =   24
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label lblLabels 
         Caption         =   "Espesor:"
         Height          =   255
         Index           =   15
         Left            =   3630
         TabIndex        =   23
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label lblLabels 
         Caption         =   "Plegat:"
         Height          =   255
         Index           =   3
         Left            =   2430
         TabIndex        =   22
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ample:"
         Height          =   255
         Index           =   2
         Left            =   5100
         TabIndex        =   21
         Top             =   1665
         Width           =   630
      End
      Begin VB.Label lblLabels 
         Caption         =   "Solapa:"
         Height          =   255
         Index           =   0
         Left            =   3015
         TabIndex        =   20
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Familia "
         Height          =   285
         Left            =   2205
         TabIndex        =   11
         Top             =   300
         Width           =   750
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Subfamilia "
         Height          =   285
         Left            =   4065
         TabIndex        =   10
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Mat:         Col:          Ad:"
         Height          =   1020
         Left            =   390
         TabIndex        =   9
         Top             =   450
         Width           =   345
      End
   End
   Begin VB.CommandButton botorefrescar 
      Height          =   480
      Left            =   14970
      Picture         =   "formestocseguretat.frx":484E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Refrescar"
      Top             =   75
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   4245
      Left            =   255
      TabIndex        =   0
      Top             =   825
      Width           =   15420
      _ExtentX        =   27199
      _ExtentY        =   7488
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detall de la linia d'estoc de seguretat."
      Height          =   4155
      Left            =   195
      TabIndex        =   36
      Top             =   5160
      Width           =   15510
      Begin MSDBGrid.DBGrid reixadetall 
         Bindings        =   "formestocseguretat.frx":4DD8
         Height          =   3810
         Left            =   90
         OleObjectBlob   =   "formestocseguretat.frx":4DED
         TabIndex        =   37
         Top             =   270
         Width           =   15285
      End
   End
   Begin VB.CommandButton exportaraxls 
      BackColor       =   &H00F0F0F0&
      Height          =   465
      Left            =   14370
      Picture         =   "formestocseguretat.frx":57BC
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Exportar a Excel la sel.lecció"
      Top             =   90
      Width           =   600
   End
   Begin VB.Label etcalculant 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculant... Un moment sisplau..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   405
      Left            =   5175
      TabIndex        =   41
      Top             =   90
      Visible         =   0   'False
      Width           =   6240
   End
End
Attribute VB_Name = "formestocseguretat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ultimfiltre As Integer
Dim whereultimfiltre As String
Dim camps(100, 4) As String
Dim iniconfigreixa As String
Dim rstseguretat As Recordset
Dim dbconsulta As Database
Dim rstdetall As Recordset
Dim veditant As String
Function cadbl(ByVal valo As Variant) As Double
  Dim vs As String
  If Not IsNumeric(valo) Or atrim(valo) = "" Then valo = 0
  vs = IIf(simboldecimal = ",", ".", ",")
  valo = substituir(atrim(valo), vs, simboldecimal)
  cadbl = CDbl(atrim(valo))
End Function
Private Sub beliminar_Click()
   Dim vid As Long
   If reixa.row < 1 Then MsgBox "Primer escull una linia per eliminar", vbCritical, "Error": Exit Sub
   If MsgBox("Segur que vols eliminar aquesta linia de control d 'estoc de seguretat?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   vid = cadbl(reixa.TextMatrix(reixa.row, 0))
   dbtmp.Execute "delete * from estocseguretat where id=" + atrim(vid)
   wait 1
   borrarelfiltre
   
End Sub

Private Sub botoafegir_Click()
   veditant = ""
   ensenyar_frameseleccio True
   lblLabels(5).Visible = True
   kgestoc.Visible = True
   etcanviskgestoc.Visible = False
  
   wait 1
   borrarelfiltre
End Sub
Sub ensenyar_frameseleccio(vensenyar As Boolean)
   Dim objecte As Object
    For Each objecte In formestocseguretat
     If objecte.Name <> "datadetall" Then
      If objecte.Name <> "frameseleccio" And objecte.Container.Name <> "frameseleccio" Then
        objecte.Enabled = Not vensenyar
      End If
     End If
   Next
   'MsgBox famad
   frameseleccio.Visible = vensenyar
End Sub
Private Sub botorefrescar_Click()
   etcalculant.Visible = True: DoEvents
   calcular_estocdeseguretat
   borrarelfiltre
   etcalculant.Visible = False: DoEvents
End Sub
Sub calcular_estocdeseguretat()
  Dim consultamicres As String
  Dim criterifamilia As String
  Dim criteridebusqueda As String
  
  calcular_estocs
End Sub
Sub calcular_estocs()
   Dim rst As Recordset
   Set rst = dbconsulta.OpenRecordset("select * from estocseguretat")
   dbconsulta.Execute "delete * from estocseguretat_linies"
   If Not rst.EOF Then rst.MoveLast: rst.MoveFirst
   While Not rst.EOF
     etcalculant = "Calculant ---> " + atrim(rst.AbsolutePosition + 1) + "/" + atrim(rst.RecordCount) + "        Un moment sisplau..."
     DoEvents
     calcular_estoc_aterra_i_assignat rst
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Function calcular_assignat(rst As Recordset, vComandes As String, vMtrsalgrup As Double, vDescripciogrup As String) As Double
  Dim rst2 As Recordset
  Set rst2 = dbtmp.OpenRecordset("select *  from parcials where idpalet=" + atrim(cadbl(rst!idpalet)) + " and idbobina=" + atrim(cadbl(rst!idbobina)) + " and not utilitzada")
  While Not rst2.EOF
    If cadbl(rst2!comanda) >= 3000 Then
       calcular_assignat = calcular_assignat + cadbl(rst2!metres)
       vComandes = vComandes + " " + atrim(rst2!comanda)
        Else
          If cadbl(rst2!comanda) > 1000 Then
            vDescripciogrup = vDescripciogrup + "  G-" + atrim(rst2!comanda)
            vMtrsalgrup = vMtrsalgrup + cadbl(rst2!metres)
          End If
    End If
    rst2.MoveNext
  Wend
  Set rst2 = Nothing
End Function
Function substituir(cadena As String, buscar As String, canviar As String) As String
   If buscar = canviar Then GoTo fi
   While InStr(1, cadena, buscar) > 0
    comença = InStr(1, cadena, buscar) - 1
    If comença < 1 Then substituir = cadena: Exit Function
    acaba = comença + Len(buscar) + 1
    cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
   Wend
fi:
   substituir = cadena
   'MsgBox linia
End Function
Function convertirKgaMtrs(vmtrs As Double, rstcomprat As Recordset) As Double
   'els metres els passo en negatiu perquè aquesta funcio els vol en negatiu si son kilos i positius si son metres
    vmtrs = vmtrs * -1
    convertirKgaMtrs = compramat.conversiokilos(cadbl(rstcomprat!codimaterial), cadbl(rstcomprat!ample), vmtrs, IIf(cadbl(rstcomprat!grmm2) > 0, cadbl(rstcomprat!grmm2), cadbl(rstcomprat!micres)), atrim(rstcomprat!semielaborat), cadbl(rstcomprat!solapa))
    convertirKgaMtrs = Redondejar(convertirKgaMtrs, 0)
End Function
Function convertirKgaMtrs_bobines(vmtrs As Double, rstcomprat As Recordset) As Double
   'els metres els passo en negatiu perquè aquesta funcio els vol en negatiu si son kilos i positius si son metres
    vmtrs = vmtrs * -1
    convertirKgaMtrs_bobines = compramat.conversiokilos(cadbl(rstcomprat!codimatprognou), cadbl(rstcomprat!ample), vmtrs, IIf(cadbl(rstcomprat!grmsm2) > 0, cadbl(rstcomprat!grmsm2), cadbl(rstcomprat!micres)), atrim(rstcomprat!semielaborat), cadbl(rstcomprat!solapa))
    convertirKgaMtrs_bobines = Redondejar(convertirKgaMtrs_bobines, 0)
End Function
Sub buscarmaterialdelareserva(rst As Recordset, rstmat As Recordset)
    Dim d As String
    d = " familia=" + atrim(rst!familia)
    d = d + " and subfamilia=" + atrim(rst!subfamilia)
    d = d + " and familiacol=" + atrim(rst!familiacol)
    d = d + " and subfamiliacol=" + atrim(rst!subfamiliacol)
    d = d + " and familiaad=" + atrim(rst!familiaad)
    d = d + " and subfamiliaad=" + atrim(rst!subfamiliaad)
   Set rstmat = dbtmp.OpenRecordset("select * from materials where " + d)
End Sub
Function convertirKgaMtrs_reserva(vmtrs As Double, rstcomprat As Recordset) As Double
    Dim rstmat As Recordset
    buscarmaterialdelareserva rstcomprat, rstmat
    If rstmat.EOF Then Exit Function
   'els metres els passo en negatiu perquè aquesta funcio els vol en negatiu si son kilos i positius si son metres
    vmtrs = vmtrs * -1
    convertirKgaMtrs_reserva = compramat.conversiokilos(cadbl(rstmat!codi), cadbl(rstcomprat!ample), vmtrs, cadbl(rstcomprat!espesor), atrim(rstcomprat!semielaborat), cadbl(rstcomprat!solapa))
    convertirKgaMtrs_reserva = Redondejar(convertirKgaMtrs_reserva, 0)
End Function

Function calcular_compratlinkat(rst As Recordset) As Double
   Dim rstlk As Recordset
   Set rstlk = dbcompres.OpenRecordset("select sum(kgcompra) as Tlk from comandesxlinia where numcomanda<>0 and idliniacompra=" + atrim(rst!idliniacompra) + " group by idliniacompra")
   calcular_compratlinkat = Redondejar(cadbl(rstlk!Tlk), 0)
End Function
Sub calcular_estoc_aterra_i_assignat(rst As Recordset)
  Dim criterifamilia As String
  Dim consultamicres As String
  Dim criteridebusqueda As String
  Dim vTaterra As Double
  Dim vTassignat As Double
  Dim rst2 As Recordset
  Dim rstreserva As Recordset
  Dim rstcompra As Recordset
  Dim vValues As String
  Dim vAssignat As Double
  Dim vReservat As Double
  Dim vTreservat As Double
  Dim vCompratlinkat As Double
  Dim vCompratestoc As Double
  Dim vTcompratestoc As Double
  Dim vTcompratlk As Double
  Dim vTerra As Double
  Dim vComandes As String
  Dim vDescripciogrup As String
  Dim vMtrsalgrup As Double
  
  criterifamilia = crear_criteri_familia(rst)
  If atrim(rst!espesor) <> "" Then
     consultamicres = " and  micres=" + atrim(cadbl(rst!espesor))
    Else: consultamicres = " and  micres>1"
  End If
  
  'guardo les bobines filtrades
  criteridebusqueda = criterifamilia + altrescriteris(rst, "palets.") + consultamicres + " and (ample>=" + passaradecimalpunt(atrim(rst!AmpleMin)) + " AND ample<=" + (passaradecimalpunt(atrim(rst!amplemax))) + ") "
  Set rst2 = dbtmp.OpenRecordset("SELECT Palets.*, materials.familia, materials.subfamilia, materials.proveidor, Bobines.Idbobina, Bobines.disponible as mtrsdisponibles FROM Bobines RIGHT JOIN (Palets INNER JOIN materials ON Palets.codimatprognou = materials.codi) ON Bobines.Idpalet = Palets.Idpalet Where " + criteridebusqueda, dbOpenSnapshot, dbReadOnly)
  
  While Not rst2.EOF
    vAssignat = 0
    vTerra = 0
    vComandes = 0
    vDescripciogrup = ""
    vComandes = ""
    vMtrsalgrup = 0
   ' If rst2!idpalet = 48170 Then Stop
    vAssignat = Redondejar(convertirKgaMtrs_bobines(calcular_assignat(rst2, vComandes, vMtrsalgrup, vDescripciogrup) * -1, rst2), 0)
    vMtrsalgrup = Redondejar(convertirKgaMtrs_bobines(vMtrsalgrup * -1, rst2), 0)
    If (vAssignat > 0 Or vMtrsalgrup > 0) Or cadbl(rst2!mtrsdisponibles) > 0 Then
      If vAssignat = 0 And cadbl(rst2!mtrsdisponibles) > 0 Then vTerra = Redondejar(convertirKgaMtrs_bobines(cadbl(rst2!mtrsdisponibles) * -1, rst2), 0)
      vValues = atrim(cadbl(rst!ID)) + "," + atrim(cadbl(rst2!idpalet)) + "," + atrim(cadbl(rst2!idbobina)) + "," + atrim(vTerra) + "," + atrim(vAssignat) + ",'" + treure_apostruf(vComandes) + "'," + atrim(cadbl(vMtrsalgrup)) + ",'" + atrim(treure_apostruf(vDescripciogrup)) + "'"
      dbconsulta.Execute "insert into estocseguretat_linies (id_estocseguretat,palet,bobina,disponible,assignat,observacions,assignatsagrup,observacionsgrup) values (" + vValues + ")"
      vTassignat = vTassignat + vAssignat
      vTaterra = vTaterra + vTerra
    End If
    rst2.MoveNext
  Wend
  
  'guardo les compres Linkat i Estoc
  criteridebusqueda = criterifamilia + altrescriteris(rst, "") + consultamicres + " and (ample>=" + passaradecimalpunt(atrim(rst!AmpleMin)) + " AND ample<=" + (passaradecimalpunt(atrim(rst!amplemax))) + ") "
 ' Clipboard.Clear
 ' Clipboard.SetText "SELECT * from liniescompra Where " + criteridebusqueda
  Set rst2 = dbcompres.OpenRecordset("SELECT * from liniescompra Where not totentregat and " + criteridebusqueda, dbOpenSnapshot, dbReadOnly)
  While Not rst2.EOF
    vCompratlinkat = calcular_compratlinkat(rst2)
    vCompratestoc = rst2!quantitatkg - vCompratlinkat 'passar kg a metres
    'vCompratlinkat = convertirKgaMtrs(vCompratlinkat, rst2)
    'vCompratestoc = convertirKgaMtrs(vCompratestoc, rst2)  'passar kg a metres
    vValues = atrim(cadbl(rst!ID)) + "," + atrim(rst2!idliniacompra) + "," + atrim(vCompratlinkat) + "," + atrim(vCompratestoc)
    dbconsulta.Execute "insert into estocseguretat_linies (id_estocseguretat,idliniacompra,estoccompratlk,estoccompratestoc) values (" + vValues + ")"
    vTcompratlk = vTcompratlk + vCompratlinkat
    vTcompratestoc = vTcompratestoc + vCompratestoc
    rst2.MoveNext
  Wend
  

  'guardo les reserves filtrades
  consultamicres = substituir(consultamicres, "micres", "espesor")
  criteridebusqueda = criterifamilia + altrescriteris(rst, "") + consultamicres + " and (ample>=" + passaradecimalpunt(atrim(rst!AmpleMin)) + " AND ample<=" + (passaradecimalpunt(atrim(rst!amplemax))) + ") "
  Set rstreserva = dbtmp.OpenRecordset("SELECT * from reserves Where metresreservats>0 and " + criteridebusqueda, dbOpenSnapshot, dbReadOnly)
  While Not rstreserva.EOF
    vReservat = Redondejar(convertirKgaMtrs_reserva(cadbl(rstreserva!metresreservats) * -1, rstreserva), 0)
    vValues = atrim(cadbl(rst!ID)) + "," + atrim(cadbl(rstreserva!idreserva)) + "," + atrim(vReservat)
    dbconsulta.Execute "insert into estocseguretat_linies (id_estocseguretat,idreserva,reservat) values (" + vValues + ")"
    vTreservat = vTreservat + vReservat
    rstreserva.MoveNext
  Wend
  
  rst.Edit
  rst!estocassignat = vTassignat
  rst!estocreservat = vTreservat
  rst!estocterra = vTaterra
  rst!estoccompratlk = vTcompratlk
  rst!estoccompratestoc = vTcompratestoc
  rst!estoccomprat = vTcompratlk + vTcompratestoc
  rst!diferencialestocseguretat = (vTaterra + vTcompratestoc) - cadbl(rst!kgestoc)
  rst.Update
  Set rst2 = Nothing
End Sub
Function aatrim(va As Variant) As String
   aatrim = atrim(va)
   If aatrim = "" Then aatrim = "N"
End Function
Function altrescriteris(rst As Recordset, vdonveelmigelaborat As String) As String
  Dim criteri As String
  Dim desc As String
  desc = atrim(rst!nomfamilia) + "   "
  If atrim(itl) <> "" Then criteri = criteri + " and " + IIf(vdonveelmigelaborat <> "", vdonveelmigelaborat, "") + "semielaborat='" + aatrim(itl) + "' "
  If InStr(1, "PEAD", desc) > 0 Or InStr(1, "PEMD", desc) > 0 Or InStr(1, "PEBD", desc) > 0 Then
    If atrim(rst!carestractat) <> "" Then criteri = criteri + " and carestractat='" + aatrim(rst!carestractat) + "' "
  End If
  If atrim(rst!obert) <> "" Then criteri = criteri + " and obert='" + aatrim(rst!obert) + "' "
  If atrim(rst!plegat) <> "" And cadbl(rst!plegat) > 0 Then criteri = criteri + " and plegat=" + atrim(cadbl(rst!plegat))
  If atrim(rst!solapa) <> "" And cadbl(rst!solapa) > 0 Then criteri = criteri + " and solapa=" + atrim(cadbl(rst!solapa))
  criteri = criteri + " and " + IIf(cabool(rst!microperforat), "", "not") + " microperforat"
  altrescriteris = criteri
End Function

Private Sub bvariesfamilies_Click(Index As Integer)
   If List1.Visible = True Then List1.Visible = False: Exit Sub
   carregar_subfamilies subfammat
   List1.Visible = True
   List1.Left = subfammat.Left
   List1.Top = subfammat.Top
End Sub

Private Sub Command1_Click()
  Dim vsql As String
  Dim rst As Recordset
  vsql = crearconsulta
  If vsql = "" Then Exit Sub
  Set rst = dbtmp.OpenRecordset("select * from estocseguretat where " + vsql)
  If Not rst.EOF Then
      MsgBox "Aquest estoc de seguretat ja existeix, elimina'l primer i despres genera'l.", vbCritical, "Error"
      GoTo fi
  End If
  guardar_registre
  MsgBox "Registre afegit.", vbInformation, "Atenció"
  ensenyar_frameseleccio False
  veditant = ""
fi:
  Set rst = Nothing
    wait 1
   borrarelfiltre
End Sub
Sub guardar_registre()
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from estocseguretat ")
  If veditant <> "" Then
       Set rst = dbtmp.OpenRecordset("select * from estocseguretat Where id = " + atrim(veditant))
       If rst.EOF Then MsgBox "Error guardant la modificacio.", vbCritical, "Error": Exit Sub
       rst.Edit
     Else: rst.AddNew
  End If
  If fammat <> "" Then rst!familia = cadbl(fammat.ItemData(fammat.ListIndex))
  If subfammat <> "" Then rst!subfamilia = cadbl(subfammat.ItemData(subfammat.ListIndex))
  If famcol <> "" Then rst!familiacol = cadbl(famcol.ItemData(famcol.ListIndex))
  If subfamcol <> "" Then rst!subfamiliacol = cadbl(subfamcol.ItemData(subfamcol.ListIndex))
  If famad <> "" Then rst!familiaad = cadbl(atrim(famad.ItemData(famad.ListIndex)))
  If subfamad <> "" Then rst!subfamiliaad = cadbl(subfamad.ItemData(subfamad.ListIndex))
  rst!nomfamilia = fammat
  rst!nomsubfamilia = subfammat
  rst!nomfamiliacol = famcol
  rst!nomsubfamiliacol = subfamcol
  rst!nomfamiliaaditiu = famad
  rst!nomsubfamiliaaditiu = subfamad
  rst!plegat = cadbl(cadbl(iplegat))
  rst!solapa = cadbl(cadbl(isolapa))
  rst!carestractat = atrim(icares)
  rst!obert = atrim(iobert)
  rst!microperforat = IIf(imicrop.Value = 1, True, False)
  rst!semielaborat = atrim(itl)
  rst!espesor = cadbl(iespesor)
  
  rst!AmpleMin = cadbl(iamplemin)
  rst!amplemax = cadbl(iamplemax)
  
  rst!kgestoc = cadbl(kgestoc)
  rst.Update
  Set rst = Nothing
End Sub

Function crearconsulta() As String
  Dim vsql As String
  If fammat = "" Or subfammat = "" Or famcol = "" Or subfamcol = "" Or famad = "" Or subfamad = "" Then
     MsgBox "Falta posar families o subfamilies.", vbCritical, "Error": Exit Function
  End If
  If fammat <> "" Then vsql = "familia=" + atrim(fammat.ItemData(fammat.ListIndex))
  If subfammat <> "" Then vsql = vsql + " and subfamilia=" + atrim(subfammat.ItemData(subfammat.ListIndex))
  If famcol <> "" Then vsql = vsql + " and familiacol=" + atrim(famcol.ItemData(famcol.ListIndex))
  If subfamcol <> "" Then vsql = vsql + " and subfamiliacol=" + atrim(subfamcol.ItemData(subfamcol.ListIndex))
  If famad <> "" Then vsql = vsql + " and familiaad=" + atrim(famad.ItemData(famad.ListIndex))
  If subfamad <> "" Then vsql = vsql + " and subfamiliaad=" + atrim(subfamad.ItemData(subfamad.ListIndex))
  
  If fammat = "" Or itl = "" Or icares = "" Or iobert = "" Or iplegat = "" Or isolapa = "" Or iespesor = "" Or iamplemin = "" Or iamplemax = "" Or kgestoc = "" Then
    MsgBox "Hi ha camps obligatoris que no estan posats.", vbCritical, "Error"
    vsql = ""
    GoTo fi
  End If
  vsql = vsql + " and plegat=" + passaradecimalpunt(atrim(cadbl(iplegat)))
  vsql = vsql + " and solapa=" + passaradecimalpunt(atrim(cadbl(isolapa)))
  vsql = vsql + " and carestractat='" + atrim(icares) + "'"
  vsql = vsql + " and obert='" + atrim(iobert) + "'"
  vsql = vsql + " and microperforat=" + IIf(imicrop.Value = 1, "True", "False")
  vsql = vsql + " and semielaborat='" + atrim(itl) + "'"
  vsql = vsql + " and espesor=" + passaradecimalpunt(atrim(cadbl(iespesor)))
  
  vsql = vsql + " and amplemin=" + passaradecimalpunt(atrim(cadbl(iamplemin)))
  vsql = vsql + " and amplemax=" + passaradecimalpunt(atrim(cadbl(iamplemax)))
  
  'vsql = vsql + " and kgestoc=" + atrim(cadbl(kgestoc))
  
fi:
  crearconsulta = vsql
End Function

Private Sub Command2_Click()
  ensenyar_frameseleccio False
  veditant = ""
End Sub

Sub carregarvalorsvariables()
    Dim i As Byte
    camps(i, 1) = "id": camps(i, 2) = "string": camps(i, 3) = "id":  i = i + 1
    camps(i, 1) = "nomfamilia": camps(i, 2) = "string": camps(i, 3) = "Familia": i = i + 1
    camps(i, 1) = "nomsubfamilia": camps(i, 2) = "string": camps(i, 3) = "Subfamilia": i = i + 1
    camps(i, 1) = "nomfamiliacol": camps(i, 2) = "string": camps(i, 3) = "Color": i = i + 1
    camps(i, 1) = "nomsubfamiliacol": camps(i, 2) = "string": camps(i, 3) = "SubColor": i = i + 1
    camps(i, 1) = "nomfamiliaaditiu": camps(i, 2) = "string": camps(i, 3) = "Aditiu": i = i + 1
    camps(i, 1) = "nomsubfamiliaaditiu": camps(i, 2) = "string": camps(i, 3) = "SubAditiu": i = i + 1
    camps(i, 1) = "amplemin": camps(i, 2) = "double": camps(i, 3) = "AmpleMin": i = i + 1
    camps(i, 1) = "amplemax": camps(i, 2) = "double": camps(i, 3) = "AmpleMax": i = i + 1
    camps(i, 1) = "espesor": camps(i, 2) = "double": camps(i, 3) = "Micres": i = i + 1
    camps(i, 1) = "semielaborat": camps(i, 2) = "string": camps(i, 3) = "T/L": i = i + 1
    camps(i, 1) = "carestractat": camps(i, 2) = "string": camps(i, 3) = "Tractat": i = i + 1
    camps(i, 1) = "obert": camps(i, 2) = "string": camps(i, 3) = "Obert": i = i + 1
    camps(i, 1) = "microperforat": camps(i, 2) = "string": camps(i, 3) = "MicroP": i = i + 1
    camps(i, 1) = "plegat": camps(i, 2) = "double": camps(i, 3) = "Plegat": i = i + 1
    camps(i, 1) = "solapa": camps(i, 2) = "double": camps(i, 3) = "Solapa": i = i + 1
    camps(i, 1) = "kgestoc": camps(i, 2) = "double": camps(i, 3) = "E.Seguretat": i = i + 1
    camps(i, 1) = "estocassignat": camps(i, 2) = "double": camps(i, 3) = "E.Assignat": i = i + 1
    camps(i, 1) = "estocreservat": camps(i, 2) = "double": camps(i, 3) = "E.Reservat": i = i + 1
    camps(i, 1) = "estocterra": camps(i, 2) = "double": camps(i, 3) = "E.Terra": i = i + 1
    camps(i, 1) = "estoccomprat": camps(i, 2) = "double": camps(i, 3) = "E.Comprat": i = i + 1
    camps(i, 1) = "estoccompratlk": camps(i, 2) = "double": camps(i, 3) = "E.Compratlk": i = i + 1
    camps(i, 1) = "estoccompratestoc": camps(i, 2) = "double": camps(i, 3) = "E.CompratEstoc": i = i + 1
    camps(i, 1) = "diferencialestocseguretat": camps(i, 2) = "double": camps(i, 3) = "Diferencial": i = i + 1
    camps(i, 1) = "observacions": camps(i, 2) = "string": camps(i, 3) = "Observacions": i = i + 1
    
End Sub

Sub bxrcontrolagafafocus(i As Integer)
  Dim cntrl As Control
  Set cntrl = Screen.ActiveControl
  If cntrl.Text <> "" Then
     If cntrl.Text = camps(cadbl(filtre(i).Tag), 3) Then cntrl.Text = ""
     cntrl.ForeColor = QBColor(0)
   Else:
      If cntrl.Name = "filtre" Then
       cntrl.Text = camps(cadbl(filtre(i).Tag), 3)
       cntrl.ForeColor = &H808080
      End If
  End If
End Sub

Sub buscaritemdata(ccombo As ComboBox, vitem As Long)
     Dim i As Integer
     i = 0
     While i < ccombo.ListCount
        If ccombo.ItemData(i) = vitem Then
            ccombo.ListIndex = i
            GoTo fi
        End If
        i = i + 1
     Wend
fi:
End Sub
Sub netejar_camps_escullirmaterial()
  subfammat.Clear: subfamcol.Clear: subfamad.Clear
  fammat = "": subfammat = "": famcol = "": subfamcol = "": famad = "": subfamad = ""
  itl = "": icares = "": iobert = "": iplegat = "": isolapa = "": iespesor = "": iamplemin = "": iamplemax = "": kgestoc = ""
  List1.Visible = False
End Sub

Private Sub Command3_Click()
  Dim vid As Double
  Dim rst As Recordset
  
  netejar_camps_escullirmaterial
  vid = cadbl(reixa.TextMatrix(reixa.row, numcol("id")))
  veditant = atrim(vid)
  Set rst = dbtmp.OpenRecordset("select * from estocseguretat Where id = " + atrim(vid))
  If rst.EOF Then MsgBox "Error", vbCritical, "Error": Exit Sub
  ensenyar_frameseleccio True
  lblLabels(5).Visible = False
  kgestoc.Visible = False
  etcanviskgestoc.Visible = True
  
  'carrego materials
  If cadbl(rst!familia) Then
    buscaritemdata fammat, cadbl(rst!familia)
    subfammat.SetFocus
    carregar_subfamilies
    If cadbl(rst!subfamilia) > 0 Then buscaritemdata subfammat, cadbl(rst!subfamilia)
  End If
  
  'carrego colorants
  If cadbl(rst!familiacol) Then
    buscaritemdata famcol, cadbl(rst!familiacol)
    subfamcol.SetFocus
    carregar_subfamilies
    If cadbl(rst!subfamiliacol) Then buscaritemdata subfamcol, cadbl(rst!subfamiliacol)
  End If
  
  
  'carrego colorants
  If cadbl(rst!familiaad) Then
    buscaritemdata famad, cadbl(rst!familiaad)
    subfamad.SetFocus
    carregar_subfamilies
    If cadbl(rst!subfamiliaad) Then buscaritemdata subfamad, cadbl(rst!subfamiliaad)
  End If
  
  iplegat = atrim(rst!plegat)
  isolapa = atrim(rst!solapa)
  icares = atrim(rst!carestractat)
  iobert = atrim(rst!obert)
  imicrop.Value = IIf(rst!microperforat, 1, 0)
  itl = atrim(rst!semielaborat)
  iespesor = atrim(rst!espesor)
  
  iamplemin = atrim(rst!AmpleMin)
  iamplemax = atrim(rst!amplemax)
  
  kgestoc = atrim(rst!kgestoc)
  Set rst = Nothing
  frameseleccio.Refresh
  
End Sub

Private Sub exportaraxls_Click()
   generar_xls
End Sub
Sub generar_xls()
   Dim i As Byte
   Dim rst As Recordset
   Dim linia As String
   Dim vvalor As String
   
   Set rst = dbconsulta.OpenRecordset("select * from estocseguretat " + IIf(whereultimfiltre <> "", " where " + whereultimfiltre, "") + " order by familia,amplemin", dbOpenSnapshot, dbReadOnly)
   If rst.EOF Then MsgBox "No hi ha dades per exportar", vbCritical, "Error": Exit Sub
   Open "c:\temp\consultaestocseguretat.csv" For Output As #1
   If Not rst.EOF Then
    i = 1
    While camps(i, 1) <> ""
       If camps(i, 4) <> "No" Then
         linia = linia + IIf(linia = "", "", ";") + camps(i, 3)
       End If
       i = i + 1
    Wend
    Print #1, linia
   End If
   While Not rst.EOF
    linia = ""
    i = 1
    While camps(i, 1) <> ""
      If camps(i, 4) <> "No" Then
        vvalor = atrim(rst.Fields(camps(i, 1)))
        If UCase(vvalor) = "FALSO" Then vvalor = "N"
        If UCase(vvalor) = "VERDADERO" Then vvalor = "S"
        linia = linia + IIf(linia = "", "", ";") + vvalor
      End If
      i = i + 1
    Wend
    Print #1, linia
    rst.MoveNext
   Wend
   Close #1
   wait 2
   obrir_document "c:\temp\consultaestocseguretat.csv"
      
End Sub

Private Sub filtre_GotFocus(Index As Integer)
  bxrcontrolagafafocus Index
  ultimfiltre = Index
End Sub

Private Sub filtre_LostFocus(Index As Integer)
 Dim noufiltre As String
  
  If Index = 998 Then whereultimfiltre = "": Exit Sub
  noufiltre = crearfiltre
  If filtre(ultimfiltre).Text = "" Then
    filtre(ultimfiltre).Text = camps(cadbl(filtre(ultimfiltre).Tag), 3)
    filtre(ultimfiltre).ForeColor = &H808080
  End If
  If noufiltre <> whereultimfiltre Or Index = 999 Then
     If noufiltre <> "" Then poblarlareixa noufiltre
     If noufiltre = "" Then borrarelfiltre
  End If
  If Index = 999 And noufiltre = "" Then
     poblarlareixa
  End If
  ratoli "normal"
  reixa.Visible = True
  whereultimfiltre = noufiltre
  botorefrescar.Tag = noufiltre ' el guardo pel llistat
  
End Sub
Function crearwere(i As Integer) As String
   Dim w As String
   Dim j As Integer
   If filtre(i) = "" Then Exit Function
   j = cadbl(filtre(i).Tag)
   If camps(j, 2) = "date" Then
      If IsDate(filtre(i)) Then
         crearwere = camps(j, 1) + "=#" + format(filtre(i), "mm/dd/yy") + "# "
      End If
      Exit Function
   End If
   If InStr(1, camps(j, 2), "string") > 0 Or camps(j, 1) = "comanda" Then
         crearwere = possarweres(camps(j, 1), "LIKE", treure_apostruf(filtre(i)))
       Exit Function
   End If
   If InStr(1, filtre(i), ">") > 0 Or InStr(1, filtre(i), "<") > 0 Or InStr(1, filtre(i), "=") > 0 Then
      crearwere = camps(j, 1) + filtre(i)
     Else: crearwere = camps(j, 1) + "=" + passaradecimalpunt(atrim(cadbl(filtre(i))))
   End If
   

End Function
Function possarweres(ByVal camp As String, condicio As String, ByVal filtre As String) As String
  Dim re As String
'camps(j, 1) + " LIKE '*" + treure_apostruf(filtre(i)) + "*'"
  filtre = filtre + ","
  If camp = "nomclient" And cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1)) > 0 Then camp = "codiclient"
  While InStr(1, filtre, ",") > 0 And filtre <> ""
    If camp <> "codiclient" Then
       re = IIf(re <> "", re + " or ", "") + camp + " like '*" + Mid(filtre, 1, InStr(1, filtre, ",") - 1) + "*'"
      Else: re = IIf(re <> "", re + " or ", "") + camp + " =" + atrim(cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1))) + ""
    End If
    filtre = Mid(filtre, InStr(1, filtre, ",") + 1)
  Wend
  If re <> "" Then re = "(" + re + ")"
  possarweres = re
End Function
Function crearfiltre() As String
  Dim i As Integer
  Dim were As String
  Dim w As String
  For i = 0 To filtre.Count - 1
    If filtre(i).Text <> camps(cadbl(filtre(i).Tag), 3) And Not (camps(cadbl(filtre(i).Tag), 1) = "comanda" And cadbl(filtre(i)) > 0) Then
      w = crearwere(i)
      If were = "" Then
         were = w
        Else: If w <> "" Then were = were + " and " + w
      End If
    End If
  Next i
  crearfiltre = were
End Function

Private Sub Form_Load()
  Set dbconsulta = dbtmp
  carregarvalorsvariables
  iniconfigreixa = "reixaestocseguretat.ini"
  carregartamanyform
  carregar_combo_families
  configreixa
  poblarlareixa
End Sub
Sub carregar_combo_families()
  Dim rstfam As Recordset
  
  Set rstfam = dbtmpb.OpenRecordset("select * from familiesmaterials where codi>499")
  fammat.Clear
  While Not rstfam.EOF
    fammat.AddItem atrim(rstfam!descripcio)
    fammat.ItemData(fammat.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
  Set rstfam = dbtmpb.OpenRecordset("select * from familiescolorants where codi>499")
  famcol.Clear
  While Not rstfam.EOF
    famcol.AddItem atrim(rstfam!descripcio)
    famcol.ItemData(famcol.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
  Set rstfam = dbtmpb.OpenRecordset("select * from familiesaditius where codi>499")
  famad.Clear
  While Not rstfam.EOF
    famad.AddItem atrim(rstfam!descripcio)
    famad.ItemData(famad.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
End Sub
Sub carregar_subfamilies(Optional combof As Control)
  Dim rstsub As Recordset
  Dim combo As Control
  Dim subfamilia As String
  
  Set combo = formestocseguretat.ActiveControl
  If Not combof Is Nothing Then Set combo = combof
  If formestocseguretat.Controls(combo.Tag).ListIndex = -1 And combof Is Nothing Then MsgBox "Primer has d'escullir la familia": Exit Sub
  'If combo.ListIndex = -1 Then combo.Clear: Exit Sub
  If combo.Name = "subfammat" And fammat.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(fammat.ItemData(fammat.ListIndex))): subfamilia = "subfamiliesmaterials"
  If combo.Name = "subfamcol" And famcol.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(famcol.ItemData(famcol.ListIndex))): subfamilia = "subfamiliescolorants"
  If combo.Name = "subfamad" And famad.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(famad.ItemData(famad.ListIndex))): subfamilia = "subfamiliesaditius"
    combo.Clear

  If subfamilia <> "" Then
     Set rstsub = dbtmpb.OpenRecordset("select codi,descripcio from " + subfamilia + " where " + r) '+ " and descripcio like '*" + treure_apostrof(subfammat.Text) + "*'")
    Else: Exit Sub
  End If
  
  While Not rstsub.EOF
    combo.AddItem atrim(rstsub!descripcio)
    combo.ItemData(combo.NewIndex) = cadbl(rstsub!codi)
    List1.AddItem atrim(rstsub!descripcio)
    List1.ItemData(List1.NewIndex) = cadbl(rstsub!codi)
    rstsub.MoveNext
  Wend
  
  
End Sub
Sub carregartamanyform()
  If cadbl(llegir_ini("TamanyFormEstocseguretat", "ample", iniconfigreixa)) > 0 Then
   formestocseguretat.Tag = "canvianttamany"
   formestocseguretat.Width = llegir_ini("TamanyFormEstocseguretat", "ample", iniconfigreixa)
   formestocseguretat.Height = llegir_ini("TamanyFormEstocseguretat", "alt", iniconfigreixa)
   formestocseguretat.Tag = ""
  End If
End Sub

Private Sub Form_Resize()
If formestocseguretat.Height - reixa.Top - 800 < 1 Then Exit Sub
   reixa.Width = formestocseguretat.Width - 500
   formestocseguretat.Height = 10000
   botorefrescar.Left = formestocseguretat.Width - 1000
   exportaraxls.Left = botorefrescar.Left - (exportaraxls.Width + 10)
   If formestocseguretat.Tag <> "canvianttamany" Then
    escriure_ini "TamanyFormEstocseguretat", "ample", atrim(formestocseguretat.Width), iniconfigreixa
    escriure_ini "TamanyFormEstocseguretat", "alt", atrim(formestocseguretat.Height), iniconfigreixa
   End If
End Sub

Private Sub iamplemax_LostFocus()
   ' iamplemax = treurelacoma(iamplemax)
End Sub

Private Sub iamplemin_LostFocus()
   ' iamplemin = treurelacoma(iamplemin)
End Sub

Private Sub iespesor_LostFocus()
   'iespesor = treurelacoma(iespesor)
End Sub
Function treurelacoma(v As String) As String
  If InStr(1, v, ",") > 0 Then v = substituir(v, ",", ".")
  treurelacoma = v
End Function

Private Sub iplegat_LostFocus()
'iplegat = treurelacoma(iplegat)
End Sub

Private Sub isolapa_LostFocus()
  '  isolapa = treurelacoma(isolapa)
End Sub

Private Sub kgestoc_Change()
 ' kgestoc = treurelacoma(kgestoc)
End Sub

Private Sub reixa_Click()
   reixadetall.Visible = False
   If reixa.col = numcol("E.Terra") Then possar_info_terra
   If reixa.col = numcol("E.Reservat") Then possar_info_reserva
   If reixa.col = numcol("E.Comprat") Then possar_info_comprat
   If reixa.col = numcol("E.Compratlk") Then possar_info_comprat_lk False
   If reixa.col = numcol("E.CompratEstoc") Then possar_info_comprat_lk True
   If reixa.col = numcol("E.Assignat") Then possar_info_assignat
End Sub
Sub possar_info_comprat()
   Dim vid As Double
   Dim vsql As String
   vid = cadbl(reixa.TextMatrix(reixa.row, numcol("id")))
   If vid = 0 Then Exit Sub
   vsql = "select idliniacompra from estocseguretat_linies where idliniacompra>0 and id_estocseguretat=" + atrim(vid)
   Set rstdetall = dbtmp.OpenRecordset(vsql)
   If rstdetall.EOF Then Exit Sub
   If cadbl(rstdetall!idliniacompra) > 0 Then
    Set rstdetall = dbtmp.OpenRecordset("SELECT capcalera.numcomanda as [NºCompra], capcalera.nomprov as [Proveïdor], liniescompra.codimaterial AS CodiMat, liniescompra.nommaterial AS Nom_Material, liniescompra.quantitatkg AS Kg_Comprats FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra Where idliniacompra in (" + vsql + ")") ' + atrim(cadbl(rstdetall!idliniacompra)))
    If Not rstdetall.EOF Then
        Set datadetall.Recordset = rstdetall
        reixadetall.Refresh
    End If
    reixadetall.Visible = True
   End If
   Set rstdetall = Nothing
End Sub
Sub possar_info_comprat_lk(vestoc As Boolean)
   Dim vid As Double
   Dim vsql As String
   
   vid = cadbl(reixa.TextMatrix(reixa.row, numcol("id")))
   If vid = 0 Then Exit Sub
   vsql = "select idliniacompra from estocseguretat_linies where idliniacompra>0 and id_estocseguretat=" + atrim(vid)
   Set rstdetall = dbtmp.OpenRecordset(vsql)
   If rstdetall.EOF Then Exit Sub
   If cadbl(rstdetall!idliniacompra) > 0 Then
    Set rstdetall = dbtmp.OpenRecordset("SELECT comandavisual as [NºComanda],kgcompra as [Kg_Comprats] from comandesxlinia where " + IIf(vestoc, "numcomanda=0 and ", "numcomanda>0 and ") + " idliniacompra in (" + vsql + ")") ' + atrim(cadbl(rstdetall!idliniacompra)))
    If Not rstdetall.EOF Then
        Set datadetall.Recordset = rstdetall
        reixadetall.Refresh
        reixadetall.Visible = True
    End If
    
   End If
   Set rstdetall = Nothing
End Sub
Sub possar_info_reserva()
   Dim vreserva As Double
   vreserva = cadbl(reixa.TextMatrix(reixa.row, numcol("id")))
   If vreserva = 0 Then Exit Sub
   Set rstdetall = dbtmp.OpenRecordset("SELECT  percomandaoclient.numcomanda as [NºComanda], percomandaoclient.metres as [Metres] FROM (Estocseguretat_linies INNER JOIN Reserves ON Estocseguretat_linies.idreserva = Reserves.idreserva) INNER JOIN percomandaoclient ON Reserves.idreserva = percomandaoclient.idreserva where numcomanda>0 and id_estocseguretat=" + atrim(vreserva))
   If Not rstdetall.EOF Then
    Set datadetall.Recordset = rstdetall
    reixadetall.Refresh
   End If
   reixadetall.Visible = True
   Set rstdetall = Nothing
End Sub
Sub possar_info_terra()
   Dim vid As Double
   vid = cadbl(reixa.TextMatrix(reixa.row, numcol("id")))
   If vid = 0 Then Exit Sub
   Set rstdetall = dbtmp.OpenRecordset("SELECT  palet as [NºPalet],bobina as [NºBobina], disponible as [Kg_Disponible], assignatsagrup as [Kg_grup],observacionsgrup as [Obs_del grup]   from estocseguretat_linies where palet>0 and id_estocseguretat=" + atrim(vid) + " and (assignat=0 or assignatsagrup>0)  order by palet,bobina")
   If Not rstdetall.EOF Then
    Set datadetall.Recordset = rstdetall
    reixadetall.Refresh
   End If
   reixadetall.Visible = True
   Set rstdetall = Nothing
End Sub
Sub possar_info_assignat()
   Dim vid As Double
   vid = cadbl(reixa.TextMatrix(reixa.row, numcol("id")))
   If vid = 0 Then Exit Sub
   Set rstdetall = dbtmp.OpenRecordset("SELECT  palet as [NºPalet],bobina as [NºBobina], assignat as [Kg_Assignat], observacions as [Observació] from estocseguretat_linies where palet>0 and id_estocseguretat=" + atrim(vid) + " and assignat>0 order by palet,bobina")
   If Not rstdetall.EOF Then
    Set datadetall.Recordset = rstdetall
    reixadetall.Refresh
   End If
   reixadetall.Visible = True
   Set rstdetall = Nothing
End Sub



Private Sub reixa_DblClick()
    Dim vobs As String
    Dim vid As Double
    vid = cadbl(reixa.TextMatrix(reixa.row, numcol("id")))
    If reixa.col = numcol("Observacions") Then
     vobs = reixa.Text
     vobs = InputBox("Entra la observació d'aquesta linia." + Chr(10) + "ESCRIU UN ESPAI PER BORRAR TOT EL CONTINGUT.", "Observacions", vobs)
     If StrPtr(vobs) = Empty Then Exit Sub
     dbconsulta.Execute "update estocseguretat set observacions='" + atrim(treure_apostruf(vobs)) + "' where id=" + atrim(vid)
     reixa = vobs
    End If
    If reixa.col = numcol("AmpleMin") Then
     vobs = reixa.Text
     vobs = InputBox("Entra l'amplada mínima d'aquesta linia." + Chr(10) + "ESCRIU UN ESPAI PER BORRAR TOT EL CONTINGUT.", "Observacions", vobs)
     If StrPtr(vobs) = Empty Then Exit Sub
     dbconsulta.Execute "update estocseguretat set amplemin='" + atrim(cadbl(vobs)) + "' where id=" + atrim(vid)
     reixa = atrim(cadbl(vobs))
    End If
    If reixa.col = numcol("AmpleMax") Then
     vobs = reixa.Text
     vobs = InputBox("Entra l'amplada màxima d'aquesta linia." + Chr(10) + "ESCRIU UN ESPAI PER BORRAR TOT EL CONTINGUT.", "Observacions", vobs)
     If StrPtr(vobs) = Empty Then Exit Sub
     dbconsulta.Execute "update estocseguretat set amplemax='" + atrim(cadbl(vobs)) + "' where id=" + atrim(vid)
     reixa = atrim(cadbl(vobs))
    End If
    If reixa.col = numcol("E.Seguretat") Then
     vobs = reixa.Text
     vobs = InputBox("Entra l'estoc de seguretat d'aquesta linia." + Chr(10) + "ESCRIU UN ESPAI PER BORRAR TOT EL CONTINGUT.", "Observacions", vobs)
     If StrPtr(vobs) = Empty Then Exit Sub
     dbconsulta.Execute "update estocseguretat set kgestoc='" + atrim(cadbl(vobs)) + "' where id=" + atrim(vid)
     reixa = atrim(cadbl(vobs))
    End If
End Sub

Private Sub reixa_LostFocus()
   guardar_amples_reixa
End Sub

Private Sub reixadetall_DblClick()
  If reixadetall.Columns(0).DataField = "NºComanda" Then cridarcomandes reixadetall.Columns(0)
  If reixadetall.Columns(0).DataField = "NºCompra" Then cridarcomandescompra reixadetall.Columns(0)
  If reixadetall.Columns(0).DataField = "NºPalet" Then cridarpalet reixadetall.Columns(0), reixadetall.Columns(1)
End Sub
Sub cridarpalet(palet As Double, bobina As Double)
  Form1.palets.RecordSource = "select * from palets where idpalet=" + atrim(palet)
  Form1.palets.Refresh
  If Not Form1.palets.Recordset.EOF Then
    Form1.palets.Recordset.MoveLast: Form1.palets.Recordset.MoveFirst
    Form1.bobines.Recordset.FindFirst "idbobina=" + atrim(bobina)
    formestocseguretat.Hide
    Form1.SetFocus
  End If
End Sub
Sub cridarcomandescompra(comanda As Double)
  Dim rstcompres As Recordset
  On Error GoTo obrircomandes
  escriure_ini "Planificacio", "comandacompraxrobrir", atrim(comanda), "comandes.ini"
  AppActivate "Comandes de compra.", True
  
  
  On Error Resume Next
  Exit Sub
obrircomandes:
  On Error Resume Next
   Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", "comandes.ini")) + "compres.exe", vbNormalFocus
End Sub
Sub cridarcomandes(comanda As Double)
 On Error GoTo obrircomandes
  escriure_ini "Planificacio", "comandaxrobrir", atrim(comanda), "comandes.ini"
  AppActivate "Manteniment de Comandes"
  
  On Error Resume Next
  Exit Sub
obrircomandes:
  On Error Resume Next
   Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", "comandes.ini")) + "comandes.exe - comandes", vbNormalFocus
End Sub

Private Sub subfamad_Change()
carregar_subfamilies
End Sub

Private Sub subfamad_DropDown()
carregar_subfamilies
End Sub

Private Sub subfamcol_Change()
carregar_subfamilies
End Sub

Private Sub subfamcol_DropDown()
carregar_subfamilies
End Sub

Private Sub subfammat_Change()
  carregar_subfamilies
End Sub

Private Sub subfammat_DropDown()
  carregar_subfamilies
End Sub
Sub configreixa(Optional nocarregaramples As Boolean)
  Dim rst As Recordset
  Dim col As Long
  Dim enes As Byte
  Dim i As Long
  If Not nocarregaramples Then descarregarfiltres
  reixa.LeftCol = 0
  If reixa.Rows > 1 Then reixa.TopRow = 1
  Set rst = dbconsulta.OpenRecordset("select * from estocseguretat order by familia,amplemin", dbOpenSnapshot, dbReadOnly)
  col = 0
  enes = 0
  reixa.Cols = rst.Fields.Count
  For i = 0 To rst.Fields.Count - 1
    If camps(i, 4) <> "N" And camps(i, 1) <> "" Then

       reixa.ColAlignment(col) = 2
       reixa.TextMatrix(0, col) = camps(i, 3)
       
       If Not nocarregaramples Then colocarfiltre col, i
       col = col + 1
        Else: enes = enes + 1
    End If
  Next i

  reixa.ColWidth(0) = 0
  
  If enes > 0 Then reixa.Cols = reixa.Cols - enes
  If Not nocarregaramples Then carregar_amples_reixa
  Set rst = Nothing
End Sub
Sub descarregarfiltres()
  Dim i As Byte
  For i = 1 To filtre.Count - 1
   Unload filtre.Item(i)
  Next i
End Sub

Sub carregar_amples_reixa()
 Dim ample As String
 Dim X As Long
 Dim j As Integer
 If iniconfigreixa <> "" Then ' existeix("c:\windows\" + iniconfigreixa) Then
 
  X = reixa.Left + 35
  For j = 0 To reixa.Cols - 1
   ample = llegir_ini("AmplesReixa", UCase(reixa.TextMatrix(0, j)), iniconfigreixa)
   If ample <> "{[}]" Then
    reixa.ColWidth(j) = cadbl(ample)
    If X < reixa.Width Then
     filtre(j).Left = X
     filtre(j).Width = cadbl(ample)
     filtre(j).Visible = IIf(ample < 50, False, True)
     filtre(j).ForeColor = &H808080
      Else: If filtre.Count < j - 1 Then filtre(j).Visible = False
    End If
    X = X + cadbl(ample)
   End If
 Next j
End If
filtre(0).Width = filtre(0).Width - 50
filtre(0).Left = filtre(0).Left + 50
End Sub

Function numcol(nom As String) As Byte
   numcol = 0
   For i = 0 To reixa.Cols - 1
     If reixa.TextMatrix(0, i) = nom Then numcol = i
   Next i
   
End Function

Sub colocarfiltre(col As Long, i As Long)
  If filtre.Count <= col Then Load filtre(col)
  filtre(col).Text = camps(i, 3)
  filtre(col).Tag = i
End Sub

Sub poblarlareixa(Optional vwere As String)
  Dim i As Double
  Dim fila As Integer
  Dim col As Byte
  Dim rst As Recordset
  Dim rstreclamades As Recordset
  Dim apuntxrimprimir As Double
  Dim tenimmaterial As Boolean
  Dim tenimclixes As Boolean
  Dim textetaula As String
  
  ratoli "espera"
  reixa.Visible = False
  reixa.Clear
  reixa.BackColor = QBColor(15)
  configreixa IIf(vwere <> "", True, False)
  reixa.Rows = 1
  Set rst = dbconsulta.OpenRecordset("select * from estocseguretat " + IIf(vwere <> "", " where " + vwere, "") + " order by familia,amplemin", dbOpenSnapshot, dbReadOnly)
  If rst.EOF Then MsgBox "No hi ha cap registre.": Exit Sub
  fila = 0
  reixa.Tag = "poblant"
  While Not rst.EOF
   fila = fila + 1
   reixa.Rows = fila + 1
   reixa.row = fila
   col = 0
   For i = -1 To rst.Fields.Count - 1
    If camps(i + 1, 1) <> "" And camps(i + 1, 4) <> "N" Then
      
      reixa.TextMatrix(fila, col) = IIf(IsNull(rst.Fields(camps(i + 1, 1))), "", rst.Fields(camps(i + 1, 1)))
     
      
format:   ' apartir d'aqui aplico el format a la casella
      'posso el format del camp dataimpresio
      If camps(i + 1, 2) = "date" Then
        If camps(i + 1, 1) <> "dataimpresio" And camps(i + 1, 1) <> "dataoperari" Then
          If reixa.TextMatrix(fila, col) = "0:00:00" Then
            reixa.TextMatrix(fila, col) = ""
                Else: reixa.TextMatrix(fila, col) = format(reixa.TextMatrix(fila, col), "dd/mm/yy")
          End If
            Else: reixa.TextMatrix(fila, col) = format(reixa.TextMatrix(fila, col), "dd/mm hh:nn")
        End If
      End If
      
       'posso el format del camp tempsimpresio
      If camps(i + 1, 1) <> "Temps" And cadbl(reixa.TextMatrix(fila, col)) < 0 Then
           reixa.col = col
           'reixa.TextMatrix(fila, col) = cadbl(reixa.TextMatrix(fila, col)) * -1
           reixa.CellBackColor = &H8080FF    'vermell
      End If
      
      
       'posso el format del camp tipusdeadhesius de laminadora
      If camps(i + 1, 1) = "tipuscola" Then
         If Mid(reixa.TextMatrix(fila, col), 1, 1) = "@" Then
           reixa.col = col
           reixa.row = fila
           colorcelda = cadbl(Mid(reixa.TextMatrix(fila, col), 2, InStr(2, reixa.TextMatrix(fila, col), " ") - 1))
           reixa.CellBackColor = colorcelda
           reixa.TextMatrix(fila, col) = Mid(reixa.TextMatrix(fila, col), InStr(2, reixa.TextMatrix(fila, col), " "))
           reixa.col = numcol("Nom Client")
           reixa.CellBackColor = colorcelda
         End If
      End If
      col = col + 1
    End If
   Next i
   rst.MoveNext
  Wend
  'registres = atrim(rst.RecordCount) + " Registres"
  Set rst = Nothing
  reixa.Tag = ""
  reixa.Visible = True
  reixa.row = 1
  ratoli "normal"
  
End Sub

Private Sub treurefiltre_Click()
borrarelfiltre
End Sub
Sub borrarelfiltre()
 configreixa False
 poblarlareixa
End Sub

Sub guardar_amples_reixa()
Dim j As Integer
If iniconfigreixa <> "" Then
  For j = 0 To reixa.Cols - 1
   escriure_ini "AmplesReixa", UCase(reixa.TextMatrix(0, j)), atrim(reixa.ColWidth(j)), iniconfigreixa
 Next j
End If
End Sub
Function crear_criteri_familia(rst As Recordset, Optional vcont As Byte) As String
   Dim d As String
   If cadbl(rst!familia) >= 0 Then
      d = " familia=" + atrim(rst!familia)
    Else: d = " familia>0"
   End If
   If rst!subfamilia > 0 Then vcont = vcont + 1: d = d + " and subfamilia=" + atrim(rst!subfamilia)
   If rst!familiacol > 0 Then vcont = vcont + 1: d = d + " and familiacol=" + atrim(rst!familiacol)
   If rst!subfamiliacol > 0 Then vcont = vcont + 1: d = d + " and subfamiliacol=" + atrim(rst!subfamiliacol)
   If rst!familiaad > 0 Then vcont = vcont + 1: d = d + " and familiaad=" + atrim(rst!familiaad)
   If rst!subfamiliaad > 0 Then vcont = vcont + 1: d = d + " and subfamiliaad=" + atrim(rst!subfamiliaad)
   crear_criteri_familia = d
End Function
