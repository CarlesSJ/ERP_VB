VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formdevolucions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devolució de materials"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11895
   Icon            =   "formdevolucions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CheckTots 
      Caption         =   "Veure tots els albarans"
      Height          =   195
      Left            =   1680
      TabIndex        =   8
      Top             =   90
      Width           =   2490
   End
   Begin VB.CommandButton Command2 
      Height          =   330
      Left            =   11415
      Picture         =   "formdevolucions.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   180
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Height          =   330
      Left            =   11070
      Picture         =   "formdevolucions.frx":1654
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Alta  Registres"
      Top             =   180
      Width           =   345
   End
   Begin VB.CommandButton consultar 
      Height          =   360
      Left            =   1020
      Picture         =   "formdevolucions.frx":1BDE
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Buscar Registres"
      Top             =   90
      Width           =   420
   End
   Begin VB.CommandButton eliminar 
      Height          =   360
      Left            =   585
      Picture         =   "formdevolucions.frx":2168
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   90
      Width           =   420
   End
   Begin VB.CommandButton alta 
      Height          =   360
      Left            =   150
      Picture         =   "formdevolucions.frx":26F2
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Alta  Registres"
      Top             =   90
      Width           =   420
   End
   Begin VB.ListBox llistabobines 
      BackColor       =   &H00EAD9CE&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   9540
      TabIndex        =   1
      Top             =   555
      Width           =   2295
   End
   Begin VB.Data datadevolucions 
      Caption         =   "datadevolucions"
      Connect         =   "Access"
      DatabaseName    =   "\\SERVERPRODU\Dades\progcomandes\dades\Palets.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5490
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "devoluciomaterial"
      Top             =   90
      Visible         =   0   'False
      Width           =   2580
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "formdevolucions.frx":2C7C
      Height          =   3600
      Left            =   135
      OleObjectBlob   =   "formdevolucions.frx":2C96
      TabIndex        =   0
      Top             =   540
      Width           =   9300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bobines afectades"
      Height          =   255
      Left            =   9660
      TabIndex        =   2
      Top             =   300
      Width           =   1635
   End
End
Attribute VB_Name = "formdevolucions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
   Dim rst As Recordset
   Dim vcodip As Double
   Dim vnomp As String
   Dim valbarainp As String
   Dim vdata_dev As String
   
   escullir_proveidor vcodip, vnomp
   If vcodip = 0 Then Exit Sub
   valbarainp = InputBox("Entra el numero d'albarà d'INPLACSA (Sortida d'expedicions).", "Albarà recullida")
   If atrim(valbarainp) = "" Then Exit Sub
   Set rst = dbtmp.OpenRecordset("select * from devoluciomaterial where albaradevolucio='" + atrim(valbarainp) + "'")
   If Not rst.EOF Then MsgBox "Aquest albarà ja està utilitzat en una devolució.", vbCritical, "Error": GoTo fi
   vdata_dev = InputBox("Entra la data de devolució del material. dd/mm/yy ", "Data devolució", format(Now, "dd/mm/yy"))
   If Not IsDate(vdata_dev) Then Exit Sub
   datadevolucions.Recordset.AddNew
   datadevolucions.Recordset!albaradevolucio = valbarainp
   datadevolucions.Recordset!codiproveidor = vcodip
   datadevolucions.Recordset!nomproveidor = vnomp
   datadevolucions.Recordset!datadevolucio = vdata_dev
   datadevolucions.Recordset.Update
   datadevolucions.Refresh
   datadevolucions.Recordset.MoveLast
fi:
   Set rst = Nothing
End Sub
Sub escullir_proveidor(vcodip As Double, vnomp As String)
  Load formseleccio
  formseleccio.sortirs.Tag = "filtre"
  Set formseleccio.Data1.Recordset = dbcomandes.OpenRecordset("select codi,nom from proveidors order by nom")
  formseleccio.Caption = "Escull_PROVEIDOR"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).Width = 1000
  'formseleccio.DBGrid2.Columns(1).Width = 2000
  'formseleccio.Width = 500
  formseleccio.Show 1
  If seleccioret = 1 Then
   vcodip = formseleccio.DBGrid2.Columns("codi")
   vnomp = formseleccio.DBGrid2.Columns("nom")
  End If
  Unload formseleccio
End Sub

Private Sub CheckTots_Click()
    carregar_devolucions
End Sub

Private Sub Command1_Click()
   Dim vsql As String
   Dim vcodip As Double
   Dim vidparcial As Double
   Dim v As Long
   v = datadevolucions.Recordset!id
   datadevolucions.Refresh
   datadevolucions.Recordset.FindFirst "id=" + atrim(v)
   vcodip = cadbl(datadevolucions.Recordset!codiproveidor)
   vsql = "SELECT  * from escullir_linies_devoluciomaterial "
   vsql = vsql + " WHERE codi=" + atrim(vcodip) + " AND comanda='300' and id=null "
   Load formseleccio
   formseleccio.Command3.Tag = ""
  formseleccio.sortirs.Tag = ""
  formseleccio.Command2.Tag = "4"
 ' Clipboard.Clear
 ' Clipboard.SetText vsql
  Set formseleccio.Data1.Recordset = dbtmp.OpenRecordset(vsql)
  formseleccio.Caption = "Escull_Parcial_300"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Visible = False
  formseleccio.DBGrid2.Columns(2).Visible = False
  formseleccio.DBGrid2.Columns(3).Visible = False
  formseleccio.DBGrid2.Columns(4).Width = 1000
'  formseleccio.DBGrid2.Columns(4).Caption = "Palet"
'  formseleccio.DBGrid2.Columns(5).Caption = "Bob"
'  formseleccio.DBGrid2.Columns(6).Caption = "Palet"
  'formseleccio.DBGrid2.Columns(1).Width = 2000
  'formseleccio.Width = 500
  formseleccio.Show 1
  If seleccioret = 1 Then
   vidparcial = cadbl(formseleccio.Data1.Recordset.Fields("Idparcial"))
   If vidparcial > 0 Then
        dbtmp.Execute "Insert into devoluciomaterial_linies (Idcapcalera,idparcial300) values (" + atrim(datadevolucions.Recordset!id) + "," + atrim(vidparcial) + ")"
   End If
  End If
  Unload formseleccio
  carregar_bobines_parcials
End Sub

Private Sub Command2_Click()
  If MsgBox("Segur que vols treure aquest palet de la devolució?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
  dbtmp.Execute "delete * from devoluciomaterial_linies where idcapcalera=" + atrim(datadevolucions.Recordset!id) + " and idparcial300=" + atrim(llistabobines.ItemData(llistabobines.ListIndex))
  carregar_bobines_parcials
End Sub

Private Sub consultar_Click()
  Dim vcamp As String
  Dim v As String
  vcamp = "albaradevolucio"
  If reixa.col > 0 Then vcamp = reixa.Columns(reixa.col).DataField
  v = InputBox("Entra el valor que vols buscar a " + atrim(vcamp) + ":", "Buscar")
  If consultar.Tag = "" Then consultar.Tag = datadevolucions.RecordSource
  datadevolucions.RecordSource = consultar.Tag + " and " + vcamp + " like '*" + v + "*'"
  datadevolucions.Refresh
  
End Sub

Private Sub datadevolucions_Reposition()
   carregar_bobines_parcials
End Sub

Private Sub eliminar_Click()
   If datadevolucions.Recordset.EOF Then MsgBox "No hi ha cap registre seleccionat.", vbCritical, "Error": Exit Sub
   If UCase(InputBox("Escriu [ELIMINAR] per eliminar aquesta devolució de material.", "ELIMINAR")) = "ELIMINAR" Then
       datadevolucions.Database.Execute "delete * from devoluciomaterial_linies where idcapcalera=" + atrim(datadevolucions.Recordset!id)
       datadevolucions.Recordset.Delete
       llistabobines.Clear
   End If
End Sub

Sub carregar_bobines_parcials()
   Dim rst As Recordset
   
   llistabobines.Clear
   If datadevolucions.Recordset.EOF Then Exit Sub
   
   Set rst = dbtmp.OpenRecordset("SELECT devoluciomaterial_linies.Idcapcalera, parcials.id,Parcials.idpalet, Parcials.idbobina, Parcials.metres FROM Parcials RIGHT JOIN devoluciomaterial_linies ON Parcials.id = devoluciomaterial_linies.idparcial300 WHERE devoluciomaterial_linies.Idcapcalera=" + atrim(cadbl(datadevolucions.Recordset!id)) + " order by idpalet,idbobina;")
   While Not rst.EOF
      llistabobines.AddItem justificar(atrim(cadbl(rst!idpalet)), 6, "D") + "/" + justificar(atrim(cadbl(rst!idbobina)), 3, "E") + justificar(cadbl(rst!metres), 7, "D") + "m"
      llistabobines.ItemData(llistabobines.NewIndex) = cadbl(rst!id)
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub Form_Load()
  datadevolucions.DatabaseName = rutadelfitxer(cami) + "palets.mdb"
  carregar_devolucions
End Sub
Sub carregar_devolucions()

  datadevolucions.RecordSource = "select * from devoluciomaterial where " + IIf(CheckTots.Value <> 1, " (numerofacturasap='' or numerofacturasap=null) and datafacturasap= null", "id>0")
  datadevolucions.Refresh
End Sub

Private Sub reixa_DblClick()
   Dim v As String
   On Error GoTo fi
   If reixa.Columns(reixa.col).DataField = "datafacturaSAP" Then
       v = InputBox("Entra la data de la factura del PROVEIDOR referent a aquesta devolució." + vbNewLine + "format d'entrada: DD/MM/YY", "Num Factura devolució.")
       If v <> "" Then reixa.Text = v: datadevolucions.Recordset.Move 0
   End If
   If reixa.Columns(reixa.col).DataField = "numerofacturaSAP" Then
       v = InputBox("Entra el numero de factura del PROVEIDOR referent a aquesta devolució.", "Num Factura devolució.")
       If atrim(v) <> "" Then reixa.Text = v: datadevolucions.Recordset.Move 0
   End If
   Exit Sub
fi:
    MsgBox "Error de dades, no es guarden els canvis.", vbCritical, "Error"
End Sub

Private Sub reixa_KeyPress(KeyAscii As Integer)
reixa_DblClick
End Sub
