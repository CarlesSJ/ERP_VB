VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formtarifesperreferencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tarifes per referència"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1995
      Top             =   540
   End
   Begin VB.CommandButton alta 
      Height          =   375
      Left            =   255
      Picture         =   "mantenimentreftarifes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Alta  Registres"
      Top             =   570
      Width           =   360
   End
   Begin VB.CommandButton eliminar 
      Height          =   375
      Left            =   645
      Picture         =   "mantenimentreftarifes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   570
      Width           =   360
   End
   Begin VB.CommandButton bbuscarref 
      Height          =   375
      Left            =   1035
      Picture         =   "mantenimentreftarifes.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Buscar Referencia"
      Top             =   570
      Width           =   360
   End
   Begin VB.Data datatarifes 
      Caption         =   "datatarifes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2115
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tarifes_referencies"
      Top             =   6030
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.ComboBox comboclient 
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      Text            =   "Escull Client"
      Top             =   195
      Width           =   4320
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "mantenimentreftarifes.frx":109E
      Height          =   4860
      Left            =   120
      OleObjectBlob   =   "mantenimentreftarifes.frx":10B4
      TabIndex        =   0
      Top             =   1005
      Width           =   7155
   End
   Begin VB.Label etcomentari 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H005C31DD&
      Height          =   255
      Left            =   1665
      TabIndex        =   6
      Top             =   750
      Width           =   5460
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
      Height          =   240
      Left            =   375
      TabIndex        =   2
      Top             =   225
      Width           =   1080
   End
End
Attribute VB_Name = "formtarifesperreferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
    Dim vref As String
    Dim vtarifa As String
    If comboclient.Tag <> "" Then
        vref = InputBox("Escriu la referència d'INPLACSA.", "Referència INP")
        If vref = "" Then Exit Sub
        datatarifes.Recordset.FindFirst "refinplacsa='" + atrim(vref) + "'"
        If Not datatarifes.Recordset.NoMatch Then MsgBox "Aquesta referència ja existeix.", vbCritical, "Error": Exit Sub
        vtarifa = InputBox("Escriu la tarifa del client.", "Tarifa")
        If StrPtr(vtarifa) = 0 Then Exit Sub
        If atrim(vtarifa) = "" Then vtarifa = "0"
        datatarifes.Recordset.AddNew
        datatarifes.Recordset!codiclient = cadbl(comboclient.Tag)
        datatarifes.Recordset!coditarifa = vtarifa
        datatarifes.Recordset!refinplacsa = vref
        datatarifes.Recordset.Update
        datatarifes.Refresh
        datatarifes.Recordset.FindFirst "refinplacsa='" + vref + "'"
    End If
'    refrescar_reixa
End Sub

Private Sub bbuscarref_Click()
   Dim vref As String
   vref = InputBox("Entra la referència del client que vols buscar.", "Buscar referència")
   datatarifes.Recordset.FindFirst "refinplacsa ='" + atrim(vref) + "'"
   If datatarifes.Recordset.NoMatch Then MsgBox "No s'ha trobat aquesta referència.", vbInformation, "Atenció"
End Sub

Private Sub comboclient_Click()
   If comboclient.Text = "- Escullir client -" Then
      bbuscarref.SetFocus
      Timer1.Enabled = True
      
       Else: comboclient.Tag = comboclient.Text
   End If
   refrescar_reixa
End Sub
Sub emplenarcombo()
  Dim rst As Recordset
   comboclient.Clear
   comboclient.AddItem "- Escullir client -"
   Set rst = dbtmp.OpenRecordset("select distinct(grupdeclient) as grup from clients  ")
   While Not rst.EOF
      If atrim(rst!grup) <> "" Then comboclient.AddItem UCase(rst!grup)
      rst.MoveNext
   Wend
End Sub
Private Sub comboclient_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub
Sub refrescar_reixa()
  datatarifes.RecordSource = "select * from tarifes_referencies where codiclient='" + atrim(comboclient.Tag) + "'"
  datatarifes.Refresh
End Sub
Sub triarclient()
  Load formseleccio
  formseleccio.Caption = "Escullir client"
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = datatarifes.DatabaseName
  formseleccio.Data1.RecordSource = "select codi,nom from clients order by nom"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   comboclient.Tag = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   comboclient = atrim(formseleccio.Data1.Recordset!nom)
  End If
  Unload formseleccio
  
End Sub

Private Sub eliminar_Click()
  If datatarifes.Recordset.EOF Then Exit Sub
  If MsgBox("SEGUR QUE VOLS ELIMINAR LA REFRENCIA " + datatarifes.Recordset!refinplacsa + "?", vbCritical + vbYesNo) = vbYes Then
      datatarifes.Recordset.Delete
    End If
End Sub

Private Sub Form_Click()
  MsgBox comboclient.Tag
End Sub

Private Sub Form_Load()
  datatarifes.DatabaseName = cami
  datatarifes.RecordSource = "select * from tarifes_referencies where codiclient='-'"
  datatarifes.Refresh
  emplenarcombo
End Sub

Private Sub reixa_DblClick()
   Dim v As String
   If reixa.Columns(reixa.col).DataField = "coditarifa" Then
       vcoditarifa = InputBox("Escriu el nou codi de tarifa", "Canvi de codi", reixa.Text)
   End If
   If reixa.Columns(reixa.col).DataField = "inactiva" Then
       If MsgBox("Vols canviar l'estat d'aquesta tarifa?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
           datatarifes.Recordset.Edit
           datatarifes.Recordset!inactiva = Not datatarifes.Recordset!inactiva
           datatarifes.Recordset.Update
           datatarifes.Recordset.Move 0
       End If
   End If
   If reixa.Columns(reixa.col).DataField = "impost_regimenfiscal" Then
           v = InputBox("Escriu el REGIM FISCAL." + vbNewLine + " valors(A,B,C,D,E,F,G,H,I,J,K,L,M)", "Atenció")
           If StrPtr(v) = 0 Then Exit Sub
           v = UCase(v)
           If InStr(1, "ABCDEFGHIJKLM ", v) = 0 Then MsgBox "Aquest valor no es vàlid", vbCritical, "Error": Exit Sub
           datatarifes.Recordset.Edit
           datatarifes.Recordset!impost_regimenfiscal = v
           datatarifes.Recordset.Update
           datatarifes.Recordset.Move 0
   End If
   If reixa.Columns(reixa.col).DataField = "impost_claveproducto" Then
           v = UCase(InputBox("Escriu la CLAVE DE PRODUCTO." + vbNewLine + " valors(A,B,C)", "Atenció"))
           If StrPtr(v) = 0 Then Exit Sub
           v = UCase(v)
           If InStr(1, "ABC ", v) = 0 Then MsgBox "Aquest valor no es vàlid", vbCritical, "Error": Exit Sub
           datatarifes.Recordset.Edit
           datatarifes.Recordset!impost_claveproducto = v
           datatarifes.Recordset.Update
           datatarifes.Recordset.Move 0
   End If
   
End Sub

Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If reixa.Columns(reixa.col).DataField = "impost_regimenfiscal" Then
       etcomentari = "Excempt per: [E] ús mèdic;[F] Mèdic Ziper;[H] Agricola"
         Else: etcomentari = ""
   End If
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
   triarclient
   refrescar_reixa
End Sub
