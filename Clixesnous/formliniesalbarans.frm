VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form formliniesalbarans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment de Linies d'Albarans"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6480
   ScaleWidth      =   14430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cobservacioalbara 
      Height          =   285
      Left            =   2385
      MaxLength       =   80
      TabIndex        =   19
      ToolTipText     =   "Linia que s'afegirà al final de l'albarà de clixés quan es puji al SAP"
      Top             =   960
      Width           =   11595
   End
   Begin VB.ComboBox Comboquifactura 
      Height          =   315
      ItemData        =   "formliniesalbarans.frx":0000
      Left            =   1710
      List            =   "formliniesalbarans.frx":000A
      TabIndex        =   13
      Top             =   585
      Width           =   1395
   End
   Begin VB.TextBox ccodifacturacio 
      Height          =   285
      Left            =   11295
      TabIndex        =   11
      Top             =   600
      Width           =   2685
   End
   Begin VB.ComboBox combonomclientfact 
      Height          =   315
      Left            =   4470
      TabIndex        =   8
      Top             =   585
      Width           =   5070
   End
   Begin VB.Data albarans 
      Caption         =   "albarans"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"formliniesalbarans.frx":0020
      Top             =   90
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame framebotons2 
      Height          =   585
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   14355
      Begin VB.Timer rellotgepostit 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   7365
         Top             =   60
      End
      Begin Crystal.CrystalReport llistat 
         Left            =   2025
         Top             =   105
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton imprimir 
         Height          =   360
         Left            =   960
         Picture         =   "formliniesalbarans.frx":00C7
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   13890
         Picture         =   "formliniesalbarans.frx":0651
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Sortir"
         Top             =   135
         Width           =   390
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   120
         Picture         =   "formliniesalbarans.frx":0BDB
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton guardar 
         Height          =   360
         Left            =   540
         Picture         =   "formliniesalbarans.frx":1165
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   375
      End
      Begin MSComCtl2.DTPicker picker 
         Height          =   315
         Left            =   3105
         TabIndex        =   7
         Top             =   75
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/mm/yy"
         Format          =   662831107
         CurrentDate     =   41303
      End
   End
   Begin VB.Frame framealb 
      Caption         =   "        Albarans"
      DragIcon        =   "formliniesalbarans.frx":16EF
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5115
      Left            =   30
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   1290
      Width           =   14355
      Begin VB.Frame postit 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   11415
         TabIndex        =   16
         Top             =   1365
         Visible         =   0   'False
         Width           =   2700
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   $"formliniesalbarans.frx":1C79
            Height          =   990
            Left            =   120
            TabIndex        =   17
            Top             =   75
            Width           =   2490
         End
      End
      Begin VB.CommandButton beliminararxiupdf 
         Height          =   390
         Left            =   13935
         Picture         =   "formliniesalbarans.frx":1D39
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Eliminar arxiu PDF"
         Top             =   870
         Width           =   360
      End
      Begin VB.CommandButton bveurepdfalbara 
         DragIcon        =   "formliniesalbarans.frx":22C3
         Height          =   390
         Left            =   13950
         OLEDropMode     =   1  'Manual
         Picture         =   "formliniesalbarans.frx":284D
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Arxiu PDF"
         Top             =   450
         Width           =   345
      End
      Begin MSDBGrid.DBGrid reixaalbarans 
         Bindings        =   "formliniesalbarans.frx":2DD7
         Height          =   4785
         Left            =   135
         OleObjectBlob   =   "formliniesalbarans.frx":2DEA
         TabIndex        =   6
         Tag             =   "albarans"
         Top             =   225
         Width           =   13770
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Observació per l'albarà del SAP:"
      Height          =   285
      Left            =   45
      TabIndex        =   18
      Top             =   1005
      Width           =   2355
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa que factura:"
      Height          =   285
      Left            =   60
      TabIndex        =   12
      Top             =   600
      Width           =   1620
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Comanda Clixes:            (Client)"
      Height          =   420
      Left            =   9795
      TabIndex        =   10
      Top             =   555
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Facturar Clixes a:"
      Height          =   285
      Left            =   3180
      TabIndex        =   9
      Top             =   615
      Width           =   1485
   End
End
Attribute VB_Name = "formliniesalbarans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub beliminararxiupdf_Click()
  Dim vnumalbara As String
  If albarans.Recordset.EOF Then Exit Sub
  vnumalbara = atrim(albarans.Recordset!num_alb)
  vfitxerdesti = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Albarans" + "\v" + atrim(ordremodificacio) + "\" + treuresimbols(vnumalbara) + ".pdf"
  If MsgBox("Segur que vols desvincular el pdf de l'albarà " + atrim(vnumalbara), vbInformation + vbYesNo + vbDefaultButton2, "Eliminar PDF") = vbYes Then
    albarans.Recordset.FindFirst "num_alb='" + atrim(vnumalbara) + "'"
    While Not albarans.Recordset.NoMatch
            albarans.Recordset.Edit
            albarans.Recordset!albarapdfvinculat = False
            albarans.Recordset.Update
            albarans.Recordset.FindNext "num_alb='" + atrim(vnumalbara) + "'"
    Wend
    albarans.Recordset.Move 0
    Kill vfitxerdesti
  End If
End Sub

Private Sub bveurepdfalbara_Click()
   Dim vnumalbara As String
   Dim vfitxerdesti As String
   If albarans.Recordset.EOF Then Exit Sub
   vnumalbara = atrim(albarans.Recordset!num_alb)
   vfitxerdesti = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Albarans" + "\v" + atrim(ordremodificacio) + "\" + treuresimbols(vnumalbara) + ".pdf"
   obrir_document vfitxerdesti
End Sub

Private Sub bveurepdfalbara_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   draganddropcopiaralbara atrim(data.Files(1))
End Sub

Private Sub ccodifacturacio_LostFocus()
  formclixes.modificacions.Recordset.Edit
  formclixes.modificacions.Recordset!codifacturacioclixes = atrim(ccodifacturacio)
  formclixes.modificacions.Recordset.Update
End Sub

Private Sub cobservacioalbara_LostFocus()
  formclixes.modificacions.Recordset.Edit
  formclixes.modificacions.Recordset!observacionsfacturaclixes = atrim(cobservacioalbara)
  formclixes.modificacions.Recordset.Update
End Sub

Private Sub combonomclientfact_DropDown()
   If Comboquifactura = "" Then MsgBox "Escull primer l'empresa que facturarà.", vbCritical, "Error": Exit Sub
   Load formseleccio
   formseleccio.sortirs.tag = "filtre"
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select codiSAP,nif,nomclient from clients_codisSAP" + IIf(Comboquifactura = "Inplacsa", "", "Plasel")
   formseleccio.width = 10000
   formseleccio.refrescar
    formseleccio.DBGrid2.Columns(0).width = 1200
    formseleccio.DBGrid2.Columns(1).width = 1200
   formseleccio.DBGrid2.Columns(2).width = 5000
   formseleccio.colocar_botofiltre 2
   formseleccio.Show 1
   
    If seleccioret = 1 Then
            formclixes.modificacions.Recordset.Edit
            formclixes.modificacions.Recordset!codiclientfactclixes = formseleccio.DBGrid2.Columns("codiSAP")
            formclixes.modificacions.Recordset.Update
            
            combonomclientfact = formseleccio.DBGrid2.Columns("nomclient")
    End If
     If seleccioret = 9 Then
         formclixes.modificacions.Recordset.Edit
         formclixes.modificacions.Recordset!codiclientfactclixes = "0"
         formclixes.modificacions.Recordset.Update
         combonomclientfact = ""
    End If
    formseleccio.Data1.RecordSource = ""
    formseleccio.Data1.Refresh
    Unload formseleccio
    SendKeys "{TAB}"
End Sub

Private Sub combonomclientfact_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub combonomclientfact_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Comboquifactura_Click()
   formclixes.modificacions.Recordset.Edit
   formclixes.modificacions.Recordset!empresafacturadora = Mid(Comboquifactura, 1, 1)
   formclixes.modificacions.Recordset.Update
End Sub

Private Sub Comboquifactura_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub Comboquifactura_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub eliminar_Click()
 If MsgBox("Segur que vols borrar aquesta linia?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
      
                If Not albarans.Recordset.EOF Then
                       albarans.Recordset.Delete
                       albarans.Recordset.MoveFirst
                       gravar_canvis
                End If
    
  End If
End Sub
Sub gravar_canvis()
  Dim bk As Long
  bk = albarans.Recordset!ordre
  If Not gravar_reixa(reixaalbarans) Then Exit Sub
   
'   If clixes.Recordset.EditMode > 0 Then clixes.Recordset.Update
   'modificar_Click
   On Error GoTo erro
   If albarans.Recordset.EditMode > 0 Then albarans.Recordset.Update
   reixaalbarans.SetFocus
   albarans.Recordset.FindFirst "ordre=" + atrim(cadbl(bk))
   Exit Sub
erro:
   ' MsgBox err.Description
    
End Sub
Function gravar_reixa(reixa As DBGrid) As Boolean
    Dim fila As Double
    gravar_reixa = True
     fila = reixa.row
     reixa.SetFocus
     If Not albarans.Recordset.EOF Then
       On Error GoTo error
       albarans.Recordset.Move 0
       On Error GoTo 0
     End If
'     SendKeys "{down}"
     DoEvents
'     If reixa.row <> fila + 1 Then gravar_reixa = False
     reixa.row = fila
    Exit Function
error:
   MsgBox "Hi ha hagut algun error en algun camp, revisa que sigui tot correcte.", vbCritical, "Error"
End Function

Sub carregar_clientfacturaciopredeterminat()
Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("SELECT Clixes.id_treball, clients.grupdeclient FROM Clixes LEFT JOIN clients ON Clixes.codiclienttemporal = clients.codi where id_treball=" + atrim(id_treball))
   If rst.EOF Then Exit Sub
   If cadbl(formclixes.modificacions.Recordset!codiclientfactclixes) = 0 And rst!grupdeclient = "ARDO" Then
      ' tot lo de ardo es factura aquest codi
       combonomclientfact = "ARDO FOODS NV"
       formclixes.modificacions.Recordset.Edit
       formclixes.modificacions.Recordset!codiclientfactclixes = 43000007419#
       formclixes.modificacions.Recordset.Update
       MsgBox "He posat el codi 7419 ARDO FOODS NV com a facturació de clixes del GRUP ARDO, si no es així pots canviar-lo quan vulguis.", vbInformation, "Atenció"
   End If
   If cadbl(formclixes.modificacions.Recordset!codiclientfactclixes) = 0 And cadbl(formclixes.clixes.Recordset!codiclienttemporal) = 6511 Then
      ' M TERESA LLAURADO HA DEMANAT QUE ES FACTURI A TORRAS RAFI
       combonomclientfact = "PRODUCTES M TORRAS RAFI SL"
       formclixes.modificacions.Recordset.Edit
       formclixes.modificacions.Recordset!codiclientfactclixes = 43000006934#
       formclixes.modificacions.Recordset.Update
       MsgBox "He posat el CLIENT de facturació a PRODUCTES M TORRAS RAFI perquè el client M TERESA LLAURADO ho ha demanat, si no es així pots canviar-lo quan vulguis.", vbInformation, "Atenció"
   End If
   posarnomclientfacturacio
 
End Sub
Private Sub Form_Load()
  Dim idtreball As String
  Dim ordremodifi As String
  idtreball = atrim(formclixes.modificacions.Recordset!id_treball)
  ordremodifi = atrim(formclixes.modificacions.Recordset!ordre)
  albarans.DatabaseName = formclixes.clixes.DatabaseName
  albarans.RecordSource = "SELECT Clixes_albarans.*, Clixes_detallsalb.descripcio FROM Clixes_detallsalb INNER JOIN Clixes_albarans ON Clixes_detallsalb.id_detall = Clixes_albarans.id_detall where id_treball=" + idtreball + " and ordremodificacio=" + ordremodifi + " order by clixes_albarans.ordre"
  albarans.Refresh
  'posarnomclientfacturacio
  carregar_clientfacturaciopredeterminat
  ensenyarpostit
End Sub
Sub ensenyarpostit()
   postit.visible = True
   rellotgepostit.Enabled = True
End Sub
Sub buscarcodiclientalesfactures()
  Dim rst As Recordset
  Dim numcomanda As String
  If formclixes.llistadecomandespendents.ListCount = 0 Then Exit Sub
  numcomanda = atrim(Mid(formclixes.llistadecomandespendents.List(0), 1, 6))
  Set rst = dbcomandes.OpenRecordset("select codicomptable from comandes_extres where comanda=" + atrim(cadbl(numcomanda)))
  If Not rst.EOF Then
   formclixes.modificacions.Recordset.Edit
   formclixes.modificacions.Recordset!codiclientfactclixes = cadbl(rst!codicomptable)
   formclixes.modificacions.Recordset.Update
  End If
  Set rst = Nothing
End Sub
Sub posarnomclientfacturacio()
  Dim rst As Recordset
  If cadbl(formclixes.modificacions.Recordset!codiclientfactclixes) = 0 Then buscarcodiclientalesfactures
  Set rst = dbcomandes.OpenRecordset("select codisap,nomclient from clients_codisSAP where codisap=" + atrim(cadbl(formclixes.modificacions.Recordset!codiclientfactclixes)))
  If rst.EOF Then combonomclientfact.Text = "NO HE TROBAT EL CODI COMPTABLE": GoTo fi
  combonomclientfact = atrim(rst!codisap) + " - " + atrim(rst!nomclient)
fi:
   ccodifacturacio = atrim(formclixes.modificacions.Recordset!codifacturacioclixes)
   cobservacioalbara = atrim(formclixes.modificacions.Recordset!observacionsfacturaclixes)
   If atrim(formclixes.modificacions.Recordset!empresafacturadora) <> "" Then
      Comboquifactura = IIf(atrim(formclixes.modificacions.Recordset!empresafacturadora) = "I", "Inplacsa", "Plasel")
        Else
          Comboquifactura = "Inplacsa"
          Comboquifactura_Click
   End If
   Set rst = Nothing
End Sub

Private Sub Form_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   draganddropcopiaralbara atrim(data.Files(1))
End Sub

Private Sub framealb_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   draganddropcopiaralbara atrim(data.Files(1))
End Sub

Private Sub guardar_Click()
  gravar_canvis
End Sub

Private Sub imprimir_Click()
  formclixes.crear_taules_tmp
   imprimir_albarans
End Sub

Sub ensenyar_picker(reixa As DBGrid, taula As data)
  If reixa.Columns(reixa.col).Locked = True Then Exit Sub
  r = ""
  If taula.Recordset.Fields(reixa.Columns(reixa.col).DataField).Type = 8 Then
       picker.visible = True
       If IsDate(reixa) Then
          picker.Value = reixa
         Else: picker.Value = Now
       End If
       picker.Move reixa.Container.Left + reixa.Columns(reixa.col).Left, reixa.Container.Top + reixa.Columns(reixa.col).Top
       picker.SetFocus
       SendKeys ("%{down}")
       picker.tag = reixa.Name
         Else: picker.visible = False
  End If
   
End Sub


Private Sub picker_CloseUp()
formliniesalbarans.Controls(picker.tag) = picker.Value
 formliniesalbarans.Controls(picker.tag).SetFocus
End Sub

Function albmaxordre(dbctrl As Control) As Integer
   Dim rs As Recordset
    'If albarans.Recordset.EOF Then albmaxordre = 0
   Set rs = dbctrl.Recordset.Clone
   If Not rs.EOF Then
    ' rs.Sort = "ordre"
    ' rs.Requery
     rs.MoveLast
     albmaxordre = cadbl(rs!ordre)
    Else: albmaxordre = 0
   End If
End Function

Private Sub reixaalbarans_ButtonClick(ByVal ColIndex As Integer)
  If reixaalbarans.Columns(ColIndex).Locked Then Exit Sub
   If Not validarliniaalb Then Exit Sub
   Load formseleccio
   formseleccio.Data1.DatabaseName = formclixes.clixes.DatabaseName
   formseleccio.Data1.RecordSource = "select * from clixes_detallsalb"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns("id_detall").width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           reixaalbarans.Columns("id_detall") = formseleccio.DBGrid2.Columns("id_detall")
        End If
   End If
  ' If seleccioret = 9 Then
  '         reixaalbarans.Columns("id_detall") = Null
  ' End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
'   albarans.Refresh
 'guardar_reg reixaalbarans, albarans
   'If albarans.Recordset.EditMode = 0 Then albarans.Recordset.Edit
   'albarans.Recordset.Update
   gravar_reixa reixaalbarans
End Sub
Function validarliniaalb() As Boolean
  Dim bk As Integer
  validarliniaalb = True
  If Not IsDate(reixaalbarans.Columns("data")) Then
      MsgBox "Data erronea canvia-la sisplau"
      reixaalbarans.col = reixaalbarans.Columns("data").ColIndex
      validarliniaalb = False
      Exit Function
  End If
  'On Error GoTo fi
  bk = reixaalbarans.Columns("ordre")
  If albarans.Recordset.EditMode = 0 Then albarans.Recordset.Edit
  albarans.UpdateControls
  albarans.Recordset.Update
  
  If bk > 0 Then albarans.Recordset.FindFirst "ordre=" + atrim(bk)
  Exit Function
fi:
  validarliniaalb = False
End Function

Private Sub reixaalbarans_DblClick()
If Not gravar_reixa(reixaalbarans) Then Exit Sub
   ensenyar_picker reixaalbarans, albarans
   comprovarsiestafacturat
End Sub
Sub comprovarsiestafacturat()
  Dim bk As Double
  If reixaalbarans.Columns("data").Locked Or albarans.Recordset.EOF Then Exit Sub
 If reixaalbarans.Columns(reixaalbarans.col).DataField = "facturat" Then
     If reixaalbarans.Columns("facturat") = "Sí" Then
         reixaalbarans.Columns("facturat") = False
        Else:
          If MsgBox("Vols passar tots els pendents de facturar a facturats?", vbInformation + vbYesNo, "Atenció") = vbYes Then
            dbclixes.Execute "update clixes_albarans set facturat=true,lotambelqueshafacturat=0 where id_treball=" + atrim(cadbl(formclixes.clixes.Recordset!id_treball)) + " and ordremodificacio=" + atrim(cadbl(formclixes.modificacions.Recordset!ordre))
            bk = albarans.Recordset!ordre
            albarans.Refresh
            reixaalbarans.Refresh
            albarans.Recordset.FindFirst "ordre=" + atrim(bk)
              Else: reixaalbarans.Columns("facturat") = True
          End If
     End If
     'If albarans.Recordset.EditMode = 0 Then albarans.Recordset.Edit
     'albarans.Recordset.Update
     gravar_reixa reixaalbarans
                 

 End If
End Sub

Sub draganddropcopiaralbara(vfitxer As String)
  Dim vnumalbara As String
  If InStr(1, UCase(vfitxer), ".PDF") = 0 Then MsgBox "El fitxer ha de ser PDF", vbCritical, "Error": Exit Sub
  If albarans.Recordset.EOF Then Exit Sub
  albarans.Recordset.FindFirst "albarapdfvinculat=false"
  If albarans.Recordset.NoMatch Then MsgBox "No hi ha cap albarà pendent de vincular PDF. Primer crea la linia/es corresponent abans de vincular-lo.", vbCritical, "Error": Exit Sub
  vnumalbara = atrim(albarans.Recordset!num_alb)
  If existeix(vfitxer) Then
       vnumalbara = InputBox("Entra el Nº d'albarà que vols vincular", "Vincular Albarà", vnumalbara)
       If vnumalbara <> "" Then copiarfitxersiexisteixalbara vfitxer, vnumalbara
   End If
End Sub
Sub copiarfitxersiexisteixalbara(vfitxerorigen As String, vnumalbara As String)
    Dim vfitxerdesti As String
    vfitxerdesti = ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Albarans" + "\v" + atrim(ordremodificacio) + "\" + treuresimbols(vnumalbara) + ".pdf"
    albarans.Recordset.FindFirst "num_alb='" + atrim(vnumalbara) + "'"
    If Not albarans.Recordset.NoMatch Then
        formclixes.crearcarpeta ruta_documentacio_clixes + "\" + Format(id_treball, "00000")
        formclixes.crearcarpeta ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Albarans"
        formclixes.crearcarpeta ruta_documentacio_clixes + "\" + Format(id_treball, "00000") + "\Albarans" + "\v" + atrim(ordremodificacio)
        On Error GoTo error
        Copiar_Fitxer vfitxerorigen, vfitxerdesti
        While Not albarans.Recordset.NoMatch
            albarans.Recordset.Edit
            albarans.Recordset!albarapdfvinculat = True
            albarans.Recordset.Update
            albarans.Recordset.FindNext "num_alb='" + atrim(vnumalbara) + "'"
        Wend
        albarans.Recordset.Move 0
        MsgBox "Procés acabat", vbInformation, "Copiar albarà"
           Else: MsgBox "Aquest albarà no existeix en aquest treball", vbCritical, "Error"
    End If
    Exit Sub
error:
    
End Sub

Private Sub reixaalbarans_Error(ByVal DataError As Integer, Response As Integer)
If 16389 = DataError Then
      Response = 0
      MsgBox "No hi ha una descripció sel.leccionada", vbCritical, "Atenció"
      reixaalbarans.SetFocus
   End If
End Sub

Private Sub reixaalbarans_GotFocus()
  'Dim cancel As Boolean
  reixaalbarans.Columns("facturat").Locked = True
  'If albarans.Recordset.EOF And atrim(combonomclientfact) = "" Then MsgBox "Primer has d'escullir el client on facturar", vbCritical, "Atenció": cancel = True: GoTo fi
  'If albarans.Recordset.EOF And atrim(Comboquifactura) = "" Then MsgBox "Primer has d'escullir l'empresa on facturar", vbCritical, "Atenció": cancel = True: GoTo fi
'fi:
'  If cancel Then Comboquifactura.SetFocus
  'If albarans.Recordset.EditMode > 0 Then albarans.UpdateRecord: albarans.Recordset.Edit ': activarframes True
End Sub

Private Sub reixaalbarans_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then comprovarsiestafacturat
End Sub

Private Sub reixaalbarans_OnAddNew()
   Dim gran As Integer
   gran = albmaxordre(albarans)
   albarans.Recordset!id_treball = formclixes.clixes.Recordset!id_treball
   albarans.Recordset!ordremodificacio = formclixes.modificacions.Recordset!ordre
   albarans.Recordset!ordre = gran + 1
   reixaalbarans.Columns("ordre") = gran + 1
End Sub

Private Sub reixaalbarans_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If (reixaalbarans.col = 0 Or reixaalbarans.col = 1) And reixaalbarans.Text = "" Then reixaalbarans.Text = copiarvaloranterior(reixaalbarans.Columns(reixaalbarans.col).DataField)
   possarbotonspdf
End Sub
Sub possarbotonspdf()
   beliminararxiupdf.Enabled = False
   bveurepdfalbara.Enabled = False
   If albarans.Recordset.EOF Then Exit Sub
   If albarans.Recordset!albarapdfvinculat Then
      beliminararxiupdf.Enabled = True
      bveurepdfalbara.Enabled = True
   End If
End Sub
Function copiarvaloranterior(camp As String) As String
  Dim rstcp As Recordset
  Set rstcp = albarans.Recordset.Clone
  If rstcp.EOF Then Exit Function
  rstcp.MoveLast ': rstcp.MovePrevious
  If rstcp.EOF Then Exit Function
  If atrim(rstcp!ordre) = reixaalbarans.Columns("ordre") Then rstcp.MovePrevious
  copiarvaloranterior = atrim(rstcp.Fields(camp))
End Function

Private Sub rellotgepostit_Timer()
  postit.visible = False
  rellotgepostit.Enabled = flase
End Sub
Function demanar_comanda_clixes_client(vcodiclient As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select impagats from clients where codi=" + atrim(vcodiclient))
   If Not rst.EOF Then demanar_comanda_clixes_client = rst!impagats
   Set rst = Nothing
End Function

Private Sub sortir_Click()
    If combonomclientfact = "" Then MsgBox "Quan sapigues a qui es facturen els clixes pensa a possar-ho, Gràcies.", vbInformation, "Atenció"
    If ccodifacturacio = "" And demanar_comanda_clixes_client(formclixes.clixes.Recordset!codiclienttemporal) Then
        If Not albarans.Recordset.EOF Then
         If albarans.Recordset.RecordCount > 0 Then
            MsgBox "Aquest client demana que hi hagi un numero de comanda de clixes a l'albarà, hauries de possar-lo deseguida que puguis.", vbInformation, "Atenció"
         End If
        End If
    End If
    Unload formliniesalbarans
End Sub
Sub imprimir_albarans()
  Dim rstimp As Recordset
  Dim facturats As Boolean
  facturats = True
  If MsgBox("Vols imprimir els albarans NO FACTURATS?" + Chr(10) + Chr(13) + "Si prems No es faran els FACTURATS", vbYesNo + vbDefaultButton1, "Escull") = vbYes Then
      facturats = False
  End If
  Set rstimp = dbclixes.OpenRecordset("tmp_clixes_capcalera")
  If Not rstimp.EOF Then
     While Not rstimp.EOF
       rstimp.Delete
       rstimp.MoveNext
     Wend
  End If
  With formclixes
  rstimp.AddNew
  rstimp!id_treball = .clixes.Recordset!id_treball
  rstimp!arxiuclixe = .clixes.Recordset!arxiu
  rstimp!datainici = .modificacions.Recordset!dataobertura
  rstimp!formaimp = .formaimpresio
  rstimp!estatclixe = .clixes.Recordset!estatclixe
  rstimp!client = .nomclient
  rstimp!marca = .marcaproducte
  rstimp!linia = .liniaproducte
  'rstimp!representant = .nomrepresentant
  rstimp!proveidor = .nomproveidor
  'rstimp!montadora = clixes.Recordset!montadora
  rstimp!codibarres = .clixes.Recordset!codidebarres
  rstimp!dataentrega = IIf(IsDate(.modificacions.Recordset!datatancament), .modificacions.Recordset!datatancament, .modificacions.Recordset!datatancamentprevista)
  rstimp!observacions = .modificacions.Recordset!observacions
  rstimp!sistemaimpresio = .modificacions.Recordset!sistemadimpresio
  rstimp!bandesclixes = .modificacions.Recordset!bandes
  rstimp.Update
  
  Set rstimp = dbclixes.OpenRecordset("tmp_clixes_albarans_linies")
  If Not rstimp.EOF Then
     While Not rstimp.EOF
       rstimp.Delete
       rstimp.MoveNext
     Wend
  End If
  albarans.Recordset.MoveFirst
  While Not albarans.Recordset.EOF
   If albarans.Recordset!facturat = facturats Then
    rstimp.AddNew
    rstimp!id_treball = albarans.Recordset!id_treball
    rstimp!data = albarans.Recordset!data
    rstimp!numalb = albarans.Recordset!num_alb
    rstimp!quantitat = albarans.Recordset!quantitat
    rstimp!descripcio = albarans.Recordset!descripcio
    rstimp!import = albarans.Recordset!import
    rstimp!facturat = IIf(albarans.Recordset!facturat, "Sí", "No")
    rstimp.Update
   End If
    albarans.Recordset.MoveNext
  Wend
  
  
   llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "clixes_albarans.rpt"
 llistat.DataFiles(0) = camiclixes
 If facturats Then
    llistat.Formulas(0) = "facturatsono='Albarans Facturats'"
   Else
     llistat.Formulas(0) = "facturatsono='Albarans NO Facturats'"
 End If
 llistat.DiscardSavedData = True
 llistat.Destination = crptToPrinter
 wait 1
 llistat.Action = 1
  End With
End Sub
