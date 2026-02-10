VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formaniloximpresores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment d'anilox d'impresores"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   10860
   Icon            =   "formaniloximpresores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox checknoactius 
      Caption         =   "Només       No Actius"
      Height          =   480
      Left            =   9480
      TabIndex        =   14
      Top             =   690
      Width           =   1260
   End
   Begin VB.Frame finformacio 
      Caption         =   "Informació de l'Anilox"
      Enabled         =   0   'False
      Height          =   6780
      Left            =   5400
      TabIndex        =   7
      Top             =   510
      Width           =   5415
      Begin VB.CommandButton Command5 
         Height          =   450
         Left            =   1110
         Picture         =   "formaniloximpresores.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Informació de l'estat."
         Top             =   225
         Width           =   435
      End
      Begin VB.Data datamatriculaanilox 
         Caption         =   "datamatriculaanilox"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   360
         Left            =   1200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   135
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.CommandButton Command4 
         Height          =   450
         Left            =   510
         Picture         =   "formaniloximpresores.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Eliminar matricula"
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton Command3 
         Height          =   450
         Left            =   60
         Picture         =   "formaniloximpresores.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Afegir Matricula"
         Top             =   210
         Width           =   435
      End
      Begin VB.Data datainfoanilox 
         Caption         =   "datainfoanilox"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   360
         Left            =   930
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2685
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.CommandButton Command2 
         Height          =   450
         Left            =   480
         Picture         =   "formaniloximpresores.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Eliminar la informació sel.leccionada"
         Top             =   3240
         Width           =   435
      End
      Begin VB.CommandButton Command1 
         Height          =   450
         Left            =   30
         Picture         =   "formaniloximpresores.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Afegir informació d'anilox"
         Top             =   3240
         Width           =   435
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "formaniloximpresores.frx":213C
         Height          =   2850
         Left            =   30
         OleObjectBlob   =   "formaniloximpresores.frx":2155
         TabIndex        =   8
         Top             =   3750
         Width           =   5265
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "formaniloximpresores.frx":2B54
         Height          =   2505
         Left            =   45
         OleObjectBlob   =   "formaniloximpresores.frx":2B72
         TabIndex        =   11
         Top             =   705
         Width           =   5280
      End
      Begin VB.Label etinfanilox 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Height          =   405
         Left            =   1575
         TabIndex        =   16
         Top             =   285
         Visible         =   0   'False
         Width           =   2460
      End
   End
   Begin VB.CommandButton sortir 
      Height          =   390
      Left            =   10350
      Picture         =   "formaniloximpresores.frx":3C29
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Alta  Registres"
      Top             =   75
      Width           =   390
   End
   Begin VB.CommandButton alta 
      Height          =   450
      Left            =   75
      Picture         =   "formaniloximpresores.frx":41B3
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Alta  Registres"
      Top             =   90
      Width           =   435
   End
   Begin VB.CommandButton eliminar 
      Height          =   450
      Left            =   945
      Picture         =   "formaniloximpresores.frx":473D
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   90
      Width           =   435
   End
   Begin VB.CommandButton modificar 
      Height          =   450
      Left            =   510
      Picture         =   "formaniloximpresores.frx":4CC7
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Modificar Registres"
      Top             =   90
      Width           =   435
   End
   Begin VB.CommandButton guardar 
      Height          =   450
      Left            =   1380
      Picture         =   "formaniloximpresores.frx":5251
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Acceptar els canvis (F1)."
      Top             =   90
      Width           =   435
   End
   Begin VB.ComboBox opcions 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Data dataaniloxos 
      Caption         =   "dataaniloxos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   6180
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   15
      Visible         =   0   'False
      Width           =   2880
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "formaniloximpresores.frx":57DB
      Height          =   6675
      Left            =   60
      OleObjectBlob   =   "formaniloximpresores.frx":57F2
      TabIndex        =   5
      Top             =   600
      Width           =   5280
   End
   Begin VB.Menu mmanteniment 
      Caption         =   "Manteniments"
      Begin VB.Menu mtipificacionsinfo 
         Caption         =   "Tipificacions d'Informació"
      End
   End
   Begin VB.Menu mimpresio 
      Caption         =   "Impresió"
      Begin VB.Menu llanilox 
         Caption         =   "Llistat d'Aniloxos"
      End
   End
End
Attribute VB_Name = "formaniloximpresores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
  Dim vmaq As Integer
  Dim vnommaq As String
  Dim vid As Long
  If dataaniloxos.Recordset.EditMode > 0 Then Exit Sub
  escullir_maquina vmaq, vnommaq
  If vmaq = 0 Then Exit Sub
  dataaniloxos.Recordset.AddNew
  dataaniloxos.Recordset!nommaquina = vnommaq
  dataaniloxos.Recordset!nummaquina = vmaq
  vid = dataaniloxos.Recordset!ID
  dataaniloxos.Recordset.Update
  dataaniloxos.Refresh
  dataaniloxos.Recordset.FindFirst "id=" + atrim(vid)
  'dataaniloxos.Recordset.Bookmark = dataaniloxos.Recordset.LastModified
  If dataaniloxos.Recordset.NoMatch Then Exit Sub
  dataaniloxos.Recordset.Edit
  checknoactius.Value = 0
  finformacio.Enabled = True
  DBGrid1.AllowUpdate = True
  DBGrid1.SetFocus
  DBGrid1.col = 1
End Sub
Sub escullir_maquina(vmaq As Integer, vnommaq As String)
Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.data1.DatabaseName = cami
  formseleccio.data1.RecordSource = "select codi,descripcio  from maquines where maquina='I' and donadadebaixa=null"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 500
  formseleccio.DBGrid2.Columns(1).Width = 2000
  formseleccio.Show 1
  If seleccioret = 1 Then
        If Not formseleccio.data1.Recordset.EOF Then
           vmaq = cadbl(formseleccio.DBGrid2.Columns("codi"))
           vnommaq = atrim(formseleccio.DBGrid2.Columns("descripcio"))
        End If
  End If
  formseleccio.data1.RecordSource = ""
  formseleccio.data1.Refresh
  Unload formseleccio

   
End Sub

Private Sub checknoactius_Click()
'  Command1.Enabled = Not Command1.Enabled
'  Command2.Enabled = Not Command2.Enabled
'  Command3.Enabled = Not Command3.Enabled
'  Command4.Enabled = Not Command4.Enabled
   posarinformacio
End Sub

Private Sub Command1_Click()
  'If datainfoanilox.Recordset!actiu = False Then MsgBox "Aquest anilox està donat de baixa", vbCritical, "Donat de baixa": Exit Sub
  novainformacio
End Sub
Sub novainformacio()
   Dim informacio As String
   Dim rsta As Recordset
   Dim Data As String
   Dim vdensitatesquerra As Double
   Dim vdensitatdreta As Double
   Dim vdensitatcentre As Double
   Dim vtextedensitat As String
   Dim vobs As String
   Dim vidactual As String
   
   vidactual = atrim(datamatriculaanilox.Recordset!matricula)
   Set rsta = datamatriculaanilox.Database.OpenRecordset("select * from aniloxos_informacio")
   While Not IsDate(Data)
      Data = InputBox("Entra la data que vols utilitzar. dd/mm/yy", "Data entrada", Format(Now, "dd/mm/yy"))
      If Not IsDate(Data) Then MsgBox "Aquesta data no es correcte", vbCritical, "Atenció": GoTo fi
   Wend
   informacio = escullir_informacio
   If informacio = "" Then Exit Sub
   If informacio = "DATA ENTRADA DE L'ANILOX" Then
        datainfoanilox.Recordset.FindFirst "informacio = ""DATA ENTRADA DE L'ANILOX"""
        If Not datainfoanilox.Recordset.NoMatch Then MsgBox "Ja està entrada una data d'entrada en aquest anilox no pots duplicar-lo, si de cas borra l'anterior.", vbCritical, "Atenció": Exit Sub
   End If
   If informacio = "LECTURA DE VOLUM" Then
       vdensitatesquerra = cadbl(InputBox("Entra la Lectura de l'Esquerra: ", "Lectura del volum"))
       If vdensitatesquerra = 0 Then MsgBox "Lectura a zero no vàlida", vbCritical, "Error": Exit Sub
       vdensitatcentre = cadbl(InputBox("Entra la Lectura de l'Centre: ", "Lectura del volum"))
       If vdensitatcentre = 0 Then MsgBox "Lectura a zero no vàlida", vbCritical, "Error": Exit Sub
       vdensitatdreta = cadbl(InputBox("Entra la Lectura de l'Dreta: ", "Lectura del volum"))
       If vdensitatdreta = 0 Then MsgBox "Lectura a zero no vàlida", vbCritical, "Error": Exit Sub
       
       
       vtextedensitat = " E:" + atrim(vdensitatesquerra) + " C:" + atrim(vdensitatcentre) + " D:" + atrim(vdensitatdreta)
       
   End If
   If informacio <> "OBSERVACIO LLIURE" Then
        rsta.AddNew
        rsta!IDanilox = dataaniloxos.Recordset!ID
        If informacio = "NETEJA AMB LASER" Then Data = Data + " " + Format(Now, "hh:nn")
        rsta!Data = CVDate(Data)
        rsta!matricula = datamatriculaanilox.Recordset!matricula
        rsta!matricula_inplacsa = datamatriculaanilox.Recordset!matricula_inplacsa
        rsta!informacio = informacio + vtextedensitat
        rsta.Update
   End If
   'si es la baixa de l'anilox passo DATA ENTRADA DE L'ANILOX a inactiu/ el faig servir per controlar si està actiu o no
   If informacio = "BAIXA DE L´ANILOX" Or informacio = "ENVIAT A RECTIFICAR" Then
       datainfoanilox.Database.Execute "update aniloxos_informacio set actiu=false where matricula='" + atrim(datamatriculaanilox.Recordset!matricula) + "' AND informacio=""DATA ENTRADA DE L'ANILOX"""
   End If
   If informacio = "TORNAT DE RECTIFICAR" Then
       datainfoanilox.Database.Execute "update aniloxos_informacio set actiu=true where matricula='" + atrim(datamatriculaanilox.Recordset!matricula) + "' AND informacio=""DATA ENTRADA DE L'ANILOX"""
   End If
   If informacio = "OBSERVACIO LLIURE" Then
       vobs = InputBox("Escriu la observació que vols entrar" + Chr(10) + "50 caràcters", "Observació lliure")
       If vobs <> "" Then
        rsta.AddNew
        rsta!IDanilox = dataaniloxos.Recordset!ID
        rsta!Data = Now
        rsta!matricula = datamatriculaanilox.Recordset!matricula
        rsta!matricula_inplacsa = datamatriculaanilox.Recordset!matricula_inplacsa
        rsta!informacio = Mid(vobs, 1, 50)
        rsta.Update
       End If
   End If
   datamatriculaanilox.Refresh
   datainfoanilox.Refresh
   datamatriculaanilox.Recordset.FindFirst "matricula='" + atrim(vidactual) + "'"
   If Not datainfoanilox.Recordset.EOF Then datainfoanilox.Recordset.MoveLast
fi:
   Set rsta = Nothing
End Sub
Function escullir_informacio()
Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.data1.DatabaseName = cami
  formseleccio.data1.RecordSource = "select descripcio  from tipusmantenimentanilox where descripcio<>""DATA ENTRADA DE L'ANILOX"" order by nomodificables"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 3000
  formseleccio.Show 1
  
   If seleccioret = 1 Then
        
        If Not formseleccio.data1.Recordset.EOF Then
           escullir_informacio = formseleccio.DBGrid2.Columns("descripcio")
        End If
   End If
    If seleccioret = 9 Then
        escullir_informacio = ""
   End If
   formseleccio.data1.RecordSource = ""
   formseleccio.data1.Refresh
   Unload formseleccio

   'codimuntadora.SetFocus
End Function

Private Sub Command2_Click()
   If datainfoanilox.Recordset!informacio = "DATA ENTRADA DE L'ANILOX" Then MsgBox "Aquesta informació no la pots eliminar, hauries d'eliminar la matricula", vbCritical, "Atenció": Exit Sub
   If MsgBox("Segur que vols borrar aquesta linia d'informació?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   If datainfoanilox.Recordset!informacio = "BAIXA DE L´ANILOX" Then
      MsgBox "Com que elimines la data de BAIXA DE L'ANILOX es passarà a actiu.", vbInformation, "Alta anilox"
      datainfoanilox.Database.Execute "update aniloxos_informacio set actiu=true where matricula='" + atrim(datamatriculaanilox.Recordset!matricula) + "' AND informacio=""DATA ENTRADA DE L'ANILOX"""
   End If
   datainfoanilox.Recordset.Delete
   datainfoanilox.Refresh
End Sub

Private Sub Command3_Click()
   novamatricula
   
   
End Sub
Sub novamatricula()
   Dim matricula As String
   Dim matricula_inp As String
   Dim rsta As Recordset
   Dim Data As String
   Dim vsituacio As String
   
   If datamatriculaanilox.Recordset.RecordCount >= dataaniloxos.Recordset!quantitat Then
     MsgBox "No pots entrar mes matricules que la quantitats d'aniloxos d'aquesta liniatura", vbCritical, "Atenció"
     Exit Sub
   End If
   matricula = InputBox("Entra la matrícula de l'Anilox.", "Nou anilox")
   If matricula = "" Then Exit Sub
   matricula_inp = InputBox("Entra la Referència de l'Anilox que utilitzaras a Inplacsa." + Chr(10) + "Ex: 1,2,3,...", "Nou anilox")
   If StrPtr(matricula_inp) = 0 Then Exit Sub
   If Not IsNumeric(matricula_inp) Then MsgBox "Aquest valor de Referència no es vàlid": Exit Sub
   
   Set rsta = datamatriculaanilox.Database.OpenRecordset("select * from aniloxos_informacio where idanilox=" + atrim(dataaniloxos.Recordset!ID) + " and actiu=true and matricula_inplacsa='" + atrim(matricula_inp) + "'")
   If Not rsta.EOF Then MsgBox "Aquest numero de referencia per aquest anilox no pots utilitzar-lo, està repetit.", vbCritical, "Atenció": Exit Sub
   
   Set rsta = datamatriculaanilox.Database.OpenRecordset("select * from aniloxos_informacio where matricula='" + atrim(matricula) + "'")
   If Not rsta.EOF Then MsgBox "Aquesta matricula ja existeix en algun anilox no pots utilitzar-lo", vbCritical, "Atenció": Exit Sub
   While Not IsDate(Data)
      Data = InputBox("Entra la data d'entrada que vols utilitzar. dd/mm/yy", "Data entrada", Format(Now, "dd/mm/yy"))
      If Not IsDate(Data) Then MsgBox "Aquesta data no es correcte", vbCritical, "Atenció"
   Wend
   While vsituacio <> "M" And vsituacio <> "C"
     vsituacio = UCase(InputBox("Entra la situació on està aquest anilox." + Chr(10) + " (M)Màquina o (C)Caixa.", "Situacio"))
   Wend
   
   rsta.AddNew
   rsta!IDanilox = dataaniloxos.Recordset!ID
   rsta!Data = CVDate(Data)
   rsta!matricula = matricula
   rsta!matricula_inplacsa = cadbl(matricula_inp)
   rsta!informacio = "DATA ENTRADA DE L'ANILOX"
   rsta!situacio = vsituacio
   rsta.Update
   datamatriculaanilox.Refresh
   datamatriculaanilox.Recordset.FindFirst "matricula='" + matricula + "'"
End Sub

Private Sub Command4_Click()
   If UCase(InputBox("Segur que vols borrar aquesta matricula i tota la informació?" + Chr(10) + "També pots possar una informació amb donat de baixa." + Chr(10) + " Escriu [eliminar] per eliminar matricula i informació.", "Eliminar matricula")) <> "ELIMINAR" Then Exit Sub
   datamatriculaanilox.Database.Execute "delete * from aniloxos_informacio where idanilox=" + atrim(dataaniloxos.Recordset!ID) + " and matricula='" + atrim(datamatriculaanilox.Recordset!matricula) + "'"
   datamatriculaanilox.Refresh
End Sub

Private Sub Command5_Click()
   Dim vnummatricula As String
   Dim v As String
   Dim rst As Recordset
   
   If datamatriculaanilox.Recordset.EOF Then Exit Sub
   vnummatricula = datamatriculaanilox.Recordset!matricula
   Set rst = dbtmp.OpenRecordset("select observacio from aniloxos_informacio where matricula='" + vnummatricula + "' and informacio=""DATA ENTRADA DE L'ANILOX""")
   If rst.EOF Then Exit Sub
   If atrim(rst!observacio) <> "" Then
        If MsgBox("Vols modificar-la?" + Chr(10) + atrim(rst!observacio), vbInformation + vbDefaultButton2 + vbYesNo) = vbYes Then
           v = InputBox("Detall de l'estat de l'anilox.", "Observació", atrim(rst!observacio))
        End If
         Else
         v = InputBox("Detall de l'estat de l'anilox.", "Observació", atrim(rst!observacio))
   End If
   If StrPtr(v) = 0 Then Exit Sub 'si s'ha apretat cancelar
   rst.Edit
   rst!observacio = v
   rst.Update
End Sub

Private Sub Command6_Click()
   
End Sub

Private Sub dataaniloxos_Reposition()
   posarinformacio
End Sub
Sub posarinformacio()
  datamatriculaanilox.RecordSource = "select distinct matricula,matricula_inplacsa,situacio,diesneteja,metresneteja, actiu from aniloxos_informacio where idanilox=" + atrim(cadbl(dataaniloxos.Recordset!ID)) + " and actiu=" + IIf(checknoactius.Value = 1, "False", "True") + " and informacio=""DATA ENTRADA DE L'ANILOX"""
  datamatriculaanilox.Refresh
 
End Sub
Private Sub datamatriculaanilox_Reposition()
   posarinfodematricula
End Sub

Sub posarinfodematricula()
 If Not datamatriculaanilox.Recordset.EOF Then
     datainfoanilox.RecordSource = "select data,informacio,actiu from aniloxos_informacio where matricula='" + atrim(datamatriculaanilox.Recordset!matricula) + "'"
       Else
        datainfoanilox.RecordSource = "select data,informacio,actiu from aniloxos_informacio where matricula='{[}]'"
  End If
  datainfoanilox.Refresh
End Sub
Private Sub DBGrid1_ButtonClick(ByVal ColIndex As Integer)
 If dataaniloxos.Recordset.EditMode > 0 Then
  opcions.Visible = True
  'opcions.Left = DBGrid1.Columns(ColIndex).Left + 30
  opcions.Top = DBGrid1.RowTop(DBGrid1.row) + DBGrid1.Top
  opcions.Left = DBGrid1.Columns(ColIndex).Left + DBGrid1.Left
  'opcions.Left = DBGrid1.RowTop(DBGrid1.Row) + DBGrid1.Left
  opcions.SetFocus
  SendKeys ("%{DOWN}")
 End If
End Sub

Private Sub DBGrid3_DblClick()
   Dim matricula_inp As String
   Dim vsituacio As String
   If DBGrid3.col = 0 Then
    matricula_inp = InputBox("Entra la Matricula de l'anilox." + Chr(10) + "Aquesta operació pot afectar a l'historia d'utilització d'aquest anilox només s'ha de fer si es extrictament necessari.", "Anilox")
    If StrPtr(matricula_inp) = 0 Then Exit Sub
    'If Not IsNumeric(matricula_inp) Then MsgBox "Aquest valor no val com a referència.": Exit Sub
    If InStr(1, matricula_inp, "-") > 0 Then MsgBox "La matricula no pot tenir un [-] guió." + Chr(10) + "La màquina de neteja laser no ho permet.", vbCritical + vbOKOnly, "Error": Exit Sub
    If atrim(matricula_inp) <> "" Then
     If atrim(matricula_inp) <> atrim(datamatriculaanilox.Recordset!matricula) Then
      If MsgBox("El canvi d 'una matricula influirà en tota la seva historia d'utilització." + Chr(10) + "Faré un canvi de tot l'historial guardat.", vbCritical + vbDefaultButton2 + vbYesNo, "A T E N C I Ó") = vbYes Then
          ratoli "espera"
          dataaniloxos.Database.Execute "update aniloxos_informacio set matricula='" + atrim(matricula_inp) + "' where matricula='" + atrim(datamatriculaanilox.Recordset!matricula) + "'"
          For i = 1 To 8
            dataaniloxos.Database.Execute "update aniloxtimeline set matricula" + atrim(i) + "='" + atrim(matricula_inp) + "' where matricula" + atrim(i) + "='" + atrim(datamatriculaanilox.Recordset!matricula) + "'"
            DoEvents
          Next i
          ratoli "normal"
      End If
     End If
    End If
    dataaniloxos.Recordset.Move 0
   End If
   If DBGrid3.col = 1 Then
    matricula_inp = InputBox("Entra la Referència que utilitzareu a Inplacsa.", "Anilox")
    If StrPtr(matricula_inp) = 0 Then Exit Sub
    If Not IsNumeric(matricula_inp) Then MsgBox "Aquest valor no val com a referència.": Exit Sub
    dataaniloxos.Database.Execute "update aniloxos_informacio set matricula_inplacsa=" + matricula_inp + " where matricula='" + atrim(datamatriculaanilox.Recordset!matricula) + "'"
    dataaniloxos.Recordset.Move 0
   End If
   If DBGrid3.col = 4 Then
    While vsituacio <> "M" And vsituacio <> "C"
     vsituacio = UCase(InputBox("Entra la situació on està aquest anilox." + Chr(10) + " (M)Màquina o (C)Caixa.", "Situacio"))
    Wend
    dataaniloxos.Database.Execute "update aniloxos_informacio set situacio='" + vsituacio + "' where matricula='" + atrim(datamatriculaanilox.Recordset!matricula) + "'"
    dataaniloxos.Recordset.Move 0
   End If
   If DBGrid3.col = 2 Then
    vsituacio = InputBox("Entra quants dies vols que passin abans de netejar aquest anilox." + " S'avisarà per metres o per dies.", "Dies neteja anilox")
    If StrPtr(vsituacio) = 0 Then Exit Sub
    If atrim(vsituacio) <> "0" Then If cadbl(vsituacio) = 0 Then Exit Sub
    dataaniloxos.Database.Execute "update aniloxos_informacio set diesneteja=" + vsituacio + " where matricula='" + atrim(datamatriculaanilox.Recordset!matricula) + "'"
    dataaniloxos.Recordset.Move 0
   End If
   If DBGrid3.col = 3 Then
    vsituacio = InputBox("Entra cada quants metres vols netejar aquest anilox." + " S'avisarà per metres o per dies.", "Metres neteja anilox")
    If cadbl(vsituacio) = 0 Then Exit Sub
    dataaniloxos.Database.Execute "update aniloxos_informacio set metresneteja=" + vsituacio + " where matricula='" + atrim(datamatriculaanilox.Recordset!matricula) + "'"
    dataaniloxos.Recordset.Move 0
   End If

End Sub

Private Sub DBGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If DBGrid3.Columns(DBGrid3.col).Caption = "Dies Netejar" Then
      etinfanilox.Visible = True
      etinfanilox = "Possa 0 dies per netejar al canviar."
        Else: etinfanilox.Visible = False
  End If
End Sub

Private Sub eliminar_Click()
   If dataaniloxos.Recordset.EditMode > 0 Then MsgBox "Estas editant...": Exit Sub
   If datamatriculaanilox.Recordset.RecordCount > 0 Then MsgBox "No pots eliminar un anilox si hi ha matricules relacionades." + Chr(10) + "Primer elimina les matricules.", vbCritical, "Eliminar matricules": Exit Sub
   If UCase(InputBox("Segur que vols eliminar aquest anilox?" + Chr$(10) + "També eliminaras totes les matricules relacionades" + Chr(10) + "ESCRIU [ELIMINAR] PER ACCEPTAR.", "Atenció")) <> "ELIMINAR" Then Exit Sub
   dataaniloxos.Recordset.Delete
   dataaniloxos.Refresh
End Sub

Private Sub Form_Load()
   possarnommaquines
   datamatriculaanilox.DatabaseName = cami
   datainfoanilox.DatabaseName = cami
   dataaniloxos.DatabaseName = cami
   dataaniloxos.RecordSource = "select * from aniloxos order by nummaquina"
   dataaniloxos.Refresh
   
   
  
   
End Sub
Sub possarnommaquines()
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from maquines where maquina='I' and donadadebaixa=null")
  While Not rst.EOF
   
    opcions.AddItem atrim(rst!descripcio)
    opcions.ItemData(opcions.NewIndex) = cadbl(rst!codi)
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub

Private Sub guardar_Click()
 If dataaniloxos.Recordset.EditMode = 0 Then Exit Sub
   DBGrid1.AllowUpdate = False
  dataaniloxos.Recordset.Update
  dataaniloxos.Refresh
  finformacio.Enabled = False
  dataaniloxos.Recordset.Bookmark = dataaniloxos.Recordset.LastModified
End Sub

Private Sub maqalt1_Change()
Dim i As Byte
   For i = 0 To maquinaalt1.ListCount - 1
     If maquinaalt1.ItemData(i) = cadbl(maqalt1) Then maquinaalt1.ListIndex = i
   Next i
End Sub

Private Sub maqalt2_Change()
Dim i As Byte
   For i = 0 To maquinaalt2.ListCount - 1
     If maquinaalt2.ItemData(i) = cadbl(maqalt2) Then maquinaalt2.ListIndex = i
   Next i
End Sub

Private Sub maqalt3_Change()
   Dim i As Byte
   For i = 0 To maquinaalt3.ListCount - 1
     If maquinaalt3.ItemData(i) = cadbl(maqalt3) Then maquinaalt3.ListIndex = i
   Next i
End Sub

Private Sub maquinaalt1_Click()
  maqalt1 = maquinaalt1.ItemData(maquinaalt1.ListIndex)
End Sub

Private Sub maquinaalt2_Click()
  maqalt2 = maquinaalt2.ItemData(maquinaalt2.ListIndex)
End Sub

Private Sub maquinaalt3_Click()
maqalt3 = maquinaalt3.ItemData(maquinaalt3.ListIndex)
End Sub

Private Sub llanilox_Click()
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.report
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistataniloxosilasevavida.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "comandes.mdb"
  'oreport.RecordSelectionFormula = "{aniloxosinformacio.id} in (SELECT First({aniloxos_informacio.id]) From {aniloxos_informacio} GROUP BY {aniloxos_informacio.matricula}"
  
  oreport.DiscardSavedData
   
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
  

End Sub


Private Sub modificar_Click()
  If dataaniloxos.Recordset.EditMode > 0 Then Exit Sub
   DBGrid1.AllowUpdate = True
  dataaniloxos.Recordset.Edit
'  checknoactius.Value = 0
  finformacio.Enabled = True
  DBGrid1.SetFocus
End Sub

Private Sub tipuscilindre_Change()

End Sub

Private Sub triar_Click()
   If seleccioimpresora.ListIndex = -1 Then Exit Sub
   dataaniloxos.RecordSource = "select * from cilindres where nummaquinaprincipal=" + atrim(seleccioimpresora.ItemData(seleccioimpresora.ListIndex))
   dataaniloxos.Refresh
End Sub

Private Sub seleccioimpresora_Change()

End Sub

Private Sub seleccioimpresora_Click()

End Sub
Sub guardarvalorcombo()

If dataaniloxos.Recordset.EditMode > 0 And opcions.ListIndex <> -1 Then
   DBGrid1.Text = opcions.Text
   dataaniloxos.Recordset!nummaquina = opcions.ItemData(opcions.ListIndex)
End If

opcions.Visible = False
DBGrid1.col = 1
DoEvents
DBGrid1.SetFocus
SendKeys "{TAB}"

End Sub

Private Sub mtipificacionsinfo_Click()
   Load formaltarep
  formaltarep.Caption = "Tipificacions Manteniment Anilox"
  formaltarep.data1.DatabaseName = cami
  formaltarep.data1.RecordSource = "select Descripcio from tipusmantenimentanilox where nomodificables=false"
  formaltarep.refrescar
  'formaltarep.DBGrid1.Columns(0).Visible = False
  formaltarep.DBGrid1.Refresh
  formaltarep.Show 1
End Sub

Private Sub opcions_Click()
guardarvalorcombo
End Sub

Private Sub opcions_LostFocus()
  opcions.Visible = False
End Sub

Private Sub sortir_Click()
  Unload Me
End Sub
