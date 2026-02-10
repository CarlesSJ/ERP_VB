VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formaltamaquines 
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   Icon            =   "formaltamaquines.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox colsbloc 
      Height          =   240
      Left            =   30
      TabIndex        =   9
      Text            =   "46"
      Top             =   465
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.ComboBox opcions 
      Height          =   315
      ItemData        =   "formaltamaquines.frx":0442
      Left            =   450
      List            =   "formaltamaquines.frx":045E
      TabIndex        =   8
      Text            =   "Extrussora"
      Top             =   1095
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4125
      Top             =   825
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "formaltamaquines.frx":04C1
      Height          =   5340
      Left            =   75
      Negotiate       =   -1  'True
      OleObjectBlob   =   "formaltamaquines.frx":04D1
      TabIndex        =   6
      Top             =   585
      Width           =   10455
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   75
      TabIndex        =   0
      Top             =   -75
      Width           =   10395
      Begin VB.CommandButton bextres 
         BackColor       =   &H00F1B75F&
         Caption         =   "Detalls extres"
         Height          =   330
         Left            =   7275
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1905
      End
      Begin VB.CommandButton bcontrasenya 
         BackColor       =   &H0080FFFF&
         Caption         =   "Assignar contrasenya"
         Height          =   315
         Left            =   7065
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   195
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CommandButton donardebaixaalta 
         Height          =   360
         Left            =   1500
         Picture         =   "formaltamaquines.frx":1A6C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Donar de Baixa o Alta la màquina"
         Top             =   165
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "formaltamaquines.frx":1FF6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   945
         Picture         =   "formaltamaquines.frx":2580
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   525
         Picture         =   "formaltamaquines.frx":2B0A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consulta Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   9915
         Picture         =   "formaltamaquines.frx":3094
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   9465
         Picture         =   "formaltamaquines.frx":361E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label estattaula 
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
         Left            =   3825
         TabIndex        =   7
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3645
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1350
      Visible         =   0   'False
      Width           =   1965
   End
End
Attribute VB_Name = "formaltamaquines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Sub refrescar()
 Dim tipusdato As Byte
 Dim grandoto As Integer
 Dim espais As Byte
 Data1.Refresh
 DBGrid1.ClearFields
 DBGrid1.Refresh
 DBGrid1.ReBind

 DBGrid1.AllowUpdate = False
 On Error GoTo fi
 For i = 0 To 50
   tipusdato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Type
   grandato = Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Size
   If grandato < 5 Then grandato = 5
   DBGrid1.Columns(i).Width = grandato * 100
   DBGrid1.Columns(i).Caption = UCase(DBGrid1.Columns(i).Caption)
   If Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Name = "maquina" Or Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Name = "rebmicromacro" Then
      DBGrid1.Columns(i).Button = True
   End If
   If Data1.Recordset.Fields(DBGrid1.Columns(i).DataField).Name = "donadadebaixa" Then
      DBGrid1.Columns(i).Width = grandato * 150
   End If
 Next i
fi:
End Sub

Private Sub alta_Click()
  alta_registre
End Sub

Private Sub bcontrasenya_Click()
   Dim vcontrasenya As String
   Dim vcontrasenya2 As String
   If Data1.Recordset.EOF Then Exit Sub
   If MsgBox("Segur que vols assigna una nova contrasenya a l'operari " + Data1.Recordset!descripcio + "? " + Chr(10) + IIf(Data1.Recordset!maquina = "T", "ELS TOREROS 4 DIGITS NUMERICS.", ""), vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   While InStr(1, vcontrasenya, "'") Or vcontrasenya = ""
      vcontrasenya = atrim(InputBoxEx("Entra la contrasenya." + Chr(10) + "No s'adment l'apostruf com a valor", "Contrasenya", , , , , , SPassword, 10))
      If atrim(Data1.Recordset!maquina) = "T" Then
           If Len(vcontrasenya) <> 4 Or cadbl(vcontrasenya) = 0 Then
              MsgBox "La contrasenya dels TORERUS ha de ser de 4 digits NUMERICS i no pot ser 0", vbCritical, "Error"
              vcontrasenya = ""
           End If
      End If
   Wend
   vcontrasenya2 = atrim(InputBoxEx("Verificació de la contrasenya.", "Verificació Contrasenya", , , , , , SPassword, 10))
   If vcontrasenya <> vcontrasenya2 Then MsgBox "Les contrasenyes no coincideixen.", vbCritical, "Error": Exit Sub
   dbtmp.Execute "delete * from operaris_contrasenyes where operari=" + atrim(Data1.Recordset!codi) + " and seccio='" + atrim(Data1.Recordset!maquina) + "'"
   dbtmp.Execute "insert into operaris_contrasenyes (operari,seccio,contrasenya) values (" + atrim(Data1.Recordset!codi) + ",'" + atrim(Data1.Recordset!maquina) + "','" + vcontrasenya + "')"
   
End Sub

Private Sub bextres_Click()
  If bextres.Tag = "" Then
    bextres.Tag = "1"
    formmaquinesextres.Show
    formmaquinesextres.Top = 6600
    formmaquinesextres.Left = 13600
    SetForegroundWindow formmaquinesextres.hwnd
      Else
        Unload formmaquinesextres
        bextres.Tag = ""
  End If
  
End Sub

Private Sub Command1_Click()
  acceptar
End Sub


Private Sub consultar_Click()
'buscar_registre
End Sub

Private Sub Command2_Click()
  
End Sub

Private Sub Data1_Reposition()
 If DBGrid1.Tag = "" Then DBGrid1.Tag = Data1.RecordSource
 If InStr(1, LCase(Data1.RecordSource), "from maquines") > 0 Then donardebaixaalta.Visible = True

End Sub

Private Sub DBGrid1_ButtonClick(ByVal ColIndex As Integer)
 If Data1.Recordset.Fields(ColIndex).SourceField = "rebmicromacro" And DBGrid1.Columns("maquina") <> "R" Then
   MsgBox "Només es pot escullir Micro/Macro a les Rebobinadores", vbCritical, "Error"
   Exit Sub
 End If
 If Data1.Recordset.EditMode > 0 Then
  opcions.Visible = True
  carregar_opcions Data1.Recordset.Fields(ColIndex).SourceField
  'opcions.Left = DBGrid1.Columns(ColIndex).Left + 30
  opcions.Top = DBGrid1.RowTop(DBGrid1.row) + DBGrid1.Top
  opcions.Left = DBGrid1.Columns(ColIndex).Left + DBGrid1.Left
  'opcions.Left = DBGrid1.RowTop(DBGrid1.Row) + DBGrid1.Left
  opcions.SetFocus
  SendKeys ("%{DOWN}")
 End If
  'While Screen.ActiveControl.Name = "opcions"
  '  DoEvents
  'Wend
  'If Data1.Recordset.EditMode > 0 Then DBGrid1.Col = 1: DoEvents: DBGrid1.SetFocus
End Sub
Sub carregar_opcions(vnomcamp As String)
   If vnomcamp = "maquina" Then
        opcions.Clear
        opcions.AddItem "Extrussora"
        opcions.AddItem "Impressora"
        opcions.AddItem "Laminadora"
        opcions.AddItem "Rebobinadora"
        opcions.AddItem "Soldadora"
        opcions.AddItem "Muntadora"
        opcions.AddItem "Clixes Repàs"
        opcions.AddItem "Torerus"
         opcions.Tag = "1" 'es per retallar nomes la inicial quan s'esculli la opcio
   End If
   If vnomcamp = "rebmicromacro" Then
       opcions.Clear
       opcions.AddItem "MicroCalent"
       opcions.AddItem "MicroFred"
       opcions.AddItem "Macro"
       opcions.AddItem "MicroCalent/Macro"
       opcions.AddItem "MicroFred/Macro"
       opcions.AddItem "Tots"
       opcions.AddItem " "
       opcions.Tag = "" 'es per NO retallar nomes la inicial quan s'esculli la opcio
   End If
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
  Data1.RecordSource = DBGrid1.Tag + " order by " + DBGrid1.Columns(ColIndex).DataField
  refrescar
End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If InStr(1, Me.Caption, "operari") = 0 Then
   If KeyCode > 57 And DBGrid1.Columns(DBGrid1.col).Button Then DBGrid1_ButtonClick (0)
   'KeyCode = 0
 End If
 If KeyCode = 113 And DBGrid1.AllowUpdate Then
    If colsbloc <> "" And InStr(1, colsbloc, atrim(DBGrid1.col)) Then
     DBGrid1.Text = triar_mesura
     SendKeys "{TAB}"
    End If
 End If
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If bextres.Tag = "1" Then SetForegroundWindow formmaquinesextres.hwnd: DoEvents
End Sub

Private Sub donardebaixaalta_Click()
   Dim resp As String
   Dim maq As Long
   Dim sec As String
   If Data1.Recordset.EOF Then MsgBox "Escull una màquina per donar de baixa.": Exit Sub
   If Data1.Recordset.EditMode > 0 Then Data1.Recordset.Update
   maq = Data1.Recordset!codi
   sec = Data1.Recordset!maquina
   resp = InputBox("Entra la data de baixa de la màquina [" + UCase(atrim(Data1.Recordset!descripcio)) + "]" + Chr(13) + Chr(10) + " o [ALTA] per anular la data de baixa.", "Baixa de màquina")
   If IsDate(resp) Then
       Data1.Database.Execute "update  maquines set donadadebaixa=#" + Format(resp, "mm/dd/yy") + "# where codi=" + atrim(Data1.Recordset!codi) + " and maquina='" + atrim(sec) + "'"
      Else
        If UCase(resp) = "ALTA" Then
          Data1.Database.Execute "update  maquines set donadadebaixa=null where codi=" + atrim(Data1.Recordset!codi) + " and maquina='" + atrim(sec) + "'"
        End If
        
   End If
   Data1.Refresh
   Data1.Recordset.FindFirst "codi=" + atrim(maq) + " and maquina='" + atrim(sec) + "'"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{TAB}": KeyCode = 0
 If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
 If KeyAscii > 50 Then KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub Form_Load()
  centerscreen Me
  'DBGrid1.ReBind
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload formmaquinesextres
  seleccioret = 0
End Sub

Private Sub Label1_Click()

End Sub

Private Sub modificar_Click()
 If Not Data1.Recordset.EOF Then
   DBGrid1.Enabled = True
   Data1.Recordset.Edit
   DBGrid1.AllowUpdate = True
   DBGrid1.col = 0
   DBGrid1.SetFocus
 End If
End Sub

Sub guardarvalorcombo()

If Data1.Recordset.EditMode > 0 Then
   If opcions.Tag = "1" Then
      DBGrid1.Text = Mid(opcions.List(opcions.ListIndex), 1, 1)
       Else: DBGrid1.Text = atrim(opcions.List(opcions.ListIndex))
   End If
End If
If DBGrid1.Text = "T" Then MsgBox "Recorda de posar una contrasenya pels usuaris de Torerus.", vbInformation, "Atenció"
opcions.Visible = False
DBGrid1.col = 1
DoEvents
DBGrid1.SetFocus
SendKeys "{TAB}"

End Sub

Private Sub opcions_Click()
  guardarvalorcombo
End Sub

Private Sub opcions_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then KeyCode = 0: guardarvalorcombo
End Sub

Private Sub opcions_KeyPress(KeyAscii As Integer)
  If InStr(1, "EISLR", Chr$(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub opcions_LostFocus()
opcions.Visible = False
End Sub

Private Sub sortir_Click()
 Unload Me
End Sub

Sub acceptar()
  seleccioret = 1
  Me.Hide
End Sub

Private Sub Text1_Change()
Text1.Tag = "1"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 65 And Not DBGrid1.AllowUpdate Then alta_registre: KeyCode = 0
'If KeyCode = 69 Then buscar_registre
If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then gravar_registre
'If KeyCode = 13 Then SendKeys "{TAB}": KeyCode = 0

End Sub
Sub alta_registre()
 If DBGrid1.AllowUpdate = False Then
      'DBGrid1.Enabled = True
      'DBGrid1.SetFocus
      'DBGrid1.AllowUpdate = True
      'Data1.Recordset.AddNew
      'DoEvents
      'DBGrid1.SetFocus
          If Data1.Recordset.EditMode > 0 Then Exit Sub
          If InStr(1, UCase(Data1.RecordSource), " FROM MAQUINES") > 0 Then
             Data1.Database.Execute "delete * from maquines where maquina=' ' or maquina='' or maquina=null"
          End If
          If InStr(1, UCase(Data1.RecordSource), " FROM OPERARIS") > 0 Then
             Data1.Database.Execute "delete * from operaris where maquina=' ' or maquina='' or maquina=null"
          End If
          Data1.Recordset.AddNew
          Data1.Recordset!maquina = " "
          Data1.Recordset.Update
         ' Data1.Refresh
          Data1.Recordset.Bookmark = Data1.Recordset.LastModified
          Data1.Recordset.Edit
          DBGrid1.AllowUpdate = True
          DBGrid1.SetFocus

 End If
End Sub
Sub gravar_registre()
On Error GoTo err
 If Data1.Recordset.EditMode > 0 Then
      Data1.Recordset.Update
      DBGrid1.AllowUpdate = False
      Data1.Recordset.Bookmark = Data1.Recordset.LastModified
 End If

Exit Sub
err:
  MsgBox "Hi ha hagut un error al gravar les dades. Potser algun valor està duplicat a la clau. Torna-ho a provar", vbCritical, "Atenció"
End Sub
Sub cancelar_registre()
  If Data1.Recordset.EditMode > 0 Then
   Data1.Recordset.CancelUpdate
   Else: Unload Me
  End If
    
  DBGrid1.AllowUpdate = False
End Sub

Private Sub eliminar_Click()
 On Error GoTo err
  If MsgBox("Segur que vols Eliminar?", vbYesNo + vbCritical, "Atenció") = 6 Then
    Data1.Recordset.Delete
    Data1.Recordset.MoveNext
    If Data1.Recordset.EOF Then Data1.Recordset.MovePrevious
  End If
 Exit Sub
err:
  MsgBox "No s'ha pogut eliminar possiblement perque tingui registres relacionats. O bé no hi ha res per eliminar."
End Sub

Private Sub Timer1_Timer()
 estattaula.Caption = textestattaula(Data1.EditMode)
  If estattaula.ForeColor <> QBColor(0) Then
     estattaula.ForeColor = QBColor(0)
    Else: estattaula.ForeColor = QBColor(14)
  End If
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
  If Len(DBGrid1.Text) >= Data1.Recordset.Fields(DBGrid1.Columns(DBGrid1.col).DataField).Size And KeyAscii > 55 And UCase(DBGrid1.Columns(DBGrid1.col).DataField) <> "MAQUINA" Then KeyAscii = 0
  If UCase(DBGrid1.Columns(DBGrid1.col).DataField) = "DATADEBAIXA" Then MsgBox "FES SERVIR EL BOTÓ DE BAIXA PER DONAR LA DATA DE BAIXA": KeyAscii = 0: Exit Sub
  If InStr(1, colsbloc, DBGrid1.col) Then
     KeyAscii = 0

 End If
 If UCase(DBGrid1.Columns(DBGrid1.col).DataField) = "MAQUINA" And Data1.Recordset.EditMode > 0 Then
   If InStr(1, "EISLR", Chr$(KeyAscii)) = 0 Then
      KeyAscii = 0
     Else: DBGrid1.Text = ""
   End If
 End If
End Sub
Function triar_mesura() As String
  Load formseleccio
  formseleccio.Caption = "Selecciona un Valor"
  formseleccio.Data1.DatabaseName = Data1.DatabaseName
  formseleccio.Data1.RecordSource = "select * from mesureslineals"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   triar_mesura = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
End Function

