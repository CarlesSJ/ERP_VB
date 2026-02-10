VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formseleccio 
   Caption         =   "c"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "formseleccio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "formseleccio.frx":058A
      Height          =   5310
      Left            =   90
      OleObjectBlob   =   "formseleccio.frx":059A
      TabIndex        =   5
      Top             =   645
      Width           =   4500
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   75
      TabIndex        =   0
      Top             =   -60
      Width           =   4515
      Begin VB.CommandButton Command2 
         Caption         =   "Cap"
         Height          =   375
         Left            =   3210
         TabIndex        =   6
         ToolTipText     =   "No sel.leccionar res."
         Top             =   180
         Width           =   405
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2790
         Picture         =   "formseleccio.frx":0F6D
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "filtre"
         ToolTipText     =   "Buscar"
         Top             =   180
         Width           =   405
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   75
         TabIndex        =   3
         Top             =   195
         Width           =   2655
      End
      Begin VB.CommandButton sortirs 
         Height          =   375
         Left            =   4050
         Picture         =   "formseleccio.frx":14F7
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Tancar"
         Top             =   180
         Width           =   405
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   3630
         Picture         =   "formseleccio.frx":1A81
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Acceptar"
         Top             =   180
         Width           =   405
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1275
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1965
   End
End
Attribute VB_Name = "formseleccio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iniconfigreixa As String

Sub refrescar()
 Dim tipusdato As Byte
 Dim grandoto As Integer
 Dim espais As Byte
 
 Data1.Refresh
 'If Data1.Recordset.EOF Then GoTo fi
 DBGrid2.Refresh
 'DBGrid2.ReBind
 DBGrid2.AllowUpdate = False
' On Error GoTo fi
 For i = 0 To DBGrid2.Columns.Count - 1
   tipusdato = Data1.Recordset.Fields(DBGrid2.Columns(i).DataField).Type
   grandato = Data1.Recordset.Fields(DBGrid2.Columns(i).DataField).Size
   If grandato < 5 Then grandato = 5
   DBGrid2.Columns(i).Width = grandato * 100
   DBGrid2.Columns(i).caption = UCase(DBGrid2.Columns(i).caption)
    
 Next i
 carregar_amples_reixa
fi:
If formseleccio.Tag = "1" Then DBGrid2.Columns(0).Width = 0
End Sub

Private Sub Command1_Click()
  acceptar
End Sub

Private Sub Command2_Click()
seleccioret = 9
  Me.Hide
End Sub

Private Sub Command3_Click()
 Dim colu As Byte
 colu = DBGrid2.col
 If Command3.Tag <> "filtre" Then
  If Text1.Tag = "1" Then
   Data1.Recordset.FindFirst (DBGrid2.Columns(DBGrid2.col).DataField + " like '*" + Text1.Text + "*'")
   Text1.Tag = ""
    Else: Data1.Recordset.FindNext (DBGrid2.Columns(DBGrid2.col).DataField + " like '*" + Text1.Text + "*'"): Text1.Tag = ""
  End If
   Else
      Data1.RecordSource = possarfiltre
      'MsgBox Data1.RecordSource
      Data1.Refresh
      refrescar
   End If
  DBGrid2.Visible = True
  DBGrid2.SetFocus
  DBGrid2.col = colu
End Sub
Function possarfiltre()
   Dim va As String
   Dim res As String
   Dim andowhere As String
   Dim nomdelcamp As String
   va = formseleccio.Tag
   andowhere = " where "
   If InStr(1, UCase(va), " WHERE ") Then andowhere = " and "
   res = va
   nomdelcamp = Data1.Recordset.Fields(DBGrid2.col).SourceTable + "." + Data1.Recordset.Fields(DBGrid2.col).SourceField
   If InStr(1, va, "order") Then res = Mid(va, 1, InStr(1, va, "order") - 1)
   res = res + andowhere + (nomdelcamp + " like '*" + Text1.Text + "*'")
   If InStr(1, va, "order") Then res = res + Mid(va, InStr(1, va, " order"))
   possarfiltre = res
End Function
Private Sub Data1_Reposition()
 If DBGrid2.Tag = "" Then DBGrid2.Tag = Data1.RecordSource

End Sub

Private Sub DBGrid2_DblClick()
acceptar
End Sub

Private Sub DBGrid2_HeadClick(ByVal ColIndex As Integer)
  Dim consulta As String
  On Error GoTo fi
  If InStr(1, DBGrid2.Tag, "order by") > 0 Then Exit Sub
  Data1.RecordSource = DBGrid2.Tag + " order by " + DBGrid2.Columns(ColIndex).DataField
'  MsgBox Data1.RecordSource
  'Data1.Refresh
  refrescar
  DBGrid2.col = ColIndex
fi:
End Sub
Sub substituir(variable As String, buscar As String, canviar As String)
   Dim linia As String
   linia = variable
   comença = InStr(1, linia, buscar) - 1
   If comença < 1 Then Exit Sub
   acaba = comença + Len(buscar) + 1
   linia = Mid(linia, 1, comença) + canviar + Mid(linia, acaba)
   'MsgBox linia
   variable = linia
End Sub
Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 If KeyCode = 38 And DBGrid2.Row = 0 Then Text1.SetFocus
End Sub

Private Sub DBGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then acceptar
End Sub

Private Sub Form_Activate()
 Dim taula As String
If DBGrid2.Columns.Count > 1 Then DBGrid2.col = 1
DBGrid2.SetFocus
DoEvents
Text1.SetFocus
Data1.RecordSource = Data1.RecordSource + " "
If formseleccio.caption = "c" Then
 If InStr(1, LCase(Data1.RecordSource), " from ") <> 0 Then
    formseleccio.caption = "Busqueda de " + UCase(Mid(Data1.RecordSource, InStr(1, LCase(Data1.RecordSource), "from ") + 5, InStr(1, Mid(Data1.RecordSource, InStr(1, LCase(Data1.RecordSource), "from ") + 5), " ")))
    taula = UCase(Mid(Data1.RecordSource, InStr(1, LCase(Data1.RecordSource), "from ") + 5, InStr(1, Mid(Data1.RecordSource, InStr(1, LCase(Data1.RecordSource), "from ") + 5), " ")))
  Else: formseleccio.caption = "Busqueda de " + UCase(Data1.RecordSource)
        taula = UCase(Data1.RecordSource)
 End If
End If
If cadbl(Text1.Tag) > 0 Then DBGrid2.col = cadbl(Text1.Tag): Text1.Tag = ""
If Me.Tag = "" Then Me.Tag = Data1.RecordSource
iniconfigreixa = "reixasel" + Trim(taula) + ".ini"

If Command2.Tag <> "" Then
   DBGrid2.col = cadbl(Command2.Tag): Command2.Tag = ""
   guardar_amples_reixa
 Else: carregar_amples_reixa
End If
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 13 Or KeyCode = 112) And Not DBGrid2.AllowUpdate And Screen.ActiveControl.Name = "DBGrid2" Then acceptar
If (KeyCode = 13 Or KeyCode = 112) And Screen.ActiveForm.Name = "formseleccio" And Screen.ActiveControl.Name = "Text1" Then
   Command3_Click
End If
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
End Sub

Private Sub Form_Load()
  centerscreen Me
 
End Sub

Private Sub Form_Resize()
On Error Resume Next
DBGrid2.Width = formseleccio.Width - 500
Frame1.Width = formseleccio.Width - 300
DBGrid2.Height = formseleccio.Height - DBGrid2.Top - 700
sortirs.Left = Frame1.Width - sortirs.Width - 75
End Sub

Private Sub Form_Unload(Cancel As Integer)
seleccioret = 0
guardar_amples_reixa
End Sub


Sub acceptar()
  seleccioret = 1
  Me.Hide
End Sub

Sub carregar_amples_reixa()
 Dim ample As String
 Dim nom As String
 If existeix("c:\windows\" + iniconfigreixa) Then
  For j = 0 To DBGrid2.Columns.Count - 1
   ample = llegir_ini("AmplesReixa", UCase(DBGrid2.Columns(j).DataField), iniconfigreixa)
   nom = llegir_ini("NomsReixa", UCase(DBGrid2.Columns(j).DataField), iniconfigreixa)
   
   If ample <> "{[}]" Then
      DBGrid2.Columns(j).Width = cadbl(ample)
      DBGrid2.Columns(j).caption = nom
    Else: Exit Sub
   End If
 Next j
 If cadbl(llegir_ini("Amplesformulari", "ample", iniconfigreixa)) > 1000 Then
  
  formseleccio.Width = cadbl(llegir_ini("Amplesformulari", "ample", iniconfigreixa))
  formseleccio.Height = cadbl(llegir_ini("Amplesformulari", "alt", iniconfigreixa))
  formseleccio.Top = cadbl(llegir_ini("Posicioformulari", "top", iniconfigreixa))
  formseleccio.Left = cadbl(llegir_ini("Posicioformulari", "left", iniconfigreixa))
    Else
      formseleccio.Top = (Screen.Height / 2) - (bobinesdentrada.Height / 2)
      formseleccio.Left = (Screen.Width / 2) - (bobinesdentrada.Width / 2)
 End If
 Form_Resize
End If
End Sub
Sub guardar_amples_reixa()
If iniconfigreixa <> "" Then
  For j = 0 To DBGrid2.Columns.Count - 1
   escriure_ini "AmplesReixa", UCase(DBGrid2.Columns(j).DataField), atrim(DBGrid2.Columns(j).Width), iniconfigreixa
   escriure_ini "NomsReixa", UCase(DBGrid2.Columns(j).DataField), atrim(DBGrid2.Columns(j).caption), iniconfigreixa
 Next j
End If
escriure_ini "Amplesformulari", "ample", atrim(formseleccio.Width), iniconfigreixa
escriure_ini "Amplesformulari", "alt", atrim(formseleccio.Height), iniconfigreixa
escriure_ini "Posicioformulari", "left", atrim(formseleccio.Left), iniconfigreixa
escriure_ini "Posicioformulari", "top", atrim(formseleccio.Top), iniconfigreixa
End Sub

Private Sub sortirs_Click()
 Unload Me
End Sub

Private Sub Text1_Change()
Text1.Tag = "1"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then DBGrid2.SetFocus

  'If KeyCode = 13 Then Command3_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub
