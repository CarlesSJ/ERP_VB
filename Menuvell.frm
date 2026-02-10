VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manteniment de Clients"
   ClientHeight    =   8010
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   525
         Picture         =   "Menu.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Consulta Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton gravar 
         Height          =   450
         Left            =   9375
         Picture         =   "Menu.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Modificacio Registres"
         Top             =   225
         Width           =   450
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   1425
         Picture         =   "Menu.frx":0690
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Eliminacio Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "Menu.frx":09A2
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Alta  Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.Data clients 
         Caption         =   "                     Clients"
         Connect         =   "Access"
         DatabaseName    =   "c:\misdoc~1\commandes\comandes.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   390
         Left            =   3975
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "clients"
         Top             =   225
         Width           =   3840
      End
      Begin VB.CommandButton sortir 
         Height          =   450
         Left            =   9900
         Picture         =   "Menu.frx":0DD4
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Alta  Registres"
         Top             =   225
         Width           =   450
      End
      Begin VB.CommandButton consultar 
         Height          =   360
         Left            =   975
         Picture         =   "Menu.frx":12D6
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Consulta Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.Label estattaula 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3810
         TabIndex        =   81
         Top             =   300
         Width           =   105
      End
   End
   Begin VB.Frame areadatos 
      Enabled         =   0   'False
      Height          =   7290
      Left            =   75
      TabIndex        =   2
      Top             =   675
      Width           =   10455
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   8700
         Top             =   150
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   5400
         Picture         =   "Menu.frx":1658
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   3300
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2250
         Picture         =   "Menu.frx":1A1E
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   3300
         Width           =   315
      End
      Begin VB.TextBox Text42 
         Height          =   285
         Left            =   3975
         MousePointer    =   3  'I-Beam
         TabIndex        =   29
         Top             =   3300
         Width           =   1365
      End
      Begin VB.TextBox Text41 
         Height          =   285
         Left            =   975
         TabIndex        =   28
         Top             =   3300
         Width           =   1215
      End
      Begin VB.TextBox Text40 
         DataField       =   "obssol1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   37
         Top             =   6675
         Width           =   7335
      End
      Begin VB.TextBox Text39 
         DataField       =   "obssol2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   43
         Top             =   6915
         Width           =   7335
      End
      Begin VB.TextBox Text38 
         DataField       =   "obsreb1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   36
         Top             =   6075
         Width           =   7335
      End
      Begin VB.TextBox Text37 
         DataField       =   "obsreb2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   42
         Top             =   6315
         Width           =   7335
      End
      Begin VB.TextBox Text36 
         DataField       =   "obslam1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   35
         Top             =   5475
         Width           =   7335
      End
      Begin VB.TextBox Text35 
         DataField       =   "obslam2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   41
         Top             =   5715
         Width           =   7335
      End
      Begin VB.TextBox Text34 
         DataField       =   "obsimp1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   34
         Top             =   4875
         Width           =   7335
      End
      Begin VB.TextBox Text33 
         DataField       =   "obsimp2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   40
         Top             =   5115
         Width           =   7335
      End
      Begin VB.TextBox Text32 
         DataField       =   "obsext1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   33
         Top             =   4275
         Width           =   7335
      End
      Begin VB.TextBox Text9 
         DataField       =   "obsext2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   39
         Top             =   4515
         Width           =   7335
      End
      Begin VB.TextBox Text7 
         DataField       =   "observacions1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   32
         Top             =   3675
         Width           =   7335
      End
      Begin VB.TextBox Text8 
         DataField       =   "observacions2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   765
         TabIndex        =   38
         Top             =   3915
         Width           =   7335
      End
      Begin VB.TextBox Text31 
         DataField       =   "horaridesc"
         DataSource      =   "clients"
         Height          =   285
         Left            =   7920
         TabIndex        =   24
         Top             =   2310
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Certificat Qualitat"
         DataField       =   "certqualitat"
         DataSource      =   "clients"
         Height          =   255
         Left            =   8355
         TabIndex        =   31
         Top             =   3915
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Albara Valorat"
         DataField       =   "albvalorat"
         DataSource      =   "clients"
         Height          =   255
         Left            =   8355
         TabIndex        =   30
         Top             =   3675
         Width           =   1935
      End
      Begin VB.TextBox Text30 
         DataField       =   "numproveidor"
         DataSource      =   "clients"
         Height          =   285
         Left            =   7920
         TabIndex        =   25
         Top             =   2595
         Width           =   1215
      End
      Begin VB.TextBox Text29 
         DataSource      =   "clients"
         Height          =   285
         Left            =   7560
         TabIndex        =   45
         Top             =   3315
         Width           =   2775
      End
      Begin VB.TextBox Text28 
         DataField       =   "formapag"
         DataSource      =   "clients"
         Height          =   285
         Left            =   6885
         TabIndex        =   27
         Top             =   3315
         Width           =   615
      End
      Begin VB.TextBox Text27 
         DataSource      =   "clients"
         Height          =   285
         Left            =   7560
         TabIndex        =   44
         Top             =   2955
         Width           =   2775
      End
      Begin VB.TextBox Text26 
         DataField       =   "representant"
         DataSource      =   "clients"
         Height          =   285
         Left            =   6885
         TabIndex        =   26
         Top             =   2955
         Width           =   615
      End
      Begin VB.TextBox Text25 
         DataField       =   "email"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   2880
         Width           =   4455
      End
      Begin VB.TextBox Text24 
         DataField       =   "obsfax2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Top             =   2595
         Width           =   3855
      End
      Begin VB.TextBox Text23 
         DataField       =   "fax2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   2595
         Width           =   1815
      End
      Begin VB.TextBox Text22 
         DataField       =   "obsfax1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   2760
         TabIndex        =   13
         Top             =   2310
         Width           =   3855
      End
      Begin VB.TextBox Text21 
         DataField       =   "fax1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   2310
         Width           =   1815
      End
      Begin VB.TextBox Text20 
         DataField       =   "obstel2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   2760
         TabIndex        =   11
         Top             =   2025
         Width           =   3855
      End
      Begin VB.TextBox Text19 
         DataField       =   "telefon2"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Top             =   2025
         Width           =   1815
      End
      Begin VB.TextBox Text18 
         DataField       =   "obstel1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   2760
         TabIndex        =   9
         Top             =   1740
         Width           =   3855
      End
      Begin VB.TextBox Text17 
         DataField       =   "telefon1"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   1740
         Width           =   1815
      End
      Begin VB.TextBox Text16 
         DataField       =   "faxe"
         DataSource      =   "clients"
         Height          =   285
         Left            =   6720
         TabIndex        =   23
         Top             =   2025
         Width           =   3495
      End
      Begin VB.TextBox Text15 
         DataField       =   "telefone"
         DataSource      =   "clients"
         Height          =   285
         Left            =   6720
         TabIndex        =   22
         Top             =   1740
         Width           =   3495
      End
      Begin VB.TextBox Text14 
         DataField       =   "provinciae"
         DataSource      =   "clients"
         Height          =   285
         Left            =   7920
         TabIndex        =   21
         Top             =   1455
         Width           =   2295
      End
      Begin VB.TextBox Text13 
         DataField       =   "codipostale"
         DataSource      =   "clients"
         Height          =   285
         Left            =   6720
         TabIndex        =   20
         Top             =   1455
         Width           =   1095
      End
      Begin VB.TextBox Text12 
         DataField       =   "poblacioe"
         DataSource      =   "clients"
         Height          =   285
         Left            =   6720
         TabIndex        =   19
         Top             =   1170
         Width           =   3495
      End
      Begin VB.TextBox Text11 
         DataField       =   "domicilie"
         DataSource      =   "clients"
         Height          =   285
         Left            =   6720
         TabIndex        =   18
         Top             =   885
         Width           =   3495
      End
      Begin VB.TextBox Text10 
         DataField       =   "nome"
         DataSource      =   "clients"
         Height          =   285
         Left            =   6720
         TabIndex        =   17
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox Text6 
         DataField       =   "poblacio"
         DataSource      =   "clients"
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   1455
         Width           =   4455
      End
      Begin VB.TextBox Text5 
         DataField       =   "codipostal"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   1455
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         DataField       =   "poblacio"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   1170
         Width           =   5775
      End
      Begin VB.TextBox Text3 
         DataField       =   "domicili"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   885
         Width           =   5775
      End
      Begin VB.TextBox Text2 
         DataField       =   "nom"
         DataSource      =   "clients"
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox Text1 
         DataField       =   "codi"
         DataSource      =   "clients"
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Arxiu Ult:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   26
         Left            =   3075
         TabIndex        =   77
         Top             =   3300
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Arxiu Exp:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   6
         Left            =   150
         TabIndex        =   76
         Top             =   3300
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Obs."
         DataSource      =   "clients"
         Height          =   255
         Index           =   25
         Left            =   75
         TabIndex        =   70
         Top             =   6675
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Soldad."
         DataSource      =   "clients"
         Height          =   255
         Index           =   24
         Left            =   75
         TabIndex        =   69
         Top             =   6960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Obs."
         DataSource      =   "clients"
         Height          =   255
         Index           =   23
         Left            =   75
         TabIndex        =   68
         Top             =   6075
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Rebo."
         DataSource      =   "clients"
         Height          =   255
         Index           =   22
         Left            =   75
         TabIndex        =   67
         Top             =   6360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Obs."
         DataSource      =   "clients"
         Height          =   255
         Index           =   21
         Left            =   75
         TabIndex        =   66
         Top             =   5475
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Lami."
         DataSource      =   "clients"
         Height          =   255
         Index           =   20
         Left            =   75
         TabIndex        =   65
         Top             =   5760
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Obs."
         DataSource      =   "clients"
         Height          =   255
         Index           =   19
         Left            =   75
         TabIndex        =   64
         Top             =   4875
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Impr."
         DataSource      =   "clients"
         Height          =   255
         Index           =   18
         Left            =   75
         TabIndex        =   63
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Obs."
         DataSource      =   "clients"
         Height          =   255
         Index           =   17
         Left            =   75
         TabIndex        =   62
         Top             =   4275
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Extru."
         DataSource      =   "clients"
         Height          =   255
         Index           =   16
         Left            =   75
         TabIndex        =   61
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Obs.  "
         DataSource      =   "clients"
         Height          =   255
         Index           =   15
         Left            =   75
         TabIndex        =   60
         Top             =   3675
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Codi:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   5
         Left            =   75
         TabIndex        =   59
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Num Prov:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   4
         Left            =   6765
         TabIndex        =   58
         Top             =   2670
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Horari Entrega"
         DataSource      =   "clients"
         Height          =   255
         Index           =   14
         Left            =   6765
         TabIndex        =   57
         Top             =   2385
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Forma Pag:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   13
         Left            =   5760
         TabIndex        =   56
         Top             =   3390
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Representant:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   12
         Left            =   5760
         TabIndex        =   55
         Top             =   3030
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail"
         DataSource      =   "clients"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   54
         Top             =   2955
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Fax2"
         DataSource      =   "clients"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   53
         Top             =   2670
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Fax1"
         DataSource      =   "clients"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   52
         Top             =   2385
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Telf2"
         DataSource      =   "clients"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   51
         Top             =   2100
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Telf1"
         DataSource      =   "clients"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   50
         Top             =   1815
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cp/Pr:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   49
         Top             =   1530
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Pob:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   1170
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Dom:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   885
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Nom:"
         DataSource      =   "clients"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub camp1_Change()

End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub alta_Click()
alta_registre
End Sub

Private Sub clients_Reposition()
  carregar_lookups
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Sub triarrepresentant()
  Load formseleccio
  formseleccio.Data1.DatabaseName = clients.DatabaseName
  formseleccio.Data1.RecordSource = "select * from clients"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   MsgBox formseleccio.Data1.Recordset!codi
  End If
  Unload formseleccio
  
End Sub

Private Sub Command3_Click()

End Sub

Private Sub consultar_Click()
  buscant = True
  alta_registre
  deixartotblanc
  
End Sub

Private Sub eliminar_Click()
 On Error GoTo err
  If MsgBox("Segur que vols Eliminar?", vbYesNo + vbCritical, "Atenció") = 6 Then
    clients.Recordset.Delete
    clients.Recordset.MoveNext
    If clients.Recordset.EOF Then clients.Recordset.MovePrevious
  End If
 Exit Sub
err:
  MsgBox "No s'ha pogut eliminar possiblement perque tingui registres relacionats. O bé no hi ha res per eliminar."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 65 Then alta_registre: KeyCode = 0
If KeyCode = 69 Then buscar_registre
If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then gravar_registre
If KeyCode = 13 Then SendKeys "{TAB}": KeyCode = 0

End Sub
Sub buscar_registre()

End Sub
Sub alta_registre()
 If areadatos.Enabled = False Then
      areadatos.Enabled = True
      clients.Recordset.AddNew
      DoEvents
      Text1.Enabled = True
      'busco el mes gran i el poso a codi +1
      If Not buscant Then
        Set rsttmp = dbtmp.OpenRecordset("select max(codi) as [grancodi] from clients")
        If Not rsttmp.EOF Then
          Text1 = atrim(cadbl(rsttmp!grancodi) + 1)
         Else: Text1 = "1"
        End If
      End If
      Text1.SetFocus
 End If
End Sub
Sub gravar_registre()
 If areadatos.Enabled And Not buscant Then
    Text1.Enabled = False
    sortir.SetFocus
    DoEvents
    If Screen.ActiveControl.Name = "sortir" Then
      clients.Recordset.Update
      areadatos.Enabled = False
      clients.Recordset.Bookmark = clients.Recordset.LastModified
    End If
 End If
 If buscant Then finalitzarbusqueda
End Sub
Sub cancelar_registre()
  If clients.Recordset.EditMode > 0 Then
   clients.Recordset.CancelUpdate
   areadatos.Enabled = False
   Text1.Enabled = False
   buscant = False
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
End Sub

Private Sub Form_Load()
centerscreen Me
clients.DatabaseName = "c:\misdoc~1\commandes\comandes.mdb"
clients.RecordSource = "clients"
clients.Refresh
Set dbtmp = OpenDatabase(clients.DatabaseName)
possarvalordcamps
End Sub

Private Sub frame2_Click()


End Sub

Private Sub gravar_Click()
gravar_registre
End Sub

Private Sub modificar_Click()
   areadatos.Enabled = True
   clients.Recordset.Edit
   Text2.SetFocus
End Sub

Private Sub sortir_Click()
 Unload Me
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_LostFocus()
 If Not buscant And clients.Recordset.EditMode > 0 Then
   Set rsttmp = dbtmp.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(Text1.Text)))
   If rsttmp.RecordCount > 0 Then MsgBox "Aquest codi ja existeix haurieu de canviar-lo": If areadatos.Enabled Then Text1.SetFocus
 End If
End Sub

Private Sub Timer1_Timer()
  estattaula.Caption = textestattaula(clients.EditMode)
  If estattaula.ForeColor <> QBColor(0) Then
     estattaula.ForeColor = QBColor(0)
    Else: estattaula.ForeColor = QBColor(14)
  End If
End Sub


Sub recorregutregistres()
 Dim objecte As Object
 queryorder = ""
 querywhere = ""
 'On Error Resume Next
 For Each objecte In Form1
    If TypeOf objecte Is TextBox Then
      If objecte.DataField <> "" Then ' Si Texto es igual "Hola".
        If objecte.Text <> "" Then evaluarcontingut objecte.DataField, objecte.Text, clients.Recordset.Fields(objecte.DataField).Type
     End If
    End If
Next

End Sub


Function evaluarcontingut(camp As String, valor As String, tipusdato As Byte) As String
  Dim rest As String
  rest = ""
  evaluarcontingut = ""
  If triarordre(camp, valor) Then Exit Function
  If tipusdato = 10 Then
   If InStr(1, valor, "*") Or InStr(1, valor, "?") Then
      rest = " like '" + valor + "'"
     Else
       If InStr(1, valor, ">") Or InStr(1, valor, "<") Or InStr(1, valor, "=") Then
           rest = "'" + valor + "'"
        Else: rest = "=" + "'" + valor + "'"
       End If
   End If
  End If
  If tipusdato <> 10 Then
    If InStr(1, valor, ">") Or InStr(1, valor, "<") Or InStr(1, valor, "=") Then
           rest = atrim(cadbl(valor))
        Else: rest = "=" + atrim(cadbl(valor))
    End If
  End If
  rest = camp + rest
  evaluarcontingut = rest
  
  If querywhere = "" Then
     querywhere = rest
    Else
     querywhere = querywhere + " and " + rest + " "
  End If
End Function

Function triarordre(camp As String, valorord As String) As Boolean
  Dim ord As String
  triarordre = False
  If InStr(1, valorord, "<<") Then ord = camp + " " + " ASC"
  If InStr(1, valorord, ">>") Then ord = camp + " " + " DESC"
  If ord <> "" Then
      triarordre = True
    Else: Exit Function
  End If
  If queryorder = "" Then
     queryorder = ord
   Else: queryorder = queryorder + ", " + ord
  End If
  
End Function
Sub finalitzarbusqueda()
 recorregutregistres
 If clients.Recordset.EditMode > 0 Then clients.Recordset.CancelUpdate
 buscant = False
 Text1.Enabled = True
 areadatos.Enabled = False
 If queryorder <> "" Then queryorder = " Order By " + queryorder
 If querywhere <> "" Then querywhere = " Where " + querywhere
 clients.RecordSource = "select * from clients " + querywhere + queryorder
 clients.Refresh
End Sub

Sub deixartotblanc()
 For Each objecte In Form1
    If TypeOf objecte Is TextBox Then
      If objecte.DataField <> "" Then ' Si Texto es igual "Hola".
        objecte.Text = ""
     End If
    End If
Next

End Sub

Sub carregar_lookups()

End Sub
Sub possarvalordcamps()
 For Each objecte In Form1
    If TypeOf objecte Is TextBox Then
      If objecte.DataField <> "" Then
         objecte.MaxLength = clients.Recordset.Fields(objecte.DataField).Size
      End If
    End If
Next

End Sub
