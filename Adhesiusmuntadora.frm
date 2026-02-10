VERSION 5.00
Begin VB.Form formadhesiusmuntadora 
   Caption         =   "Manteniment d'Adhesius Muntadora"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   Icon            =   "Adhesiusmuntadora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fdades 
      Caption         =   "Dades Bàsiques"
      Enabled         =   0   'False
      Height          =   3120
      Left            =   60
      TabIndex        =   16
      Top             =   720
      Width           =   6420
      Begin VB.TextBox Text2 
         DataField       =   "inicialsfoam"
         DataSource      =   "dataadhesius"
         Height          =   300
         Left            =   5820
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1425
         Width           =   495
      End
      Begin VB.ComboBox selsubfamilia 
         DataField       =   "subfamilia"
         DataSource      =   "dataadhesius"
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   2240
         Width           =   4710
      End
      Begin VB.ComboBox selfamilia 
         DataField       =   "familia"
         DataSource      =   "dataadhesius"
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Top             =   1825
         Width           =   4710
      End
      Begin VB.ComboBox selproveidor 
         DataField       =   "nomproveidor"
         DataSource      =   "dataadhesius"
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   610
         Width           =   4710
      End
      Begin VB.TextBox Text1 
         DataField       =   "gruix"
         DataSource      =   "dataadhesius"
         Height          =   300
         Left            =   1140
         TabIndex        =   8
         Top             =   2655
         Width           =   1800
      End
      Begin VB.TextBox descripcioinplacsa 
         DataField       =   "descripcioinplacsa"
         DataSource      =   "dataadhesius"
         Height          =   300
         Left            =   1140
         TabIndex        =   4
         Top             =   1425
         Width           =   4095
      End
      Begin VB.TextBox direccio 
         DataField       =   "descripcioprov"
         DataSource      =   "dataadhesius"
         Height          =   300
         Left            =   1140
         TabIndex        =   3
         Top             =   1025
         Width           =   4710
      End
      Begin VB.TextBox codi 
         BackColor       =   &H00FFFFFF&
         DataField       =   "codiintern"
         DataSource      =   "dataadhesius"
         Height          =   300
         Left            =   1140
         TabIndex        =   1
         Top             =   210
         Width           =   2505
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicials:"
         Height          =   480
         Left            =   5280
         TabIndex        =   24
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Gruix (cm)"
         Height          =   300
         Left            =   180
         TabIndex        =   23
         Top             =   2685
         Width           =   825
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Subfamilia"
         Height          =   300
         Left            =   180
         TabIndex        =   22
         Top             =   2265
         Width           =   825
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Familia"
         Height          =   360
         Left            =   180
         TabIndex        =   21
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Desc. Inplacsa"
         Height          =   480
         Left            =   180
         TabIndex        =   20
         Top             =   1290
         Width           =   885
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Desc. Prov:"
         Height          =   300
         Left            =   180
         TabIndex        =   19
         Top             =   1035
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveïdor:"
         Height          =   300
         Left            =   180
         TabIndex        =   18
         Top             =   645
         Width           =   4860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi Intern:"
         Height          =   300
         Left            =   180
         TabIndex        =   17
         Top             =   270
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   6420
      Begin VB.CommandButton consultar 
         Height          =   450
         Left            =   5340
         Picture         =   "Adhesiusmuntadora.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda de Registres"
         Top             =   150
         Width           =   450
      End
      Begin VB.CommandButton sortir 
         Height          =   450
         Left            =   5820
         Picture         =   "Adhesiusmuntadora.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Sortir"
         Top             =   150
         Width           =   450
      End
      Begin VB.Data dataadhesius 
         Caption         =   "Adhesius"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   390
         Left            =   1950
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "adhesiusmuntadora"
         Top             =   210
         Width           =   3360
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "Adhesiusmuntadora.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   1425
         Picture         =   "Adhesiusmuntadora.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton gravar 
         Height          =   360
         Left            =   960
         Picture         =   "Adhesiusmuntadora.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Guardar Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   525
         Picture         =   "Adhesiusmuntadora.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Modificació Registres"
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
         TabIndex        =   13
         Top             =   300
         Width           =   105
      End
   End
End
Attribute VB_Name = "formadhesiusmuntadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
alta_registre
'framemsg.Visible = False
End Sub
Sub alta_registre()
 If dataadhesius.Recordset.EditMode = 0 Then
      fdades.Enabled = True
      dataadhesius.Recordset.AddNew
      
        codi.SetFocus
     Else: MsgBox "No pots afegir si estàs editant...", vbCritical, "Atenció"
 End If
End Sub

Private Sub borrar_Click(Index As Integer)
  
End Sub

Private Sub Command1_Click()
  If dataadhesius.Recordset.EditMode = 2 Then MsgBox "Estas afegint un proveïdor nou... Primer guarda els canvis i despres edita ": Exit Sub
  
  framemsg.Visible = Not framemsg.Visible
End Sub

Private Sub consultar_Click()
   Dim b As String
   'framemsg.Visible = False
   b = InputBox("Entra la Descripcio Inplacsa a buscar ", "Busqueda")
   If b <> "" Then
     dataadhesius.RecordSource = "select * from adhesiusmuntadora where descripcioinplacsa like '*" + atrim(b) + "*'"
     dataadhesius.Refresh
     b = ""
    
   End If
   If Not dataadhesius.Recordset.EOF Then dataadhesius.Recordset.MoveLast: dataadhesius.Recordset.MoveFirst
End Sub

Private Sub eliminar_Click()
  eliminarproveidor
  framemsg.Visible = False
End Sub
Sub eliminarproveidor()
 On Error GoTo err
  If UCase(InputBox("Segur que vols Eliminar aquest adhesiu?" + Chr(10) + Chr(13) + " escriu [ELIMINAR] per confirmar-ho.", "Eliminar proveidor")) = "ELIMINAR" Then
    dataadhesius.Recordset.Delete
    dataadhesius.Recordset.MoveNext
    If dataadhesius.Recordset.EOF Then dataadhesius.Recordset.MovePrevious
  End If
 Exit Sub
err:
  MsgBox "No s'ha pogut eliminar possiblement perque tingui registres relacionats. O bé no hi ha res per eliminar."
End Sub

Private Sub Form_Activate()
   Dim camp As Control
   Dim nomdelcamp As String
   On Error Resume Next
   For Each camp In Me
      nomdelcamp = mirarnomdelcamp(camp)
      If nomdelcamp <> "" Then
        If dataadhesius.Recordset.Fields(nomdelcamp).Type = 10 Then
            camp.MaxLength = dataadhesius.Recordset.Fields(nomdelcamp).Size
        End If
      End If
   Next
End Sub
Function mirarnomdelcamp(camp As Control) As String
   On Error Resume Next
   mirarnomdelcamp = camp.DataField
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then cancelar_registre
  If KeyCode = 112 Then gravarregistre
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
 Sub cancelar_registre()
   If dataadhesius.Recordset.EditMode > 0 Then
       dataadhesius.Recordset.CancelUpdate
       fdades.Enabled = False
   End If
 End Sub
Private Sub Form_Load()
  dataadhesius.DatabaseName = cami
  dataadhesius.RecordSource = "select * from adhesiusmuntadora  order by codiintern"
  dataadhesius.Refresh
  
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub gravar_Click()
  gravarregistre
  'framemsg.Visible = False
End Sub
Sub gravarregistre()
If dataadhesius.EditMode > 0 Then
    dataadhesius.Recordset.Update
  End If
  fdades.Enabled = False
End Sub
Private Sub modificar_Click()
  If dataadhesius.Recordset.EditMode = 0 Then
     dataadhesius.Recordset.Edit
     fdades.Enabled = True
     DoEvents
     codi.SetFocus
  End If
  
End Sub

Private Sub proveidors_Reposition()
  If Not dataadhesius.Recordset.EOF Then
   dataadhesius.Caption = "Adhesius:  " + atrim(cadbl(dataadhesius.Recordset.AbsolutePosition) + 1) + " de " + atrim(dataadhesius.Recordset.RecordCount)
   'actualitzamsgs
     Else: dataadhesius.Caption = "Adhesius"
  End If
End Sub


Private Sub selmsg_Click(Index As Integer)
  
End Sub

Private Sub selfamilia_DropDown()
 triar_familia "famadhesiusmunt"
End Sub

Private Sub selproveidor_DropDown()
  triar_proveidor
End Sub
Sub triar_proveidor()
  Load formseleccio
  formseleccio.Data1.DatabaseName = dataadhesius.DatabaseName
  formseleccio.Data1.RecordSource = "select * from proveidors"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   dataadhesius.Recordset!codiproveidor = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   selproveidor.Text = atrim(formseleccio.Data1.Recordset!nom)
   
  End If
  Unload formseleccio
  
End Sub
Sub triar_familia(nomtaula As String)
  Load formseleccio
  formseleccio.Data1.DatabaseName = dataadhesius.DatabaseName
  formseleccio.Data1.RecordSource = "select * from " + nomtaula
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   If nomtaula = "famadhesiusmunt" Then
     dataadhesius.Recordset!idfam = atrim(cadbl(formseleccio.Data1.Recordset!codi))
     selfamilia.Text = atrim(formseleccio.Data1.Recordset!descripcio)
       Else
        dataadhesius.Recordset!idsubfam = atrim(cadbl(formseleccio.Data1.Recordset!codi))
        selsubfamilia.Text = atrim(formseleccio.Data1.Recordset!descripcio)
   End If
  End If
  Unload formseleccio
  
End Sub


Private Sub selsubfamilia_DropDown()
triar_familia "subfamadhesiusmunt"
End Sub

Private Sub sortir_Click()
  Unload Me
  Menu.SetFocus
End Sub

