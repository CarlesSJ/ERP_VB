VERSION 5.00
Begin VB.Form fFotogravadors 
   Caption         =   "Manteniment de Fotogravadors"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "Fotogravadors.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameFotogravadors 
      Caption         =   "Fotogravadors"
      Enabled         =   0   'False
      Height          =   5775
      Left            =   30
      TabIndex        =   8
      Top             =   600
      Width           =   4485
      Begin VB.ComboBox cidioma 
         DataField       =   "idioma"
         DataSource      =   "Fotogravadors"
         Height          =   315
         ItemData        =   "Fotogravadors.frx":058A
         Left            =   3720
         List            =   "Fotogravadors.frx":059A
         TabIndex        =   40
         Text            =   "Combo1"
         Top             =   510
         Width           =   585
      End
      Begin VB.TextBox txtFields 
         DataField       =   "email"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   12
         Left            =   975
         MaxLength       =   255
         TabIndex        =   38
         Top             =   2430
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "observacions"
         DataSource      =   "Fotogravadors"
         Height          =   705
         Index           =   14
         Left            =   75
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   4770
         Width           =   4230
      End
      Begin VB.CheckBox chkFields 
         Caption         =   "Actiu"
         DataField       =   "actiu"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   12
         Left            =   3450
         TabIndex        =   35
         Top             =   150
         Width           =   915
      End
      Begin VB.TextBox txtFields 
         DataField       =   "obsfax"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   11
         Left            =   990
         MaxLength       =   255
         TabIndex        =   34
         Top             =   4275
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "obstel2"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   10
         Left            =   990
         MaxLength       =   100
         TabIndex        =   32
         Top             =   3630
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "obstel1"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   9
         Left            =   990
         MaxLength       =   100
         TabIndex        =   30
         Top             =   3030
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "fax"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   8
         Left            =   990
         MaxLength       =   20
         TabIndex        =   28
         Top             =   3960
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "telefon2"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   7
         Left            =   990
         MaxLength       =   20
         TabIndex        =   26
         Top             =   3345
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "telefon1"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   6
         Left            =   990
         MaxLength       =   20
         TabIndex        =   24
         Top             =   2730
         Width           =   3345
      End
      Begin VB.TextBox txtFields 
         DataField       =   "provincia"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   5
         Left            =   990
         MaxLength       =   255
         TabIndex        =   22
         Top             =   2145
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "localitat"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   4
         Left            =   990
         MaxLength       =   255
         TabIndex        =   20
         Top             =   1815
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "codipostal"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   3
         Left            =   990
         TabIndex        =   18
         Top             =   1500
         Width           =   960
      End
      Begin VB.TextBox txtFields 
         DataField       =   "direccio"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   2
         Left            =   990
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1185
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "nomfotogravador"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   1
         Left            =   990
         MaxLength       =   20
         TabIndex        =   14
         Top             =   855
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "codi"
         DataSource      =   "Fotogravadors"
         Height          =   285
         Index           =   0
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   945
      End
      Begin VB.Label Idioma 
         Caption         =   "Idioma"
         Height          =   240
         Left            =   3105
         TabIndex        =   41
         Top             =   540
         Width           =   510
      End
      Begin VB.Label lblLabels 
         Caption         =   "Email:"
         Height          =   255
         Index           =   12
         Left            =   105
         TabIndex        =   39
         Top             =   2430
         Width           =   1050
      End
      Begin VB.Label lblLabels 
         Caption         =   "Observacions:"
         Height          =   255
         Index           =   14
         Left            =   90
         TabIndex        =   36
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "ObsFax:"
         Height          =   255
         Index           =   11
         Left            =   75
         TabIndex        =   33
         Top             =   4275
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "ObsTel2:"
         Height          =   255
         Index           =   10
         Left            =   75
         TabIndex        =   31
         Top             =   3630
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "ObsTel1:"
         Height          =   255
         Index           =   9
         Left            =   60
         TabIndex        =   29
         Top             =   3030
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fax"
         Height          =   255
         Index           =   8
         Left            =   75
         TabIndex        =   27
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Telf:2"
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   25
         Top             =   3345
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Telf1:"
         Height          =   255
         Index           =   6
         Left            =   75
         TabIndex        =   23
         Top             =   2730
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Provincia:"
         Height          =   255
         Index           =   5
         Left            =   75
         TabIndex        =   21
         Top             =   2145
         Width           =   1100
      End
      Begin VB.Label lblLabels 
         Caption         =   "Localitat:"
         Height          =   255
         Index           =   4
         Left            =   75
         TabIndex        =   19
         Top             =   1815
         Width           =   1100
      End
      Begin VB.Label lblLabels 
         Caption         =   "Codi Postal:"
         Height          =   255
         Index           =   3
         Left            =   75
         TabIndex        =   17
         Top             =   1500
         Width           =   1100
      End
      Begin VB.Label lblLabels 
         Caption         =   "Direcció:"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   1185
         Width           =   1100
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nom"
         Height          =   255
         Index           =   1
         Left            =   75
         TabIndex        =   13
         Top             =   855
         Width           =   1100
      End
      Begin VB.Label lblLabels 
         Caption         =   "Codi:"
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   11
         Top             =   540
         Width           =   1100
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   4470
      Begin VB.Data Fotogravadors 
         Caption         =   "Fotogravadors"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   360
         Left            =   2235
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Fotogravadors"
         Top             =   180
         Width           =   1260
      End
      Begin VB.CommandButton consultar 
         Height          =   360
         Left            =   3495
         Picture         =   "Fotogravadors.frx":05AE
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda de Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "Fotogravadors.frx":0B38
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Alta  Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   975
         Picture         =   "Fotogravadors.frx":10C2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Eliminacio Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   540
         Picture         =   "Fotogravadors.frx":164C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Consulta Registres"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton sortir 
         Height          =   390
         Left            =   4005
         Picture         =   "Fotogravadors.frx":1BD6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Alta  Registres"
         Top             =   135
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Height          =   390
         Left            =   1410
         Picture         =   "Fotogravadors.frx":2160
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   390
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   300
      Index           =   2
      Left            =   90
      TabIndex        =   10
      Top             =   1605
      Width           =   405
   End
   Begin VB.Label status 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   135
      TabIndex        =   7
      Top             =   5985
      Width           =   4470
   End
   Begin VB.Label autonum 
      Height          =   135
      Left            =   0
      TabIndex        =   6
      Top             =   1335
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "fFotogravadors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub colsbloc_Change()

End Sub

Private Sub alta_Click()
Dim gran As Long
frameFotogravadors.Enabled = True
Fotogravadors.RecordSource = "select * from Fotogravadors order by codi"
Fotogravadors.Recordset.MoveLast
If Not Fotogravadors.Recordset.EOF Then gran = Fotogravadors.Recordset!codi
gran = gran + 1
Fotogravadors.Recordset.AddNew
frameFotogravadors.Enabled = True
Fotogravadors.Recordset!codi = gran
txtFields(0) = gran
txtFields(1).SetFocus
End Sub

Private Sub Command2_Click()
  If Not existeixrang Then
      espesors.Recordset.AddNew
       espesors.Recordset!codi = Fotogravadors.Recordset!codi
       espesors.Recordset!micres = cadbl(micres)
       espesors.Recordset!grmsm2 = cadbl(grmm2)
      espesors.Recordset.Update
     Else: MsgBox "Aquest espesor ja existeix."
  End If
End Sub
Function existeixrang() As Boolean
  existeixrang = False
  If cadbl(micres) > 0 And espesors.Recordset.RecordCount > 0 Then
   espesors.Recordset.FindFirst "micres=" + atrim(cadbl(micres))
   If Not espesors.Recordset.NoMatch Then
      existeixrang = True
   End If
  End If
  If (cadbl(micres) = 0 And cadbl(grmm2) > 0) And espesors.Recordset.RecordCount > 0 Then
   espesors.Recordset.FindFirst "grmsm2=" + atrim(cadbl(grmm2))
   If Not espesors.Recordset.NoMatch Then
      existeixrang = True
   End If
  End If
  
End Function

Private Sub Command1_Click()
   If Fotogravadors.Recordset.EditMode > 0 Then
      Fotogravadors.Recordset.Update
      Fotogravadors.Recordset.Bookmark = Fotogravadors.Recordset.LastModified
   End If
   frameFotogravadors.Enabled = False
End Sub

Private Sub consultar_Click()
   Dim b As String
   b = InputBox("Entra la descripcio a buscar", "Busqueda")
   If b <> "" Then
      Fotogravadors.RecordSource = "select * from Fotogravadors where nomfotogravador like '*" + b + "*'"
      Fotogravadors.Refresh
   End If
End Sub

Private Sub eliminar_Click()
  If MsgBox("Segur que vols borrar aquest Fotogravador?, PENSA QUE TOTES LES RELACIONS DE CLIXES AMB ELL QUEDERAN ORFES", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
     If InputBox("Escriu la paraula [ELIMINAR] per fer efectiu l'eliminació", "Control de seguretat") = "ELIMINAR" Then
         Fotogravadors.Recordset.Delete
         Fotogravadors.Refresh
     End If
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then gravar_registre
End Sub
Sub gravar_registre()
  If Fotogravadors.Recordset.EditMode > 0 Then
     Fotogravadors.Recordset.Update
     frameFotogravadors.Enabled = False
     Fotogravadors.Recordset.Bookmark = Fotogravadors.Recordset.LastModified
  End If
End Sub
Sub cancelar_registre()
 
 If Fotogravadors.EditMode > 0 Then Fotogravadors.Recordset.CancelUpdate
 frameFotogravadors.Enabled = False
End Sub

Private Sub Form_Load()
  Fotogravadors.DatabaseName = camiclixes
End Sub

Private Sub Fotogravadors_Reposition()
frameFotogravadors.Enabled = False
  
End Sub

Private Sub modificar_Click()
   If Not Fotogravadors.Recordset.EOF Then
     Fotogravadors.Recordset.Edit
     frameFotogravadors.Enabled = True
     txtFields(1).SetFocus
   End If
End Sub

Private Sub sortir_Click()
 Unload fFotogravadors
End Sub

Private Sub Timer1_Timer()

End Sub
Sub triar_proveidor()
  Load formseleccio
  formseleccio.Data1.DatabaseName = Fotogravadors.DatabaseName
  formseleccio.Data1.RecordSource = "select * from proveidors"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   txtFields(1).Text = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   Fotogravadors.Recordset!proveidor = txtFields(1).Text
   nomproveidor.caption = atrim(formseleccio.Data1.Recordset!nom)
  End If
  Unload formseleccio
  
End Sub



