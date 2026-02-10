VERSION 5.00
Begin VB.Form frmdibuix 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   23145
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   23145
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5550
      TabIndex        =   77
      Top             =   22500
      Width           =   5550
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   4289
         TabIndex        =   81
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actuali&zar"
         Height          =   300
         Left            =   2973
         TabIndex        =   80
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   1657
         TabIndex        =   79
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   300
         Left            =   341
         TabIndex        =   78
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Data datPrimaryRS 
      Align           =   2  'Align Bottom
      Caption         =   " "
      Connect         =   "Access"
      DatabaseName    =   "M:\progcomandes\dades\Treballs.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from [dibuix] Order by [numtreball]"
      Top             =   22800
      Width           =   5550
   End
   Begin VB.TextBox txtFields 
      DataField       =   "dataaprovacio"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   69
      Left            =   2040
      TabIndex        =   76
      Top             =   22140
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "aprovat"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   68
      Left            =   2040
      MaxLength       =   15
      TabIndex        =   74
      Top             =   21820
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "elaborat"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   67
      Left            =   2040
      MaxLength       =   15
      TabIndex        =   72
      Top             =   21500
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "magnificat"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   66
      Left            =   2040
      TabIndex        =   70
      Top             =   21180
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "codibarres"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   65
      Left            =   2040
      MaxLength       =   13
      TabIndex        =   68
      Top             =   20860
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "mida"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   64
      Left            =   2040
      TabIndex        =   66
      Top             =   20540
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "situacioextrem"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   63
      Left            =   2040
      TabIndex        =   64
      Top             =   20220
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "situacio"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   62
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   62
      Top             =   19900
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "pautacelula"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   61
      Left            =   2040
      TabIndex        =   60
      Top             =   19580
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "muntarbocaambboca"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   60
      Left            =   2040
      TabIndex        =   58
      Top             =   19260
      Width           =   3375
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "muntarcaraidors"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   59
      Left            =   2040
      TabIndex        =   56
      Top             =   18940
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "motius"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   58
      Left            =   2040
      TabIndex        =   54
      Top             =   18620
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "corro"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   57
      Left            =   2040
      TabIndex        =   52
      Top             =   18300
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "desenvol"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   56
      Left            =   2040
      TabIndex        =   50
      Top             =   17980
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ll8"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   55
      Left            =   2040
      TabIndex        =   48
      Top             =   17660
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ll7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   54
      Left            =   2040
      TabIndex        =   46
      Top             =   17340
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ll6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   53
      Left            =   2040
      TabIndex        =   44
      Top             =   17020
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ll5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   52
      Left            =   2040
      TabIndex        =   42
      Top             =   16700
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ll4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   51
      Left            =   2040
      TabIndex        =   40
      Top             =   16380
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ll3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   50
      Left            =   2040
      TabIndex        =   38
      Top             =   16060
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ll2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   49
      Left            =   2040
      TabIndex        =   36
      Top             =   15740
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ll1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   48
      Left            =   2040
      TabIndex        =   34
      Top             =   15420
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tt8"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   47
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   32
      Top             =   15100
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tt7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   46
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   30
      Top             =   14780
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tt6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   45
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   28
      Top             =   14460
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tt5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   44
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   26
      Top             =   14140
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tt4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   43
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   24
      Top             =   13820
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tt3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   42
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   22
      Top             =   13500
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tt2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   41
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   20
      Top             =   13180
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tt1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   40
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   18
      Top             =   12860
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "l8"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   39
      Left            =   2040
      TabIndex        =   16
      Top             =   12540
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "l7"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   38
      Left            =   2040
      TabIndex        =   14
      Top             =   12220
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "l6"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   37
      Left            =   2040
      TabIndex        =   12
      Top             =   11900
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "l5"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   36
      Left            =   2040
      TabIndex        =   10
      Top             =   11580
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "l4"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   35
      Left            =   2040
      TabIndex        =   8
      Top             =   11260
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "l3"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   34
      Left            =   2040
      TabIndex        =   6
      Top             =   10940
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "l2"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   33
      Left            =   2040
      TabIndex        =   4
      Top             =   10620
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "l1"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   32
      Left            =   2040
      TabIndex        =   2
      Top             =   10300
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "producte"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   17
      Left            =   6300
      MaxLength       =   10
      TabIndex        =   0
      Top             =   5500
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "dataaprovacio:"
      Height          =   255
      Index           =   69
      Left            =   120
      TabIndex        =   75
      Top             =   22140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "aprovat:"
      Height          =   255
      Index           =   68
      Left            =   120
      TabIndex        =   73
      Top             =   21820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "elaborat:"
      Height          =   255
      Index           =   67
      Left            =   120
      TabIndex        =   71
      Top             =   21500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "magnificat:"
      Height          =   255
      Index           =   66
      Left            =   120
      TabIndex        =   69
      Top             =   21180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "codibarres:"
      Height          =   255
      Index           =   65
      Left            =   120
      TabIndex        =   67
      Top             =   20860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "mida:"
      Height          =   255
      Index           =   64
      Left            =   120
      TabIndex        =   65
      Top             =   20540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "situacioextrem:"
      Height          =   255
      Index           =   63
      Left            =   120
      TabIndex        =   63
      Top             =   20220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "situacio:"
      Height          =   255
      Index           =   62
      Left            =   120
      TabIndex        =   61
      Top             =   19900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "pautacelula:"
      Height          =   255
      Index           =   61
      Left            =   120
      TabIndex        =   59
      Top             =   19580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "muntarbocaambboca:"
      Height          =   255
      Index           =   60
      Left            =   120
      TabIndex        =   57
      Top             =   19260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "muntarcaraidors:"
      Height          =   255
      Index           =   59
      Left            =   120
      TabIndex        =   55
      Top             =   18940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "motius:"
      Height          =   255
      Index           =   58
      Left            =   120
      TabIndex        =   53
      Top             =   18620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "corro:"
      Height          =   255
      Index           =   57
      Left            =   120
      TabIndex        =   51
      Top             =   18300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "desenvol:"
      Height          =   255
      Index           =   56
      Left            =   120
      TabIndex        =   49
      Top             =   17980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ll8:"
      Height          =   255
      Index           =   55
      Left            =   120
      TabIndex        =   47
      Top             =   17660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ll7:"
      Height          =   255
      Index           =   54
      Left            =   120
      TabIndex        =   45
      Top             =   17340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ll6:"
      Height          =   255
      Index           =   53
      Left            =   120
      TabIndex        =   43
      Top             =   17020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ll5:"
      Height          =   255
      Index           =   52
      Left            =   120
      TabIndex        =   41
      Top             =   16700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ll4:"
      Height          =   255
      Index           =   51
      Left            =   120
      TabIndex        =   39
      Top             =   16380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ll3:"
      Height          =   255
      Index           =   50
      Left            =   120
      TabIndex        =   37
      Top             =   16060
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ll2:"
      Height          =   255
      Index           =   49
      Left            =   120
      TabIndex        =   35
      Top             =   15740
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ll1:"
      Height          =   255
      Index           =   48
      Left            =   120
      TabIndex        =   33
      Top             =   15420
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "tt8:"
      Height          =   255
      Index           =   47
      Left            =   120
      TabIndex        =   31
      Top             =   15100
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "tt7:"
      Height          =   255
      Index           =   46
      Left            =   120
      TabIndex        =   29
      Top             =   14780
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "tt6:"
      Height          =   255
      Index           =   45
      Left            =   120
      TabIndex        =   27
      Top             =   14460
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "tt5:"
      Height          =   255
      Index           =   44
      Left            =   120
      TabIndex        =   25
      Top             =   14140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "tt4:"
      Height          =   255
      Index           =   43
      Left            =   120
      TabIndex        =   23
      Top             =   13820
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "tt3:"
      Height          =   255
      Index           =   42
      Left            =   120
      TabIndex        =   21
      Top             =   13500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "tt2:"
      Height          =   255
      Index           =   41
      Left            =   120
      TabIndex        =   19
      Top             =   13180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "tt1:"
      Height          =   255
      Index           =   40
      Left            =   120
      TabIndex        =   17
      Top             =   12860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "l8:"
      Height          =   255
      Index           =   39
      Left            =   120
      TabIndex        =   15
      Top             =   12540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "l7:"
      Height          =   255
      Index           =   38
      Left            =   120
      TabIndex        =   13
      Top             =   12220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "l6:"
      Height          =   255
      Index           =   37
      Left            =   120
      TabIndex        =   11
      Top             =   11900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "l5:"
      Height          =   255
      Index           =   36
      Left            =   120
      TabIndex        =   9
      Top             =   11580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "l4:"
      Height          =   255
      Index           =   35
      Left            =   120
      TabIndex        =   7
      Top             =   11260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "l3:"
      Height          =   255
      Index           =   34
      Left            =   120
      TabIndex        =   5
      Top             =   10940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "l2:"
      Height          =   255
      Index           =   33
      Left            =   120
      TabIndex        =   3
      Top             =   10620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "l1:"
      Height          =   255
      Index           =   32
      Left            =   120
      TabIndex        =   1
      Top             =   10300
      Width           =   1815
   End
End
Attribute VB_Name = "frmdibuix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAgregar_Click()
  datPrimaryRS.Recordset.AddNew
End Sub

Private Sub cmdEliminar_Click()
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
End Sub

Private Sub cmdActualizar_Click()
  datPrimaryRS.UpdateRecord
  datPrimaryRS.Recordset.Bookmark = datPrimaryRS.Recordset.LastModified
End Sub

Private Sub cmdCerrar_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub datPrimaryRS_Error(DataErr As Integer, Response As Integer)
  'Aquí es dónde puede situar el código de tratamiento de error
  'Si desea ignorar los errores, quite el comentario de la siguiente línea
  'Si desea capturarlos, agregue código aquí para controlarlos
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0  'Despreciar el error
End Sub

Private Sub datPrimaryRS_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'Mostrar la posición actual de registro para dynasets y snapshots
  datPrimaryRS.Caption = "Record: " & (datPrimaryRS.Recordset.AbsolutePosition + 1)
End Sub

Private Sub datPrimaryRS_Validate(Action As Integer, Save As Integer)
  'Aquí se sitúa el código de validación
  'Este evento se invoca cuando ocurre la siguiente acción
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
      Screen.MousePointer = vbDefault
  End Select
  Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub txtFields_Change(Index As Integer)

End Sub
