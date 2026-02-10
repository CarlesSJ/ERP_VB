VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form missatgespeucomandescompra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Misatges al peu de la comanda de compra"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8460
   Icon            =   "missatgespeucomandescompra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   8325
      Begin VB.CommandButton modificar 
         Height          =   450
         Left            =   540
         Picture         =   "missatgespeucomandescompra.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Modificació Registres"
         Top             =   135
         Width           =   450
      End
      Begin VB.Data msgpeu 
         Caption         =   "Missatges Peu"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "msgpeucomanda"
         Top             =   150
         Width           =   4500
      End
      Begin VB.CommandButton sortir 
         Height          =   450
         Left            =   7710
         Picture         =   "missatgespeucomandescompra.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Sortir"
         Top             =   135
         Width           =   450
      End
      Begin VB.CommandButton alta 
         Height          =   450
         Left            =   90
         Picture         =   "missatgespeucomandescompra.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   135
         Width           =   450
      End
      Begin VB.CommandButton eliminar 
         Height          =   450
         Left            =   975
         Picture         =   "missatgespeucomandescompra.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   135
         Width           =   450
      End
   End
   Begin VB.Frame Frame 
      Enabled         =   0   'False
      Height          =   2520
      Left            =   30
      TabIndex        =   4
      Top             =   720
      Width           =   8340
      Begin VB.Data descripciomsg 
         Caption         =   "desc"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1470
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "descripcionsmsgpeu"
         Top             =   1860
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.TextBox nom 
         DataField       =   "descripcio"
         DataSource      =   "msgpeu"
         Height          =   375
         Left            =   645
         TabIndex        =   6
         Top             =   210
         Width           =   3780
      End
      Begin VB.CommandButton gravar 
         Height          =   450
         Left            =   4455
         Picture         =   "missatgespeucomandescompra.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Guardar Registres"
         Top             =   135
         Width           =   450
      End
      Begin MSDBGrid.DBGrid reixa 
         Bindings        =   "missatgespeucomandescompra.frx":213C
         Height          =   1710
         Left            =   75
         OleObjectBlob   =   "missatgespeucomandescompra.frx":2154
         TabIndex        =   7
         Top             =   660
         Width           =   8025
      End
      Begin VB.Label Label1 
         Caption         =   "Nom:"
         Height          =   345
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "missatgespeucomandescompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
  msgpeu.Recordset.AddNew
  Frame.Enabled = True
  descripciomsg.RecordSource = "select * from descripcionsmsgpeu where idmsg=0"
  descripciomsg.Refresh
  
  nom.SetFocus
End Sub

Private Sub eliminar_Click()
If msgpeu.Recordset.EOF Then Exit Sub
  If MsgBox("Estas segur que vols borrar aquest missatge?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
       If msgpeu.Recordset.EditMode > 0 Then msgpeu.Recordset.CancelUpdate
       descripciomsg.Database.Execute "delete * from descripcionsmsgpeu where idmsg=" + atrim(cadbl(msgpeu.Recordset!ID))
       msgpeu.Recordset.Delete
       msgpeu.Refresh
       Frame.Enabled = False
  End If
End Sub

Private Sub Form_Load()
  msgpeu.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
  descripciomsg.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
  msgpeu.Refresh
  msgpeu.Recordset.MoveLast
  msgpeu.Recordset.MoveFirst
End Sub

Private Sub gravar_Click()
  If msgpeu.Recordset.EditMode > 0 Then
     msgpeu.Recordset.Update
     msgpeu.Recordset.Bookmark = msgpeu.Recordset.LastModified
     msgpeu.Recordset.Edit
  End If
End Sub

Private Sub modificar_Click()
  If msgpeu.Recordset.EOF Then Exit Sub
  msgpeu.Recordset.Edit
  Frame.Enabled = True
  'descripciomsg.RecordSource = "select * from descripcionsmsgpeu where id=0"
  'descripciomsg.Refresh
  nom.SetFocus
End Sub

Private Sub msgpeu_Reposition()
   If msgpeu.Recordset.EOF Then msgpeu.Caption = "Missatges Peu": Exit Sub
    descripciomsg.RecordSource = "select * from descripcionsmsgpeu where idmsg=" + atrim(msgpeu.Recordset!ID) + " order by ordre"
    descripciomsg.Refresh
   msgpeu.Caption = "Missatge:  " + atrim(cadbl(msgpeu.Recordset.AbsolutePosition) + 1) + " de " + atrim(msgpeu.Recordset.RecordCount)
  
End Sub

Private Sub reixa_KeyDown(KeyCode As Integer, Shift As Integer)
  If Len(reixa.Text) > 79 And KeyCode > 31 Then
    MsgBox "Nomes es gravaran 80 caràcters per linia": KeyCode = 0: Exit Sub
    reixa.Text = Mid(reixa.Text, 1, 80)
  End If
End Sub

Private Sub reixa_OnAddNew()
  If msgpeu.Recordset.EditMode = 1 Then
   descripciomsg.Recordset!idmsg = msgpeu.Recordset!ID
     Else: MsgBox "Primer grava el nom"
  End If
End Sub

Private Sub sortir_Click()
  Unload missatgespeucomandescompra
End Sub
