VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formstopped 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informació de comandes Stopped"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12180
   Icon            =   "Formcomandesstopped.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bbuscaritem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11760
      Picture         =   "Formcomandesstopped.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   360
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
      Left            =   8250
      Picture         =   "Formcomandesstopped.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      Width           =   735
   End
   Begin VB.TextBox nomdelclient 
      Height          =   360
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   75
      Width           =   6690
   End
   Begin VB.CommandButton bhistoric 
      Enabled         =   0   'False
      Height          =   360
      Left            =   11790
      Picture         =   "Formcomandesstopped.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Veure l'historic"
      Top             =   5715
      Width           =   360
   End
   Begin VB.CommandButton bexportar 
      Height          =   360
      Left            =   11760
      Picture         =   "Formcomandesstopped.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Exportar informació"
      Top             =   1200
      Width           =   360
   End
   Begin VB.CommandButton alta 
      Height          =   360
      Left            =   11760
      Picture         =   "Formcomandesstopped.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Alta  Registres"
      Top             =   465
      Width           =   360
   End
   Begin VB.CommandButton eliminar 
      Height          =   360
      Left            =   11760
      Picture         =   "Formcomandesstopped.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   840
      Width           =   360
   End
   Begin VB.Data datastopped 
      Caption         =   "datastopped"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   525
      Left            =   4020
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "refclient_stopped"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2820
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "Formcomandesstopped.frx":26C6
      Height          =   5685
      Left            =   15
      OleObjectBlob   =   "Formcomandesstopped.frx":26DC
      TabIndex        =   7
      Top             =   435
      Width           =   11730
   End
   Begin VB.Label Label1 
      Caption         =   "Client:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   630
      TabIndex        =   5
      Top             =   90
      Width           =   1005
   End
End
Attribute VB_Name = "formstopped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
   If cadbl(nomdelclient.Tag) = 0 Then MsgBox "Primer has d'escullir un client.", vbCritical, "Error": Exit Sub
   If datastopped.Recordset.EditMode > 0 Then Exit Sub
   datastopped.Recordset.AddNew
   datastopped.Recordset!Data = Now
   datastopped.Recordset!numclient = cadbl(nomdelclient.Tag)
   datastopped.Recordset.Update
   datastopped.Recordset.Bookmark = datastopped.Recordset.LastModified
   reixa.Refresh
   reixa.col = 1
   reixa.SetFocus
   
End Sub

Private Sub bbuscaritem_Click()
   Dim vitem As String
   vitem = InputBox("Entra la referencia que vols buscar d'aquest client.", "Buscar referencia")
   datastopped.Recordset.FindFirst "refclient='" + treure_apostruf(vitem) + "'"
End Sub

Private Sub bexportar_Click()
  If nomdelclient = "" Then MsgBox "Escull primer un client", vbCritical, "Atenció": Exit Sub
    exportarinformaciodesactivades
End Sub

Private Sub bhistoric_Click()
  If bhistoric.BackColor = QBColor(12) Then
    datadesactivades.RecordSource = "select * from informaciodesactivades where actiu order by data desc"
    datadesactivades.Refresh
    bhistoric.BackColor = alta.BackColor
    alta.Enabled = True: eliminar.Enabled = True
    Exit Sub
  End If
  If bhistoric.BackColor <> QBColor(12) Then
    datadesactivades.RecordSource = "select * from informaciodesactivades where not actiu order by data desc"
    datadesactivades.Refresh
    bhistoric.BackColor = QBColor(12)
    alta.Enabled = False: eliminar.Enabled = False
    Exit Sub
  End If
End Sub

Private Sub Command3_Click()
  triarclient
  actualitzar_reixa
End Sub
Sub actualitzar_reixa()
   datastopped.RecordSource = "select * from refclient_stopped where numclient=" + atrim(cadbl(nomdelclient.Tag))
   datastopped.Refresh
End Sub
Sub triarclient()
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = datastopped.DatabaseName
  formseleccio.Data1.RecordSource = "select * from clients"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   nomdelclient.Tag = atrim(cadbl(formseleccio.Data1.Recordset!codi))
   nomdelclient = atrim(formseleccio.Data1.Recordset!nom)
  End If
  Unload formseleccio
  
End Sub

Private Sub eliminar_Click()
   If datastopped.Recordset.EOF Then Exit Sub
   If UCase(InputBox("Si vols eliminar-la entra la contrasenya d'eliminació", "Eliminar")) = "INPLACSA" Then GoTo eliminar
   GoTo fi
eliminar:
     datastopped.Recordset.Delete
fi:
   datastopped.Refresh
End Sub

Private Sub Form_Load()
   datastopped.DatabaseName = cami
   datastopped.RecordSource = "select * from refclient_stopped"
   datastopped.Refresh
End Sub
Sub exportarinformaciodesactivades()
   Dim vfitxer As String
   Dim rst As Recordset
   vfitxer = "c:\temp\informaciostopped.csv"
   If existeix(vfitxer) Then Kill vfitxer
   Open vfitxer For Output As #3
   Set rst = datastopped.Recordset
   If Not (rst.EOF And rst.BOF) Then rst.MoveFirst
   vlinia = "DATA;REFCLIENT;NOM_CLIENT;DESCRIPCIO"
   Print #3, vlinia
   While Not rst.EOF
      vlinia = Format(rst!Data, "dd/mm/yy") + ";" + treuresimbols(atrim(rst!refclient)) + ";" + treuresimbols(nomdelclient) + ";" + treuresimbols(atrim(rst!observacio))
      Print #3, vlinia
      rst.MoveNext
   Wend
   Close #3
   Set rst = Nothing
   If existeix(vfitxer) Then obrir_document vfitxer
End Sub
