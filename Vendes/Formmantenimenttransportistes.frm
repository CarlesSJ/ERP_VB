VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Formmantenimenttransportistes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment de transportistes"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12630
   Icon            =   "Formmantenimenttransportistes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data datatransportistes 
      Caption         =   "Transportistes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      Caption         =   "E-mails recullida per país"
      Height          =   2715
      Left            =   7530
      TabIndex        =   1
      Top             =   690
      Width           =   4845
      Begin VB.CommandButton Command5 
         Height          =   390
         Left            =   510
         Picture         =   "Formmantenimenttransportistes.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   225
         Width           =   405
      End
      Begin VB.CommandButton Command3 
         Height          =   390
         Left            =   90
         Picture         =   "Formmantenimenttransportistes.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   405
      End
      Begin VB.Data Dataemails 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   510
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Transportistes_emailsperpais"
         Top             =   2010
         Visible         =   0   'False
         Width           =   1530
      End
      Begin MSDBGrid.DBGrid reixa 
         Bindings        =   "Formmantenimenttransportistes.frx":109E
         Height          =   1995
         Left            =   60
         OleObjectBlob   =   "Formmantenimenttransportistes.frx":10B3
         TabIndex        =   13
         Top             =   645
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dades"
      Height          =   2700
      Left            =   165
      TabIndex        =   0
      Top             =   690
      Width           =   7335
      Begin VB.CheckBox checkvisible 
         Caption         =   "Visible"
         DataSource      =   "datatransportistes"
         Enabled         =   0   'False
         Height          =   240
         Left            =   6180
         TabIndex        =   16
         Top             =   195
         Width           =   1095
      End
      Begin VB.TextBox cCamp 
         DataField       =   "observaciopredeterminada"
         DataSource      =   "datatransportistes"
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1290
         TabIndex        =   6
         Top             =   1875
         Width           =   5715
      End
      Begin VB.TextBox cCamp 
         BackColor       =   &H0080FFFF&
         DataField       =   "email_copiaexpedicions"
         DataSource      =   "datatransportistes"
         Height          =   315
         Index           =   3
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox cCamp 
         BackColor       =   &H0080FFFF&
         DataField       =   "email_expedicions"
         DataSource      =   "datatransportistes"
         Height          =   315
         Index           =   2
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox cCamp 
         DataField       =   "descripcio"
         DataSource      =   "datatransportistes"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   3
         Top             =   690
         Width           =   3930
      End
      Begin VB.TextBox cCamp 
         DataField       =   "codi"
         DataSource      =   "datatransportistes"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1245
         TabIndex        =   2
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "<--- Doble clic per editar."
         Height          =   300
         Left            =   4470
         TabIndex        =   17
         Top             =   1275
         Width           =   1860
      End
      Begin VB.Label Label1 
         Caption         =   "Observacions:"
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   1935
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Email Copia Expedicions:"
         Height          =   495
         Index           =   3
         Left            =   90
         TabIndex        =   10
         Top             =   1410
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Email Expedicions:"
         Height          =   510
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   975
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Nom transport:"
         Height          =   300
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Codi:"
         Height          =   300
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   315
         Width           =   465
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   $"Formmantenimenttransportistes.frx":1AAC
      Height          =   270
      Left            =   90
      TabIndex        =   12
      Top             =   75
      Width           =   12450
   End
End
Attribute VB_Name = "Formmantenimenttransportistes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cCamp_DblClick(Index As Integer)
   Dim v As String
   If Index = 2 Or Index = 3 Then
     v = InputBox("Escriu l'email correcte.", "Canvi email", cCamp(Index))
     If StrPtr(v) = 0 Then Exit Sub
     If InStr(1, v, "@") > 0 Or v = "" Then
        datatransportistes.Recordset.Edit
        datatransportistes.Recordset.Fields(cCamp(Index).DataField) = v
        datatransportistes.Recordset.Update
        datatransportistes.Recordset.Move 0
     End If
   End If
End Sub

Private Sub Command3_Click()
   Dim vpais As String
   Dim vemails As String
   vpais = escullir_pais
   vemails = InputBox("Entra el e-mail o e-mails separats per ; dels contactes per aquest país." + vbNewLine + "L'E-MAIL S'ENVIARÀ A LES DIRECCIONS D'ENVIAMENT PRINCIPAL DEL TRANSPORTISTA I A MES A MES A AQUESTA DEPENENT DEL PAÍS D'ENVIAMENT.", "E-Mail")
   If InStr(1, vemails, "@") > 0 Then
    Dataemails.Recordset.AddNew
    Dataemails.Recordset!id_transport = datatransportistes.Recordset!codi
    Dataemails.Recordset!codipais = vpais
    Dataemails.Recordset!emailscontacte = vemails
    Dataemails.Recordset.Update
   End If
End Sub

Private Sub Command5_Click()
   If Dataemails.Recordset.EOF Then MsgBox "Primer escull un email per eliminar.", vbCritical, "Eliminar": Exit Sub
   If MsgBox("Segur que vols eliminar aquesta adreça?", vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
       Dataemails.Recordset.Delete
       datatransportistes.Move 0
   End If
End Sub

Private Sub datatransportistes_Reposition()
   Dataemails.RecordSource = "select * from Transportistes_emailsperpais where id_transport=" + atrim(datatransportistes.Recordset!codi)
   Dataemails.Refresh
   checkvisible.Value = IIf(datatransportistes.Recordset!visible, 1, 0)
End Sub

Private Sub Form_Load()
   Dataemails.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
   datatransportistes.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
   datatransportistes.RecordSource = "select * from transportistes order by visible desc,descripcio"
   datatransportistes.Refresh
   
End Sub

Function escullir_pais() As String
   Load formseleccio
   formseleccio.caption = "Escull un pais"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codipais,nompais from paisos "
  formseleccio.refrescar
'  formseleccio.DBGrid2.Columns(2).Width = 900
  formseleccio.width = 9000
  'formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  
   If seleccioret = 1 Then
       escullir_pais = formseleccio.DBGrid2.Columns("codipais")
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Function
