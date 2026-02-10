VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formdesactivades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informació de les comandes desactivades"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12180
   Icon            =   "FormInformacioDesactivades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bhistoric 
      Height          =   360
      Left            =   11790
      Picture         =   "FormInformacioDesactivades.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Veure l'historic"
      Top             =   5715
      Width           =   360
   End
   Begin VB.CommandButton bexportar 
      Height          =   360
      Left            =   11760
      Picture         =   "FormInformacioDesactivades.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Exportar informació"
      Top             =   1170
      Width           =   360
   End
   Begin VB.CommandButton alta 
      Height          =   360
      Left            =   11760
      Picture         =   "FormInformacioDesactivades.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Alta  Registres"
      Top             =   15
      Width           =   360
   End
   Begin VB.CommandButton eliminar 
      Height          =   360
      Left            =   11760
      Picture         =   "FormInformacioDesactivades.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   390
      Width           =   360
   End
   Begin VB.Data datadesactivades 
      Caption         =   "datadesactivades"
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
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   2820
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "FormInformacioDesactivades.frx":1BB2
      Height          =   6165
      Left            =   15
      OleObjectBlob   =   "FormInformacioDesactivades.frx":1BCD
      TabIndex        =   0
      Top             =   15
      Width           =   11730
   End
End
Attribute VB_Name = "formdesactivades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
   If datadesactivades.Recordset.EditMode > 0 Then Exit Sub
   datadesactivades.Recordset.AddNew
   datadesactivades.Recordset!Data = Now
   datadesactivades.Recordset.Update
   datadesactivades.Recordset.Bookmark = datadesactivades.Recordset.LastModified
   reixa.Refresh
   reixa.col = 2
   reixa.SetFocus
   
End Sub

Private Sub bexportar_Click()
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

Private Sub eliminar_Click()
   If datadesactivades.Recordset.EOF Then Exit Sub
   If datadesactivades.Recordset!tipus = "P" Then
       MsgBox "Aquesta linia només es pot eliminar desde compres o assignació ja que afecta a ells.", vbCritical, "Atenció"
       If UCase(InputBox("Si vols eliminar-la igualment entra la contrasenya d'eliminació", "Eliminar")) = "INPLACSA" Then GoTo eliminar
       GoTo fi
   End If
   If MsgBox("Segur que vols eliminar aquest registre?", vbCritical + vbYesNo + vbDefaultButton2, "Eliminar") = vbYes Then
eliminar:
     datadesactivades.Recordset.Edit
     datadesactivades.Recordset!actiu = False
     datadesactivades.Recordset.Update
   End If
fi:
   datadesactivades.Refresh
End Sub

Private Sub Form_Load()
   datadesactivades.DatabaseName = cami
   datadesactivades.RecordSource = "select * from informaciodesactivades where actiu order by data desc"
   datadesactivades.Refresh
End Sub
Sub exportarinformaciodesactivades()
   Dim vfitxer As String
   Dim rst As Recordset
   vfitxer = "c:\temp\informaciodesactivades.csv"
   If existeix(vfitxer) Then Kill vfitxer
   Open vfitxer For Output As #3
   Set rst = dbtmp.OpenRecordset("select * from informaciodesactivades where actiu order by data")
   vlinia = "DATA;COMANDA/REF;NOM_CLIENT;DESCRIPCIO"
   Print #3, vlinia
   While Not rst.EOF
      vlinia = Format(rst!Data, "dd/mm/yy") + ";" + treuresimbols(atrim(rst!comandaoreferencia)) + ";" + treuresimbols(atrim(rst!nomclient)) + ";" + treuresimbols(atrim(rst!descripcio))
      Print #3, vlinia
      rst.MoveNext
   Wend
   Close #3
   Set rst = Nothing
   If existeix(vfitxer) Then obrir_document vfitxer
End Sub
