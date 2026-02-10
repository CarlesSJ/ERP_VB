VERSION 5.00
Begin VB.Form Avisos 
   Caption         =   "Avisos"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   Icon            =   "Avisos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1785
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   675
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.DirListBox directoris 
      Height          =   315
      Left            =   3030
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.FileListBox fitxers 
      Height          =   285
      Left            =   2115
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   6840
   End
End
Attribute VB_Name = "Avisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim ruta_relativa As String
 centerscreen Me
  List1.Clear
  ruta_relativa_client = carpeta_del_client
  Me.Caption = ruta_relativa_client
  If existeix(ruta_relativa_docs + "\" + ruta_relativa_client + "\avisos") Then
    fitxers.Path = ruta_relativa_docs + "\" + ruta_relativa_client + "\avisos"
    fitxers.Refresh
    i = 0
    While i < fitxers.ListCount
     r = fitxers.List(i)
     List1.AddItem (UCase(r))
     i = i + 1
    Wend
   End If
   ratoli "normal"

End Sub

Private Sub List1_DblClick()
  Avisos.Caption = "Avisos        OBRINT..."
  
  obrir_document Chr$(34) + ruta_relativa_docs + "\" + ruta_relativa_client + "\avisos\" + List1 + Chr$(34)
  ' r = ruta_relativa_docs + "\" + ruta_relativa_client + "\avisos\" + List1
  '  b = "cmd /c "
  ' If existeix("c:\windows\command\start.exe") Then b = "start "
  ' r = Shell(b + Chr$(34) + r + Chr$(34), vbMinimizedFocus)
   wait (2)
  Avisos.Caption = "Avisos"
End Sub
