VERSION 5.00
Begin VB.Form escullirimpresora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Escullir Impresora"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   Icon            =   "escullirimpresora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   4620
      TabIndex        =   3
      Top             =   2085
      Width           =   1290
   End
   Begin VB.CommandButton bacceptar 
      Caption         =   "Acceptar"
      Height          =   360
      Left            =   3030
      TabIndex        =   2
      Top             =   2085
      Width           =   1290
   End
   Begin VB.ListBox llistaimpresores 
      Height          =   1620
      Left            =   135
      TabIndex        =   1
      Top             =   315
      Width           =   5835
   End
   Begin VB.Label nomimpresora 
      Height          =   165
      Left            =   375
      TabIndex        =   6
      Top             =   2145
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label driverimpresora 
      Height          =   165
      Left            =   150
      TabIndex        =   5
      Top             =   2295
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label portimpresora 
      Height          =   165
      Left            =   210
      TabIndex        =   4
      Top             =   2070
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom de la Impresora"
      Height          =   270
      Left            =   315
      TabIndex        =   0
      Top             =   60
      Width           =   2370
   End
End
Attribute VB_Name = "escullirimpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bacceptar_Click()
  Me.Tag = "acceptar"
  Me.Hide
End Sub

Private Sub Command1_Click()
  Me.Tag = "cancelar"
  Me.Hide
End Sub

Private Sub Form_Load()
   'emplenarllistaimpresores
End Sub
Sub emplenarllistaimpresores(Optional nomdelaescullida As String)
   Dim p As Printer
   Dim indexescullida As Integer
   llistaimpresores.Clear
   For Each p In Printers
     llistaimpresores.AddItem p.DeviceName
     If UCase(p.DeviceName) = UCase(nomdelaescullida) Then indexescullida = llistaimpresores.NewIndex
   Next p
   If llistaimpresores.ListCount > 0 Then
      dadesimpresora indexescullida
   End If
End Sub
Sub dadesimpresora(indexescullida As Integer)
   Dim p As Printer
   llistaimpresores.ListIndex = indexescullida
   For Each p In Printers
     If UCase(p.DeviceName) = UCase(llistaimpresores.Text) Then
        nomimpresora = p.DeviceName
        portimpresora = p.Port
        driverimpresora = p.DriverName
        Exit Sub
     End If
   Next p
End Sub

Private Sub llistaimpresores_Click()
  dadesimpresora llistaimpresores.ListIndex
End Sub
