VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formvalorsdelta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valors delta"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9510
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton eliminar 
      Height          =   450
      Left            =   675
      Picture         =   "formvalorsdelta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Eliminacio Registres"
      Top             =   105
      Width           =   465
   End
   Begin VB.CommandButton alta 
      Height          =   450
      Left            =   195
      Picture         =   "formvalorsdelta.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Alta  Registres"
      Top             =   105
      Width           =   465
   End
   Begin VB.CommandButton sortir 
      Height          =   525
      Left            =   8460
      Picture         =   "formvalorsdelta.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Imprimir ticket valors delta."
      Top             =   45
      Width           =   570
   End
   Begin VB.Data datadelta 
      Caption         =   "datadelta"
      Connect         =   "Access"
      DatabaseName    =   "W:\progcomandes\dades\baixes.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   2670
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "impresores_valorsdelta"
      Top             =   3465
      Visible         =   0   'False
      Width           =   2745
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "formvalorsdelta.frx":109E
      Height          =   3165
      Left            =   1380
      OleObjectBlob   =   "formvalorsdelta.frx":10B2
      TabIndex        =   1
      Top             =   630
      Width           =   7920
   End
   Begin VB.ListBox llistalectures 
      Height          =   3180
      Left            =   135
      TabIndex        =   0
      Top             =   615
      Width           =   1020
   End
End
Attribute VB_Name = "formvalorsdelta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
