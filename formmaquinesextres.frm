VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formmaquinesextres 
   BackColor       =   &H00F1B75F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalls extres de la màquina"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Data dataextres 
      Caption         =   "dataextres"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1035
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   15
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Detall tinters"
      Height          =   3240
      Left            =   180
      TabIndex        =   17
      Top             =   1500
      Width           =   4215
      Begin VB.Data datatintes 
         Caption         =   "datatintes"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   870
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2385
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.CommandButton bactualitzar 
         BackColor       =   &H0025EFAD&
         Caption         =   "Actualitzar tintes"
         Height          =   300
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Height          =   2490
         Left            =   150
         OleObjectBlob   =   "formmaquinesextres.frx":0000
         TabIndex        =   20
         Top             =   585
         Width           =   4005
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1290
         TabIndex        =   19
         Top             =   225
         Width           =   330
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Nº Tinters màx:"
         Height          =   270
         Index           =   2
         Left            =   165
         TabIndex        =   18
         Top             =   300
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Amplada material"
      Height          =   1245
      Left            =   2160
      TabIndex        =   10
      Top             =   120
      Width           =   2220
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1020
         TabIndex        =   14
         Top             =   630
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1020
         TabIndex        =   13
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label4 
         BackColor       =   &H00EAD9CE&
         Caption         =   "mm"
         Height          =   270
         Index           =   2
         Left            =   1725
         TabIndex        =   16
         Top             =   675
         Width           =   360
      End
      Begin VB.Label Label4 
         BackColor       =   &H00EAD9CE&
         Caption         =   "mm"
         Height          =   270
         Index           =   1
         Left            =   1725
         TabIndex        =   15
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Ample Màx:"
         Height          =   270
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   375
         Width           =   1065
      End
      Begin VB.Label Label2 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Ample Mín:"
         Height          =   270
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Top             =   645
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Preu hora"
      Height          =   1260
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   1830
      Begin VB.TextBox cpreuhora 
         Height          =   285
         Index           =   2
         Left            =   915
         TabIndex        =   6
         Top             =   825
         Width           =   435
      End
      Begin VB.TextBox cpreuhora 
         Height          =   285
         Index           =   1
         Left            =   915
         TabIndex        =   5
         Top             =   540
         Width           =   435
      End
      Begin VB.TextBox cpreuhora 
         Height          =   285
         Index           =   0
         Left            =   915
         TabIndex        =   4
         Top             =   255
         Width           =   435
      End
      Begin VB.Label Label6 
         BackColor       =   &H00EAD9CE&
         Caption         =   "€/h"
         Height          =   270
         Index           =   0
         Left            =   1365
         TabIndex        =   9
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label5 
         BackColor       =   &H00EAD9CE&
         Caption         =   "€/h"
         Height          =   270
         Index           =   0
         Left            =   1365
         TabIndex        =   8
         Top             =   555
         Width           =   360
      End
      Begin VB.Label Label4 
         BackColor       =   &H00EAD9CE&
         Caption         =   "€/h"
         Height          =   270
         Index           =   0
         Left            =   1365
         TabIndex        =   7
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label3 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Tarifa 3:"
         Height          =   270
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   855
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Tarifa 2:"
         Height          =   270
         Index           =   0
         Left            =   195
         TabIndex        =   2
         Top             =   593
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Tarifa 1:"
         Height          =   270
         Index           =   0
         Left            =   195
         TabIndex        =   1
         Top             =   330
         Width           =   735
      End
   End
End
Attribute VB_Name = "formmaquinesextres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
  'MsgBox Trim(Me.Left) + " - " + Trim(Me.Top)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  formaltamaquines.bextres.Tag = ""
End Sub

Private Sub Text3_LostFocus()
  'If Text3.Tag <> Text3 Then bactualitzar.Visible
End Sub
