VERSION 5.00
Begin VB.Form formqualitatimpresio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Qualitat de la impresió"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   ClipControls    =   0   'False
   Icon            =   "formqualitatimpresio.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Dolenta"
      Height          =   960
      Left            =   3105
      Picture         =   "formqualitatimpresio.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   75
      Width           =   1290
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Bona"
      Height          =   960
      Left            =   1605
      Picture         =   "formqualitatimpresio.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Molt Bona"
      Height          =   960
      Left            =   90
      Picture         =   "formqualitatimpresio.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   1290
   End
End
Attribute VB_Name = "formqualitatimpresio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   escullirqualitat 3
End Sub

Private Sub Command2_Click()
  escullirqualitat 2
End Sub
Sub escullirqualitat(vqualitat As Byte)
   dbtmpb.Execute "update impressorestot set qualitatimpresio=" + atrim(vqualitat) + " where comanda=" + atrim(cadbl(Form1.comanda))
   Unload formqualitatimpresio
End Sub

Private Sub Command3_Click()
   escullirqualitat 1
End Sub

