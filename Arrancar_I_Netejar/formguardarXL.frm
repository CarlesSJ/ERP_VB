VERSION 5.00
Begin VB.Form formguardarXL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guardar XL"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   Icon            =   "formguardarXL.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEE4D7&
      Caption         =   "Escanejar codi XL de barres de la Estanteria"
      Height          =   1170
      Left            =   855
      TabIndex        =   1
      Top             =   2025
      Width           =   4290
      Begin VB.TextBox CBEstanteria 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   585
         TabIndex        =   3
         Top             =   360
         Width           =   3150
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Escanejar codi XL de barres de la Bossa"
      Height          =   1170
      Left            =   840
      TabIndex        =   0
      Top             =   435
      Width           =   4290
      Begin VB.TextBox CBBossa 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   630
         TabIndex        =   2
         Top             =   345
         Width           =   3150
      End
   End
   Begin VB.Label etoperari 
      Height          =   180
      Left            =   75
      TabIndex        =   4
      Top             =   45
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "formguardarXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBBossa_GotFocus()
    CBBossa.SelStart = 0
    CBBossa.SelLength = Len(CBBossa)
End Sub

Private Sub CBBossa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then comprovarxls
End Sub

Private Sub CBBossa_LostFocus()
   comprovarxls
End Sub

Private Sub CBEstanteria_GotFocus()
    CBEstanteria.SelStart = 0
    CBEstanteria.SelLength = Len(CBEstanteria)
End Sub

Private Sub CBEstanteria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then comprovarxls
End Sub
Sub comprovarxls()
  ' If CBBossa = "" Then
  '   CBBossa.SetFocus
  '      Else:
  '        CBEstanteria.SetFocus
  ' End If
   If atrim(CBBossa) <> "" Then
        If UCase(CBBossa) <> UCase(form1.etarxiuperdefecte.Tag) Then
           CBBossa.BackColor = QBColor(12)
             Else: CBBossa.BackColor = &H25EFAD: CBEstanteria.SetFocus
        End If
          Else: CBBossa.BackColor = QBColor(15)
          
   End If
   
   If atrim(CBEstanteria) <> "" Then
        If UCase(CBEstanteria) <> UCase(form1.etarxiuperdefecte.Tag) Then
           CBEstanteria.BackColor = QBColor(12)
             Else: CBEstanteria.BackColor = &H25EFAD
        End If
          Else: CBEstanteria.BackColor = QBColor(15)
   End If
   If UCase(CBBossa) = UCase(CBEstanteria) Then
      DoEvents
      wait 1
      guardarxl
      Unload formguardarXL
   End If
   
   
End Sub
Sub guardarxl()
      
'    dbbaixes.Execute "update muntadorescilindres set [opendreçar]=" + atrim(etoperari) + ", [dataendreça]=now where id in (" + form1.bguardarxl.Tag + ")"
    dbbaixes.Execute "update muntadorescilindres set [opendreçar]=" + atrim(etoperari) + ", [dataendreça]=now where numcomanda=" + atrim(cadbl(form1.etcomanda))
End Sub
Private Sub Timer1_Timer()
  
End Sub

Private Sub CBEstanteria_LostFocus()
   comprovarxls
End Sub

Private Sub Form_Activate()
  CBBossa.SetFocus
  
End Sub

