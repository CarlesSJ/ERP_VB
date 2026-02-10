VERSION 5.00
Begin VB.Form formselecciotipuscompra 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Escullir tipus compra"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Height          =   285
      Left            =   5220
      Picture         =   "selecciotipuscompra.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   15
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00F1B75F&
      Caption         =   "Material varis"
      Height          =   645
      Left            =   3525
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   180
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00ED823A&
      Caption         =   "Tintes"
      Height          =   645
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   1575
   End
   Begin VB.CommandButton boto1 
      BackColor       =   &H00FDDECE&
      Caption         =   "Bobines Material"
      Height          =   645
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "formselecciotipuscompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub boto1_Click()
   Dim rst As Recordset
   formselecciotipuscompra.tag = ""
   Set rst = dbtmp.OpenRecordset("select * from proveidors where codi=" + atrim(comandescompra.capcalera.Recordset!codiproveidor))
   If Not rst.EOF Then
       If atrim(rst!tipusCQ) = "" Then rst.Edit: rst!tipusCQ = "L": rst.Update
       If atrim(rst!tipusproveidorIMPOST) = "" Then MsgBox "NO POTS FER COMPRES DE FILM AMB AQUEST PROVEIDOR PERQUÈ ENCARA NO TE POSSAT EL TIPUS DE PROVEIDOR D'IMPOST AL MANTENIMENT DE PROVEIDORS.", vbCritical, "ERROR": GoTo fi
   End If
   formselecciotipuscompra.tag = "M"
fi:
   formselecciotipuscompra.Hide
   Set rst = Nothing
End Sub

Private Sub Command1_Click()
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from proveidors where codi=" + atrim(comandescompra.capcalera.Recordset!codiproveidor))
   If Not rst.EOF Then If atrim(rst!tipusCQ) = "" Then rst.Edit: rst!tipusCQ = "L": rst.Update
   formselecciotipuscompra.tag = "T"
   formselecciotipuscompra.Hide
   Set rst = Nothing
End Sub

Private Sub Command2_Click()
formselecciotipuscompra.tag = "V"
   formselecciotipuscompra.Hide
End Sub

Private Sub Command3_Click()
  formselecciotipuscompra.tag = ""
   formselecciotipuscompra.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Command3_Click
End Sub
