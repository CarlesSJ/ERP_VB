VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form formverificarlotsinkmaker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verificar lots dels dosificadors de Inkmaker"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   5835
      Left            =   165
      TabIndex        =   1
      Top             =   1095
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10292
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox ccodi 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   735
      TabIndex        =   0
      Top             =   285
      Width           =   2310
   End
   Begin VB.Label etdosificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005C31DD&
      Height          =   285
      Left            =   210
      TabIndex        =   2
      Top             =   825
      Width           =   4125
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   3090
      Picture         =   "formverificaciolotsinkmaker.frx":0000
      Stretch         =   -1  'True
      Top             =   330
      Width           =   450
   End
End
Attribute VB_Name = "formverificarlotsinkmaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub comprovarcodiescanejat()
  If etdosificador = "" Then
      ccodi = UCase(ccodi)
      If Mid(ccodi, 1, 1) = "I" Then
        etdosificador = ccodi
         Else: etdosificador = ""
      End If
        Else
          afegiralareixa UCase(etdosificador), UCase(ccodi)
          etdosificador = ""
  End If
  ccodi = ""
End Sub
Function escorrectelotinkmaker(vdosificador As String, vcodi As String) As Boolean
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim vllauna As String
   Set rst = dbtintes.OpenRecordset("SELECT * fROM Componentsbase where numdosificador=" + Mid(vdosificador, 2))
   If Not rst.EOF Then
      Set rstc = dbtintes.OpenRecordset("select * from detallnumeroslotsbase where idcomponent=" + atrim(rst!idcomponent) + " order by data desc")
      If Not rstc.EOF Then
         vllauna = atrim(rstc!numerodelot)
         If vllauna = vcodi Then escorrectelotinkmaker = True
      End If
   End If
   Set rst = Nothing
   Set rstc = Nothing
End Function
Sub afegiralareixa(vdosificador As String, vcodi As String)
   reixa.AddItem vdosificador + Chr(9) + vcodi
   If escorrectelotinkmaker(vdosificador, vcodi) Then
       reixa.Row = reixa.Rows - 1
       reixa.col = 0
       reixa.CellBackColor = QBColor(10)
       reixa.col = 1
       reixa.CellBackColor = QBColor(10)
         Else
           reixa.Row = reixa.Rows - 1
           reixa.col = 0
           reixa.CellBackColor = QBColor(12)
           reixa.col = 1
           reixa.CellBackColor = QBColor(12)
    End If
End Sub
Private Sub ccodi_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      comprovarcodiescanejat
      KeyAscii = 0
  End If
End Sub

Private Sub Form_Activate()
  ccodi.SetFocus
End Sub

Private Sub Form_Load()
  reixa.ColWidth(0) = 2000
  reixa.ColWidth(1) = 2000
  reixa.TextMatrix(0, 0) = "Dosificador"
  reixa.TextMatrix(0, 1) = "Lot escanejat"

End Sub

Private Sub Image1_Click()
  ccodi.SetFocus
End Sub
