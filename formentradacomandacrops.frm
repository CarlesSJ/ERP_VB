VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form formentradacomandacrops 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada de comandes de Crop's"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3675
   Icon            =   "formentradacomandacrops.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Sortir"
      Height          =   480
      Left            =   1905
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Aplicar Canvis"
      Height          =   480
      Left            =   525
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   5820
      Left            =   135
      TabIndex        =   0
      Top             =   960
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   10266
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   3570
      Begin VB.CommandButton alta 
         Height          =   420
         Left            =   3060
         Picture         =   "formentradacomandacrops.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   255
         Width           =   405
      End
      Begin VB.TextBox cpedidoclient 
         Height          =   285
         Left            =   1185
         TabIndex        =   2
         Top             =   390
         Width           =   1845
      End
      Begin VB.TextBox ccomanda 
         Height          =   285
         Left            =   135
         MaxLength       =   6
         TabIndex        =   1
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label eterror 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   75
         TabIndex        =   8
         Top             =   645
         Width           =   3420
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Comanda      Nº Pedido del client"
         Height          =   255
         Left            =   165
         TabIndex        =   7
         Top             =   150
         Width           =   3090
      End
   End
End
Attribute VB_Name = "formentradacomandacrops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
   afegir_pedido_a_reixa
End Sub

Private Sub ccomanda_Change()
  If Len(ccomanda) = 6 Then cpedidoclient.SetFocus
  eterror = ""
End Sub

Private Sub ccomanda_LostFocus()
   demanar_pedido
End Sub
Sub demanar_pedido()
  Dim rst As Recordset
  Dim ultimnumcomanda As String
  ultimnumcomanda = atrim(cadbl(ccomanda))
  Set rst = dbtmp.OpenRecordset("select client,comandaclient,producte from comandes where comanda=" + atrim(cadbl(ccomanda)))
  If Not rst.EOF Then
       eterror = ""
       If InStr(1, rst!producte, "PC") > 0 Then msg = "La comanda " + atrim(ultimnumcomanda) + " no es la principal.": GoTo nocrops
       If cadbl(rst!client) <> 6841 Then msg = "La comanda " + atrim(ultimnumcomanda) + " no es de Crop's": GoTo nocrops
       If atrim(rst!comandaclient) <> "" Then msg = "La comanda " + atrim(ultimnumcomanda) + " ja te el pedido entrat.": GoTo nocrops
    Else
nocrops:
      If Screen.ActiveControl.Name = "cpedidoclient" Then ccomanda.SetFocus
      ccomanda = ""
      eterror = msg
  End If
End Sub

Private Sub Command1_Click()
  Dim i As Integer
   Dim vnumc As String
   Dim vnumpedido As String
   
   For i = 1 To reixa.Rows - 1
      reixa.row = i
      reixa.col = 0
      vnumc = reixa
      reixa.col = 1
      vnumpedido = reixa
      If cadbl(vnumc) > 0 Then modificarpedido_i_imprimirprimerapaginacomanda cadbl(vnumc), atrim(vnumpedido)
   Next i
End Sub
Sub modificarpedido_i_imprimirprimerapaginacomanda(vnumc As Double, vnumpedido As String)
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(vnumc))
   If Not rst.EOF Then Set rst = dbtmp.OpenRecordset("select comanda from comandes where (comanda=" + atrim(rst!comanda) + " or comanda=" + atrim(rst!linkcomanda1) + " or comanda=" + atrim(rst!linkcomanda2) + ") and comanda>0")
   While Not rst.EOF
      dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(cadbl(rst!comanda)) + ",'" + nomordinador + "','comandaclient','','" + treure_apostruf(vnumpedido) + "')"
      dbtmp.Execute "update comandes set comandaclient='" + treure_apostruf(vnumpedido) + "' where comanda=" + atrim(cadbl(rst!comanda))
      rst.MoveNext
   Wend
   'dbtmp.Execute "insert into comandes_controlcanvis (comanda,usuari,campafectat,valoranterior,valoractual) values (" + atrim(vnumc) + ",'" + nomordinador + "','Imprimir1afulla_Nºpedido','','')"
   
   formcomandes.llistar_comanda False, atrim(vnumc), True
   
   
End Sub
Private Sub Command2_Click()
   Unload formentradacomandacrops
   
End Sub

Private Sub cpedidoclient_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then afegir_pedido_a_reixa
End Sub
Sub afegir_pedido_a_reixa()
  Dim msg As String
  If cpedidoclient = "" Then eterror = "El pedido del client està en blanc.": cpedidoclient.SetFocus: Exit Sub
  If jaexisteixelpedido(cadbl(ccomanda)) Then
      msg = "La comanda " + atrim(cadbl(ccomanda)) + " ja està entrada a la reixa"
       Else: reixa.AddItem ccomanda + Chr(9) + cpedidoclient
  End If
  ccomanda = ""
  cpedidoclient = ""
  ccomanda.SetFocus
  eterror = msg
End Sub
Function jaexisteixelpedido(numc As Double) As Boolean
   Dim i As Integer
   reixa.col = 0
   For i = 1 To reixa.Rows - 1
      reixa.row = i
      If cadbl(reixa) = numc Then jaexisteixelpedido = True: Exit Function
   Next i
End Function
Private Sub Form_Activate()
   ccomanda.SetFocus
End Sub

Private Sub Form_Load()
  reixa.Cols = 2
  reixa.ColWidth(0) = 1000
  reixa.ColWidth(1) = 1800
  reixa.Rows = 0
  reixa.AddItem "NºComanda" + Chr(9) + "Nº Pedido Client"
  reixa.Rows = 2
  reixa.FixedRows = 1
  reixa.FixedCols = 0
  reixa.Rows = 1
  
  
  
End Sub
