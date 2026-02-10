VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formconsumllaunes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consums Llaunes"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton novallauna 
      Caption         =   "Nova Llauna"
      Height          =   615
      Left            =   180
      Picture         =   "formconsumllaunes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir l'etiqueta de bobina d'entrada parcial."
      Top             =   3615
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Height          =   465
      Left            =   1800
      Picture         =   "formconsumllaunes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1395
      Width           =   690
   End
   Begin VB.CommandButton bcancelar 
      BackColor       =   &H008080FF&
      Caption         =   "Cancelar"
      Height          =   450
      Left            =   4395
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3735
      Width           =   1410
   End
   Begin VB.CommandButton bdacord 
      BackColor       =   &H0080FF80&
      Caption         =   "d'Acord"
      Height          =   450
      Left            =   2955
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3735
      Width           =   1410
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   2355
      Left            =   2505
      TabIndex        =   5
      Top             =   1260
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4154
      _Version        =   393216
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nom tinta del tinter"
      Height          =   615
      Left            =   105
      TabIndex        =   3
      Top             =   45
      Width           =   5775
      Begin VB.Label nomtinta 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   5595
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Llaunes"
      Height          =   2880
      Left            =   105
      TabIndex        =   0
      Top             =   690
      Width           =   1680
      Begin VB.TextBox numllauna 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   105
         TabIndex        =   2
         Top             =   270
         Width           =   1410
      End
      Begin VB.ListBox llistallaunes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   735
         Width           =   1470
      End
   End
End
Attribute VB_Name = "formconsumllaunes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bcancelar_Click()
  Unload Me
End Sub

Private Sub bdacord_Click()
 gravar_reixa cadbl(formconsumllaunes.tag)
 Unload Me
End Sub

Private Sub Command1_Click()
    Dim i As Byte
    If numllauna <> "" Then
       afegirllauna numllauna, nomtinta.tag
       numllauna = ""
       Exit Sub
    End If
    For i = 0 To llistallaunes.ListCount - 1
       If llistallaunes.Selected(i) Then
          afegirllauna llistallaunes.List(i), nomtinta.tag
       End If
    Next i
    
End Sub
Function jahiesalallista(numllauna As String) As Boolean
    Dim i As Byte
    For i = 2 To reixa.Rows
      If reixa.TextMatrix(i - 1, 0) = numllauna Then jahiesalallista = True
    Next i
End Function
Sub afegirllauna(numllauna As String, coditinta As String)
   numllauna = UCase(numllauna)
   If jahiesalallista(numllauna) Then MsgBox "La llauna " + atrim(numllauna) + " ja ès a la llista.", vbCritical, "Error": Exit Sub
    If Not escorrectelatinta(numllauna, coditinta) Then
       MsgBox "La tinta de la llauna " + atrim(numllauna) + " no es correspont amb la del tinter o la llauna no està activa.", vbCritical, "Atenció"
       Exit Sub
        Else
          afegirlallaunaalallista numllauna
    End If
End Sub
Sub afegirlallaunaalallista(numllauna As String)
   Dim resp As String
   resp = InputBox("Entra els Kg de tinta gastats per aquest tinter.", "Tinta gastada")
   If cadbl(resp) > 0 Then
    reixa.AddItem numllauna + Chr(9) + resp
   End If
End Sub
Function escorrectelatinta(numllauna As String, coditinta As String) As Boolean
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna,llaunes.activa FROM tintes LEFT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta WHERE llaunes.numllauna='" + atrim(numllauna) + "' and tintes.codi='" + atrim(coditinta) + "';")
  If Not rst.EOF Then
     If rst!activa Then
         escorrectelatinta = True
     End If
  End If
End Function
Private Sub Form_Activate()
   If nomtinta = "" Then
    nomtinta = nomdelatinta(nomtinta.tag)
    configurar_reixa
    carregar_reixa cadbl(formconsumllaunes.tag)
    carregar_llistallaunes nomtinta.tag
   End If
End Sub

Sub carregar_llistallaunes(coditinta As String)
   Dim rst As Recordset
   llistallaunes.Clear
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna FROM tintes LEFT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta WHERE (((Llaunes.situacio)='Imp') AND ((Llaunes.activa)=True) and tintes.codi='" + atrim(coditinta) + "');")
   While Not rst.EOF
     llistallaunes.AddItem atrim(rst!numllauna)
     rst.MoveNext
   Wend
   If llistallaunes.ListCount = 0 Then MsgBox "No hi ha cap tinta disponibles per aquesta comanda", vbCritical, "Atenció": Exit Sub
End Sub
Function nomdelatinta(coditinta As String) As String
   Dim rst As Recordset
   
   Set rst = dbtintes.OpenRecordset("select descripcio from tintes where codi='" + atrim(coditinta) + "'")
   If Not rst.EOF Then
       nomdelatinta = atrim(rst!descripcio)
         Else: nomdelatinta = ""
   End If
   Set rst = Nothing
End Function

Sub configurar_reixa()
   reixa.ColWidth(0) = 1300
   reixa.ColWidth(1) = 1600
   reixa.TextMatrix(0, 0) = "Nº Llauna"
   reixa.TextMatrix(0, 1) = "Kg Gastats"
   reixa.FixedCols = 1
End Sub
Sub netejar_reixa()
   reixa.Cols = 2
   reixa.Rows = 1
   
End Sub
Sub carregar_reixa(numc As Double)
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, historiallauna.* FROM Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna where tipusmoviment='I' and comanda=" + atrim(numc) + ";")
   netejar_reixa
   While Not rst.EOF
    reixa.AddItem atrim(rst!numllauna) + Chr(9) + atrim(rst!kg)
    'reixa.TextMatrix(reixa.Rows, 0) = rst!numllauna
    'reixa.TextMatrix(reixa.Rows, 1) = rst!kg
    rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub gravar_reixa(numc As Double)
   Dim rst As Recordset
   Dim fila As Byte
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, historiallauna.* FROM Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna where tipusmoviment='I' and comanda=" + atrim(numc) + ";")
   'netejar_reixa
   fila = 1
   While fila < reixa.Rows
    rst.FindFirst "numllauna='" + atrim(reixa.TextMatrix(fila, 0)) + "'"
    If rst.NoMatch Then
         If cadbl(reixa.TextMatrix(fila, 1)) > 0 Then
                rst.AddNew
                rst!tipusmoviment = "I"
                rst!idnumllauna = iddenumllauna(atrim(reixa.TextMatrix(fila, 0)))
                rst!Data = Now
                rst!comanda = numc
                rst!kg = cadbl(reixa.TextMatrix(fila, 1))
                rst.Update
         End If
        Else
          If cadbl(reixa.TextMatrix(fila, 1)) = 0 Then
            rst.Delete
             Else
                rst.Edit
                rst.kg = cadbl(reixa.TextMatrix(fila, 1))
                rst.Update
          End If
    End If
    calcularkgdisponiblesllauna atrim(reixa.TextMatrix(fila, 0))
    fila = fila + 1
   Wend
   Set rst = Nothing
End Sub
Function iddenumllauna(numllauna As String) As Integer
  Dim rst As Recordset
  iddenumllauna = 0
  Set rst = dbtintes.OpenRecordset("select id from llaunes where numllauna='" + atrim(numllauna) + "'")
  If Not rst.EOF Then iddenumllauna = rst!id
  Set rst = Nothing
End Function

Private Sub novallauna_Click()

   'crear llauna nova passant els parametres de idtinta i idrefproveidor de la anterior
   'crear la historia nova amb "K" i guardant la iddela llauna anterior
   
End Sub

Private Sub reixa_DblClick()
  Dim resp As String
  resp = InputBox("Entra els kilos consumits d'aquesta tinta", "Tinta gastada", reixa.Text)
  If cadbl(resp) > 0 Or resp = "0" Then reixa.Text = atrim(cadbl(resp))
End Sub
