VERSION 5.00
Begin VB.Form triarversiotreball 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Escullir versió del treball"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   ControlBox      =   0   'False
   Icon            =   "triarversiotreball.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Versions"
      Height          =   3315
      Left            =   105
      TabIndex        =   3
      Top             =   735
      Width           =   7275
      Begin VB.ListBox llistaversions 
         Columns         =   1
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   105
         TabIndex        =   4
         Top             =   210
         Width           =   7050
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   90
      TabIndex        =   0
      Top             =   -30
      Width           =   7290
      Begin VB.CommandButton Command1 
         Caption         =   "Cap Treball"
         Height          =   450
         Left            =   5490
         Picture         =   "triarversiotreball.frx":058A
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Cap treball seleccionat"
         Top             =   195
         Width           =   1140
      End
      Begin VB.CommandButton sortir 
         Height          =   450
         Left            =   6690
         Picture         =   "triarversiotreball.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Sortir a Menú"
         Top             =   195
         Width           =   465
      End
      Begin VB.CommandButton Command11 
         Height          =   450
         Left            =   2325
         Picture         =   "triarversiotreball.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Acceptar els canvis (F1)."
         Top             =   165
         Width           =   465
      End
      Begin VB.TextBox ntreball 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   945
         TabIndex        =   1
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Treball"
         Height          =   225
         Left            =   60
         TabIndex        =   2
         Top             =   315
         Width           =   930
      End
   End
End
Attribute VB_Name = "triarversiotreball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If MsgBox("Segur que no vols escullir cap treball per aquesta comanda?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
     ntreball = "-1"
     Me.Hide
  End If
End Sub

Private Sub Command11_Click()
   carregarversions cadbl(ntreball)
End Sub

Private Sub llistaversions_DblClick()
   Me.Hide
End Sub

Private Sub sortir_Click()
  ntreball = ""
  Me.Hide
End Sub

Sub carregarversions(numt As Double)
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim ultimacomanda As String
   Dim estatclixes As String
   Dim impresio As String
   Dim verrortintersrepetits As String
   llistaversions.Clear
   
   Set rst = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(numt))
   While Not rst.EOF
      ultimacomanda = ""
      verrortintersrepetits = ""
      Set rstc = dbtmp.OpenRecordset("select max(comanda) as maxcomanda,last(datacomanda) as datac,last(proximaseccio) as proxima,last(numordremodificacio) as vordre  from comandes where numtreball=" + atrim(numt) + " and (numordremodificacio=" + atrim(cadbl(rst!ordre)) + " or numordremodificacio=0) and (proximaseccio<>'I' and proximaseccio<>'E') order by max(comanda) DESC")
      
      
      If Not rstc.EOF Then
          If rstc!vordre <> 0 Then
            ultimacomanda = atrim(rstc!maxcomanda) + " " + atrim(rstc!proxima) + " " + atrim(rstc!datac)
          End If
      End If
      estatclixes = "  " + formcomandes.estatdelclixe(numt, cadbl(rst!ordre))
      estatclixes = Mid(estatclixes, InStr(1, estatclixes, " - ") + 3)
      impresio = "M"
      If ultimacomanda = "" And cadbl(rst!ordre) = 1 Then impresio = "N"
      If (estatclixes = "CLIXES ENTRATS" Or estatclixes = "REPOSICIÓ DEL CLIXE") And InStr(1, "IE", atrim(rstc!proxima)) = 0 Then impresio = "R"
      If (estatclixes <> "CLIXES ENTRATS" And estatclixes <> "REPOSICIÓ DEL CLIXE") Then
         If IsNull(rstc!proxima) Then
            If rst!ordre = 1 Then impresio = "N"
             Else
                If InStr(1, "IE", atrim(rstc!proxima)) = 0 Then
                     impresio = "N"
                       Else: impresio = "R"
                End If
         End If
      End If
      If (estatclixes = "RETORNEM CLIXES") Then
          If ultimacomanda = "" Then
             impresio = "N"
               Else: impresio = "R"
          End If
      End If
      verrortintersrepetits = revisartintersrepetits(numt, cadbl(rst!ordre))
      If verrortintersrepetits = "" Then
         verrortintersrepetits = estatclixes
           Else: llistaversions.BackColor = QBColor(12)
      End If
      llistaversions.AddItem impresio + " " + justificar("v" + atrim(cadbl(rst!ordre)), 6) + "|" + justificar(verrortintersrepetits, 25) + "| " + ultimacomanda
      llistaversions.ItemData(llistaversions.NewIndex) = cadbl(rst!ordre)
      rst.MoveNext
   Wend
   Set rst = Nothing
   Set rstc = Nothing
End Sub

Function revisartintersrepetits(vnumtreball As Double, vordre As Double) As String
   Dim rst As Recordset
   Dim vultim As Integer
   Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(vnumtreball) + " and ordremodificacio=" + atrim(vordre) + " order by ordretinter")
   vultim = 0
   While Not rst.EOF
      If vultim = cadbl(rst!ordretinter) Then
         revisartintersrepetits = " (OJU! Tinter repetit)": GoTo fi
      End If
      vultim = cadbl(rst!ordretinter)
      rst.MoveNext
   Wend
fi:
   Set rst = Nothing

End Function
