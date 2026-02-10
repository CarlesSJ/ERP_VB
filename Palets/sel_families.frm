VERSION 5.00
Begin VB.Form sel_families 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sel.leccionar families"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   ControlBox      =   0   'False
   Icon            =   "sel_families.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cfiltrar 
      Height          =   315
      ItemData        =   "sel_families.frx":058A
      Left            =   150
      List            =   "sel_families.frx":0597
      TabIndex        =   12
      Text            =   "Cap"
      Top             =   2190
      Width           =   2100
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Cancelar"
      Height          =   510
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2010
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Acceptar"
      Height          =   510
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2010
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1860
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   5820
      Begin VB.TextBox micres 
         Height          =   285
         Left            =   1380
         TabIndex        =   14
         Text            =   "0"
         Top             =   1500
         Width           =   675
      End
      Begin VB.ComboBox fammat 
         Height          =   315
         Left            =   495
         TabIndex        =   6
         Top             =   450
         Width           =   2580
      End
      Begin VB.ComboBox subfammat 
         Height          =   315
         Left            =   3105
         TabIndex        =   5
         Tag             =   "fammat"
         Top             =   450
         Width           =   2490
      End
      Begin VB.ComboBox subfamcol 
         Height          =   315
         Left            =   3105
         TabIndex        =   4
         Tag             =   "famcol"
         Top             =   780
         Width           =   2490
      End
      Begin VB.ComboBox famcol 
         Height          =   315
         Left            =   495
         TabIndex        =   3
         Top             =   780
         Width           =   2580
      End
      Begin VB.ComboBox subfamad 
         Height          =   315
         Left            =   3105
         TabIndex        =   2
         Tag             =   "famad"
         Top             =   1095
         Width           =   2490
      End
      Begin VB.ComboBox famad 
         Height          =   315
         Left            =   495
         TabIndex        =   1
         Top             =   1110
         Width           =   2580
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "( zero no filtra)"
         Height          =   240
         Left            =   2130
         TabIndex        =   16
         Top             =   1500
         Width           =   1650
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Micres"
         Height          =   285
         Left            =   795
         TabIndex        =   15
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Familia "
         Height          =   285
         Left            =   1995
         TabIndex        =   9
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Subfamilia "
         Height          =   285
         Left            =   3570
         TabIndex        =   8
         Top             =   225
         Width           =   2115
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Mat:         Col:          Ad:"
         Height          =   1020
         Left            =   180
         TabIndex        =   7
         Top             =   390
         Width           =   345
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar prestatgeries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   315
      TabIndex        =   13
      Top             =   1980
      Width           =   1650
   End
End
Attribute VB_Name = "sel_families"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   selecciomicres = 0
   nomfiltrefam = crear_nomfiltrefam
   nomfiltrefam = nomfiltrefam + IIf(cadbl(micres) > 0, " Micres=" + micres, "")
   selecciofam = crear_criteri_familia
   If selecciofam = "sortir" Then GoTo fi
   If cadbl(micres) > 0 Then selecciomicres = cadbl(micres)
   filtrarprestatge = cfiltrar.Text
   MsgBox "El llistat trigarà una bona estona... " + Chr(10) + Chr(13) + "Prem Acceptar per començar.", vbInformation, "Atenció"
   Unload sel_families
fi:
End Sub
Function crear_nomfiltrefam() As String
  crear_nomfiltrefam = fammat + " " + subfammat + " " + famcol + " " + subfamcol + " " + famad + " " + subfamad
End Function
Function crear_criteri_familia() As String
   Dim d As String
   If fammat.ListIndex >= 0 Then
      d = " familia=" + atrim(fammat.ItemData(fammat.ListIndex))
    Else: d = " familia>0"
   End If
   If subfammat.ListIndex >= 0 Then d = d + " and subfamilia=" + atrim(subfammat.ItemData(subfammat.ListIndex))
   If famcol.ListIndex >= 0 Then d = d + " and familiacol=" + atrim(famcol.ItemData(famcol.ListIndex))
   If subfamcol.ListIndex >= 0 Then d = d + " and subfamiliacol=" + atrim(subfamcol.ItemData(subfamcol.ListIndex))
   If famad.ListIndex >= 0 Then d = d + " and familiaad=" + atrim(famad.ItemData(famad.ListIndex))
   If subfamad.ListIndex >= 0 Then d = d + " and subfamiliaad=" + atrim(subfamad.ItemData(subfamad.ListIndex))
   If d = "" Then d = " familia=0 "
   crear_criteri_familia = d
   
End Function
Private Sub Command2_Click()
  selecciofam = "NO"
  Unload sel_families
End Sub

Private Sub Form_Load()
    nomfiltrefam = ""
    filtrarprestatge = ""
    carregar_combo_families
End Sub
Sub carregar_combo_families()
  Dim rstfam As Recordset
  
  Set rstfam = dbtmpb.OpenRecordset("select * from familiesmaterials where codi>499")
  fammat.Clear
  While Not rstfam.EOF
    fammat.AddItem atrim(rstfam!descripcio)
    fammat.ItemData(fammat.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
  Set rstfam = dbtmpb.OpenRecordset("select * from familiescolorants where codi>499")
  famcol.Clear
  While Not rstfam.EOF
    famcol.AddItem atrim(rstfam!descripcio)
    famcol.ItemData(famcol.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
  Set rstfam = dbtmpb.OpenRecordset("select * from familiesaditius where codi>499")
  famad.Clear
  While Not rstfam.EOF
    famad.AddItem atrim(rstfam!descripcio)
    famad.ItemData(famad.NewIndex) = cadbl(rstfam!codi)
    rstfam.MoveNext
  Wend
End Sub

Sub carregar_subfamilies(Optional combof As Control)
  Dim rstsub As Recordset
  Dim combo As Control
  Dim subfamilia As String
  
  Set combo = sel_families.ActiveControl
  If Not combof Is Nothing Then Set combo = combof
  If sel_families.Controls(combo.Tag).ListIndex = -1 And combof Is Nothing Then MsgBox "Primer has d'escullir la familia": Exit Sub
  'If combo.ListIndex = -1 Then combo.Clear: Exit Sub
  If combo.Name = "subfammat" And fammat.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(fammat.ItemData(fammat.ListIndex))): subfamilia = "subfamiliesmaterials"
  If combo.Name = "subfamcol" And famcol.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(famcol.ItemData(famcol.ListIndex))): subfamilia = "subfamiliescolorants"
  If combo.Name = "subfamad" And famad.ListIndex <> -1 Then r = " codifam=" + atrim(cadbl(famad.ItemData(famad.ListIndex))): subfamilia = "subfamiliesaditius"
    combo.Clear

  If subfamilia <> "" Then
     Set rstsub = dbtmpb.OpenRecordset("select codi,descripcio from " + subfamilia + " where " + r) '+ " and descripcio like '*" + treure_apostrof(subfammat.Text) + "*'")
    Else: Exit Sub
  End If
  
  While Not rstsub.EOF
    combo.AddItem atrim(rstsub!descripcio)
    combo.ItemData(combo.NewIndex) = cadbl(rstsub!codi)
    rstsub.MoveNext
  Wend
  
  
End Sub

Private Sub subfamad_DropDown()
   carregar_subfamilies
End Sub

Private Sub subfamcol_DropDown()
  carregar_subfamilies
End Sub

Private Sub subfammat_DropDown()
   carregar_subfamilies
End Sub
