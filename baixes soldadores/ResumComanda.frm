VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormResumComanda 
   Caption         =   "Resum de la Comanda"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12045
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8745
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command29 
      Height          =   870
      Left            =   11145
      Picture         =   "ResumComanda.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Document descriptiu de com ha de ser la sortida"
      Top             =   1530
      Width           =   690
   End
   Begin VB.Frame Framedadescomanda 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Dades de la comanda"
      Height          =   2970
      Left            =   60
      TabIndex        =   1
      Top             =   75
      Width           =   11910
      Begin VB.ComboBox combosacsocaixes 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "ResumComanda.frx":07FE
         Left            =   9945
         List            =   "ResumComanda.frx":0808
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "Caixes"
         ToolTipText     =   "Sacs o Caixes"
         Top             =   2385
         Width           =   1815
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   20
         Left            =   7380
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1800
         Width           =   3270
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "obssol1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   19
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2400
         Width           =   8325
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "costatobertsol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   18
         Left            =   6195
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1785
         Width           =   675
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "microperforatsol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   17
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1830
         Width           =   675
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "simulteneitatsol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   16
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1140
         Width           =   675
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "TAC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   15
         Left            =   4755
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1155
         Width           =   675
      End
      Begin VB.TextBox camps 
         Height          =   300
         Index           =   14
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1050
         Width           =   2880
      End
      Begin VB.TextBox camps 
         Height          =   300
         Index           =   13
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1485
         Width           =   3240
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   12
         Left            =   10545
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   11
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1065
         Width           =   540
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   10
         Left            =   10545
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   780
         Width           =   900
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   9
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   765
         Width           =   540
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   8
         Left            =   10545
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   465
         Width           =   900
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   7
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   450
         Width           =   540
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "fuellebocasol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   7770
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   465
         Width           =   675
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "fuellebasesol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   450
         Width           =   675
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "solapasol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   450
         Width           =   675
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "longitudsol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   465
         Width           =   675
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "ampleplegsol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   465
         Width           =   675
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "amplesol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   465
         Width           =   675
      End
      Begin VB.TextBox camps 
         Alignment       =   2  'Center
         DataField       =   "producte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sacs o Caixes->"
         Height          =   255
         Index           =   17
         Left            =   8700
         TabIndex        =   56
         Top             =   2490
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "El client demana:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   7350
         TabIndex        =   39
         Top             =   1575
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacions"
         Height          =   255
         Index           =   15
         Left            =   195
         TabIndex        =   37
         Top             =   2115
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Costat Obert?"
         Height          =   255
         Index           =   14
         Left            =   6075
         TabIndex        =   35
         Top             =   1590
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MicroPerforat"
         Height          =   255
         Index           =   13
         Left            =   4620
         TabIndex        =   33
         Top             =   1635
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Simulteneitat:"
         Height          =   240
         Index           =   12
         Left            =   6090
         TabIndex        =   31
         Top             =   945
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TAC:"
         Height          =   240
         Index           =   11
         Left            =   4875
         TabIndex        =   29
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   8820
         TabIndex        =   27
         Top             =   1380
         Width           =   2940
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Màquina:"
         Height          =   240
         Index           =   9
         Left            =   195
         TabIndex        =   24
         Top             =   1530
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipus Soldadura:"
         Height          =   240
         Index           =   8
         Left            =   180
         TabIndex        =   23
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mesura:"
         Height          =   240
         Index           =   7
         Left            =   10785
         TabIndex        =   9
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Espesor:"
         Height          =   240
         Index           =   6
         Left            =   9270
         TabIndex        =   8
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "F.Boca:"
         Height          =   240
         Index           =   5
         Left            =   7755
         TabIndex        =   7
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "F.Base:"
         Height          =   240
         Index           =   4
         Left            =   6240
         TabIndex        =   6
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Solapa:"
         Height          =   240
         Index           =   3
         Left            =   4725
         TabIndex        =   5
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Llarg:"
         Height          =   240
         Index           =   2
         Left            =   3210
         TabIndex        =   4
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ample/Plegat:"
         Height          =   240
         Index           =   1
         Left            =   1530
         TabIndex        =   3
         Top             =   255
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Producte:"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   255
         Width           =   780
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acceptar"
      Height          =   390
      Left            =   10200
      TabIndex        =   0
      Top             =   8115
      Width           =   1740
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEE4D7&
      Caption         =   "Accessoris"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5550
      Left            =   90
      TabIndex        =   41
      Top             =   3105
      Width           =   11895
      Begin MSFlexGridLib.MSFlexGrid reixa 
         Bindings        =   "ResumComanda.frx":081A
         Height          =   4350
         Left            =   135
         TabIndex        =   45
         Top             =   420
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   7673
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H008080FF&
         Caption         =   "-  Eliminar Accessori"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4815
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1365
         Width           =   1860
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H006BEBB1&
         Caption         =   "+  Afegir Accessori"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4815
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   510
         Width           =   1860
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dades de l'accessori"
         Height          =   4650
         Left            =   6780
         TabIndex        =   42
         Top             =   285
         Width           =   4965
         Begin VB.CommandButton Command4 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Llegir Codi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   3480
            Picture         =   "ResumComanda.frx":0833
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Escanejar Refinplacsa"
            Top             =   3510
            Width           =   1440
         End
         Begin VB.CommandButton bfotos 
            Caption         =   "Ubicació"
            Height          =   420
            Index           =   2
            Left            =   1725
            Style           =   1  'Graphical
            TabIndex        =   51
            Tag             =   "U"
            Top             =   225
            Width           =   795
         End
         Begin VB.CommandButton bfotos 
            Caption         =   "Posició"
            Height          =   420
            Index           =   1
            Left            =   930
            Style           =   1  'Graphical
            TabIndex        =   50
            Tag             =   "P"
            Top             =   225
            Width           =   795
         End
         Begin VB.CommandButton bfotos 
            BackColor       =   &H005C31DD&
            Caption         =   "Foto"
            Height          =   420
            Index           =   0
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   49
            Tag             =   "F"
            Top             =   225
            Width           =   795
         End
         Begin VB.Label etdadesaccessori 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   1275
            Left            =   150
            TabIndex        =   54
            Top             =   3225
            Width           =   3600
         End
         Begin VB.Label etrefinplacsa 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2460
            TabIndex        =   53
            Top             =   3285
            Width           =   2445
         End
         Begin VB.Shape Shape1 
            Height          =   2940
            Left            =   75
            Top             =   210
            Width           =   4065
         End
         Begin VB.Image fotoaccessori 
            Height          =   2760
            Left            =   150
            Stretch         =   -1  'True
            Top             =   270
            Width           =   3930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Foto del accessori."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A6A58E&
            Height          =   330
            Left            =   1020
            TabIndex        =   48
            Top             =   1245
            Width           =   2430
         End
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Verd: ID escanejat"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   2745
         TabIndex        =   47
         Top             =   240
         Width           =   1890
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Vermell: ID no escanejat"
         ForeColor       =   &H005C31DD&
         Height          =   255
         Left            =   435
         TabIndex        =   46
         Top             =   240
         Width           =   1890
      End
   End
End
Attribute VB_Name = "FormResumComanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" _
                Alias "SendMessageA" _
                (ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long

Dim vrutaaccessoris As String

Private Sub bfotos_Click(Index As Integer)
bfotos(0).BackColor = &H8000000F: bfotos(1).BackColor = &H8000000F: bfotos(2).BackColor = &H8000000F
   bfotos(Index).BackColor = &H5C31DD
   carregar_fotoactiva
End Sub
Function fotoactiva() As String
  Dim i As Byte
  For i = 0 To 2
   If bfotos(i).BackColor = &H5C31DD Then fotoactiva = bfotos(i).tag
  Next i
End Function
Function numaccessori(vidccessoriutilitzat As Long) As Long
  Dim rst As Recordset
  numaccessori = 0
  Set rst = dbtmpb.OpenRecordset("select * from soldadores_accessorisutilitzats where id=" + atrim(vidccessoriutilitzat))
  If Not rst.EOF Then numaccessori = cadbl(rst!idaccessori)
  Set rst = Nothing
End Function
Sub carregar_fotoactiva()
  Dim vfoto As String
  Dim vnomfitxer As String
  Dim vnumaccessori As Long
  vfoto = fotoactiva
  vnumaccessori = numaccessori(reixa.TextMatrix(reixa.row, 0))
  vnomfitxer = vrutaaccessoris + "FotosAccessoris\" + vfoto + "_" + atrim(cadbl(vnumaccessori)) + ".jpg"
  fotoaccessori.Picture = LoadPicture("")
  fotoaccessori.tag = ""
  If existeix(vnomfitxer) Then
     On Error Resume Next
     fotoaccessori.Picture = LoadPicture(vnomfitxer)
     fotoaccessori.tag = vnomfitxer
  End If
  vnomfitxer = vrutaaccessoris + "FotosAccessoris\F_" + atrim(cadbl(vnumaccessori)) + ".jpg"
  If existeix(vnomfitxer) Then bfotos(0).FontUnderline = True Else bfotos(0).FontUnderline = False
  vnomfitxer = vrutaaccessoris + "FotosAccessoris\P_" + atrim(cadbl(vnumaccessori)) + ".jpg"
  If existeix(vnomfitxer) Then bfotos(1).FontUnderline = True Else bfotos(1).FontUnderline = False
  vnomfitxer = vrutaaccessoris + "FotosAccessoris\U_" + atrim(cadbl(vnumaccessori)) + ".jpg"
  If existeix(vnomfitxer) Then bfotos(2).FontUnderline = True Else bfotos(2).FontUnderline = False
End Sub

Private Sub combosacsocaixes_KeyDown(KeyCode As Integer, Shift As Integer)
 KeyCode = 0
End Sub

Private Sub combosacsocaixes_KeyPress(KeyAscii As Integer)
 KeyAscii = 0
End Sub

Private Sub combosacsocaixes_LostFocus()
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select * from soldadorestot where comanda=" + atrim(cadbl(Form1.comanda)))
  If Not rst.EOF Then rst.Edit: rst!sacsocaixes = combosacsocaixes: rst.Update
  Set rst = Nothing
End Sub

Private Sub Command1_Click()
Dim ComboHandle As Long
  If combosacsocaixes = "" Then
      MsgBox "Escull si utilitzaras SACS o CAIXES.", vbInformation, "ATENCIÓ"
      combosacsocaixes.SetFocus
      ComboHandle = combosacsocaixes.hwnd
      RetVal = SendMessage(ComboHandle, CB_SHOWDROPDOWN, 1, 0)
      Exit Sub
  End If
  Unload Me
End Sub

Sub carregar_dadescomanda(vnumc As Double)
  Dim rstc As Recordset
  Dim rstespesor As Recordset
  Dim i As Byte
  Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
  If Not rstc.EOF Then
    For i = 0 To camps.Count - 1
       If camps(i).DataField <> "" Then camps(i) = atrim(rstc.Fields(camps(i).DataField))
    Next i
  End If
  Set rstespesor = dbtmp.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(rstc!mesuraquantdemanada))
  posar_espesors rstc
  camps(20) = atrim(rstc!tubbaseext) + " " + atrim(rstespesor!descripcio)
  camps(13) = atrim(rstc!soldadora) + "-" + nommaquina(rstc!soldadora)
  camps(14) = tipussellat(atrim(rstc!tipusoldadura))
  
  combosacsocaixes = ""
  Set rstc = dbtmpb.OpenRecordset("select * from soldadorestot where comanda=" + atrim(cadbl(Form1.comanda)))
  If Not rstc.EOF Then combosacsocaixes = atrim(rstc!sacsocaixes)
  
  Set rstc = Nothing
  Set rstespesor = Nothing

End Sub
Sub posar_espesors(rstc As Recordset)
  Dim rstc2 As Recordset
  Dim rstc3 As Recordset
  Set rstc2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rstc!linkcomanda1))
  Set rstc3 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rstc!linkcomanda2))
  camps(7) = atrim(cadbl(rstc!espessor))
  camps(8) = nommesura(cadbl(rstc!mesuraesp))
  If Not rstc2.EOF Then
    camps(9) = atrim(cadbl(rstc2!espessor))
    camps(10) = nommesura(cadbl(rstc2!mesuraesp))
  End If
  If Not rstc3.EOF Then
    camps(11) = atrim(cadbl(rstc3!espessor))
    camps(12) = nommesura(cadbl(rstc3!mesuraesp))
  End If
  Label1(10) = "Total: " + atrim(rstc!espessorsol) + " " + nommesura(cadbl(rstc!mesuraesp))
  Set rstc2 = Nothing
  Set rstc3 = Nothing
End Sub
Function nommesura(vmesura As Integer) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from mesureslineals where codi=" + atrim(vmesura))
  If Not rst.EOF Then
      nommesura = atrim(rst!descripcio)
  End If
  Set rst = Nothing
End Function
Function tipussellat(vtipus As String) As String
  Dim rsttmp2 As Recordset
  Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from tipussoldadura where codi='" + atrim(vtipus) + "'")
  If Not rsttmp2.EOF Then
     tipussellat = atrim(rsttmp2!descripcio)
  End If
  Set rsttmp2 = Nothing
End Function

Function nommaquina(vnummaq As Double) As String
  Dim rsttmp2 As Recordset
  Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from maquines where maquina='S' and codi=" + atrim(vnummaq))
  If Not rsttmp2.EOF Then
     nommaquina = atrim(rsttmp2!descripcio)
  End If
  Set rsttmp2 = Nothing
End Function


Private Sub Command2_Click()
  Dim vidaccessori As Double
  Dim vnomaccessori As String
  Dim vrefinplacsa As String
  Dim vcontrolLot As Boolean
  escullir_accessori vidaccessori, vnomaccessori, vrefinplacsa, vcontrolLot
  If vidaccessori > 0 Then
        dbtmpb.Execute "insert into soldadores_accessorisutilitzats (comanda,nomaccessori,idaccessori,refinplacsa,lottraçabilitat) values (" + atrim(Form1.comanda) + ",'" + atrim(treure_apostruf(vnomaccessori)) + "'," + atrim(vidaccessori) + ",'" + atrim(vrefinplacsa) + "'," + IIf(vcontrolLot, "'-'", "''") + ")"
        poblar_reixa
  End If
End Sub
Sub escullir_accessori(vidaccessori As Double, vnomaccessori As String, vrefinplacsa As String, vcontrolLot As Boolean)
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select numaccessori,familia,subfamilia,descripcio_curta,control_traçabilitat from accessoris_soldadora where '" + atrim(nummaq) + "' IN ([maquinescompatibles]) order by descripcio_curta"
  formseleccio.caption = "Triar Accessori"
  formseleccio.refrescar
  formseleccio.width = 12000
  formseleccio.DBGrid2.Columns(0).width = 0
  formseleccio.DBGrid2.Columns(1).width = 2700
  formseleccio.DBGrid2.Columns(2).width = 2900
  formseleccio.DBGrid2.Columns(3).width = 5000
  formseleccio.DBGrid2.Font.Size = 12
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Show 1
  If seleccioret = 1 Then
    vidaccessori = cadbl(atrim(formseleccio.Data1.Recordset!numaccessori))
    vnomaccessori = atrim(formseleccio.Data1.Recordset!descripcio_curta)
    vrefinplacsa = ""
    If formseleccio.Data1.Recordset!control_traçabilitat = True Then vcontrolLot = True
  End If
  Unload formseleccio
    
End Sub
Private Sub Command29_Click()
  Form1.obrir_DOC_ArxiuSOL
End Sub

Private Sub Command3_Click()
  If MsgBox("Segur que vols eliminar l'accessori " + vbNewLine + reixa.TextMatrix(reixa.row, 1) + "?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  dbtmpb.Execute "delete * from  soldadores_accessorisutilitzats where id=" + atrim(reixa.TextMatrix(reixa.row, 0))
  poblar_reixa
End Sub

Private Sub Command4_Click()
  Dim v As String
  Dim rst As Recordset
  
  Dim vp1 As String
  Dim vp2 As String
  Dim vp3 As String
  
  Set rst = dbtmpb.OpenRecordset("select * from soldadores_accessorisutilitzats where id=" + atrim(reixa.TextMatrix(reixa.row, 0)))
  If Not rst.EOF Then If rst!refinplacsa <> "" Then If MsgBox("Aquest accessori ja té un codi escanejat." + vbNewLine + "Vols tornar a escanejar-lo?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then GoTo fi
  rst.Edit
  demanar_valors_accessori rst!refinplacsa, "Entra el codi de barres de l'accessori utilitzat" + vbNewLine + "ESCRIU [BORRAR] per eliminar-lo"
  demanar_valors_accessori rst!posicio, "Entra el valor de posicio"
  demanar_valors_accessori rst!pressio1, "Entra el valor de pressio1"
  demanar_valors_accessori rst!pressio2, "Entra el valor de pressio2"
  demanar_valors_accessori rst!pressio3, "Entra el valor de pressio3"
  If rst!lottraçabilitat <> "" Then
      demanar_valors_accessori rst!lottraçabilitat, "Entra el Lot de traçabilitat"
      If rst!lottraçabilitat = "" Then rst!lottraçabilitat = "-"
  End If
  While rst!lottraçabilitat = "-"   'si es "-" es que espero un lot si no no ho demano
      demanar_valors_accessori rst!lottraçabilitat, "Entra el Lot de traçabilitat"
      If rst!lottraçabilitat = "" Then rst!lottraçabilitat = "-"
  Wend
  rst.Update
  poblar_reixa
fi:
  Set rst = Nothing
End Sub
Sub demanar_valors_accessori(vcamp As Field, vnomcampvisual As String)
  Dim v As String
  v = InputBox("Entra el valor " + vnomcampvisual + ":", "Valor", atrim(vcamp))
  If StrPtr(v) = 0 Then Exit Sub
  'If UCase(atrim(vp3)) = "BORRAR" Then v = ""
  If vcamp.Type <> 10 Then
       vcamp = cadbl(v)
         Else: vcamp = atrim(v)
  End If
End Sub
Function noexisteixrefinplacsaorepetit(v As String, vidaccessori As Long) As Boolean
  Dim rst As Recordset
  Dim vsql As String
  vsql = "SELECT Accessoris_soldadora.numaccessori, Accessoris_soldadora.maquinescompatibles, Accessoris_soldadora_detall.refinplacsa, Accessoris_soldadora_detall.databaixa FROM Accessoris_soldadora_detall LEFT JOIN Accessoris_soldadora ON Accessoris_soldadora_detall.id_accessori = Accessoris_soldadora.numaccessori WHERE (((Accessoris_soldadora_detall.databaixa) Is Null))"
  Set rst = dbtmp.OpenRecordset(vsql + " and numaccessori=" + atrim(vidaccessori) + " and '" + atrim(nummaq) + "' in ([maquinescompatibles]) and refinplacsa='" + v + "'")
  If rst.EOF Then MsgBox "Aquest accessori no serveix per la màquina escullida o no existeix.", vbCritical, "Error": noexisteixrefinplacsaorepetit = True: GoTo fi
  Set rst = dbtmpb.OpenRecordset("Select * from soldadores_accessorisutilitzats where comanda=" + Form1.comanda + " and refinplacsa='" + atrim(v) + "'")
  If Not rst.EOF Then MsgBox "Aquest codi d'accessori ja s'ha entrat en un altra accessori.", vbCritical, "Error": noexisteixrefinplacsaorepetit = True: GoTo fi
  
fi:
  Set rst = Nothing
End Function

Private Sub Form_Load()
  vrutaaccessoris = "\\ord_copies\DadesProduccio\Arxius Produccio\DadesGenerals\DocumentacióAccessoris\"
  config_reixa
  poblar_reixa
End Sub
Sub poblar_reixa()
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select * from soldadores_accessorisutilitzats where comanda=" + atrim(Form1.comanda))
  If Not rst.EOF Then rst.MoveLast: rst.MoveFirst: reixa.Rows = rst.RecordCount + 1
  While Not rst.EOF
      reixa.TextMatrix(rst.AbsolutePosition + 1, 1) = rst!nomaccessori
      reixa.TextMatrix(rst.AbsolutePosition + 1, 0) = rst!id
      reixa.col = 1
      reixa.row = rst.AbsolutePosition + 1
      If atrim(rst!refinplacsa) <> "" Then
           reixa.CellBackColor = &H6BEBB1
             Else
               reixa.CellBackColor = &H8080FF
      End If
      rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Sub config_reixa()
  reixa.Cols = 2
  reixa.Rows = 2
  reixa.FixedRows = 1
  reixa.FixedCols = 0
  reixa.ColWidth(0) = 0
  reixa.TextMatrix(0, 1) = "Nom de l'accessori"
  reixa.ColWidth(1) = 4200
  reixa.ColAlignment(1) = 1
End Sub

Private Sub reixa_RowColChange()
  carregar_dades_accessori
End Sub
Sub carregar_dades_accessori()
    Dim rst As Recordset
    Dim vid As Double
    etrefinplacsa = "RefInplacsa:"
    carregar_fotoactiva
    vid = reixa.TextMatrix(reixa.row, 0)
    Set rst = dbtmpb.OpenRecordset("select * from soldadores_accessorisutilitzats where id=" + atrim(vid))
    If Not rst.EOF Then
        etrefinplacsa = "RefInplacsa=" + atrim(rst!refinplacsa)
        etdadesaccessori = "Posició: " + atrim(rst!posicio) + vbNewLine
        etdadesaccessori = etdadesaccessori + "Pressió1: " + atrim(rst!pressio1) + vbNewLine
        etdadesaccessori = etdadesaccessori + "Pressió2: " + atrim(rst!pressio2) + vbNewLine
        etdadesaccessori = etdadesaccessori + "Pressió3: " + atrim(rst!pressio3) + vbNewLine
        If atrim(rst!lottraçabilitat) <> "" Then etdadesaccessori = etdadesaccessori + "Lot Traçabilitat: " + atrim(rst!lottraçabilitat) + vbNewLine
    End If
    Set rst = Nothing
End Sub
