VERSION 5.00
Begin VB.Form Formdisposiciomaterialscomanda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disposició dels materials de la comanda"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10605
   Icon            =   "Formdisposiciomaterialscomanda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Segona Firma"
      Height          =   795
      Left            =   1170
      Picture         =   "Formdisposiciomaterialscomanda.frx":00D2
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Guardar els canvis fets"
      Top             =   6780
      Width           =   1545
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   795
      Left            =   75
      Picture         =   "Formdisposiciomaterialscomanda.frx":01CC
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Eliminar la configuració i recarregar de nou."
      Top             =   6780
      Width           =   975
   End
   Begin VB.Frame Framelamanonims 
      BackColor       =   &H00EEE4D7&
      Caption         =   "Laminadores Anónims"
      Height          =   2400
      Left            =   45
      TabIndex        =   9
      Top             =   4200
      Width           =   10320
      Begin VB.Frame Frame3 
         BackColor       =   &H00EEE4D7&
         Height          =   1290
         Index           =   0
         Left            =   30
         TabIndex        =   33
         Top             =   660
         Width           =   465
         Begin VB.CommandButton Command1 
            Height          =   450
            Left            =   60
            Picture         =   "Formdisposiciomaterialscomanda.frx":0546
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   450
            Width           =   345
         End
         Begin VB.Label etad2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "D2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   45
            TabIndex        =   36
            Top             =   960
            Width           =   375
         End
         Begin VB.Label etad1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "D1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   45
            TabIndex        =   35
            Top             =   150
            Width           =   375
         End
      End
      Begin VB.TextBox cmaterial5 
         Height          =   360
         Left            =   510
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1590
         Width           =   5370
      End
      Begin VB.TextBox cmaterial4 
         Height          =   360
         Left            =   510
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   780
         Width           =   5370
      End
      Begin VB.OptionButton bcaraaimprimir4 
         BackColor       =   &H00EEE4D7&
         Caption         =   "Cara 1"
         Height          =   210
         Index           =   0
         Left            =   5940
         TabIndex        =   19
         Top             =   660
         Width           =   4320
      End
      Begin VB.OptionButton bcaraaimprimir4 
         BackColor       =   &H00EEE4D7&
         Caption         =   "Cara 2"
         Height          =   210
         Index           =   1
         Left            =   5940
         TabIndex        =   18
         Top             =   945
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EEE4D7&
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   5835
         TabIndex        =   29
         Top             =   1350
         Width           =   4440
         Begin VB.OptionButton bcaraaimprimir5 
            BackColor       =   &H00EEE4D7&
            Caption         =   "Cara 2"
            Height          =   210
            Index           =   1
            Left            =   60
            TabIndex        =   31
            Top             =   465
            Width           =   4335
         End
         Begin VB.OptionButton bcaraaimprimir5 
            BackColor       =   &H00EEE4D7&
            Caption         =   "Cara 1"
            Height          =   210
            Index           =   0
            Left            =   75
            TabIndex        =   30
            Top             =   180
            Width           =   4320
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Anònim 2:"
         Height          =   255
         Index           =   4
         Left            =   540
         TabIndex        =   25
         Top             =   1320
         Width           =   2115
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cara a Laminar:"
         Height          =   255
         Left            =   6600
         TabIndex        =   24
         Top             =   1185
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Anònim 1:"
         Height          =   255
         Index           =   3
         Left            =   510
         TabIndex        =   22
         Top             =   510
         Width           =   2115
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Cara a Laminar:"
         Height          =   255
         Left            =   6600
         TabIndex        =   21
         Top             =   405
         Width           =   1200
      End
   End
   Begin VB.CommandButton bguardarcanvis 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Primera firma"
      Height          =   945
      Left            =   8805
      Picture         =   "Formdisposiciomaterialscomanda.frx":0AD0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Guardar els canvis fets"
      Top             =   6675
      Width           =   1545
   End
   Begin VB.Frame frameLamImp 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Laminadores i Imprès"
      Height          =   2400
      Left            =   45
      TabIndex        =   1
      Top             =   1695
      Width           =   10320
      Begin VB.Frame Frame3 
         BackColor       =   &H00EEE4D7&
         Height          =   1425
         Index           =   1
         Left            =   90
         TabIndex        =   37
         Top             =   540
         Width           =   465
         Begin VB.CommandButton Command2 
            Height          =   450
            Left            =   60
            Picture         =   "Formdisposiciomaterialscomanda.frx":0DA6
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   510
            Width           =   345
         End
         Begin VB.Label etid1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "D1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   45
            TabIndex        =   40
            Top             =   150
            Width           =   375
         End
         Begin VB.Label etid2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "D2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   45
            TabIndex        =   39
            Top             =   1095
            Width           =   375
         End
      End
      Begin VB.TextBox cmaterial3 
         Height          =   360
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1590
         Width           =   5310
      End
      Begin VB.TextBox cmaterial2 
         Height          =   360
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   615
         Width           =   5310
      End
      Begin VB.OptionButton bcaraaimprimir2 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Cara 1"
         Height          =   210
         Index           =   0
         Left            =   6060
         TabIndex        =   11
         Top             =   555
         Width           =   4215
      End
      Begin VB.OptionButton bcaraaimprimir2 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Cara 2"
         Height          =   210
         Index           =   1
         Left            =   6045
         TabIndex        =   10
         Top             =   855
         Width           =   4230
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAD9CE&
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   5955
         TabIndex        =   26
         Top             =   1275
         Width           =   4320
         Begin VB.OptionButton bcaraaimprimir3 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Cara 2"
            Height          =   210
            Index           =   1
            Left            =   45
            TabIndex        =   28
            Top             =   510
            Width           =   4230
         End
         Begin VB.OptionButton bcaraaimprimir3 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Cara 1"
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   27
            Top             =   255
            Width           =   4200
         End
      End
      Begin VB.Label Label7 
         BackColor       =   &H005C31DD&
         Caption         =   "   La cara vermella es la cara impresa..."
         Height          =   255
         Left            =   3750
         TabIndex        =   32
         Top             =   90
         Width           =   3150
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material Anònim o Laminat:"
         Height          =   255
         Index           =   2
         Left            =   645
         TabIndex        =   17
         Top             =   1320
         Width           =   2115
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cara a Laminar:"
         Height          =   255
         Left            =   6570
         TabIndex        =   16
         Top             =   1110
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material imprès:"
         Height          =   255
         Index           =   1
         Left            =   690
         TabIndex        =   14
         Top             =   345
         Width           =   2115
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cara a Laminar:"
         Height          =   255
         Left            =   6570
         TabIndex        =   13
         Top             =   330
         Width           =   1200
      End
   End
   Begin VB.Frame frameimp 
      BackColor       =   &H00FDDECE&
      Caption         =   "Impresores"
      Height          =   1155
      Left            =   75
      TabIndex        =   0
      Top             =   330
      Width           =   10320
      Begin VB.OptionButton bcaraaimprimir 
         BackColor       =   &H00FDDECE&
         Caption         =   "Cara 2"
         Height          =   210
         Index           =   1
         Left            =   5970
         TabIndex        =   7
         Top             =   750
         Width           =   4305
      End
      Begin VB.OptionButton bcaraaimprimir 
         BackColor       =   &H00FDDECE&
         Caption         =   "Cara 1"
         Height          =   210
         Index           =   0
         Left            =   5970
         TabIndex        =   6
         Top             =   465
         Width           =   4290
      End
      Begin VB.TextBox cmaterial1 
         Height          =   360
         Left            =   345
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   555
         Width           =   5580
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cara a imprimir:"
         Height          =   255
         Left            =   6450
         TabIndex        =   5
         Top             =   135
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material imprès:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   285
         Width           =   2115
      End
   End
   Begin VB.Label etverificador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      Height          =   270
      Left            =   1020
      TabIndex        =   44
      Top             =   7605
      Width           =   1965
   End
   Begin VB.Label etcreador 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      Height          =   270
      Left            =   8655
      TabIndex        =   42
      Top             =   7620
      Width           =   1965
   End
   Begin VB.Label etrefinplacsa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disposició dels materials a les seccions.  Ref: INP01I4565"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   345
      TabIndex        =   2
      Top             =   75
      Width           =   9810
   End
End
Attribute VB_Name = "Formdisposiciomaterialscomanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vultimaref As String
Dim vcarregantdades As Boolean

Private Sub bcaraaimprimir_Click(Index As Integer)
posarcolorcaraimpresa
End Sub

Private Sub bcaraaimprimir4_Click(Index As Integer)
  If Not vcarregantdades Then
      guardar_dades
      carregar_dades atrim(etrefinplacsa.Tag)
  End If
End Sub

Private Sub bcaraaimprimir5_Click(Index As Integer)
 If Not vcarregantdades Then
      guardar_dades
      carregar_dades atrim(etrefinplacsa.Tag)
  End If
End Sub

Private Sub bguardarcanvis_Click()
   If faltencampsperemplenar Then MsgBox "Falten camps per emplenar, no pots tancar sense possar-los.", vbCritical, "Error": Exit Sub
   guardar_dades nomordinador
   Unload Me
End Sub
Function faltencampsperemplenar() As Boolean
  If bcaraaimprimir(0).Value = 0 And bcaraaimprimir(1).Value = 0 Then faltencampsperemplenar = True
  If frameLamImp.Visible Then
       If bcaraaimprimir2(0).Value = 0 And bcaraaimprimir2(1).Value = 0 Then faltencampsperemplenar = True
       If bcaraaimprimir3(0).Value = 0 And bcaraaimprimir3(1).Value = 0 Then faltencampsperemplenar = True
  End If
  If Framelamanonims.Visible Then
      If bcaraaimprimir4(0).Value = 0 And bcaraaimprimir4(1).Value = 0 Then faltencampsperemplenar = True
      If bcaraaimprimir4(0).Value = 0 And bcaraaimprimir4(1).Value = 0 Then faltencampsperemplenar = True
  End If
End Function

Private Sub Command1_Click()
  If etad1 = "D1" Then
       etad1 = "D2": etad2 = "D1"
         Else: etad1 = "D1": etad2 = "D2"
  End If
  guardar_dades
End Sub

Private Sub Command2_Click()
  Dim v As Byte
  If etid1 = "D1" Then
       etid1 = "D2": etid2 = "D1"
         Else: etid1 = "D1": etid2 = "D2"
  End If
  guardar_dades
End Sub

Private Sub Command3_Click()
  Dim vultimaref As String
  If MsgBox("Segur que vols eliminar aquesta configuració i carregar valors per defecte?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  vultimaref = etrefinplacsa.Tag
  dbtmp.Execute "delete * from referencies_disposiciomaterials where refinplacsa='" + vultimaref + "'"
  wait 1
  carregar_dades vultimaref
End Sub

Private Sub etrevisor_Click()

End Sub

Private Sub Command4_Click()
   If faltencampsperemplenar Then MsgBox "Falten camps per emplenar, no pots tancar sense possar-los.", vbCritical, "Error": Exit Sub
   If etcreador = "" Then Exit Sub
   If etcreador = nomordinador Then MsgBox "No pot fer la segona firma la mateixa persona que la primera.", vbCritical, "Error": Exit Sub
   guardar_dades , nomordinador
   Unload Me
End Sub

Private Sub Form_Activate()
  
  ' etrefinplaca.tag hi posso el valor de la referencia d'inplacsa que interesa
  etrefinplacsa = "Disposició dels materials a les seccions.  Ref: " + etrefinplacsa.Tag
  If vultimaref <> etrefinplacsa.Tag Then
      vultimaref = etrefinplacsa.Tag
      carregar_dades vultimaref
  End If
  
End Sub
Function esanonim(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("SELECT productes.ruta FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi WHERE (((comandes.comanda)=" + atrim(vnumc) + "));")
   If Not rst.EOF Then
        If InStr(1, rst!ruta, "L") = 0 And InStr(1, rst!ruta, "I") = 0 Then esanonim = True
   End If
   Set rst = Nothing
End Function
Sub crear_dadesnoves(vrefinplacsa As String)
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim rstc As Recordset
   Set rst = dbtmp.OpenRecordset("select * from referencies_disposiciomaterials")
   dbtmp.Execute "delete * from referencies_disposiciomaterials where refinplacsa='" + vrefinplacsa + "'"
   rst.FindFirst "refinplacsa='" + atrim(vrefinplacsa) + "'"
   If Not rst.NoMatch Then Exit Sub
   Set rst2 = dbtmp.OpenRecordset("select * from comandes_extres where comanda<>0 and refinplacsa='" + atrim(etrefinplacsa.Tag) + "' order by comanda asc")
   If rst2.EOF Then Exit Sub
   Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rst2!comanda))
   If rstc.EOF Then Exit Sub
   Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(rstc!comanda) + " or comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + " and comanda<>0 order by comanda asc")
   Set rst2 = dbtmp.OpenRecordset("SELECT materials.descripcio, comandes.* FROM comandes LEFT JOIN materials ON comandes.materialex = materials.codi where comanda=" + atrim(rst2!comanda))
   If Not rst2.EOF Then
       If esanonim(rst2!comanda) Then vrefinplacsa = "referenciaanonim": GoTo fi
       rst.AddNew
       rst!refinplacsa = vrefinplacsa
       rst!material1 = cadbl(rst2!materialex)
       rst!nommat1 = Trim(rst2!descripcio)
       If rstc!linkcomanda1 > 0 Then
            Set rst2 = dbtmp.OpenRecordset("SELECT materials.descripcio, comandes.* FROM comandes LEFT JOIN materials ON comandes.materialex = materials.codi where comanda=" + atrim(rstc!linkcomanda1))
            If Not rst2.EOF Then
                rst!material2 = cadbl(rst2!materialex)
                rst!nommat2 = Trim(rst2!descripcio)
            End If
       End If
       If rstc!linkcomanda2 > 0 Then
            Set rst2 = dbtmp.OpenRecordset("SELECT materials.descripcio, comandes.* FROM comandes LEFT JOIN materials ON comandes.materialex = materials.codi where comanda=" + atrim(rstc!linkcomanda2))
            If Not rst2.EOF Then
                rst!material3 = cadbl(rst2!materialex)
                rst!nommat3 = Trim(rst2!descripcio)
            End If
       End If
       rst!pc2_desbmat1 = 1
       rst!pc_desbmat1 = 1
       rst!nomcreador = UCase(nomordinador)
       rst.Update
      
   End If
fi:
   Set rst = Nothing
   Set rst2 = Nothing
   Set rstc = Nothing
End Sub
Sub posarcolorcaraimpresa()
   If bcaraaimprimir(0).Value = True Then
          bcaraaimprimir2(0).BackColor = &H5C31DD
        Else: bcaraaimprimir2(0).BackColor = frameLamImp.BackColor
   End If
   If bcaraaimprimir(1).Value = True Then
          bcaraaimprimir2(1).BackColor = &H5C31DD
        Else: bcaraaimprimir2(1).BackColor = frameLamImp.BackColor
   End If
   
    If bcaraaimprimir(0).Value = True Then
          bcaraaimprimir(0).BackColor = &H5C31DD
        Else: bcaraaimprimir(0).BackColor = frameimp.BackColor
   End If
   If bcaraaimprimir(1).Value = True Then
          bcaraaimprimir(1).BackColor = &H5C31DD
        Else: bcaraaimprimir(1).BackColor = frameimp.BackColor
   End If
End Sub
Sub carregar_dades(vref As String)
   Dim rst As Recordset
   Dim vx As Double
   Dim vrefinplacsa As String
   vcarregantdades = True
   vrefinplacsa = atrim(vref)
   Set rst = dbtmp.OpenRecordset("Select * from referencies_disposiciomaterials")
   rst.FindFirst "refinplacsa='" + atrim(vrefinplacsa) + "'"
    cmaterial1.Tag = "": cmaterial1 = "": nommat1 = "": bcaraaimprimir(0).Value = False: bcaraaimprimir(1).Value = False
   If rst.NoMatch Then
dadesnoves:
        crear_dadesnoves vrefinplacsa
        If vrefinplacsa = "referenciaanonim" Then GoTo fi
        Set rst = dbtmp.OpenRecordset("Select * from referencies_disposiciomaterials")
        rst.FindFirst "refinplacsa='" + atrim(vrefinplacsa) + "'"
        If rst.NoMatch Then GoTo fi
         Else: If rst!material1 = 0 Then GoTo dadesnoves
   End If
   bcaraaimprimir(0).Value = False: bcaraaimprimir(1).Value = False
   bcaraaimprimir2(0).Value = False: bcaraaimprimir2(1).Value = False
   bcaraaimprimir3(0).Value = False: bcaraaimprimir3(1).Value = False
   bcaraaimprimir4(0).Value = False: bcaraaimprimir4(1).Value = False
   bcaraaimprimir5(0).Value = False: bcaraaimprimir5(1).Value = False
     'material 1
   cmaterial1.Tag = rst!material1
   cmaterial1 = atrim(rst!nommat1)
   If rst!caraimpresio = 1 Then bcaraaimprimir(0).Value = True
   If rst!caraimpresio = 2 Then bcaraaimprimir(1).Value = True
   posarnomscares cadbl(rst!material1), bcaraaimprimir()
   
      'material impres a laminadora
   If cadbl(rst!material2) <> 0 Then
        cmaterial2.Tag = rst!material1
        cmaterial2 = atrim(rst!nommat1)
        If rst!pc_caralaminar1 = 1 Then bcaraaimprimir2(0).Value = True
        If rst!pc_caralaminar1 = 2 Then bcaraaimprimir2(1).Value = True
        posarnomscares cadbl(rst!material1), bcaraaimprimir2()
        posarcolorcaraimpresa
     'material 2
        frameLamImp.Visible = True
        cmaterial3.Tag = rst!material2
        cmaterial3 = atrim(rst!nommat2)
        cmateriallaminat = cmaterial3
        If cadbl(rst!material3) <> 0 Then
               cmaterial3 = cmaterial3 + "+" + atrim(rst!nommat3)
               cmateriallaminat = cmaterial3
               posarnomscares cadbl(rst!material2), bcaraaimprimir3(), cadbl(rst!material3), IIf(rst!pc2_caralaminar1 = 1, 2, 1), IIf(rst!pc2_caralaminar2 = 1, 2, 1)
             Else: posarnomscares cadbl(rst!material2), bcaraaimprimir3()
        End If
        If rst!pc_caralaminar2 = 1 Then bcaraaimprimir3(0).Value = True
        If rst!pc_caralaminar2 = 2 Then bcaraaimprimir3(1).Value = True
          Else: frameLamImp.Visible = False
   End If
     
      'material 3
   If cadbl(rst!material3) <> 0 Then
        Framelamanonims.Enabled = True
        cmaterial4.Tag = rst!material2
        cmaterial4 = atrim(rst!nommat2)
        If rst!pc2_caralaminar1 = 1 Then bcaraaimprimir4(0).Value = True
        If rst!pc2_caralaminar1 = 2 Then bcaraaimprimir4(1).Value = True
        posarnomscares cadbl(rst!material2), bcaraaimprimir4()
        
        cmaterial5.Tag = rst!material3
        cmaterial5 = atrim(rst!nommat3)
        If rst!pc2_caralaminar2 = 1 Then bcaraaimprimir5(0).Value = True
        If rst!pc2_caralaminar2 = 2 Then bcaraaimprimir5(1).Value = True
        posarnomscares cadbl(rst!material3), bcaraaimprimir5()
        If Framelamanonims.Top > frameLamImp.Top Then
            Framelamanonims.Visible = True
            vx = Framelamanonims.Top
            Framelamanonims.Top = frameLamImp.Top
            frameLamImp.Top = vx
        End If
      Else: Framelamanonims.Visible = False
   End If
   If rst!pc2_desbmat1 = 1 Then
       etad1 = "D1": etad2 = "D2"
         Else: etad1 = "D2": etad2 = "D1"
  End If
  If rst!pc_desbmat1 = 1 Then
       etid1 = "D1": etid2 = "D2"
         Else: etid1 = "D2": etid2 = "D1"
  End If
  etcreador = atrim(rst!nomcreador)
  etverificador = atrim(rst!nomverificador)
fi:
   Set rst = Nothing
   vcarregantdades = False
   If vrefinplacsa = "referenciaanonim" Then Unload Me
End Sub
Sub posarnomscares(vcodimat As Double, vcontrol As Object, Optional vcodimat2 As Double, Optional vc1 As Byte, Optional vc2 As Byte)
   Dim rst As Recordset
   Dim rstcares As Recordset
   If vc1 = 0 Then vc1 = 1: vc2 = 2
   vcontrol(0).Caption = "Cara1"
   vcontrol(1).Caption = "Cara2"
   Set rstcares = dbtmp.OpenRecordset("select * from tractamentcares")
   Set rst = dbtmp.OpenRecordset("select codidescmatcara1,codidescmatcara2 from materials where codi=" + atrim(vcodimat))
   If Not rst.EOF Then
       rstcares.FindFirst "codi=" + atrim(cadbl(rst.Fields("codidescmatcara" + atrim(vc1))))
       If Not rstcares.NoMatch Then vcontrol(0).Caption = "C1: " + atrim(rstcares!descripcio)
       If cadbl(vcodimat2) > 0 Then Set rst = dbtmp.OpenRecordset("select codidescmatcara1,codidescmatcara2 from materials where codi=" + atrim(vcodimat2))
       If Not rst.EOF Then
            rstcares.FindFirst "codi=" + atrim(cadbl(rst.Fields("codidescmatcara" + atrim(vc2))))
            If Not rstcares.NoMatch Then vcontrol(1).Caption = "C2: " + atrim(rstcares!descripcio)
       End If
   End If
   Set rst = Nothing
   Set rst2 = Nothing
End Sub
Sub guardar_dades(Optional vfirma As String, Optional vsegonafirma As String)
   Dim rst As Recordset
   Dim vrefinplacsa As String
   
   vrefinplacsa = atrim(etrefinplacsa.Tag)
   'mirar si ja hi ha la referencia creada a la taula el camp es etrefinplaca.tag
     ' si no hi es crear-la
   Set rst = dbtmp.OpenRecordset("Select * from referencies_disposiciomaterials")
   rst.FindFirst "refinplacsa='" + atrim(vrefinplacsa) + "'"
   If rst.NoMatch Then MsgBox "Error no he trobat la referencia": Exit Sub
   rst.Edit
   
   rst!refinplacsa = vrefinplacsa
     'mat 1
  ' rst!material1 = cadbl(cmaterial1.Tag)
   'rst!nommat1 = cmaterial1
   rst!caraimpresio = IIf(bcaraaimprimir(0), 1, IIf(bcaraaimprimir(1), 2, 0))
   
      'mat 1 -lam
   rst!pc_caralaminar1 = IIf(bcaraaimprimir2(0), 1, IIf(bcaraaimprimir2(1), 2, 0))
   
      'mat2  que es el materail3
   'rst!material2 = cadbl(cmaterial3.Tag)
   'rst!nommat2 = cmaterial3
   rst!pc_caralaminar2 = IIf(bcaraaimprimir3(0), 1, IIf(bcaraaimprimir3(1), 2, 0))
   rst!pc2_caralaminar1 = IIf(bcaraaimprimir4(0), 1, IIf(bcaraaimprimir4(1), 2, 0))
   rst!pc2_caralaminar2 = IIf(bcaraaimprimir5(0), 1, IIf(bcaraaimprimir5(1), 2, 0))
   rst!materiallaminat = cmateriallaminat
   rst!pc_desbmat1 = IIf(etid1 = "D1", 1, 2)
   rst!pc2_desbmat1 = IIf(etad1 = "D1", 1, 2)
   If vfirma <> "" Then rst!nomcreador = vfirma
   If vsegonafirma <> "" Then rst!nomverificador = vsegonafirma
   rst.Update
   's'ha de mirar lo d'enviar per email a qui correspongui
   Set rst = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   vultimaref = ""
End Sub

