VERSION 5.00
Begin VB.Form Formdesbobinadors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment de Desbobinadors"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11955
   Icon            =   "Formdesbobinadors.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameescaneigbobina 
      BackColor       =   &H005C31DD&
      Height          =   2460
      Left            =   645
      TabIndex        =   41
      Top             =   8205
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton bacceptarbobina 
         Height          =   450
         Left            =   3570
         Picture         =   "Formdesbobinadors.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Manual"
         Height          =   450
         Left            =   1890
         TabIndex        =   46
         Top             =   1920
         Width           =   960
      End
      Begin VB.TextBox cbobina2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1875
         TabIndex        =   43
         Top             =   1200
         Width           =   2085
      End
      Begin VB.TextBox cbobina1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1890
         TabIndex        =   42
         Top             =   570
         Width           =   2085
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escaneig paper frontal i tubo de la bobina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   165
         TabIndex        =   47
         Top             =   195
         Width           =   4335
      End
      Begin VB.Image Image2 
         Height          =   435
         Left            =   3990
         Picture         =   "Formdesbobinadors.frx":1254
         Stretch         =   -1  'True
         Top             =   1290
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   435
         Left            =   4020
         Picture         =   "Formdesbobinadors.frx":1C57
         Stretch         =   -1  'True
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Repetició Bobina"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   300
         TabIndex        =   45
         Top             =   1140
         Width           =   1530
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Bobina: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   315
         TabIndex        =   44
         Top             =   600
         Width           =   1245
      End
   End
   Begin VB.Frame Framepassword 
      BackColor       =   &H005C31DD&
      Height          =   5355
      Left            =   7305
      TabIndex        =   21
      Top             =   8145
      Visible         =   0   'False
      Width           =   4260
      Begin VB.CommandButton Command2 
         Height          =   930
         Left            =   3330
         Picture         =   "Formdesbobinadors.frx":265A
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   4365
         Width           =   765
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   7
         Left            =   1200
         TabIndex        =   31
         Top             =   2490
         Width           =   1035
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   6
         Left            =   165
         TabIndex        =   30
         Top             =   2490
         Width           =   1035
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   8
         Left            =   2235
         TabIndex        =   34
         Top             =   2490
         Width           =   1035
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   5
         Left            =   2235
         TabIndex        =   33
         Top             =   1530
         Width           =   1035
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   2
         Left            =   2235
         TabIndex        =   32
         Top             =   570
         Width           =   1035
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   4
         Left            =   1200
         TabIndex        =   29
         Top             =   1530
         Width           =   1035
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   3
         Left            =   165
         TabIndex        =   28
         Top             =   1530
         Width           =   1035
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   1
         Left            =   1200
         TabIndex        =   27
         Top             =   570
         Width           =   1035
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   0
         Left            =   165
         TabIndex        =   26
         Top             =   570
         Width           =   1035
      End
      Begin VB.TextBox cpassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   135
         TabIndex        =   25
         Top             =   4350
         Width           =   3150
      End
      Begin VB.CommandButton cbotonum 
         BackColor       =   &H00C0FFC0&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3765
         Index           =   10
         Left            =   3315
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   555
         Width           =   810
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   9
         Left            =   150
         TabIndex        =   23
         Top             =   3435
         Width           =   2085
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Index           =   11
         Left            =   2235
         TabIndex        =   22
         Top             =   3420
         Width           =   1035
      End
      Begin VB.Label etkeypad 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   165
         TabIndex        =   35
         Top             =   150
         Width           =   3915
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   45
      Top             =   4875
   End
   Begin VB.Frame Framependentverificar 
      Caption         =   "Verificar "
      Height          =   1380
      Left            =   2640
      TabIndex        =   15
      Top             =   6765
      Visible         =   0   'False
      Width           =   6675
      Begin VB.ListBox llistapendentsverificar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   3435
         TabIndex        =   18
         Top             =   330
         Width           =   3000
      End
      Begin VB.TextBox cescanverify 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   75
         TabIndex        =   16
         Top             =   495
         Width           =   1710
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendents"
         Height          =   270
         Left            =   3585
         TabIndex        =   19
         Top             =   105
         Width           =   1395
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   " Escaneja Bobina"
         Height          =   270
         Left            =   90
         TabIndex        =   17
         Top             =   270
         Width           =   1395
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Verificar"
      Height          =   1425
      Left            =   195
      Picture         =   "Formdesbobinadors.frx":2B64
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6720
      Width           =   2085
   End
   Begin VB.TextBox nomoperari 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Escull Operari"
      Top             =   120
      Width           =   6075
   End
   Begin VB.CommandButton bsortir 
      Caption         =   "Ok"
      Height          =   1425
      Left            =   9585
      Picture         =   "Formdesbobinadors.frx":311C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6750
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EAD9CE&
      Height          =   6075
      Left            =   195
      TabIndex        =   0
      Top             =   630
      Width           =   11385
      Begin VB.CommandButton Command5 
         Height          =   630
         Left            =   9225
         Picture         =   "Formdesbobinadors.frx":33F2
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Escanejar etiqueta."
         Top             =   4995
         Width           =   780
      End
      Begin VB.CommandButton Command4 
         Height          =   630
         Left            =   2925
         Picture         =   "Formdesbobinadors.frx":3751
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Escanejar etiqueta."
         Top             =   4920
         Width           =   780
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00F1B75F&
         Caption         =   "Desbobinador 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Left            =   6690
         Picture         =   "Formdesbobinadors.frx":3AB0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1605
         Width           =   3060
      End
      Begin VB.CommandButton bdesb1 
         BackColor       =   &H00EEE4D7&
         Caption         =   "Desbobinador 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Left            =   570
         Picture         =   "Formdesbobinadors.frx":436D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1590
         Width           =   3060
      End
      Begin VB.Image fotoetiqueta2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1245
         Left            =   7365
         Stretch         =   -1  'True
         Top             =   4755
         Width           =   1845
      End
      Begin VB.Image fotoetiqueta1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1245
         Left            =   1050
         Stretch         =   -1  'True
         Top             =   4740
         Width           =   1845
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Foto Etiqueta 2"
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
         Left            =   7395
         TabIndex        =   38
         Top             =   4470
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Foto Etiqueta 1"
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
         Left            =   1140
         TabIndex        =   37
         Top             =   4455
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Bobina: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   6720
         TabIndex        =   14
         Top             =   1065
         Width           =   1230
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Bobina: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   645
         TabIndex        =   13
         Top             =   1065
         Width           =   1245
      End
      Begin VB.Label etcomanda2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "999999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7710
         TabIndex        =   9
         Top             =   435
         Width           =   1920
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comanda:"
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
         Left            =   6675
         TabIndex        =   8
         Top             =   555
         Width           =   1290
      End
      Begin VB.Label etcomanda1 
         BackStyle       =   0  'Transparent
         Caption         =   "999999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1965
         TabIndex        =   7
         Top             =   330
         Width           =   1845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Comanda:"
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
         Left            =   645
         TabIndex        =   6
         Top             =   405
         Width           =   1290
      End
      Begin VB.Label etbob2 
         BackStyle       =   0  'Transparent
         Caption         =   "99999/99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   7995
         TabIndex        =   4
         Top             =   975
         Width           =   1905
      End
      Begin VB.Label etbob1 
         BackStyle       =   0  'Transparent
         Caption         =   "99999/99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   1950
         TabIndex        =   3
         Top             =   975
         Width           =   2055
      End
   End
   Begin VB.Label etstatus 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   20
      Top             =   8085
      Visible         =   0   'False
      Width           =   5490
   End
   Begin VB.Label Label4 
      Caption         =   "Operari:"
      Height          =   300
      Left            =   105
      TabIndex        =   11
      Top             =   165
      Width           =   645
   End
End
Attribute VB_Name = "Formdesbobinadors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bacceptarbobina_Click()
      If Len(cbobina1) > 3 And Len(cbobina2) > 3 Then
         If bacceptarbobina.tag <> "T" Then
            MsgBox "No s'ha escanejat la bobina del canutu.", vbCritical, "Error"
             Else
                If cbobina1 <> cbobina2 Then MsgBox "Les etiquetes de bobines escanejades no corresponen", vbCritical, "Error": cbobina1 = "": cbobina2 = ""
         End If
      End If
      frameescaneigbobina.visible = False
End Sub

Private Sub bdesb1_Click()
    afegir_bobinadesbobinador 1
End Sub
Function labobinanoshaafegitalprograma(vnumdeb As Byte) As Boolean
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim vwere As String
   Set rst = dbtmpb.OpenRecordset("select * from bobinesdesbobinadors where numdesbobinador=" + atrim(vnumdeb) + " and maquina=" + atrim(nummaq) + " order by data desc")
   If Not rst.EOF Then
     labobinanoshaafegitalprograma = True
     Set rstc = dbtmpb.OpenRecordset("select * from impressores where comanda=" + atrim(rst!comanda) + " and (paletprova=" + atrim(rst!palet) + " and bobinaprova=" + atrim(rst!bobina) + ") or (paletprova2=" + atrim(rst!palet) + " and bobinaprova2=" + atrim(rst!bobina) + ")")
     If Not rstc.EOF Then labobinanoshaafegitalprograma = False: GoTo fi
     vwere = "(palet=" + atrim(rst!palet) + " and bobina=" + atrim(rst!bobina) + ") and (impressores.comanda=" + atrim(rst!comanda) + ")"
     Set rstc = dbtmpb.OpenRecordset("SELECT impressores.comanda, bobinesentimp.palet, bobinesentimp.bobina FROM (bobinesentimp INNER JOIN bobinesimp ON bobinesentimp.id = bobinesimp.Id) INNER JOIN impressores ON bobinesimp.controlid = impressores.Id WHERE " + vwere)
     If Not rstc.EOF Then labobinanoshaafegitalprograma = False: GoTo fi
   End If
fi:
   
End Function
Sub afegir_bobinadesbobinador(vnumdesb As Byte)
   Dim vbob As Double
   Dim vpalet As Double
   Dim vbobina As Double
   Dim vnumc As String
   Dim vvalors As String
   Dim vresp As String
   'If labobinanoshaafegitalprograma(vnumdesb) Then MsgBox "Aquesta bobina no està afegida a la comanda, no pots carregar el desbobinador fins que l'hagis afegit al programa.", vbCritical, "Error": Exit Sub
   
   Set rst = dbtmpb.OpenRecordset("select comanda from bobinesdesbobinadors where maquina=" + atrim(nummaq) + " order by data desc")
   If Not rst.EOF Then vnumc = rst!comanda
   
   'vresp = InputBox("Escaneja la bobina que vols afegir al Desbobinador " + atrim(vnumdesb), "Entrada bobina")
   vresp = ensenyarescaneigbobina
   If atrim(vresp) = "" Then Exit Sub
   
   'vnumc = InputBox("Entra la comanda relacionada amb aquesta bobina.", "Comanda", vnumc)
   vnumc = ensenyarkeypad("Entra la comanda relacionada:")
   
   If cadbl(vnumc) = 0 Then Exit Sub
   vpalet = cadbl(Mid(" " + vresp, 1, InStr(1, vresp + "  ", "/")))
   vbob = cadbl(Mid(vresp, InStr(1, vresp + "  ", "/") + 1))
   If Not comprovar_bobina(vpalet, vbob, cadbl(vnumc)) Then Exit Sub
   vvalors = "#" + Format(Now, "mm/dd/yy hh:nn:ss") + "#," + atrim(vpalet) + "," + atrim(vbob) + "," + atrim(vnumc) + "," + atrim(nummaq) + "," + atrim(vnumdesb)
   dbtmpb.Execute "insert into bobinesdesbobinadors (data,palet,bobina,comanda,maquina,numdesbobinador) values (" + vvalors + ")"
   carregar_desbobinadors
   
End Sub


Function comprovar_bobina(vpalet As Double, vbobina As Double, vnumc As Double)
  Dim rst As Recordset
  Dim vresp As String
  Dim vmatexacte As Boolean
  Dim vstockopacking As String
  Dim vgrupmaterialcompatible As Double
  Dim vgrup As Double
  Dim vtexte As String
  
  comprovar_bobina = True
  Set rst = dbtmpb.OpenRecordset("select proximaseccio from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then If rst!proximaseccio <> "I" Then MsgBox "Aquesta comanda no està apunt per imprimir.", vbCritical, "Error": comprovar_bobina = False: GoTo fi
  valorsdajust vnumc, vgrup, vtexte, vgrupmaterialcompatible
  If vpalet > 0 And vbobina > 0 Then
    obrestocks
    Set rst = dbtmp.OpenRecordset("SELECT comandes_extres.assignarstock as estoc, materialexacte frOM comandes_extres WHERE comanda=" + atrim(cadbl(vnumc)) + ";")
    If Not rst.EOF Then
      If rst!estoc Then vstockopacking = "E"
      If rst!materialexacte Then vmatexacte = True
    End If
  
    If vmatexacte Then
      If Not comprovar_materialexacte(vpalet, cadbl(form1.etmaterialexacte.tag)) Then
         If MsgBox("Aquest material no es exactament el que demana el client." + Chr(10) + "VOLS CONTINUAR IGUALEMENT?", vbCritical + vbDefaultButton2 + vbYesNo, "Error") = vbNo Then
           comprovar_bobina = False
           GoTo fi
         End If
       End If
    End If
    If vstockopacking <> "E" Then
       vresp = comprovarsieselmateixmaterial(vpalet, vbobina, cadbl(vnumc), vgrup, vgrupmaterialcompatible)
       If InStr(1, vresp, "#materialerror") > 0 Then
          If MsgBox("Aquest material no coincideix amb el material de la comanda." + Chr(10) + "VOLS CONTINUAR IGUALMENT?", vbCritical + vbDefaultButton2 + vbYesNo, "A T E N C I Ó") = vbNo Then
            comprovar_bobina = False
            GoTo fi
          End If
       End If
    End If
    inssql = "SELECT CDbl([comanda]) AS Expr1, Parcials.idpalet, Parcials.idbobina,parcials.orcomassignacio  From Parcials WHERE (((CDbl([orcomassignacio])<10000 and cdbl([orcomassignacio])>2000)) and idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbobina) + ");"
    Set rst = dbstocks.OpenRecordset(inssql)

    If Not rst.EOF Then
          If vstockopacking = "P" Then
             If MsgBox("Aquesta bobina es d'ESTOC en una comanda de PACKING-LIST." + Chr(10) + "VOLS CONTINUAR IGUALMENT?", vbCritical + vbDefaultButton2 + vbYesNo, "A T E N C I Ó") = vbYes Then
                comprovar_bobina = False
             End If
          End If
    End If
  End If
fi:
  Set rst = Nothing
End Function

Sub comprovarverificaciobobina()
    Dim rst As Recordset
    Dim vresp As String
    vresp = InputBox("Escaneja el codi petit del canutu per verificar que son la mateixa.", "Verificació de canutu")
    If atrim(vresp) = "" Then Exit Sub
    If vresp <> atrim(cescanverify) Then MsgBox "Els dos codis escanejats no coincideixen, revisa bé les etiquetes sisplau.", vbCritical, "Error": GoTo fi
    Set rst = dbtmpb.OpenRecordset("select * from bobinesdesbobinadors where trim([palet])+'/'+trim([bobina])='" + atrim(cescanverify) + "' and verificada=false and maquina=" + atrim(nummaq) + " order by data desc")
    If Not rst.EOF Then rst.Edit: rst!verificada = True: rst.Update
    etstatus.caption = "Bobina verificada correctament..."
    etstatus.visible = True
    Timer1.Enabled = True
fi:
    Set rst = Nothing
End Sub

Private Sub cbobina1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If UCase(Mid(cbobina1 + "  ", 1, 1)) = "T" Then
       If bacceptarbobina.tag = "T" Then
           MsgBox "Ja has escanejat la bobina del tubo has de fer la de l'etiqueta.", vbInformation, "Error"
           cbobina1 = ""
           GoTo fi
             Else: bacceptarbobina.tag = "T": cbobina1 = Mid(cbobina1, 2)
       End If
    End If
    cbobina2.SetFocus
  End If
fi:
End Sub

Private Sub cbobina2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If UCase(Mid(cbobina2 + "  ", 1, 1)) = "T" Then
       If bacceptarbobina.tag = "T" Then
           MsgBox "Ja has escanejat la bobina del tubo has de fer la de l'etiqueta.", vbInformation, "Error"
           cbobina2 = ""
           GoTo fi
             Else: bacceptarbobina.tag = "T": cbobina2 = Mid(cbobina2, 2)
       End If
    End If
    bacceptarbobina_Click
  End If
fi:

End Sub

Private Sub cbotonum_Click(Index As Integer)
If cbotonum(Index).caption = "OK" Then Framepassword.visible = False: GoTo fi
   cpassword.tag = cpassword.tag + cbotonum(Index).caption
   If Framepassword.tag = "password" Then
      cpassword = cpassword + "*"
       Else: cpassword = cpassword.tag
   End If
   cpassword.SetFocus
fi:
End Sub

Private Sub cescanverify_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      comprovarverificaciobobina
      carregarbobinespendentsverificar
      cescanverify = ""
      Framependentverificar.visible = False
  End If
End Sub

Private Sub Command1_Click()
   afegir_bobinadesbobinador 2
End Sub

Private Sub Command2_Click()
  cpassword = ""
  cpassword.tag = ""
End Sub

Private Sub Command3_Click()
   If Framependentverificar.visible = True Then Framependentverificar.visible = False: Exit Sub
   Framependentverificar.visible = True
   carregarbobinespendentsverificar
   cescanverify.SetFocus
End Sub
Sub carregarbobinespendentsverificar()
    Dim rst As Recordset
    llistapendentsverificar.Clear
    Set rst = dbtmpb.OpenRecordset("select * from bobinesdesbobinadors where verificada=false and maquina=" + atrim(nummaq) + " order by data desc")
    While Not rst.EOF
      llistapendentsverificar.AddItem atrim(rst!palet) + "/" + atrim(rst!bobina)
      rst.MoveNext
    Wend
    Set rst = Nothing
End Sub
Function ensenyarescaneigbobina() As String
   Command6.tag = ""
   cbobina1 = ""
   cbobina2 = ""
   bacceptarbobina.tag = ""
   frameescaneigbobina.Top = bdesb1.Top
   frameescaneigbobina.Left = (Formdesbobinadors.width / 2) - (frameescaneigbobina.width / 2)
   frameescaneigbobina.visible = True
   cbobina1.SetFocus
   While frameescaneigbobina.visible
         DoEvents
   Wend
   If cbobina1 = cbobina2 Then ensenyarescaneigbobina = cbobina1
   If Command6.tag = "manual" Then ensenyarescaneigbobina = ensenyarkeypad("Escaneja la Bobina:")
End Function

Function ensenyarkeypad(vmissatge As String) As String
   Framepassword.Top = 1
   etkeypad = vmissatge
   cpassword.tag = ""
   cpassword = ""
   Framepassword.Left = (Formdesbobinadors.width / 2) - (Framepassword.width / 2)
   Framepassword.visible = True
   cpassword.SetFocus
   While Framepassword.visible
         DoEvents
   Wend
    ensenyarkeypad = cpassword.tag
End Function

Private Sub Command4_Click()
  capturarEtiquetaiGuardarla etbob1
  carregar_desbobinadors
End Sub
Sub crearlacarpetaperPassarEtiquetesBobinaProveidor(vnumpalet As Double, carpetadesti As String)
   Dim carpetaprincipal As String
   Dim vcarpetatemporal As String
   Dim vubicaciocarpetadesti As String
   Dim vnomfitxer As String
   vcarpetatemporal = rutadelfitxer(llegir_ini("General", "cami", fitxerini))
   'carpetadesti = llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini")
   carpetadesti = vcarpetatemporal
   'si no puc accedir a la carpeta ho guardo en una temporal en el servidor fins que es pugui descarregar
  ' If Not existeix(carpetadesti + "cache_EtiquetesBobinesProveidor") Then carpetadesti = vcarpetatemporal 'MkDir carpetadesti + "cache_EtiquetesBobinesProveidor"
   carpetadesti = carpetadesti + "cache_EtiquetesBobinesProveidor"
   
   carpetaprincipal = "Els_" + atrim(atrim(Int(cadbl(vnumpalet) / 1000)) + "000")
   If Not existeix(carpetadesti) Then MkDir carpetadesti
   If Not existeix(carpetadesti + "\" + carpetaprincipal) Then MkDir carpetadesti + "\" + carpetaprincipal
   'If Not existeix(carpetadesti + "\" + carpetaprincipal + "\" + atrim(vnumpalet)) Then MkDir carpetadesti + "\" + carpetaprincipal + "\" + atrim(vnumpalet)
   vubicaciocarpetadesti = carpetadesti
   carpetadesti = carpetadesti + "\" + carpetaprincipal + "\"
   
 
End Sub


Private Sub Command5_Click()
    capturarEtiquetaiGuardarla etbob2
    carregar_desbobinadors
End Sub
Sub capturarEtiquetaiGuardarla(vEtiqueta As String)
  Dim vcarpetadesti As String
  Dim vpalet As Double
  If InStr(1, vEtiqueta + "  ", "/") = 0 Then Exit Sub
  vpalet = cadbl(Mid(vEtiqueta, 1, InStr(1, vEtiqueta + "  ", "/") - 1))
  If vpalet = 0 Then Exit Sub
  If existeix("c:\temp\capturaetiqueta_Ok.Jpg") Then Kill "c:\temp\capturaetiqueta_Ok.Jpg"
  Formdesbobinadors.tag = "escanejant"
  formcapturaetiqueta.Show 1
  Formdesbobinadors.tag = ""
  'si existeix el fitxer c:\temp\capturaetiqueta_OK.jpg guardar-lo
  'amb el codi de bobina a la carpeta de bobines i fer tot el proces
  If existeix("c:\temp\capturaetiqueta_Ok.Jpg") Then
        crearlacarpetaperPassarEtiquetesBobinaProveidor vpalet, vcarpetadesti
         'MsgBox vcarpetadesti
        FileCopy "c:\temp\capturaetiqueta_OK.jpg", vcarpetadesti + substituir(vEtiqueta, "/", "_") + ".jpg"
  End If
  
End Sub

Private Sub Command6_Click()
  frameescaneigbobina.visible = False
  Command6.tag = "manual"
End Sub

Private Sub Form_Activate()
   If arguments(1) <> "DESBOBINADORS" Then bsortir.visible = True
   carregar_desbobinadors
End Sub
Sub carregar_desbobinadors()
   Dim rst As Recordset
   
   etcomanda1 = "": etbob1 = ""
   etcomanda2 = "": etbob2 = ""
   fotoetiqueta1 = LoadPicture("")
   fotoetiqueta2 = LoadPicture("")
   Set rst = dbtmpb.OpenRecordset("select *  from bobinesdesbobinadors where numdesbobinador=1 and maquina=" + atrim(nummaq) + " order by data desc")
   If Not rst.EOF Then
     etcomanda1 = atrim(rst!comanda)
     etbob1 = atrim(rst!palet) + "/" + atrim(rst!bobina)
     Set fotoetiqueta1 = LoadPicture(nomfitxer_fotoetiquetabobina(etbob1))
   End If
   
   Set rst = dbtmpb.OpenRecordset("select *  from bobinesdesbobinadors where numdesbobinador=2 and maquina=" + atrim(nummaq) + " order by data desc")
   If Not rst.EOF Then
      etcomanda2 = atrim(rst!comanda)
      etbob2 = atrim(rst!palet) + "/" + atrim(rst!bobina)
      Set fotoetiqueta2 = LoadPicture(nomfitxer_fotoetiquetabobina(etbob2))
   End If
   Set rst = Nothing
   
End Sub
Function nomfitxer_fotoetiquetabobina(vbobina As String) As String
   Dim vrutafotos As String
   Dim vpalet As Double
   Dim vnomfitxer As String
   vrutafotos = llegir_ini("ruta", "ruta_etiquetes_bobinaproveidor", rutadelfitxer(cami) + "valorsprograma.ini")
   If Not existeix(vrutafotos) Then GoTo fi
   vpalet = cadbl(Mid(vbobina, 1, InStr(1, vbobina + " ", "/") - 1))
   If cadbl(vpalet) = 0 Then GoTo fi
   vrutafotos = rutadelfitxer(cami) + "cache_EtiquetesBobinesProveidor"
   vnomfitxer = vrutafotos + "\Els_" + atrim(atrim(Int(cadbl(vpalet) / 1000)) + "000") + "\" + substituir(vbobina, "/", "_") + ".jpg"
   If existeix(vnomfitxer) Then
           nomfitxer_fotoetiquetabobina = vnomfitxer
         Else
           vrutafotos = llegir_ini("ruta", "ruta_etiquetes_bobinaproveidor", rutadelfitxer(cami) + "valorsprograma.ini")
           vnomfitxer = vrutafotos + "\Els_" + atrim(atrim(Int(cadbl(vpalet) / 1000)) + "000") + "\" + substituir(vbobina, "/", "_") + ".jpg"
           If existeix(vnomfitxer) Then nomfitxer_fotoetiquetabobina = vnomfitxer
   End If
fi:
End Function
Private Sub Form_Click()
'  MsgBox "w:" + Trim(Me.width) + " h:" + Trim(Me.Height)
End Sub

Private Sub Form_Load()
   If numop = 0 Then nomoperari_Click
   
End Sub

Private Sub Frame2_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub fotoetiqueta1_DblClick()
   obrir_document nomfitxer_fotoetiquetabobina(etbob1)
   
End Sub

Private Sub fotoetiqueta2_Click()
    Shell nomfitxer_fotoetiquetabobina(etbob2)
End Sub

Private Sub nomoperari_Click()
  Dim numoptmp As Integer
  Dim nomoptmp As String
 
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select codi,descripcio from operaris where maquina='I' and actiu<>0 order by codi "
  formseleccio.caption = "Selecció d'Operari"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   numoptmp = cadbl(formseleccio.Data1.Recordset!codi)
   nomoptmp = atrim(formseleccio.Data1.Recordset!descripcio)
   'If InStr(1, nomoperari.Caption, "MARTINEZ") Then
   '    Command12.Visible = True
   '   Else: Command12.Visible = False
   'End If
  End If
  Unload formseleccio
  If numoptmp <> 0 Then
     nomoperari = nomoptmp
     numop = numoptmp
     Else: If cadbl(numop) = 0 Then MsgBox "Has d'escullir un operari per treballar": End
  End If
   
End Sub

Private Sub Timer1_Timer()
  etstatus.visible = False
End Sub
