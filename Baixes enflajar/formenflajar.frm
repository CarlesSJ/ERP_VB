VERSION 5.00
Begin VB.Form Formembolicar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manteniment embolicar"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11685
   Icon            =   "formenflajar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   594
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   779
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton boperari 
      Caption         =   "Escull OPERARI"
      Height          =   510
      Left            =   210
      TabIndex        =   4
      Top             =   135
      Width           =   1980
   End
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   11610
      Begin VB.Label etcomanda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9900
         TabIndex        =   2
         Top             =   135
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Escullir OPERARI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   2700
         TabIndex        =   1
         Top             =   225
         Width           =   6195
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8145
      Left            =   60
      TabIndex        =   3
      Top             =   735
      Width           =   11595
      Begin VB.CommandButton Command5 
         Caption         =   "Començar Base nova"
         Height          =   945
         Left            =   9525
         Picture         =   "formenflajar.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   5760
         Width           =   1980
      End
      Begin VB.Frame Framealçada3 
         Height          =   2175
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   4230
         Begin VB.TextBox calçada3 
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
            Height          =   510
            Left            =   2145
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1350
            Width           =   1065
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Treure aquest palet"
            Height          =   1125
            Left            =   255
            Picture         =   "formenflajar.frx":179C
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   675
            Width           =   1695
         End
         Begin VB.Label etoperari3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   315
            TabIndex        =   28
            Top             =   1845
            Width           =   1635
         End
         Begin VB.Label etdata3 
            Caption         =   "Label5"
            Height          =   210
            Left            =   2250
            TabIndex        =   23
            Top             =   1890
            Width           =   1905
         End
         Begin VB.Label etnumpalet3 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   480
            Left            =   300
            TabIndex        =   21
            Top             =   225
            Width           =   2745
         End
         Begin VB.Label Label4 
            Caption         =   "Cms"
            Height          =   300
            Left            =   3270
            TabIndex        =   18
            Top             =   1530
            Width           =   300
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Afegir palet"
         Height          =   1500
         Left            =   9510
         Picture         =   "formenflajar.frx":2041
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4215
         Width           =   2010
      End
      Begin VB.CommandButton bbaseacabada 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Base acabada"
         Height          =   1215
         Left            =   9555
         Picture         =   "formenflajar.frx":311D
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6795
         Width           =   1980
      End
      Begin VB.Frame Framealçada2 
         Height          =   2565
         Left            =   135
         TabIndex        =   8
         Top             =   2865
         Visible         =   0   'False
         Width           =   4230
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Treure aquest palet"
            Height          =   1500
            Left            =   225
            Picture         =   "formenflajar.frx":33F3
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   690
            Width           =   1695
         End
         Begin VB.TextBox calçada2 
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
            Height          =   510
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1215
            Width           =   1065
         End
         Begin VB.Label etoperari2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   270
            TabIndex        =   27
            Top             =   2250
            Width           =   1635
         End
         Begin VB.Label etdata2 
            Caption         =   "Label5"
            Height          =   240
            Left            =   2205
            TabIndex        =   24
            Top             =   2280
            Width           =   1890
         End
         Begin VB.Label etnumpalet2 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   540
            Left            =   360
            TabIndex        =   20
            Top             =   120
            Width           =   2595
         End
         Begin VB.Label Label3 
            Caption         =   "Cms"
            Height          =   300
            Left            =   3255
            TabIndex        =   10
            Top             =   1455
            Width           =   300
         End
      End
      Begin VB.Frame framealçada1 
         Height          =   2595
         Left            =   135
         TabIndex        =   5
         Top             =   5400
         Visible         =   0   'False
         Width           =   4230
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Treure aquest palet"
            Height          =   1500
            Left            =   225
            Picture         =   "formenflajar.frx":3C98
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   675
            Width           =   1695
         End
         Begin VB.TextBox calçada1 
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
            Height          =   510
            Left            =   2145
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1140
            Width           =   1065
         End
         Begin VB.Label etoperari1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   270
            TabIndex        =   26
            Top             =   2265
            Width           =   1635
         End
         Begin VB.Label etdata1 
            Caption         =   "Label5"
            Height          =   225
            Left            =   2235
            TabIndex        =   25
            Top             =   2235
            Width           =   1725
         End
         Begin VB.Label etnumpalet1 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   435
            TabIndex        =   19
            Top             =   180
            Width           =   2475
         End
         Begin VB.Label Label1 
            Caption         =   "Cms"
            Height          =   300
            Left            =   3270
            TabIndex        =   7
            Top             =   1395
            Width           =   300
         End
      End
      Begin VB.Label etinfocomanda 
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
         Height          =   6600
         Left            =   4485
         TabIndex        =   32
         Top             =   720
         Width           =   6900
      End
      Begin VB.Label etoperaribase 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6990
         TabIndex        =   31
         Top             =   7395
         Width           =   2355
      End
      Begin VB.Label etbaseacabada 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Base acabada: 24/12/23 15:30"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005C31DD&
         Height          =   450
         Left            =   4455
         TabIndex        =   30
         Top             =   7740
         Width           =   4965
      End
      Begin VB.Label etalçadabase 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "240 cms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   4710
         TabIndex        =   22
         Top             =   135
         Width           =   3615
      End
      Begin VB.Image Imgpalet3 
         Height          =   2775
         Left            =   4905
         Picture         =   "formenflajar.frx":453D
         Stretch         =   -1  'True
         Top             =   765
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.Image imgpalet2 
         Height          =   2775
         Left            =   4875
         Picture         =   "formenflajar.frx":A3A3
         Stretch         =   -1  'True
         Top             =   2715
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.Image imgpalet1 
         Height          =   2670
         Left            =   4830
         Picture         =   "formenflajar.frx":10209
         Stretch         =   -1  'True
         Top             =   4995
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   6060
         Left            =   4530
         Picture         =   "formenflajar.frx":1606F
         Stretch         =   -1  'True
         Top             =   2055
         Width           =   5130
      End
   End
End
Attribute VB_Name = "Formembolicar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbvendes As Database
Dim vdirenvioBaseactual As Double
Dim numop As Byte
Dim rstoperaris As Recordset

Private VScr As Integer, HScr As Integer
Private VFactor As Single, HFactor As Single
Private NoCambiar As Boolean


Private Sub bbaseacabada_Click()
  'If MsgBox("Vols guardar aquesta base?", vbDefaultButton2 + vbYesNo, "Actualitzar dades de la base") = vbYes Then
      guarda_dades_base etnumpalet1
  'End If
  vdirenvioBaseactual = 0
  borrar_distribucio
End Sub
Sub guarda_dades_base(etnumpalet As String)
   dbvendes.Execute "update embolicarpalets set [database]=#" + Format(Now, "mm/dd/yy hh:nn") + "#,operaribase=" + atrim(numop) + " where numreferenciagrup='" + etnumpalet + "'"
End Sub
Sub borrar_distribucio()
   imgpalet1.Visible = False: framealçada1.Visible = False: etnumpalet1.Visible = True
   imgpalet2.Visible = False: Framealçada2.Visible = False: etnumpalet2.Visible = True
   Imgpalet3.Visible = False: Framealçada3.Visible = False: etnumpalet3.Visible = True
   etnumpalet1 = "": calçada1 = "": etoperari1 = "": etoperari1.Tag = "": etdata1 = ""
   etnumpalet2 = "": calçada2 = "": etoperari2 = "": etoperari2.Tag = "": etdata2 = ""
   etnumpalet3 = "": calçada3 = "": etoperari3 = "": etoperari3.Tag = "": etdata3 = ""
   etoperaribase = "": etoperaribase.Tag = ""
   etbaseacabada = ""
End Sub
Private Sub boperari_Click()
 Dim numoptmp As Integer
 Dim nomoptmp As String
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select distinct codi,descripcio from operaris where (maquina='T') and actiu<>0 order by codi "
  formseleccio.Caption = "Selecció d'Operari"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   numoptmp = cadbl(formseleccio.Data1.Recordset!codi)
   nomoptmp = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  If numoptmp <> 0 Then
     nomoperari = nomoptmp
     numop = numoptmp
     Label2.Caption = nomoptmp
       For Each objecte In Me
      If objecte.Name <> "boperari" And InStr(1, objecte.Name, "Line") = 0 And objecte.Name <> "rellotge" And objecte.Name <> "llistat" And objecte.Name <> "llistatbob" Then
        objecte.Enabled = True
      End If
  Next objecte

     If cadbl(numop) = 0 Then MsgBox "Has d'escullir un operari per treballar": Exit Sub
  End If

End Sub

Private Sub calçada1_Change()
    etalçadabase = calcular_pes_delabase
End Sub

Private Sub calçada1_DblClick()
   Dim v As String
   v = InputBox("Escriu l'alçada del palet en Cms o bé els pisos que té. " + etnumpalet1, "Alçada")
   v = substituir(v, ".", ",")
   If cadbl(v) = 0 Then Exit Sub
   v = calcular_alçadaXrPisos(cadbl(v), etnumpalet1)
   If v < 0 Then v = v * -1
   calçada1 = v
   guardar_base
End Sub

Private Sub calçada2_Change()
  etalçadabase = calcular_pes_delabase
End Sub
Function calcular_pes_delabase() As String
  Dim vpes As Double
  ' el primer palet es suma 15cm del palet de fusta, del segon i tercer es suma 17 del palet de fusta mes la fusta de separació
  vpes = 15 + cadbl(calçada1) + IIf(cadbl(calçada2) > 0, cadbl(calçada2) + 17, 0) + IIf(cadbl(calçada3) > 0, cadbl(calçada3) + 17, 0)
  calcular_pes_delabase = atrim(vpes) + " Cms"
End Function

Private Sub calçada2_DblClick()
 Dim v As String
   v = InputBox("Escriu l'alçada del palet en Cms o bé els pisos que té. " + etnumpalet1 + " + " + etnumpalet2, "Alçada")
   v = substituir(v, ".", ",")
   If cadbl(v) = 0 Then Exit Sub
   v = calcular_alçadaXrPisos(cadbl(v), etnumpalet2)
   If v < 0 Then
        v = v * -1
        calçada2 = cadbl(v) - cadbl(calçada1)
          Else: calçada2 = atrim(v)
   End If
   guardar_base
End Sub

Private Sub calçada3_Change()
   etalçadabase = calcular_pes_delabase
End Sub

Private Sub calçada3_DblClick()
 Dim v As String
   v = InputBox("Escriu l'alçada del palet en Cms o bé els pisos que té. " + etnumpalet1 + " + " + etnumpalet2 + " + " + etnumpalet3, "Alçada")
   v = substituir(v, ".", ",")
   If cadbl(v) = 0 Then Exit Sub
   v = calcular_alçadaXrPisos(cadbl(v), etnumpalet3)
   If v < 0 Then
       v = v * -1
       calçada3 = cadbl(v) - cadbl(calçada1) - cadbl(calçada2)
        Else: calçada3 = atrim(v)
   End If
   guardar_base
End Sub
Function calcular_alçadaXrPisos(valçadaopisos As Double, vpalet As String) As Double
  Dim vnumc As Double
  Dim vnumpalet As Double
  Dim rst As Recordset
  If valçadaopisos > 10 Then calcular_alçadaXrPisos = valçadaopisos * -1: GoTo fi
  vnumc = cadbl(Mid(vpalet, 1, InStr(1, vpalet, "/") - 1))
  vnumpalet = cadbl(Mid(vpalet, InStr(1, vpalet, "/") + 1))
  If vnumc = 0 Then GoTo fi
  Set rst = dbcomandes.OpenRecordset("SELECT comandes.amplereb,comandes.ampleesq, productes.ruta FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi Where comanda = " + atrim(vnumc))
  If rst.EOF Then GoTo fi
  If InStr(1, rst!ruta, "S") = 0 Then
         If InStr(1, rst!ruta, "R") > 0 Then calcular_alçadaXrPisos = cadbl(rst!amplereb) * valçadaopisos
         If rst!ruta = "E" Then calcular_alçadaXrPisos = cadbl(rst!ampleesq) * valçadaopisos
         calcular_alçadaXrPisos = Redondejar(calcular_alçadaXrPisos, 0)
          Else: MsgBox "Aquesta comanda es de soldadores no puc calcular l'alçada amb els pisos", vbCritical, "Error"
  End If
fi:
  Set rst = Nothing
End Function

Private Sub Command1_Click()
If Imgpalet3.Visible Then Exit Sub
If MsgBox("Segur que vols treure aquest palet de la base?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
dbvendes.Execute "update embolicarpalets set posicioalabase=1,numreferenciagrup='" + etnumpalet2 + "' where numreferenciagrup='" + etnumpalet1 + "' and posicioalabase=2"
carregar_guardats etnumpalet1
'imgpalet2.Visible = False: Framealçada2.Visible = False: etnumpalet2.Visible = False

'guardar_base
'etnumpalet2 = ""
End Sub

Private Sub Command2_Click()
 If imgpalet2.Visible Then Exit Sub
 If MsgBox("Segur que vols treure aquest palet de la base?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
 imgpalet1.Visible = False: framealçada1.Visible = False: etnumpalet1.Visible = False
guardar_base
etnumpalet1 = ""
End Sub

Private Sub Command3_Click()
  Dim vpalet As String
  Dim valçada As Double
  Dim rstemb As Recordset
  Dim vmetres As Double
  Dim vdata As String
  Dim voperari As Double
  Dim vpaletprincipal As String
  
  If Imgpalet3.Visible Then MsgBox "No pots afegir mes palets a aquesta base.", vbCritical, "Error": Exit Sub
  If Not revisarsihihaelsmetrespossats Then Exit Sub
  vpalet = InputBox("Escaneja el palet que vols afegir." + vbNewLine + " O escriu-lo.  Ex: 214332/2", "Afegir palet")
  vpalet = atrim(substituir(vpalet + " ", ".", ""))
  vpalet = atrim(substituir(vpalet + " ", ",", ""))
  'If aquestpaletjatealbaraassignat(vpalet, vpaletprincipal) Then
  '    carregar_guardats vpaletprincipal
  '   Else: carregar_guardats IIf(etnumpalet1 <> "", etnumpalet1, vpalet)
  'End If
  carregar_guardats IIf(etnumpalet1 <> "", etnumpalet1, atrim(vpalet))
  carregar_info_comanda vpalet
  If vpalet = "" Then MsgBox "No he trobat aquest palet d'aquesta comanda com a pendent d'enviar.", vbCritical, "Error": Exit Sub
  If vpalet = "1" Then MsgBox "Aquest palet no té la mateixa direcció d'enviament que el que estas apilant.", vbCritical, "Error": Exit Sub
  If Not jaestaassignada(vpalet, rstemb) Then
        If Not rstemb.EOF Then
           vmetres = rstemb!metres: vdata = rstemb!Data: voperari = rstemb!operari
           If rstemb!posicioalabase > 1 Then
             MsgBox "Aquest palet ja està remuntat." + vbNewLine + "FES ACCEPTAR PER CARREGAR LA BASE.", vbCritical, "Error"
             carregar_guardats rstemb!numreferenciagrup
             GoTo cont
           End If
        End If
        If vdata = "" Then vdata = Format(Now, "dd/mm/yy hh:nn")
        If voperari = 0 Then voperari = numop
        If imgpalet1.Visible = False Then imgpalet1.Visible = True: framealçada1.Visible = True: etnumpalet1 = vpalet: calçada1 = atrim(vmetres): possarnomoperari etoperari1, voperari: etdata1 = vdata: GoTo guardar
        If imgpalet2.Visible = False Then imgpalet2.Visible = True: Framealçada2.Visible = True: etnumpalet2 = vpalet: calçada2 = atrim(vmetres): possarnomoperari etoperari2, voperari: etdata2 = vdata: GoTo guardar
        If Imgpalet3.Visible = False Then Imgpalet3.Visible = True: Framealçada3.Visible = True: etnumpalet3 = vpalet: calçada3 = atrim(vmetres): possarnomoperari etoperari3, voperari: etdata3 = vdata: GoTo guardar
guardar:
        guardar_base  'guarda els palets tal com estan
          Else: If vpalet <> etnumpalet1 Then MsgBox "Aquest palet " + vpalet + " ja està assignat a dins d'aquesta base.", vbCritical, "Palet assignat"
  End If
cont:
  Set rstemb = Nothing
  revisarsihihaelsmetrespossats
  carregar_guardats etnumpalet1
End Sub
Function jaestaassignada(vpalet As String, rst As Recordset) As Boolean
   Dim vnumc As Double
   Dim vnumpalet As Double
   If InStr(1, vpalet, "/") = 0 Then jaestaassignada = True: Exit Function
   vnumc = cadbl(Mid(vpalet, 1, InStr(1, vpalet, "/") - 1))
   vnumpalet = cadbl(Mid(vpalet, InStr(1, vpalet, "/") + 1))
   Set rst = dbvendes.OpenRecordset("select * from embolicarpalets where numcomanda=" + atrim(vnumc) + " and numpalet=" + atrim(vnumpalet))
   If Not rst.EOF Then
         If rst!numreferenciagrup = etnumpalet1 Then jaestaassignada = True
       '  valçada = cadbl(rst!metres)
   End If
   'Set rst = Nothing
End Function
Function aquestpaletjatealbaraassignat(vpalet As String, vpaletprincipal As String) As Boolean
   Dim vnumc As Double
   Dim vnumpalet As Double
   Dim rst As Recordset
   If atrim(vpalet) = "" Or InStr(1, vpalet, "/") = 0 Then Exit Function
   vnumc = cadbl(Mid(vpalet, 1, InStr(1, vpalet, "/") - 1))
   vnumpalet = cadbl(Mid(vpalet, InStr(1, vpalet, "/") + 1))
   Set rst = dbvendes.OpenRecordset("select * from embolicarpalets where numcomanda=" + atrim(vnumc) + " and numpalet=" + atrim(vnumpalet))
   If Not rst.EOF Then If cadbl(rst!numalbara) > 0 Then vpaletprincipal = atrim(rst!numreferenciagrup): aquestpaletjatealbaraassignat = True
   Set rst = Nothing
End Function
Function buscar_albara(vnumc As Double, vnumpalet As Double) As Double
   Dim rst As Recordset
   buscar_albara = 0
   Set rst = dbvendes.OpenRecordset("select * from bobinesent where comanda=" + atrim(vnumc) + " and numpalet=" + atrim(vnumpalet))
   If Not rst.EOF Then buscar_albara = cadbl(rst!numalbara)
   Set rst = Nothing
End Function
Sub guardar_base()
   Dim vvalues As String
   Dim vnumc As Double
   Dim vnumpalet As Double
   Dim vnumalb As Double
   Dim vdata As String
   'If Not calçada1.Visible Then Exit Sub
   dbvendes.Execute "delete * from embolicarpalets where numreferenciagrup='" + etnumpalet1 + "'"
   If etnumpalet1.Visible And cadbl(calçada1) > 0 Then
     vnumc = cadbl(Mid(etnumpalet1, 1, InStr(1, etnumpalet1, "/") - 1))
     vnumpalet = cadbl(Mid(etnumpalet1, InStr(1, etnumpalet1, "/") + 1))
     vdata = Format(etdata1, "mm/dd/yy hh:nn")
     dbvendes.Execute "delete * from embolicarpalets where numcomanda=" + atrim(vnumc) + " and numpalet=" + atrim(vnumpalet)
     If atrim(etoperari1.Tag) = "" Then etoperari1.Tag = atrim(numop)
     vnumalb = buscar_albara(vnumc, vnumpalet)
     vvalues = "#" + vdata + "#," + atrim(etoperari1.Tag) + "," + atrim(vnumc) + "," + atrim(vnumpalet) + "," + atrim(passaradecimalpunt(calçada1)) + ",1,'" + etnumpalet1 + "'," + atrim(vnumalb)
     dbvendes.Execute "insert into embolicarpalets (data,operari,numcomanda,numpalet,metres,posicioalabase,numreferenciagrup,numalbara) values (" + vvalues + ")"
   End If
   
   If calçada2.Visible And cadbl(calçada2) > 0 Then
    vnumc = cadbl(Mid(etnumpalet2, 1, InStr(1, etnumpalet2, "/") - 1))
    vnumpalet = cadbl(Mid(etnumpalet2, InStr(1, etnumpalet2, "/") + 1))
    vdata = Format(etdata2, "mm/dd/yy hh:nn")
    dbvendes.Execute "delete * from embolicarpalets where numcomanda=" + atrim(vnumc) + " and numpalet=" + atrim(vnumpalet)
    If atrim(etoperari2.Tag) = "" Then etoperari2.Tag = atrim(numop)
    vnumalb = buscar_albara(vnumc, vnumpalet)
    vvalues = "#" + vdata + "#," + atrim(etoperari2.Tag) + "," + atrim(vnumc) + "," + atrim(vnumpalet) + "," + atrim(passaradecimalpunt(calçada2)) + ",2,'" + etnumpalet1 + "'," + atrim(vnumalb)
    dbvendes.Execute "insert into embolicarpalets (data,operari,numcomanda,numpalet,metres,posicioalabase,numreferenciagrup,numalbara) values (" + vvalues + ")"
   End If
   
   If calçada3.Visible And cadbl(calçada3) > 0 Then
    vnumc = cadbl(Mid(etnumpalet3, 1, InStr(1, etnumpalet3, "/") - 1))
    vnumpalet = cadbl(Mid(etnumpalet3, InStr(1, etnumpalet3, "/") + 1))
    vdata = Format(etdata3, "mm/dd/yy hh:nn")
    dbvendes.Execute "delete * from embolicarpalets where numcomanda=" + atrim(vnumc) + " and numpalet=" + atrim(vnumpalet)
    If atrim(etoperari3.Tag) = "" Then etoperari3.Tag = atrim(numop)
    vnumalb = buscar_albara(vnumc, vnumpalet)
    vvalues = "#" + vdata + "#," + atrim(etoperari3.Tag) + "," + atrim(vnumc) + "," + atrim(vnumpalet) + "," + atrim(passaradecimalpunt(calçada3)) + ",3,'" + etnumpalet1 + "'," + atrim(vnumalb)
    dbvendes.Execute "insert into embolicarpalets (data,operari,numcomanda,numpalet,metres,posicioalabase,numreferenciagrup,numalbara) values (" + vvalues + ")"
   End If
   
End Sub
Function revisarsihihaelsmetrespossats() As Boolean
   revisarsihihaelsmetrespossats = True
   If cadbl(calçada1) = 0 And calçada1.Visible Then calçada1_DblClick
   If cadbl(calçada2) = 0 And calçada2.Visible Then calçada2_DblClick
   If cadbl(calçada3) = 0 And calçada3.Visible Then calçada3_DblClick
   If (cadbl(calçada1) = 0 And calçada1.Visible) Or (cadbl(calçada2) = 0 And calçada2.Visible) Or (cadbl(calçada3) = 0 And calçada3.Visible) Then revisarsihihaelsmetrespossats = False
End Function
Sub possarnomoperari(etoperari As Control, voperari As Double)
   rstoperaris.FindFirst "codi=" + atrim(cadbl(voperari))
   etoperari.Caption = ""
   etoperari.Tag = ""
   If rstoperaris.NoMatch Then Exit Sub
   etoperari.Caption = atrim(rstoperaris!descripcio)
   etoperari.Tag = atrim(voperari)
End Sub
Sub carregar_guardats(vpalet As String)
   Dim rst As Recordset
   borrar_distribucio
   If InStr(1, vpalet, "/") = 0 Then Exit Sub
    Set rst = dbvendes.OpenRecordset("select * from embolicarpalets where  numreferenciagrup='" + atrim(vpalet) + "' order by posicioalabase asc")
   If rst.EOF Then GoTo fi
   'If rst!numreferenciagrup <> vpalet And rst!posicioalabase > 1 Then
   '  vpalet = rst!numreferenciagrup
   '  Set rst = dbvendes.OpenRecordset("select * from embolicarpalets where  numreferenciagrup='" + atrim(vpalet) + "' order by posicioalabase asc")
   'End If
   While Not rst.EOF
     If rst!posicioalabase = 1 Then imgpalet1.Visible = True: framealçada1.Visible = True: etnumpalet1 = atrim(rst!numcomanda) + "/" + atrim(rst!numpalet): calçada1 = rst!metres: possarnomoperari etoperari1, cadbl(rst!operari): etdata1 = atrim(rst!Data)
     If rst!posicioalabase = 2 Then imgpalet2.Visible = True: Framealçada2.Visible = True: etnumpalet2 = atrim(rst!numcomanda) + "/" + atrim(rst!numpalet): calçada2 = rst!metres: possarnomoperari etoperari2, cadbl(rst!operari): etdata2 = atrim(rst!Data)
     If rst!posicioalabase = 3 Then Imgpalet3.Visible = True: Framealçada3.Visible = True: etnumpalet3 = atrim(rst!numcomanda) + "/" + atrim(rst!numpalet): calçada3 = rst!metres: possarnomoperari etoperari3, cadbl(rst!operari): etdata3 = atrim(rst!Data)
     If Not IsNull(rst!Database) Then
        etbaseacabada = "Base acabada: " + Format(rst!Database, "dd/mm/yy hh:nn")
        etbaseacabada.Tag = atrim(rst!Database)
        possarnomoperari etoperaribase, cadbl(rst!operaribase)
        etoperaribase.Tag = atrim(rst!operaribase)
     End If
     rst.MoveNext
   Wend
fi:
   Set rst = Nothing
End Sub
Private Sub Command4_Click()
   If MsgBox("Segur que vols treure aquest palet de la base?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   dbvendes.Execute "update embolicarpalets set posicioalabase=1, numreferenciagrup='" + etnumpalet3 + "' where numreferenciagrup='" + etnumpalet1 + "' and posicioalabase=3"
   carregar_guardats etnumpalet1

   'Imgpalet3.Visible = False: Framealçada3.Visible = False: etnumpalet3.Visible = False
   'guardar_base
   'etnumpalet3 = ""
End Sub
Sub carregar_info_comanda(vnumcipalet As String)
   Dim vnumc As Double
   Dim vnumpalet As Double
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim vmsg As String
'   Exit Sub
   etinfocomanda = ""
   If InStr(1, vnumcipalet, "/") = 0 Then Exit Sub
   vnumc = cadbl(Mid(vnumcipalet, 1, InStr(1, vnumcipalet, "/") - 1))
   vnumpalet = cadbl(Mid(vnumcipalet, InStr(1, vnumcipalet, "/") + 1))
   Set rst = dbbaixes.OpenRecordset("select * from bobinesent where numpalet=" + atrim(vnumpalet) + " and isnull(dataentrega)")
   If rst.EOF Then vnumcipalet = "": GoTo fi
   Set rst = dbcomandes.OpenRecordset("select client,direnvio from comandes where comanda=" + atrim(vnumc))
   If rst.EOF Then vnumcipalet = "": GoTo fi
   If rst!direnvio <> vdirenvioBaseactual And vdirenvioBaseactual <> 0 Then vnumcipalet = "1": GoTo fi
   Set rst2 = dbcomandes.OpenRecordset("SELECT marcailinia,clients.nom, Clients_envios.pais,clients_envios.poblacioe FROM (comandes LEFT JOIN clients ON comandes.client = clients.codi) LEFT JOIN Clients_envios ON comandes.direnvio = Clients_envios.id where comanda=" + atrim(vnumc))
   vdirenvioBaseactual = rst!direnvio
   etinfocomanda = atrim(vnumcipalet) + " - " + atrim(rst2!nom) + vbNewLine + atrim(rst2!poblacioe) + " [" + atrim(rst2!pais) + "]" + vbNewLine + atrim(rst2!marcailinia)
   possar_taulafullexpedicions Trim(vnumc), vmsg
   etinfocomanda = etinfocomanda + vmsg
fi:
   Set rst = Nothing
   Set rst2 = Nothing
End Sub

Sub llistatlookupde(taula As String, Optional control1 As String, Optional control2 As String, Optional camp As String, Optional altres As String)
Dim rsttmp2 As Recordset
If camp = "" Then camp = "descripcio"
If altres = "clientsextres" Then camp = camp + ",observacions1,observacions2,obsext1,obsext2,obsimp1,obsimp2,obslam1,obslam2,obsreb1,obsreb2,obssol1,obssol2"
If Len(taula) < 20 Then
    Set rsttmp2 = dbcomandes.OpenRecordset("select " + camp + " from " + taula + " where codi=" + atrim(cadbl(control1)), , ReadOnly)
   Else: Set rsttmp2 = dbtmp.OpenRecordset(taula, , ReadOnly)
End If
If Not rsttmp2.EOF Then
     control2 = atrim(rsttmp2.Fields(0))
    Else: control2 = ""
End If

End Sub
Sub possar_taulafullexpedicions(numerodecomanda As String, vmsg As String)
  Dim rste As Recordset
  Dim rstll As Recordset
  Dim rstenvio As Recordset
  Dim rstclient As Recordset
  Dim codienvio As Long
  Dim rsttmp As Recordset
  Dim vvalor As String
  
  
  Set rsttmp = dbcomandes.OpenRecordset("select * from comandesmesextres where comanda=" + atrim(numerodecomanda))
  If rsttmp.EOF Then Exit Sub
  Set rstclient = dbcomandes.OpenRecordset("select * from clients where codi=" + atrim(cadbl(rsttmp!client)))
  If cadbl(rsttmp!direnvio) > 0 Then
      Set rstenvio = dbcomandes.OpenRecordset("select * from clients_envios where id=" + atrim(cadbl(rsttmp!direnvio)))
      codienvio = cadbl(rsttmp!direnvio)
     Else
        Set rstenvio = dbcomandes.OpenRecordset("select * from clients where codi=" + atrim(cadbl(rsttmp!client)))
        codienvio = cadbl(rsttmp!client) * -1
  End If
  
  If rstenvio.EOF Then
    'no hi ha dades a direccions d'envio
      Exit Sub
  End If
  
  
 
  If cadbl(rstenvio!albaravalorat) Then vmsg = vmsg + vbNewLine + "VALORAT"
  If cadbl(rstenvio!codibarres) Then vmsg = vmsg + vbNewLine + "CODI DE BARRES"
  If cadbl(rstenvio!datafabricacio) Then vmsg = vmsg + vbNewLine + "DATA DE FABRICACIÓ"
  If cadbl(rstenvio!detallbobalpalet) Then vmsg = vmsg + vbNewLine + "DETALL BOBINES AL PALET"
  If cadbl(rstenvio!detallbobalfrontal) Then vmsg = vmsg + vbNewLine + "DETALL BOBINES AL FRONTAL"
  If cadbl(rstenvio!pesnetbrut) Then vmsg = vmsg + vbNewLine + "PES NET"
  'If atrim(rstenvio!bobinesmaxpalet) <> "" Then rstll!bobinesmaxpalet = rsttmp!bobinesmaxpalet
  If cadbl(rstenvio!alcadapalet) Then
    llistatlookupde "alcadespalets", atrim(rstenvio!alcadapalet), vvalor
    vmsg = vmsg + vbNewLine + vvalor + " CM" + IIf(cadbl(rstenvio!pesmaxpalet) > 0, "  Pes Màx. Palet: " + atrim(cadbl(rstenvio!pesmaxpalet)) + " Kg", "")
  End If
  If cadbl(rstenvio!tipuspalet) Then
    llistatlookupde "tipuspalets", atrim(rstenvio!tipuspalet), vvalor
    vmsg = vmsg + vbNewLine + vvalor
  End If
  If cadbl(rstenvio!guardarmostres) Then
    llistatlookupde "guardarmostres", atrim(rstenvio!guardarmostres), vvalor
    vmsg = vmsg + vbNewLine + vvalor
  End If
  If cadbl(rstenvio!cert_qualitat) Then
    llistatlookupde "cert_qualitat", atrim(rstenvio!cert_qualitat), vvalor
    vmsg = vmsg + vbNewLine + vvalor
  End If
  If cadbl(rstenvio!albaraalpalet) Then vmsg = vmsg + vbNewLine + "ALBARÀ AL PALET"
  If cadbl(rstenvio!packingalpalet) Then vmsg = vmsg + vbNewLine + "PACKING-LIST"
  If cadbl(rstenvio!tipusprotecciob) Then
    llistatlookupde "tipusproteccions", atrim(rstenvio!tipusprotecciob), vvalor
    vmsg = vmsg + vbNewLine + vvalor
  End If
  If cadbl(rstenvio!tipusprotecciop) Then
    llistatlookupde "tipusproteccions", atrim(rstenvio!tipusprotecciop), vvalor
    vmsg = vmsg + vbNewLine + vvalor
  End If
  If cadbl(rstenvio!tipusprotecciospr) Then
    llistatlookupde "tipusproteccions", atrim(rstenvio!tipusprotecciospr), vvalor
    vmsg = vmsg + vbNewLine + vvalor
  End If
  If cadbl(rstenvio!emb_anonim) Then
    llistatlookupde "embalatgesanonims", atrim(rstenvio!emb_anonim), vvalor
    vmsg = vmsg + vbNewLine + vvalor
  End If
   If cadbl(rstenvio!guardarmostres) Then
    llistatlookupde "guardarmostres", atrim(rstenvio!guardarmostres), vvalor
    vmsg = vmsg + vbNewLine + vvalor
  End If
  If cadbl(rstenvio!conosprotectors) Then
    llistatlookupde "conosprotectors", atrim(rstenvio!conosprotectors), vvalor
    vmsg = vmsg + vbNewLine + vvalor
  End If
  If cadbl(rstenvio!conosprotectors) Then
    llistatlookupde "conosprotectors", atrim(rstenvio!conosprotectors), vvalor
    vmsg = vmsg + vbNewLine + vvalor
  End If
  If cadbl(rstenvio!okenvio) Then vmsg = vmsg + vbNewLine + "DEMANAR OK PER ENVIAR"
  If cadbl(rstenvio!pfpaperfrontal) Then
    llistatlookupde "tipuspaperfrontal", atrim(rstenvio!pfpaperfrontal), vvalor
    vmsg = vmsg + vbNewLine + vvalor
  End If
  If atrim(rstenvio!pfpaperfrontal) <> "" Then
     If cadbl(rstenvio!pfpesnet) Then vmsg = vmsg + vbNewLine + "PES NET"
     If cadbl(rstenvio!pfdatafab) Then vmsg = vmsg + vbNewLine + "DATA FABRICACIO"
     If cadbl(rstenvio!pfpacking) Then vmsg = vmsg + vbNewLine + "PACKING-LIST"
     If cadbl(rstenvio!pfcodibarres) Then vmsg = vmsg + vbNewLine + "CODI DE BARRES" + IIf(atrim(rstenvio!estilfrontal) <> "", " (" + atrim(rstenvio!estilfrontal) + ")", "")
  End If
  If atrim(rstenvio!arxiuexp) <> "" Then
      obrir_document "\\ord_copies\pautacli\" + atrim(rstenvio!arxiuexp)
  End If
  

End Sub


Private Sub Command5_Click()
   vdirenvioBaseactual = 0
   borrar_distribucio
   etinfocomanda = ""
End Sub

Private Sub etalçadabase_Change()
   If cadbl(Mid(etalçadabase, 1, InStr(1, etalçadabase + " ", " "))) > 235 Then
        Frame2.BackColor = &HC0C0FF
          Else: Frame2.BackColor = &H8000000F
   End If
End Sub

Private Sub Form_Load()

'HScr = Screen.Width / Screen.TwipsPerPixelX
'VScr = Screen.Height / Screen.TwipsPerPixelY
'VFactor = VScr / 600
'HFactor = HScr / 800
'Factor = HScr / 800
'NoCambiar = True
'Me.Width = Me.Width * HFactor
'Me.Height = Me.Height * VFactor
'NoCambiar = False



 camicomandes = llegir_ini("General", "cami", "comandes.ini")
  cami = llegir_ini("General", "camibaixes", "comandes.ini")
  fitxerini = "comandes.ini"
  If cami = "{[}]" Then
    escriure_ini "General", "camibaixes", InputBox("Entra la ruta de baixes", "Atenció", "y:\comandes\baixes.mdb"), "comandes.ini"
  End If
  If Not existeix("c:\ordprog.ini") Then assignardecimalipunt
  centerscreen Me
  Set dbcomandes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "comandes.mdb")
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
  Set dbvendes = OpenDatabase(rutadelfitxer(cami) + "vendes.mdb")
  If llegir_ini("Baixes", "programaamaquina", fitxerini) = "1" Then
   Shell ("net time \\serverprodu /set /y")
  End If
  For Each objecte In Me
      If objecte.Name <> "boperari" And InStr(1, objecte.Name, "Line") = 0 And objecte.Name <> "rellotge" And objecte.Name <> "llistat" And objecte.Name <> "llistatbob" Then
        objecte.Enabled = False
      End If
  Next objecte
  etalçadabase = ""
  etoperaribase = ""
  etbaseacabada = ""
  Set rstoperaris = dbcomandes.OpenRecordset("select * from operaris where maquina='T' and actiu=1")
End Sub

Private Sub Form_Resize()
Dim ctl As Control
Exit Sub
On Error Resume Next
If NoCambiar Then Exit Sub
For Each ctl In Me
    ctl.Top = ctl.Top * VFactor
    ctl.Height = ctl.Height * VFactor
    ctl.Left = ctl.Left * HFactor
    ctl.Width = ctl.Width * HFactor
    ctl.FontSize = ctl.FontSize * HFactor
Next
End Sub

