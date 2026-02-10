VERSION 5.00
Begin VB.Form opcionsmatajust 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opcions del material"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3045
      Left            =   15
      TabIndex        =   15
      Top             =   2415
      Width           =   6360
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Height          =   570
         Left            =   5100
         Picture         =   "opcionsmaterialajust.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Eliminar les opcions de material"
         Top             =   1515
         Width           =   1110
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Height          =   570
         Left            =   5085
         Picture         =   "opcionsmaterialajust.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2145
         Width           =   1110
      End
      Begin VB.Frame dadesmatestoc 
         Caption         =   "Paràmetres del material d'Estoc"
         Height          =   2910
         Left            =   30
         TabIndex        =   16
         Top             =   0
         Width           =   4995
         Begin VB.ListBox llistaestoc 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2340
            Left            =   150
            TabIndex        =   17
            Top             =   315
            Width           =   4545
         End
         Begin VB.Label errormat 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   120
            TabIndex        =   18
            Top             =   2655
            Width           =   4695
         End
      End
   End
   Begin VB.TextBox numcomanda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3060
      TabIndex        =   6
      Top             =   210
      Width           =   1935
   End
   Begin VB.Frame framematespecific 
      Caption         =   "Parametres del material especific"
      Height          =   1755
      Left            =   150
      TabIndex        =   0
      Top             =   2445
      Width           =   3555
      Begin VB.TextBox numbob 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1695
         TabIndex        =   3
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox numpalet 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   315
         TabIndex        =   1
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label Label4 
         Caption         =   "Bob"
         Height          =   180
         Left            =   1770
         TabIndex        =   5
         Top             =   420
         Width           =   390
      End
      Begin VB.Label Label3 
         Caption         =   "Palet"
         Height          =   225
         Left            =   705
         TabIndex        =   4
         Top             =   390
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1560
         TabIndex        =   2
         Top             =   615
         Width           =   165
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Seccio Impresora"
      Height          =   2295
      Left            =   30
      TabIndex        =   7
      Top             =   75
      Width           =   5055
      Begin VB.Frame Frame2 
         Caption         =   "Tipus de material d'Ajust que s'utilitzarà a impresores"
         Height          =   1200
         Left            =   120
         TabIndex        =   11
         Top             =   1005
         Width           =   4740
         Begin VB.CommandButton b1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Material per  llençar (Grup 2500)"
            Height          =   735
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   300
            Width           =   1275
         End
         Begin VB.CommandButton b2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Material d'Estoc"
            Height          =   735
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   300
            Width           =   1275
         End
         Begin VB.CommandButton b3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Material Assignat especific"
            Height          =   735
            Left            =   3165
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   300
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         Height          =   705
         Left            =   135
         TabIndex        =   8
         Top             =   180
         Width           =   2685
         Begin VB.TextBox metresajust 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1545
            TabIndex        =   9
            Text            =   "1500"
            Top             =   195
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Metres d'ajust:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   135
            TabIndex        =   10
            Top             =   180
            Width           =   1380
         End
      End
   End
End
Attribute VB_Name = "opcionsmatajust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b1_Click()
 possarframes "b1"
End Sub
Sub possarframes(c As String)
   b1.BackColor = &HFFC0C0
   b2.BackColor = &HFFC0C0
   b3.BackColor = &HFFC0C0
   Me.Controls(c).BackColor = QBColor(9)
   'dadesmatestoc.Top = 2355
   Frame4.Top = 2355
   framematespecific.Top = 2355
   framematespecific.Left = 90
   Frame4.Left = 90
   dadesmatestoc.Visible = False
   framematespecific.Visible = False
   If c = "b2" Then
      dadesmatestoc.Visible = True
      dadesmatestoc.ZOrder 0
   End If
   If c = "b3" Then
      framematespecific.Visible = True
      framematespecific.ZOrder 0
   End If
   If Len(c) > 1 And (Frame4.Top > 1 Or cadbl(b1.Tag) = 0) Then b1.Tag = Mid(c, 2, 1)
End Sub
Private Sub b2_Click()
   possarframes "b2"
End Sub

Private Sub b3_Click()
  If Frame4.Top > 1 Then possarframes "b3"
End Sub

Private Sub Command1_Click()
  Dim metres As Double
  If assignarmat.assignarstock.BackColor <> &H80FF80 Then
     If errormat <> "" Then
        If MsgBox("Hi ha una error en la sel.leccio del grup de materials, no es gravaran els canvis." + vbNewLine + errormat + vbNewLine + "Vols continuar igualment?", vbExclamation + vbYesNo, "Atenció") = vbNo Then Exit Sub
     End If
       Else:
           If errormat <> "" And b1.Tag = "2" Then
                   If MsgBox("Els materials que vols utilitzar per ajustar son diferents que els assignats." + Chr(10) + "S'ASSIGNARÀ IGUALMENT PERÒ ASSEGURA QUE SIGUI CORRECTE." + vbNewLine + errormat, vbCritical + vbOKCancel, "ATENCIÓ") = vbCancel Then Exit Sub
           End If
  End If
  If assignarmat.assignarstock.BackColor = &H80FF80 Then  ' si es estoc no comprova el palet
   If cadbl(numpalet) > 0 And cadbl(numbob) > 0 Then
    metres = comprovar_metres_numpaletajust
    If metres = 0 Then
       MsgBox "El numero de palet d'ajust no està dins el packing-list, escull un altre palet/bob", vbCritical, "Atenció"
       Exit Sub
    End If
    If metres < cadbl(metresajust) Then
       MsgBox "Els metres d'ajust que vols possar son inferiors als assignats a aquesta bobina, arregla-ho i torna a canviar-los", vbCritical, "Atenció"
       Exit Sub
    End If
   End If
  End If
  guardarvalorspossats cadbl(numcomanda)
  assignarmat.botoajust.Tag = "1"
  Unload opcionsmatajust
End Sub
Function comprovar_metres_numpaletajust() As Double
  Dim rstp As Recordset
  Set rstp = dbtmp.OpenRecordset("Select * from parcials where comanda='" + atrim(cadbl(numcomanda)) + "' and idpalet=" + atrim(cadbl(numpalet)) + " and idbobina=" + atrim(cadbl(numbob)))
  If rstp.EOF Then
       comprovar_metres_numpaletajust = 0
     Else: comprovar_metres_numpaletajust = rstp!metres
  End If
End Function

Private Sub Command2_Click()
   If InputBox("Escriu [ELIMINAR] si estàs segur que vols eliminar aquestes opcions.", "Eliminiar opcions") = "ELIMINAR" Then
      dbtmp.Execute "delete * from opcionsdajust where comanda=" + atrim(cadbl(numcomanda))
      dbtmpb.Execute ("update comandes_extres set assignarstock=FALSE,mtrsassignatsestock=0 where comanda=" + atrim(CDbl(numcomanda)))
      assignarmat.botoajust.Tag = "0"
      Unload opcionsmatajust
   End If
End Sub

Private Sub Form_Activate()
possarvalorsguardats cadbl(numcomanda)
End Sub

Private Sub Form_Load()
  carregarllistastock
  assignarmat.botoajust.Tag = "0"
  numcomanda.BackColor = opcionsmatajust.BackColor
  If cadbl(numcomanda) = 999999 Then Exit Sub
  
End Sub


Sub guardarvalorspossats(numc As Double)
   Dim rstopcions As Recordset
   Dim sisaj As String
   Dim nomgrup As Double
   Dim grupvell As Double
   nomgrup = 0
   If llistaestoc.ListIndex > -1 Then nomgrup = llistaestoc.ItemData(llistaestoc.ListIndex)
   Set rstopcions = dbtmp.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(numc))
   If Not rstopcions.EOF Then grupvell = rstopcions!grupdestoc: Set rstopcions = Nothing
   dbtmp.Execute "delete * from opcionsdajust where comanda=" + atrim(numc)
   Set rstopcions = dbtmp.OpenRecordset("opcionsdajust")
   sisaj = cadbl(b1.Tag)
   rstopcions.AddNew
   rstopcions!comanda = numc
   rstopcions!grupdestoc = IIf(nomgrup = 0, grupvell, nomgrup)
   rstopcions!mtrsajust = cadbl(metresajust)
   rstopcions!paletajust = cadbl(numpalet)
   rstopcions!bobinaajust = cadbl(numbob)
   rstopcions!sistemadajust = sisaj
   rstopcions.Update
   Set rstopcions = Nothing
End Sub

Sub possarvalorsguardats(numc As Double)
   Dim rstopcions As Recordset
   Dim sisaj As String
   Set rstopcions = dbtmp.OpenRecordset("select * from opcionsdajust where comanda=" + atrim(numc))
   If Not rstopcions.EOF Then
     sisaj = atrim(cadbl(rstopcions!sistemadajust))
     If sisaj > 0 Then
      metresajust = cadbl(rstopcions!mtrsajust)
      buscargrupdestoc cadbl(rstopcions!grupdestoc)
      numpalet = cadbl(rstopcions!paletajust)
      numbob = cadbl(rstopcions!bobinaajust)
      If Frame2.Enabled And Frame4.Top > 1 Then possarframes "b" + atrim(sisaj)
      If Frame4.Top = 1 Then b1.Tag = sisaj
        Else: b3_Click
     End If
       Else: b3_Click
   End If
   
End Sub
Sub buscargrupdestoc(grup As Double)
  Dim i As Double
  Dim j As Double
  i = 0
  j = -1
  While i < llistaestoc.ListCount
    If llistaestoc.ItemData(i) = grup Then j = i
    i = i + 1
  Wend
  llistaestoc.ListIndex = j
End Sub
Sub carregarllistastock()
  Dim rstgrup As Recordset
  Set rstgrup = dbtmp.OpenRecordset("select * from grupsdepalets")
  While Not rstgrup.EOF
    llistaestoc.AddItem rstgrup!nomdelgrup
    llistaestoc.ItemData(llistaestoc.NewIndex) = cadbl(rstgrup!numerogrup)
    
    rstgrup.MoveNext
  Wend
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
  If assignarmat.botoajust.Tag <> "1" Then assignarmat.botoajust.Tag = "0"
End Sub
Function compararsielmaterialcompatibleescorrecte(rstmatc As Recordset, vgrupcompatible As Double) As Byte
   Dim rst As Recordset
   Dim rstm As Recordset
   Set rst = dbtmp.OpenRecordset("select * from grupsmaterialscompatibles where numerodegrup=" + atrim(cadbl(vgrupcompatible)))
   If Not rst.EOF Then
       Set rstm = dbtmp.OpenRecordset("select * from materials where " + atrim(rst!sqlprincipal) + atrim(sqlsubfamilies) + ")")
       If Not rstm.EOF Then compararsielmaterialcompatibleescorrecte = 1
   End If
   Set rst = Nothing
End Function
Private Sub llistaestoc_Click()
  Dim rstgrup As Recordset
   errormat = ""
  Set rstgrup = dbtmp.OpenRecordset("select * from grupsdepalets where numerogrup=" + atrim(cadbl(llistaestoc.ItemData(llistaestoc.ListIndex))))
  If Not rstgrup.EOF Then
     If comparasielmaterialcorrespon(numcomanda, cadbl(rstgrup!paletexemple), cadbl(rstgrup!codigrupmaterialscompatibles)) <> 1 Then
          errormat = "Grup de material no correcte"
           Else:
             If hihaprousmetresestoc(rstgrup!numerogrup, cadbl(assignarmat.mtrsnecessaris)) Then
                errormat = ""
                 Else: errormat = "No hi ha prous metres assignats a aquest Grup d'estoc."
             End If
     End If
  End If
  Set rstgrup = Nothing
End Sub
Function hihaprousmetresestoc(vnumestoc As Double, vmetresaquestacomanda As Double)
  Dim rstopcions As Recordset
  Dim vmetresnecessaris As Double
  Dim vmetresassignats As Double
  hihaprousmetresestoc = True
  Set rstopcions = dbtmp.OpenRecordset("SELECT opcionsdajust.grupdestoc as GrupEstoc, Sum(comandes.cantitatex) AS Tmetres FROM opcionsdajust LEFT JOIN comandes ON opcionsdajust.comanda = comandes.comanda Where (((comandes.proximaseccio) = 'I')) GROUP BY opcionsdajust.grupdestoc;")
  rstopcions.FindFirst "GrupEstoc=" + atrim(vnumestoc)
  If Not rstopcions.NoMatch Then
      vmetresnecessaris = cadbl(rstopcions!tmetres)
      Set rstopcions = dbtmp.OpenRecordset("SELECT Parcials.comanda, Sum(Parcials.metres) AS Tmetres From parcials GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + atrim(vnumestoc) + "'));")
      If Not rstopcions.EOF Then vmetresassignats = cadbl(rstopcions!tmetres)
      If (vmetresnecessaris + vmetresaquestacomanda) >= vmetresassignats Then hihaprousmetresestoc = False
  End If
  Set rstopcions = Nothing
End Function
Function comparasielmaterialcorrespon(comanda As String, numpalet As Double, Optional vcodicompatibles As Double) As Byte
   Dim rstcom As Recordset
   Dim rstpalet As Recordset
   Dim rstmaterial As Recordset
   Dim rstmaterialp As Recordset
   Dim resp As Byte
   Dim micres As Double
   Dim mesuraesp As String
   Set rstcom = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(comanda)))
   resp = 1
   If Not rstcom.EOF Then
      Set rstpalet = dbtmp.OpenRecordset("select * from palets where idpalet=" + atrim(numpalet))
      Set rstmaterial = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstcom!materialex)))
      Set rstmaterialp = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstpalet!codimatprognou)))
      If Not rstpalet.EOF Then
          If cadbl(rstpalet!micres) < 1 Then
             mesuraesp = "grmsm2"
            Else: mesuraesp = "micres"
          End If
          resp = 2
          If vcodicompatibles = 0 Then
                    If cadbl(rstmaterialp!familia) = cadbl(rstmaterial!familia) Then
                       If cadbl(rstmaterialp!subfamilia) = cadbl(rstmaterial!subfamilia) Then
                         If cadbl(rstmaterialp!familiacol) = cadbl(rstmaterial!familiacol) Then
                           If cadbl(rstmaterialp!subfamiliacol) = cadbl(rstmaterial!subfamiliacol) Then
                             If cadbl(rstmaterialp!familiaad) = cadbl(rstmaterial!familiaad) Then
                               If cadbl(rstmaterialp!subfamiliaad) = cadbl(rstmaterial!subfamiliaad) Then
                                    resp = 1
                               End If
                             End If
                           End If
                         End If
                       End If
                    End If
                     Else
                       If compararsielmaterialcompatibleescorrecte(rstmaterial, vcodicompatibles) Then resp = 1
          End If
          micres = assignarmat.micresmaterial(rstcom!mesuraesp, rstcom!espessor, rstcom!tubolam)
          If micres < 0 Then micres = micres * -1
          If resp = 1 Then
             resp = 3
             'If cadbl(rstpalet!ample) >= (cadbl(rstcom!ampleesq) - 1) Then
                 If cadbl(rstpalet!plegat) = cadbl(rstcom!plegatesq) Then
                   If cadbl(rstpalet!solapa) = cadbl(rstcom!solapa) Then
                     If cadbl(rstpalet.Fields(mesuraesp)) > (micres - 3) And cadbl(rstpalet.Fields(mesuraesp)) < (micres + 3) Then
                       'If assignarmat.aatrim(rstpalet!carestractat) = assignarmat.aatrim(rstcom!tractatex) Then
                         If rstpalet!obert = IIf(atrim(rstcom!oberturaex) = "", "N", atrim(rstcom!oberturaex)) Then
                           If cabool(rstpalet!microperforat) = cabool(rstcom!micropex) Then
                             If atrim(rstpalet!semielaborat) = atrim(rstcom!tubolam) Then
                               If cadbl(rstpalet.Fields(mesuraesp)) <> micres Then
                                If MsgBox("Aquest material te unes micres diferents que la comanda però esta dins les 3 micres de marge." + Chr(10) + "Es correcte?", vbCritical + vbYesNo, "Atenció") = vbYes Then
                                    resp = 1
                                   Else: resp = 3
                                End If
                                 Else: resp = 1
                               End If
                             End If
                           End If
                         End If
                       'End If
                     End If
                   End If
                 End If
              'End If
          End If
      End If
     Else: resp = 0
   End If
fi:
   comparasielmaterialcorrespon = resp
 'aquesta linia s'ha de treure per compravar bé el material
   'comparasielmaterialcorrespon = 1
End Function

Private Sub numcomanda_Change()
'  possarvalorsguardats cadbl(numcomanda)
End Sub

