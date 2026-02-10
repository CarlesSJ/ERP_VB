VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form entradabobina 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10905
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10905
   ScaleWidth      =   12765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   6465
      Top             =   1575
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12285
      Picture         =   "entradabobina.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sortir"
      Top             =   90
      Width           =   390
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informació de la Bobina"
      Height          =   8745
      Left            =   75
      TabIndex        =   11
      Top             =   2100
      Width           =   12300
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Imprimint Cara INTERIOR"
         Height          =   600
         Left            =   5370
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   795
         Width           =   1185
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FF00&
         Caption         =   "Imprimint Cara EXTERIOR"
         Height          =   600
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   135
         Width           =   1170
      End
      Begin VB.Label etbobina 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   315
         TabIndex        =   20
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label etmaterialimpres 
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
         Height          =   1065
         Left            =   6870
         TabIndex        =   19
         Top             =   285
         Width           =   5175
      End
      Begin VB.Label etescullircara 
         BackStyle       =   0  'Transparent
         Caption         =   "Escull quina cara s'està imprimint ===>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005C31DD&
         Height          =   420
         Left            =   1425
         TabIndex        =   17
         Top             =   630
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label etmatinterior 
         BackStyle       =   0  'Transparent
         Caption         =   "Interior: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   14
         Top             =   1020
         Visible         =   0   'False
         Width           =   6270
      End
      Begin VB.Label etmatexterior 
         BackStyle       =   0  'Transparent
         Caption         =   "Exterior: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   990
         TabIndex        =   13
         Top             =   255
         Visible         =   0   'False
         Width           =   5580
      End
      Begin VB.Image Image2 
         Height          =   840
         Left            =   75
         Picture         =   "entradabobina.frx":058A
         Stretch         =   -1  'True
         Top             =   195
         Width           =   1320
      End
      Begin VB.Image fotoetiqueta 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   6990
         Left            =   165
         Stretch         =   -1  'True
         Top             =   1575
         Width           =   11820
      End
   End
   Begin VB.CommandButton bdesb2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   4050
      Picture         =   "entradabobina.frx":0F42
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1065
      Width           =   1440
   End
   Begin VB.CommandButton bdesb1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   540
      Picture         =   "entradabobina.frx":1492
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1095
      Width           =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2175
      Top             =   120
   End
   Begin VB.CommandButton alta 
      Height          =   285
      Left            =   1110
      Picture         =   "entradabobina.frx":19E6
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Escriure el palet manualment"
      Top             =   15
      Width           =   300
   End
   Begin VB.CommandButton Command1 
      Height          =   390
      Left            =   1950
      Picture         =   "entradabobina.frx":1F70
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   285
      Width           =   390
   End
   Begin MSMask.MaskEdBox desb 
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   285
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   661
      _Version        =   327681
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox palet 
      Height          =   375
      Left            =   390
      TabIndex        =   10
      Top             =   300
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      _Version        =   327681
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox bobina 
      Height          =   375
      Left            =   1485
      TabIndex        =   1
      Top             =   285
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      _Version        =   327681
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label etrefinplacsa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2910
      TabIndex        =   18
      Top             =   120
      Width           =   3105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escaneja amb l'escaner l'etiqueta."
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   30
      TabIndex        =   7
      Top             =   705
      Width           =   2460
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bob."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1545
      TabIndex        =   4
      Top             =   45
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Palet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   615
      TabIndex        =   3
      Top             =   45
      Width           =   480
   End
   Begin VB.Label etdesb 
      BackStyle       =   0  'Transparent
      Caption         =   "Desb."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -15
      TabIndex        =   2
      Top             =   45
      Width           =   480
   End
End
Attribute VB_Name = "entradabobina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vafegintbobina As Boolean
Private Sub alta_Click()
  If Now > CVDate("15/02/2021") Then
     MsgBox "Avís" + Chr(10) + "Si entres la bobina manualment s'enviarà un e-mail a oficina informant.", vbOKOnly + vbInformation, "Atenció"
     form1.botoensenyarpacking.tag = "afegidamanualmentcaixa"
  End If
  palet.MaxLength = 5
  palet.AutoTab = True
  
  palet.SetFocus
End Sub

Private Sub bdesb1_Click()
  'bdesb1.caption = "57240/2"
  carregar_etiquetabobina bdesb1.caption
  'triarpalet bdesb1.caption
  
End Sub

Private Sub bdesb2_Click()
   carregar_etiquetabobina bdesb2.caption
'   triarpalet bdesb2.caption
   
End Sub
Sub triarpalet(vp As String)
  Dim vpalet As Double
  Dim vbobina As Double
  vafegintbobina = True
  palet.SetFocus
  vpalet = cadbl(Mid(" " + vp, 1, InStr(1, vp + "  ", "/")))
  vbobina = cadbl(Mid(vp, InStr(1, vp + "  ", "/") + 1))
  palet = Trim(vpalet)
  bobina = Trim(vbobina)
  Command1_Click
  vafegintbobina = False
End Sub

Private Sub bobina_GotFocus()
bobina.SelStart = 0
  bobina.SelLength = Len(bobina.text)
End Sub

Private Sub bobina_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then Command1_Click
End Sub

Private Sub Command1_Click()
  If InStr(1, bobina, "-") > 0 Or InStr(1, bobina, "/") > 0 Then
        palet = bobina: bobina = ""
        palet_LostFocus
  End If
  entradabobina.Hide
End Sub

Private Sub Command2_Click()
  entradabobina.Hide
End Sub

Private Sub Command3_Click()
  comprovar_coincidencia etmatexterior
   
End Sub
Sub comprovar_coincidencia(vetmaterial As Label)
   If vetmaterial.tag = etrefinplacsa.tag Then
       If vetmaterial.tag = etmaterialimpres.tag Then
                  triarpalet etbobina
                   Else: MsgBox "La cara del material que has escullit no coincideix amb el escullit per oficines.", vbCritical, "Error"
       End If
         Else: MsgBox "Les cares TRACTADES NO coincideixen amb les escullides al escanejar l'etiqueta de la bobina."
   End If
End Sub

Private Sub Command4_Click()
   comprovar_coincidencia etmatinterior
End Sub

Private Sub desb_GotFocus()
desb.SelStart = 0
  desb.SelLength = Len(desb.text)
End Sub
Function nomfitxer_fotoetiquetabobina(vbobina As String) As String
   Dim vrutafotos As String
   Dim vpalet As Double
   Dim vnomfitxer As String
   If vbobina = "" Then Exit Function
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

Sub carregar_etiquetabobina(vnumbobina As String)
 Dim vubicaciobobina As String
 etbobina = vnumbobina
 Command3.Enabled = True: Command4.Enabled = True
 fotoetiqueta.Picture = LoadPicture(""): fotoetiqueta.tag = ""
 vubicaciobobina = nomfitxer_fotoetiquetabobina(vnumbobina)
 If existeix(vubicaciobobina) Then
     fotoetiqueta.Picture = LoadPicture(vubicaciobobina)
     fotoetiqueta.tag = vubicaciobobina
       Else:
         MsgBox "No hi ha l'etiqueta escanejada d'aquesta bobina." + vbNewLine + "Sense etiqueta no es pot continuar.", vbCritical, "Atenció"
         Command3.Enabled = False
         Command4.Enabled = False
 End If
 possar_cara_tractada vnumbobina
End Sub

Sub possar_cara_tractada(vnumbobina As String)
   Dim rst As Recordset
 Dim vpalet As Double
 Dim vbobina As Double
 Dim rstt As Recordset
 Dim rstref As Recordset
 Dim rstmat As Recordset
 Dim vcaraexteriorbobina As Double
 etmaterial = ""
 etrefinplacsa.tag = ""
 etmatinterior.tag = "": etmatexterior.tag = ""
 vnumbobina = substituir(vnumbobina, "-", "/")
 etmaterialimpres = "": etmaterialimpres.tag = ""
 convertirScanambPaletiBobina vnumbobina, vpalet, vbobina
 Set rst = dbstocks.OpenRecordset("SELECT * from bobines where bobines.idpalet=" + Trim(vpalet) + " and idbobina=" + Trim(vbobina))
 If rst.EOF Then MsgBox "No he trobat la bobina.", vbCritical, "Error": GoTo fi
 vcaraexteriorbobina = cadbl(rst!caraexterior)
 Set rst = dbstocks.OpenRecordset("select * from palets where idpalet=" + atrim(vpalet))
 If rst.EOF Then MsgBox "No he trobat la bobina.", vbCritical, "Error": GoTo fi
 Set rst = dbstocks.OpenRecordset("select * from materials where codi=" + atrim(rst!codimatprognou))
 If rst.EOF Then MsgBox "No he trobat la bobina.", vbCritical, "Error": GoTo fi
 etmaterial = atrim(rst!descripcio)
 Set rstt = dbtmp.OpenRecordset("select * from tractamentcares")
 rstt.FindFirst "codi=" + atrim(rst!codidescmatcara1)
 'If Not rstt.NoMatch Then checkmaterialexterior(0).caption = atrim(rstt!descripcio): checkmaterialexterior(0).tag = atrim(rst!codidescmatcara1)
 If Not rstt.NoMatch Then
       If rst!codidescmatcara1 = vcaraexteriorbobina Then
             etmatexterior = atrim(rstt!descripcio)
             etmatexterior.tag = atrim(rst!codidescmatcara1)
               Else:
                 etmatinterior = atrim(rstt!descripcio)
                 etmatinterior.tag = atrim(rst!codidescmatcara1)
       End If
 End If
 rstt.FindFirst "codi=" + atrim(rst!codidescmatcara2)
 If Not rstt.NoMatch Then
       If rst!codidescmatcara2 = vcaraexteriorbobina Then
              etmatexterior = atrim(rstt!descripcio)
              etmatexterior.tag = atrim(rst!codidescmatcara2)
                Else
                  etmatinterior = atrim(rstt!descripcio)
                  etmatinterior.tag = atrim(rst!codidescmatcara2)
       End If
 End If
' If Not rstt.NoMatch Then checkmaterialexterior(1).caption = atrim(rstt!descripcio): checkmaterialexterior(1).tag = atrim(rst!codidescmatcara2)
 'If rst!codidescmatcara1 = vcaraexteriorbobina Then checkmaterialexterior(0).Value = 1
 'If rst!codidescmatcara2 = vcaraexteriorbobina Then checkmaterialexterior(1).Value = 1
 Set rstref = dbtmp.OpenRecordset("select * from referencies_disposiciomaterials where refinplacsa='" + atrim(etrefinplacsa) + "'")
 If Not rstref.EOF Then
      Set rstmat = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(rstref!material1))
      If Not rstmat.EOF Then
         
         If rstref!caraimpresio = 1 Then etrefinplacsa.tag = atrim(rstmat!codidescmatcara1)
         If rstref!caraimpresio = 2 Then etrefinplacsa.tag = atrim(rstmat!codidescmatcara2)
         rstt.FindFirst "codi=" + etrefinplacsa.tag
         If Not rstt.NoMatch Then etmaterialimpres = "Cara a imprimir: " + atrim(rstt!descripcio): etmaterialimpres.tag = atrim(etrefinplacsa.tag)
      End If
 End If
fi:
 etmatexterior = "Exterior: " + etmatexterior
 etmatinterior = "Interior: " + etmatinterior
 Set rst = Nothing
 Set rstt = Nothing

End Sub

Private Sub Form_Activate()

'If existeix("c:\ordprog.ini") Then etmatinterior.visible = True: etmatexterior.visible = True
   If form1.botoensenyarpacking.tag = "afegidamanualmentcaixa" Then palet.MaxLength = 5: palet.AutoTab = True
   If InStr(1, UCase(App.EXEName), "IMPRES") > 0 Then
        entradabobina.Height = 11000
        carregar_bobines_desbobinadors
         Else:
           entradabobina.Height = 1000: entradabobina.width = 3000: Command2.Left = 2900 - Command2.width
           entradabobina.Left = (Screen.width / 2) - (entradabobina.width / 2)
           entradabobina.Top = (Screen.Height / 2) - (entradabobina.Height / 2)
   End If
   bobina.SetFocus
   carregar_refinplacsa
 '   Else: desb.SetFocus
  'End If
End Sub
Sub carregar_refinplacsa()
   Dim rst As Recordset
   etrefinplacsa = ""
   Set rst = dbtmp.OpenRecordset("select * from comandes_extres where comanda=" + atrim(form1.comanda))
   If Not rst.EOF Then etrefinplacsa = atrim(rst!refinplacsa)
   Set rst = Nothing
End Sub

Sub carregar_bobines_desbobinadors()
   Dim cbob1 As String
   Dim cbob2 As String
   form1.actualitzarestatbobinesdesbobinadors
   cbob1 = llegir_ini("Bobines_Desbobinadors_" + atrim(nummaq), "Bobina1", rutadelfitxer(cami) + "valorsprograma.ini")
   cbob2 = llegir_ini("Bobines_Desbobinadors_" + atrim(nummaq), "Bobina2", rutadelfitxer(cami) + "valorsprograma.ini")
   If cbob1 = "{[}]" Then cbob1 = ""
   If cbob2 = "{[}]" Then cbob2 = ""
   bdesb1.caption = cbob1
   bdesb2.caption = cbob2
End Sub

Private Sub Label4_Click()

End Sub

Private Sub fotoetiqueta_DblClick()
  obrir_document fotoetiqueta.tag
End Sub

Private Sub palet_Change()
   If Len(palet) = palet.MaxLength Then bobina.SetFocus
End Sub

Private Sub palet_GotFocus()
  If Not vafegintbobina Then If Not comprovar_si_ferhomanualono Then bdesb1.SetFocus
  palet.SelStart = 0
  palet.SelLength = Len(palet.text)
End Sub
Function comprovar_si_ferhomanualono() As Boolean
  Dim v As String
  v = InputBox("Les bobines s'han d'entrar desde l'entrada de bobines del desbobinador si ho fas per aquí s'enviarà un email a oficines i a l'encarregat." + vbNewLine + "Escriu [BOBINA MANUAL] per continuar", "Bobina entrada manual")
  If StrPtr(v) = 0 Then Exit Function
  If UCase(v) <> "BOBINA MANUAL" Then Exit Function
  comprovar_si_ferhomanualono = True
  v = ""
  While v = ""
    v = InputBox("Escriu el motiu perquè entres la bobina desde aquí." + vbNewLine + " NO POTS DEIXAR SENSE RESPOSTA", "Motiu")
  Wend
  enviaremailgeneric "impresores@inplacsa.com", "S'ha entrat una bobina desde el llapis sense utilitzar el DESBOBINADOR.", atrim(Now) + vbNewLine + atrim(numop) + "-" + atrim(form1.nomoperari) + vbNewLine + "Comanda: " + atrim(form1.comanda) + vbNewLine + "Motiu:" + vbNewLine + atrim(v)
End Function

Private Sub palet_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     KeyCode = 0
     Sendkeys "{TAB}"
      Else: If palet.MaxLength = 10 Then KeyCode = 0
  End If
End Sub

'Sub convertirScanambPaletiBobina(vcodi As String, vpalet As Long, vbob As Long)
'   Dim vcont As Double
'   vcodi = atrim(vcodi)
'   While vcont < Len(vcodi)
 '     If Not IsNumeric(Mid(vcodi, vcont + 1, 1)) Then
 '       vpalet = cadbl(Mid(vcodi, 1, vcont))
 '       If Len(vcodi) >= vcont + 2 Then vbob = cadbl(Mid(vcodi, vcont + 2))
 ''       GoTo sortir
 '     End If
 '     vcont = vcont + 1
 '  Wend
'sortir:
'End Sub

Private Sub palet_KeyPress(KeyAscii As Integer)
   Static vhoraultimapulsacio As Date
   Timer1.Enabled = True
   'If vhoraultimapulsacio = "0:00:00" Or vhoraultimapulsacio = Null Then vhoraultimapulsacio = Now
   'If palet.MaxLength = 10 And KeyAscii <> 13 Then 'KeyAscii = 0
       'If DateDiff("s", vhoraultimapulsacio, Now) > 0 Then       palet = "": KeyAscii = 0: vhoraultimapulsacio = Empty
   'End If
   If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub palet_LostFocus()
  Dim vpalet As Double
  Dim vbob As Double
  convertirScanambPaletiBobina palet, vpalet, vbob
  If vpalet > 0 And vbob > 0 Then
     palet = atrim(vpalet)
     bobina = atrim(vbob)
     Command1_Click
  End If
End Sub

Private Sub Timer1_Timer()
  If Screen.ActiveControl.Name = "palet" And palet.MaxLength = 10 Then
     palet.text = ""
     Timer1.Enabled = False
  End If
End Sub

Private Sub Timer2_Timer()
  Static vcont As Byte
  If Len(etmatexterior) < 11 Then etescullircara.visible = False: Exit Sub
  etescullircara.visible = True
  etescullircara = Mid(etescullircara, 1, InStr(1, etescullircara, "="))
  vcont = vcont + 1
  etescullircara = etescullircara + String(vcont, "=") + ">"
  If vcont = 4 Then vcont = 0
End Sub
