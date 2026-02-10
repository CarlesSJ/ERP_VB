VERSION 5.00
Begin VB.Form Formescanejaretiqueta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Escanejar etiqueta de la bobina"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9150
   Icon            =   "escanejar_etiquetabobina.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameescullircolormaterial 
      BackColor       =   &H00EAD9CE&
      Caption         =   "Escullir Color del material"
      Height          =   4170
      Left            =   -3090
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CommandButton botocolor 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   2
         Left            =   6345
         TabIndex        =   14
         Top             =   1515
         Width           =   2415
      End
      Begin VB.CommandButton botocolor 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   1
         Left            =   3345
         TabIndex        =   13
         Top             =   1515
         Width           =   2415
      End
      Begin VB.CommandButton botocolor 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   0
         Left            =   345
         TabIndex        =   12
         Top             =   1515
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escull el color que té fisicament el material."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005C31DD&
         Height          =   915
         Left            =   780
         TabIndex        =   15
         Top             =   345
         Width           =   7530
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   8325
      Top             =   570
   End
   Begin VB.Frame Framecara 
      Caption         =   "Escullir cara exterior"
      Height          =   3840
      Left            =   240
      TabIndex        =   3
      Top             =   -45
      Visible         =   0   'False
      Width           =   8640
      Begin VB.CommandButton Command3 
         BackColor       =   &H008080FF&
         Caption         =   "No sé veure-ho"
         Height          =   705
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3075
         Width           =   5430
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H006BEBB1&
         Caption         =   "Acceptar"
         Height          =   960
         Left            =   6180
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2775
         Width           =   2295
      End
      Begin VB.CheckBox checkmaterialexterior 
         DownPicture     =   "escanejar_etiquetabobina.frx":048A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Index           =   0
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   5895
      End
      Begin VB.CheckBox checkmaterialexterior 
         DownPicture     =   "escanejar_etiquetabobina.frx":0E42
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Index           =   1
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1530
         Width           =   5895
      End
      Begin VB.Label Label2 
         Caption         =   "Cara Exterior"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6315
         TabIndex        =   9
         Top             =   420
         Width           =   2175
      End
      Begin VB.Image Image2 
         Height          =   1620
         Left            =   6330
         Picture         =   "escanejar_etiquetabobina.frx":17FA
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2370
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Escullir cara exterior"
      Height          =   420
      Left            =   3570
      TabIndex        =   6
      Top             =   1860
      Width           =   2205
   End
   Begin VB.TextBox cnumbobina 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3525
      TabIndex        =   1
      Top             =   1110
      Width           =   2715
   End
   Begin VB.CommandButton bacceptar 
      Caption         =   "Acceptar"
      Height          =   945
      Left            =   6255
      TabIndex        =   0
      Top             =   930
      Width           =   1290
   End
   Begin VB.Label etmaterial 
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
      ForeColor       =   &H00ED823A&
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8235
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Bobina:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   855
      TabIndex        =   2
      Top             =   1155
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   3780
      Picture         =   "escanejar_etiquetabobina.frx":21B2
      Top             =   285
      Width           =   1380
   End
End
Attribute VB_Name = "Formescanejaretiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function capturarEtiquetaiGuardarla(vEtiqueta As String) As Boolean
  Dim vcarpetadesti As String
  Dim vpalet As Double
  If InStr(1, vEtiqueta + "  ", "/") = 0 Then Exit Function
  vpalet = cadbl(Mid(vEtiqueta, 1, InStr(1, vEtiqueta + "  ", "/") - 1))
  If vpalet = 0 Then Exit Function
  If existeix("c:\temp\capturaetiqueta_Ok.Jpg") Then Kill "c:\temp\capturaetiqueta_Ok.Jpg"
  formcapturaetiqueta.Show 1
  'si existeix el fitxer c:\temp\capturaetiqueta_OK.jpg guardar-lo
  'amb el codi de bobina a la carpeta de bobines i fer tot el proces
  If existeix("c:\temp\capturaetiqueta_Ok.Jpg") Then
        If existeix("c:\temp\capturaetiqueta_OK.pdf") Then Kill "c:\temp\capturaetiqueta_OK.pdf"
        ConvertirFormats "c:\temp\capturaetiqueta_OK.jpg", "c:\temp\capturaetiqueta_OK.jpg", 20
        crearlacarpetaperPassarEtiquetesBobinaProveidor vpalet, vcarpetadesti
         'MsgBox vcarpetadesti
        FileCopy "c:\temp\capturaetiqueta_OK.jpg", vcarpetadesti + substituir(vEtiqueta, "/", "_") + ".jpg"
        capturarEtiquetaiGuardarla = True
  End If
  
End Function
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

Private Sub bacceptar_Click()
 Dim rst As Recordset
 Dim vcarpetadesti As String
 Dim vescanejadaOK As Boolean
 bacceptar.Enabled = False
 Command1.Enabled = False
 possar_cara_tractada
 If etmaterial = "" Then MsgBox "No hi ha una bobina vàlida escanejada.", vbCritical, "Error": GoTo fi
 cnumbobina = substituir(cnumbobina, "-", "/")
 If InStr(1, cnumbobina, "/") = 0 Then GoTo fi
 crearlacarpetaperPassarEtiquetesBobinaProveidor cadbl(Mid(cnumbobina, 1, InStr(1, cnumbobina + "  ", "/") - 1)), vcarpetadesti
 If jashaescanejatletiqueta(cnumbobina, vcarpetadesti) Then If MsgBox("Aquesta etiqueta ja està escanejada..." + vbNewLine + "Vols tornar-hi?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then cnumbobina = "": GoTo fi
 vescanejadaOK = capturarEtiquetaiGuardarla(cnumbobina)
fi:
 Set rst = Nothing
 If vescanejadaOK Then
      Command1_Click
     Else: cnumbobina.SetFocus
 End If
 bacceptar.Enabled = True
 Command1.Enabled = True
End Sub

Function jashaescanejatletiqueta(vnumetiqueta As String, vcarpetadesti As String, Optional vnomfitxeretiqueta As String) As Boolean
  Dim vruta As String
  vnumetiqueta = substituir(vnumetiqueta, "/", "_")
   If existeix(vcarpetadesti + vnumetiqueta + ".jpg") Then
        vnomfitxeretiqueta = vcarpetadesti + vnumetiqueta + ".jpg"
        jashaescanejatletiqueta = True
         Else
            vruta = llegir_ini("ruta", "ruta_etiquetes_bobinaproveidor", rutadelfitxer(cami) + "valorsprograma.ini")
            vruta = vruta + "\" + Mid(vcarpetadesti, InStr(1, vcarpetadesti, "Els_")) + vnumetiqueta + ".jpg"
            If existeix(vruta) Then jashaescanejatletiqueta = True
            vnomfitxeretiqueta = vruta
   End If
   
End Function

Private Sub botocolor_Click(Index As Integer)
  If botocolor(Index).Caption <> botocolor(0).Tag Then
       MsgBox "AQUEST COLOR NO COINCIDEIX AMB EL DEL MATERIAL, SI REALMENT ES D'AQUEST COLOR AVISA A OFICINES.", vbCritical, "ERROR DE COLOR"
         Else: botocolor(0).Tag = ""
  End If
End Sub

Private Sub checkmaterialexterior_Click(Index As Integer)
  Static vsocdins As Boolean
  If Not vsocdins Then
    vsocdins = True
   If Index = 0 Then If checkmaterialexterior(1).Value = 1 Then checkmaterialexterior(1).Value = 0
   If Index = 1 Then If checkmaterialexterior(0).Value = 1 Then checkmaterialexterior(0).Value = 0
   vsocdins = False
  End If
End Sub

Private Sub cnumbobina_GotFocus()
   If cnumbobina.Enabled Then
     cnumbobina.SetFocus
     cnumbobina.SelStart = 0
     cnumbobina.SelLength = Len(cnumbobina)
   End If
End Sub

Private Sub cnumbobina_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = 13 Then bacceptar_Click
End Sub

Private Sub cnumbobina_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then bacceptar_Click
End Sub

Sub possar_cara_tractada()
   Dim rst As Recordset
 Dim vpalet As Double
 Dim vbobina As Double
 Dim rstt As Recordset
 Dim vcaratractada As Double
 etmaterial = ""
 cnumbobina = substituir(cnumbobina, "-", "/")
 checkmaterialexterior(0).Tag = "": checkmaterialexterior(0).Value = 0
 checkmaterialexterior(1).Tag = "": checkmaterialexterior(1).Value = 0
 convertirScanambPaletiBobina cnumbobina, vpalet, vbobina
 Set rst = dbstocks.OpenRecordset("SELECT * from bobines where bobines.idpalet=" + Trim(vpalet) + " and idbobina=" + Trim(vbobina))
 If rst.EOF Then MsgBox "No he trobat la bobina.", vbCritical, "Error": GoTo fi
 vcaratractada = cadbl(rst!caraexterior)
 Set rst = dbstocks.OpenRecordset("select * from palets where idpalet=" + atrim(vpalet))
 If rst.EOF Then MsgBox "No he trobat la bobina.", vbCritical, "Error": GoTo fi
 Set rst = dbstocks.OpenRecordset("select * from materials where codi=" + atrim(rst!codimatprognou))
 If rst.EOF Then MsgBox "No he trobat la bobina.", vbCritical, "Error": GoTo fi
 etmaterial = atrim(rst!descripcio)
 Set rstt = dbcomandes.OpenRecordset("select * from tractamentcares")
 If cadbl(atrim(rst!codidescmatcara1)) = 0 Then MsgBox "Avisar a Compres que possin el tractament en aquest material.", vbCritical, "Atenció": GoTo fi
 rstt.FindFirst "codi=" + atrim(rst!codidescmatcara1)
 If Not rstt.NoMatch Then checkmaterialexterior(0).Caption = atrim(rstt!descripcio): checkmaterialexterior(0).Tag = atrim(rst!codidescmatcara1)
 rstt.FindFirst "codi=" + atrim(rst!codidescmatcara2)
 If Not rstt.NoMatch Then checkmaterialexterior(1).Caption = atrim(rstt!descripcio): checkmaterialexterior(1).Tag = atrim(rst!codidescmatcara2)
 If rst!codidescmatcara1 = vcaratractada Then checkmaterialexterior(0).Value = 1
 If rst!codidescmatcara2 = vcaratractada Then checkmaterialexterior(1).Value = 1
fi:
 Set rst = Nothing
 Set rstt = Nothing

End Sub
Private Sub Command1_Click()
  possar_cara_tractada
  If etmaterial = "" Then MsgBox "No hi ha bobina escullida.", vbCritical, "Error": Exit Sub
  Framecara.Visible = True
  Framecara.Left = 200
  Framecara.Top = 300
  
End Sub

Private Sub Command2_Click()
    Dim rst As Recordset
    Dim vpalet As Double
    Dim vbobina As Double
    Dim vcaraexterior As Double
    Dim venviarverificacio As Boolean
    If checkmaterialexterior(0).Value <> 1 And checkmaterialexterior(1).Value <> 1 Then MsgBox "No POTS ACCEPTAR SENSE ESCULLIR UNA CARA EXTERIOR O BÉ APRETAR EL BOTÓ DE <NO SÉ VEURE-HO>", vbCritical, "ATENCIÓ": Exit Sub
    convertirScanambPaletiBobina cnumbobina, vpalet, vbobina
    If checkmaterialexterior(0).Value = 1 Then vcaraexterior = cadbl(checkmaterialexterior(0).Tag)
    If checkmaterialexterior(1).Value = 1 Then vcaraexterior = cadbl(checkmaterialexterior(1).Tag)
    Set rst = dbstocks.OpenRecordset("select * from bobines where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbobina))
    If rst.EOF Then GoTo fi
    If cadbl(rst!caraexterior) <> vcaraexterior Then venviarverificacio = True
    dbstocks.Execute "update bobines set caraexterior=" + Trim(vcaraexterior) + " where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbobina)
    'If venviarverificacio Then enviar_etiquetaperverificar vpalet, vbobina
    enviar_etiquetaperverificar vpalet, vbobina
    escullir_color_material vpalet
fi:
    Framecara.Visible = False
    cnumbobina.SetFocus
    If Formescanejaretiqueta.Tag <> "escanejarbobines" Then Unload Formescanejaretiqueta
End Sub
Sub escullir_color_material(vpalet As Double)
    Dim rst As Recordset
    Dim vcolor As String
    Dim vcolor1 As String
    Dim vcolor2 As String
    
    Dim vInventat As Double
    frameescullircolormaterial.Top = 20
    frameescullircolormaterial.Left = 20
    frameescullircolormaterial.Visible = True
    botocolor(0).Tag = "": botocolor(0).Caption = "": botocolor(1).Caption = "": botocolor(2).Caption = ""
    
    Set rst = dbstocks.OpenRecordset("select codimatprognou from palets where idpalet=" + atrim(vpalet))
    If Not rst.EOF Then
        Set rst = dbcomandes.OpenRecordset("SELECT materials.codi, familiescolorants.descripcio FROM familiescolorants RIGHT JOIN materials ON familiescolorants.codi = materials.familiacol where MATERIALS.codi=" + atrim(rst!codimatprognou))
        If Not rst.EOF Then
            vcolor = Trim(UCase(rst!descripcio))
            Set rst = dbcomandes.OpenRecordset("select descripcio from familiescolorants")
            rst.MoveLast
            rst.MoveFirst
            Randomize
            
            'boto 1
            rst.MoveFirst
            vInventat = Int((rst.RecordCount - 1 + 1) * Rnd + 1)
            rst.Move vInventat
            While Trim(UCase(rst!descripcio)) = vcolor
              rst.MoveFirst
              vInventat = Int((rst.RecordCount - 1 + 1) * Rnd + 1)
              rst.Move vInventat
            Wend
            vcolor1 = Trim(UCase(rst!descripcio))
            botocolor(0).Caption = Trim(UCase(rst!descripcio))
            
            'boto 2
            rst.MoveFirst
            vInventat = Int((rst.RecordCount - 1 + 1) * Rnd + 1)
            rst.Move vInventat
            While Trim(UCase(rst!descripcio)) = vcolor Or Trim(UCase(rst!descripcio)) = vcolor1
              rst.MoveFirst
              vInventat = Int((rst.RecordCount - 1 + 1) * Rnd + 1)
              rst.Move vInventat
            Wend
            vcolor2 = Trim(UCase(rst!descripcio))
            botocolor(1).Caption = Trim(UCase(rst!descripcio))
            
            'boto 3
            rst.MoveFirst
            vInventat = Int((rst.RecordCount - 1 + 1) * Rnd + 1)
            rst.Move vInventat
            While Trim(UCase(rst!descripcio)) = vcolor Or Trim(UCase(rst!descripcio)) = vcolor1 Or Trim(UCase(rst!descripcio)) = vcolor2
              rst.MoveFirst
              vInventat = Int((rst.RecordCount - 1 + 1) * Rnd + 1)
              rst.Move vInventat
            Wend
            botocolor(2).Caption = Trim(UCase(rst!descripcio))
            
            
            vInventat = Int((3 - 1 + 1) * Rnd + 1)
            botocolor(vInventat - 1).Caption = vcolor 'posso aleatoriament el color correcte
            botocolor(0).Tag = vcolor
            While botocolor(0).Tag <> ""
               DoEvents
            Wend
        End If
    End If
    frameescullircolormaterial.Visible = False
    Set rst = Nothing
End Sub

Sub enviar_etiquetaperverificar(vpalet As Double, vbobina As Double)
    Dim vrutaetiqueta As String
    Dim vcarpetadesti As String
    Dim vmsg As String
    
    crearlacarpetaperPassarEtiquetesBobinaProveidor vpalet, vcarpetadesti
    If jashaescanejatletiqueta(atrim(vpalet) + "/" + atrim(vbobina), vcarpetadesti, vrutaetiqueta) Then
         If existeix("c:\temp\etiquetaemail.jpg") Then Kill "c:\temp\etiquetaemail.jpg"
         Copiar_Fitxer vrutaetiqueta, "c:\temp\etiquetaemail.jpg"
         vmsg = dadesetiquetapalet(vpalet, vbobina)
         enviaremail "VerificacioEtiquetaEscanejadaOperaris", "Verificació Etiqueta escanejada palet " + atrim(vpalet) + "/" + atrim(vbobina), vmsg, "c:\temp\etiquetaemail.jpg"
    End If
End Sub
Function dadesetiquetapalet(vpalet As Double, vbobina As Double) As String
    Dim vmsg As String
    Dim rst As Recordset
    Dim rstp As Recordset
    Dim rstmat As Recordset
    Dim rstcara As Recordset
    Dim rstb As Recordset
    Dim vdescripciocara As String
    
    Set rst = dbstocks.OpenRecordset("select * from palets where idpalet=" + atrim(vpalet))
    If rst.EOF Then GoTo fi
    Set rstb = dbstocks.OpenRecordset("select * from bobines where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbobina))
    If Not rst.EOF Then
        Set rstmat = dbstocks.OpenRecordset("select * from materials where codi=" + atrim(rst!codimatprognou))
        If rstmat.EOF Then GoTo fi
        Set rstp = dbstocks.OpenRecordset("select * from proveidors where codi=" + atrim(rstmat!proveidor))
        If rstp.EOF Then GoTo fi
        Set rstcara = dbcomandes.OpenRecordset("select * from tractamentcares where codi=" + atrim(rstb!caraexterior))
        If rstcara.EOF Then
                If rstb!caraexterior = 99999 Then vdescripciocara = "NO SÉ COM VEURE-HO"
             Else: vdescripciocara = atrim(rstcara!descripcio)
        End If
        vmsg = "Palet: " + atrim(vpalet) + vbNewLine + "Bobina: " + atrim(vbobina) + vbNewLine
        vmsg = vmsg + "Proveidor: " + atrim(rstp!nom) + vbNewLine + "Referencia proveidor: " + atrim(rstmat!refproducte) + vbNewLine
        vmsg = vmsg + "Descripcio producte: " + atrim(rstmat!codi) + "-" + atrim(rstmat!descripcio) + vbNewLine
        vmsg = vmsg + vbNewLine + "Cara EXTERIOR escullida per l'operari: " + atrim(vdescripciocara)
    End If
    dadesetiquetapalet = vmsg
fi:
    Set rstmat = Nothing
    Set rst = Nothing
    Set rstp = Nothing
    Set rstb = Nothing
    Set rstcara = Nothing
End Function
Private Sub Command3_Click()
    Dim vpalet As Double
    Dim vbobina As Double
    Dim vcaraexterior As Double
    Dim rst As Recordset
    Dim venviarverificacio As Boolean
    convertirScanambPaletiBobina cnumbobina, vpalet, vbobina
    Set rst = dbstocks.OpenRecordset("select * from bobines where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbobina))
    If rst.EOF Then GoTo fi
    If rst!caraexterior <> 99999 Then venviarverificacio = True
    dbstocks.Execute "update bobines set caraexterior=99999 where idpalet=" + atrim(vpalet) + " and idbobina=" + atrim(vbobina)
    enviar_etiquetaperverificar vpalet, vbobina
fi:
    Framecara.Visible = False
    cnumbobina.SetFocus
End Sub

Private Sub Form_Activate()
  cnumbobina.SetFocus
  cnumbobina.SelStart = 0
  cnumbobina.SelLength = Len(cnumbobina)
End Sub

Private Sub Timer1_Timer()
 On Error Resume Next
  If Screen.ActiveControl.Name <> "cnumbobina" Then
     cnumbobina.SetFocus
  End If
End Sub
