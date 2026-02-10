VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formproveidorsqualitat 
   Caption         =   "Control qualitat per proveidors"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   Icon            =   "formproveidorsqualitat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton exportaraxls 
      BackColor       =   &H00F0F0F0&
      Height          =   480
      Left            =   0
      Picture         =   "formproveidorsqualitat.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exportar a Excel"
      Top             =   0
      Width           =   480
   End
   Begin VB.CommandButton Command56 
      Height          =   270
      Left            =   15
      Picture         =   "formproveidorsqualitat.frx":0EA8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar totes les linies"
      Top             =   540
      Width           =   240
   End
   Begin VB.TextBox filtre 
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   540
      Width           =   630
   End
   Begin VB.CommandButton bPDF 
      Height          =   345
      Left            =   10845
      OLEDropMode     =   1  'Manual
      Picture         =   "formproveidorsqualitat.frx":1432
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   135
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.ListBox cllistatipificacions 
      BackColor       =   &H00FDDECE&
      Height          =   2760
      ItemData        =   "formproveidorsqualitat.frx":19BC
      Left            =   4440
      List            =   "formproveidorsqualitat.frx":19CF
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   2850
      Visible         =   0   'False
      Width           =   2130
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   4605
      Left            =   225
      TabIndex        =   0
      Top             =   810
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   8123
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin VB.Menu menuRoS 
      Caption         =   "menuRoS"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mSol 
         Caption         =   "Sol.licitud"
      End
      Begin VB.Menu mRec 
         Caption         =   "Recepció"
      End
   End
End
Attribute VB_Name = "formproveidorsqualitat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbqualitat As Database
Dim dbavisos As Database
Dim rstavisos As Recordset
Dim iniconfigreixa As String
Dim vweres As String
Dim vordrereixa As String
Dim vcampsPDF As String
Sub mSol_click()
   Dim vsol As String
   mSol.Tag = ""
   vsol = InputBox("Entra la data de sol.licitud:", "Sol.licitud", Format(Now, "dd/mm/yy"))
   If IsDate(vsol) Then mSol.Tag = vsol
End Sub
Sub mRec_click()
  Dim vRec As String
   mRec.Tag = ""
   vRec = InputBox("Entra la data de recepcio:", "Recepcio", Format(Now, "dd/mm/yy"))
   If IsDate(vRec) Then mRec.Tag = vRec
End Sub
Private Sub bPDF_Click()
   Dim vnomfitxer As String
   vnomfitxer = llegir_ini("ruta", "rutaQualitatProveidorsLOCAL", rutadelfitxer(cami) + "valorsprograma.ini")
   vnomfitxer = vnomfitxer + "Documentació_Proveidors\" + generar_nom_pdf
 '  MsgBox vnomfitxer
   If existeix(vnomfitxer) Then obrir_document vnomfitxer
End Sub

Private Sub bPDF_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Data.Files.Count = 0 Then Exit Sub
    If InStr(1, UCase(Data.Files(1)), ".PDF") > 0 Then
        obrir_document Data.Files(1)
        wait 1
        If MsgBox("REVISA QUE EL PDF ES EL CORRECTE." + vbNewLine + vbNewLine + "SEGUR QUE EL VOLS ASSIGNAR A [" + atrim(reixa.TextMatrix(0, reixa.col)) + "]?", vbInformation + vbDefaultButton2 + vbYesNo, "ASSIGNACIÓ DE PDF") = vbNo Then Exit Sub
        assignarPDFalacolumna atrim(reixa.TextMatrix(0, reixa.col)), Data.Files(1), generar_nom_pdf
         Else: MsgBox "AQUEST DOCUMENT NO ES UN PDF.", vbCritical, "ERROR"
    End If
End Sub
Function generar_nom_pdf() As String
   generar_nom_pdf = reixa.TextMatrix(reixa.row, reixa.Cols - 1) + "_" + bPDF.Tag + ".pdf"
End Function
Sub assignarPDFalacolumna(vcolumna As String, vfitxer As String, vnom As String)
   Dim vnomfitxer As String
   Dim vnomfitxerDRIVE  As String
   vnomfitxerDRIVE = llegir_ini("ruta", "rutaQualitatProveidorsDRIVE", rutadelfitxer(cami) + "valorsprograma.ini")
   vnomfitxer = llegir_ini("ruta", "rutaQualitatProveidorsLOCAL", rutadelfitxer(cami) + "valorsprograma.ini")
   vnomfitxer = vnomfitxer + "Documentació_Proveidors\" + vnom
   vnomfitxerDRIVE = vnomfitxerDRIVE + "Documentació_Proveidors\" '\" + vnom
   If existeix(vnomfitxer) Then Kill vnomfitxer
   FileCopy vfitxer, vnomfitxer
   If existeix(vnomfitxer) Then
        rstavisos.AddNew
        rstavisos!funcio = "copiar"
        rstavisos!Origen = vnomfitxer
        rstavisos!desti = vnomfitxerDRIVE
        rstavisos.Update
   End If
   
End Sub
Private Sub cllistatipificacions_Click()
   If Mid(cllistatipificacions, 1, 1) = "[" Then
        cllistatipificacions.Selected(cllistatipificacions.ListIndex) = False
        acceptar_llistatipificacions
   End If
   
End Sub

Private Sub cllistatipificacions_DblClick()
   acceptar_llistatipificacions
End Sub
Sub acceptar_llistatipificacions()
Dim v As String
   Dim rst As Recordset
   Dim j As Long
   Dim vpos As Long
   
   Set rst = dbqualitat.OpenRecordset("select * from tipificacionsgeneriques where tipus='" + atrim(cllistatipificacions.Tag) + "'")
   If cllistatipificacions.List(cllistatipificacions.ListIndex) = "[N o v a]" Then
       v = InputBox("Escriu el nom de la nova tipificacio.", "Nova")
       If StrPtr(v) = 0 Then Exit Sub
       v = atrim(v)
       If Mid(v + " ", 1, 1) = "-" Then MsgBox "No pot començar per guió la tipificació.": GoTo fi
       rst.FindFirst "ucase(descripcio)='" + atrim(treure_apostruf(v)) + "'"
       If rst.NoMatch Then
         rst.AddNew: rst!tipus = cllistatipificacions.Tag: rst!descripcio = v: rst.Update
         vpos = cllistatipificacions.ListCount - 2
         cllistatipificacions.AddItem v, vpos
         cllistatipificacions.ListIndex = -1
         cllistatipificacions.Selected(vpos) = True
       End If
   End If
   
   If cllistatipificacions.List(cllistatipificacions.ListIndex) = "[G u a r d a r]" Then
      v = ""
      For j = 0 To cllistatipificacions.ListCount - 1
         If cllistatipificacions.Selected(j) And Mid(cllistatipificacions.List(j), 1, 1) <> "-" Then v = v + "[" + cllistatipificacions.List(j) + "]"
      Next j
      reixa.Text = v
      dbqualitat.Execute "update proveidors_qualitat set " + cllistatipificacions.Tag + "='" + v + "' where codiproveidor=" + atrim(reixa.TextMatrix(reixa.row, reixa.Cols - 1))
      cllistatipificacions.Visible = False
   End If
fi:
   Set rst = Nothing
   If Mid(cllistatipificacions, 1, 1) = "-" Then cllistatipificacions.Selected(cllistatipificacions.ListIndex) = False
End Sub

Private Sub cllistatipificacions_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim rst As Recordset
  If KeyCode = 46 Then
       Set rst = dbqualitat.OpenRecordset("select * from proveidors_qualitat where " + cllistatipificacions.Tag + " like '*[[" + cllistatipificacions + "]]*'")
       If Not rst.EOF Then MsgBox "No es pot eliminar aquesta tipificació perquè ja la tè un proveidors assignada.", vbCritical, "Error": Exit Sub
       If MsgBox("Segur que vols eliminar la tipificació [" + atrim(cllistatipificacions) + "]?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
       dbqualitat.Execute "delete * from tipificacionsgeneriques where descripcio='" + atrim(cllistatipificacions) + "' and tipus='" + atrim(cllistatipificacions.Tag) + "'"
       cllistatipificacions.RemoveItem cllistatipificacions.ListIndex
  End If
  If KeyCode = 113 Then
    If cllistatipificacions.ListIndex = -1 Then MsgBox "Has d'escullir una tipificació per editar.", vbCritical, "Error": Exit Sub
       Set rst = dbqualitat.OpenRecordset("select * from tipificacionsgeneriques where tipus='" + atrim(cllistatipificacions.Tag) + "'")
       rst.FindFirst "descripcio='" + cllistatipificacions + "'"
       If Not rst.NoMatch Then
           v = InputBox("Modifica la tipificació." + vbNewLine + "Borra-la per eliminar aquesta tipificació.", "Modificar/Eliminar", cllistatipificacions)
           v = treure_apostruf(v)
           If StrPtr(v) <> 0 Then
            If atrim(v) <> "" Then
               rst.Edit: rst!descripcio = v: rst.Update
               cllistatipificacions.List(cllistatipificacions.ListIndex) = v
                Else: rst.Delete: cllistatipificacions.RemoveItem cllistatipificacions.ListIndex
            End If
           End If
       End If
     
  End If
  Set rst = Nothing
End Sub

Private Sub Command56_Click()
   Dim i As Byte
   For i = 0 To filtre.Count - 1: filtre(i) = "": Next i
   aplicar_filtre
End Sub

Private Sub exportaraxls_Click()
   Dim vcol As Double
   Dim vrow As Double
   Dim vlinia As String
   Dim vnomfitxer As String
   
   vnomfitxer = "c:\temp\qualitat_proveidors.csv"
   reixa.Redraw = False
   Open vnomfitxer For Output As #1
    vrow = 0
   While vrow < reixa.Rows
      vcol = 0
      While vcol < reixa.Cols
         reixa.col = vcol
         reixa.row = vrow
         vlinia = vlinia + IIf(vlinia <> "", ";", "") + IIf(reixa.CellFontUnderline, "*", "") + reixa.TextMatrix(vrow, vcol)
         vcol = vcol + 1
      Wend
      Print #1, vlinia
      vlinia = ""
      vrow = vrow + 1
   Wend
   Close 1
   If existeix(vnomfitxer) Then obrir_document vnomfitxer
   reixa.Redraw = True
End Sub

Private Sub filtre_LostFocus(Index As Integer)
  aplicar_filtre
End Sub
Sub aplicar_filtre()
   Dim rst As Recordset
   Static vultimordre As String
   Dim i As Byte
   Set rst = dbqualitat.OpenRecordset("select * from proveidors_qualitat")
   vweres = ""
   For i = 0 To filtre.Count - 1
      If atrim(filtre(i) <> "") Then
         vweres = vweres + IIf(vweres = "", "", " and ") + "trim(" + rst.Fields(i + 1).Name + ") like '*" + treure_apostruf(atrim(filtre(i))) + "*'"
      End If
   Next i
   If i < 14 Then
    vordrereixa = " order by " + rst.Fields(cadbl(vordrereixa) + 1).Name
    If vultimordre = vordrereixa Then
            vordrereixa = vordrereixa + " DESC": vultimordre = ""
    End If
   End If
   Set rst = Nothing
   If LCase(Mid(vordrereixa, 1, 6)) <> " order" Or reixa.row = 0 Then
       vultimordre = vordrereixa
       poblarlareixa
       carregar_amples_reixa
   End If

End Sub
Private Sub Form_Load()
   iniconfigreixa = "c:\windows\proveidors_qualitat.ini"
   Set dbqualitat = OpenDatabase(rutadelfitxer(cami) + "qualitat.mdb")
   Set dbavisos = OpenDatabase(rutadelfitxer(cami) + "avisosincidencies.mdb")
   Set rstavisos = dbavisos.OpenRecordset("select * from ordres_execucio_servidor")
   vcampsPDF = "[GFSI][NO GFSI] [Fi Bones practiques] [Fi Questionari] "
   vweres = ""
   amplesform
   actualitzar_registres
   poblarlareixa
   carregar_amples_reixa
   
   
End Sub
Sub possar_llista_tipificacions(vtipus As String, vtipificacions As String, Optional vnompossaropcions As Boolean)
  Dim rst As Recordset
  cllistatipificacions.Clear
  cllistatipificacions.Tag = vtipus
  
  Set rst = dbqualitat.OpenRecordset("select * from Tipificacionsgeneriques where tipus='" + cllistatipificacions.Tag + "' order by descripcio")
  While Not rst.EOF
    cllistatipificacions.AddItem rst!descripcio
     If InStr(1, UCase(vtipificacions), UCase("[" + atrim(rst!descripcio) + "]")) > 0 Then cllistatipificacions.Selected(cllistatipificacions.NewIndex) = True
    rst.MoveNext
  Wend
  If Not vnompossaropcions Then
    cllistatipificacions.AddItem "================="
    cllistatipificacions.AddItem "[N o v a]"
    cllistatipificacions.AddItem "[Editar] F2"
    cllistatipificacions.AddItem "[G u a r d a r]"
  End If
  Set rst = Nothing
  cllistatipificacions.Visible = True
  cllistatipificacions.Top = reixa.CellTop + reixa.CellHeight + reixa.Top
  If cllistatipificacions.Top + cllistatipificacions.Height > (reixa.Height + reixa.Top) Then
     cllistatipificacions.Top = (reixa.CellTop + reixa.Top) - cllistatipificacions.Height
  End If
  cllistatipificacions.Left = reixa.CellLeft + reixa.Left
  cllistatipificacions.Width = reixa.CellWidth
  If cllistatipificacions.Width < 1000 Then cllistatipificacions.Width = 2000
End Sub
Sub actualitzar_registres()
  Dim vsql As String
  Dim rst As Recordset
  Dim rstq As Recordset
  vsql = "SELECT proveidors.codi, proveidors_qualitat.codiproveidor, proveidors.tipusCQ, proveidors.dataCQ, proveidors.databaixa FROM proveidors_qualitat RIGHT JOIN proveidors ON proveidors_qualitat.codiproveidor = proveidors.codi WHERE (((proveidors_qualitat.codiproveidor) Is Null) AND ((proveidors.databaixa) Is Null));"
  Set rst = dbqualitat.OpenRecordset(vsql, , dbReadOnly)
  Set rstq = dbqualitat.OpenRecordset("select * from proveidors_qualitat")
  While Not rst.EOF
     rstq.AddNew
     rstq!codiproveidor = rst!codi
     rstq!tipuscontrolCQ = IIf(rst!tipusCQ = "L", "Certificado por Lote", IIf(rst!tipusCQ = "C", "Calidad concertada", ""))
     rstq!CQ_datacaducitat = IIf(IsDate(rst!dataCQ), rst!dataCQ, Null)
     rstq.Update
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Sub poblarlareixa()
   Dim rst As Recordset
   Dim vcol As Double
   Dim vrow As Double
   Dim vdatacasella As String
   
   Set rst = dbqualitat.OpenRecordset("select * from proveidors_qualitat " + IIf(vweres <> "", " where " + vweres, "") + vordrereixa)
   reixa.Redraw = False
   config_reixa
   vrow = 0
   While Not rst.EOF
      vrow = vrow + 1
      If reixa.Rows = vrow Then reixa.Rows = vrow + 1
      For vcol = 1 To rst.Fields.Count - 1
        reixa.TextMatrix(vrow, vcol - 1) = atrim(rst.Fields(vcol))
        If InStr(1, vcampsPDF, "[" + reixa.TextMatrix(0, vcol - 1) + "]") > 0 Then
            If hihaPDF(rst.Fields("codiproveidor"), reixa.TextMatrix(0, vcol - 1)) Then
                   reixa.col = vcol - 1: reixa.row = vrow
                   reixa.CellFontUnderline = True
                   reixa.CellForeColor = &HFFAE5E
            End If
        End If
        posar_color_casella vcol, vrow, atrim(rst.Fields(vcol))
      Next vcol
      reixa.TextMatrix(vrow, vcol - 1) = rst!codiproveidor
      rst.MoveNext
   Wend
   reixa.Redraw = True
   Set rst = Nothing
End Sub
Sub posar_color_casella(vcol As Double, vrow As Double, vdatacasella As String)
      reixa.col = vcol - 1: reixa.row = vrow: reixa.CellBackColor = formproveidorsqualitat.BackColor
        If Mid(vdatacasella + "     ", 1, 4) = "Sol:" Then
            vdatacasella = Mid(vdatacasella, 5, 9)
        End If
        If IsDate(vdatacasella) Then
             If DateDiff("d", vdatacasella, Now) > 0 Then
                    reixa.col = vcol - 1: reixa.row = vrow
                    reixa.CellBackColor = &H5C31DD 'vermell xulu
                Else
                   If DateDiff("d", vdatacasella, Now) > -15 Then
                        reixa.col = vcol - 1: reixa.row = vrow
                        reixa.CellBackColor = &H80FF&     'taronja xulu
                    Else
                        If DateDiff("d", vdatacasella, Now) > -30 Then
                               reixa.col = vcol - 1: reixa.row = vrow
                               reixa.CellBackColor = &HFFFF&     'groc xulu
                        End If
                   End If
            End If
        End If
End Sub
Function hihaPDF(vnomproveidor As String, vnomcamp As String) As Boolean
   Dim vnomfitxer As String
   Dim vnomdelpdf As String
   vnomdelpdf = vnomproveidor + "_" + vnomcamp + ".pdf"
   vnomfitxer = llegir_ini("ruta", "rutaQualitatProveidorsLOCAL", rutadelfitxer(cami) + "valorsprograma.ini")
   vnomfitxer = vnomfitxer + "Documentació_Proveidors\" + vnomdelpdf
   If existeix(vnomfitxer) Then
      hihaPDF = True
   End If
End Function
Sub config_reixa()
  reixa.Clear
   reixa.Cols = 1
   reixa.Rows = 1
   reixa.Cols = 13
   reixa.TextMatrix(0, 0) = "Nom PROVEIDOR"
   reixa.TextMatrix(0, 1) = "Tipus Productes"
   reixa.TextMatrix(0, 2) = "PRL"
   reixa.TextMatrix(0, 3) = "Tipus CQ"
   reixa.TextMatrix(0, 4) = "Caducitat tipus CQ"
   reixa.TextMatrix(0, 5) = "Rec. Questionari"
   reixa.TextMatrix(0, 6) = "Cad. Questionari"
  ' reixa.TextMatrix(0, 6) = "NO GFSI"
  ' reixa.TextMatrix(0, 7) = "Caducitat NO GFSI"
   reixa.TextMatrix(0, 7) = "Rec. Bones_Practiques"
   reixa.TextMatrix(0, 8) = "Cad. Bones_Practiques"
   reixa.TextMatrix(0, 9) = "Rec. Traçabilitat"
   reixa.TextMatrix(0, 10) = "Cad. Traçabilitat"
   reixa.TextMatrix(0, 11) = "Observacions"
   reixa.TextMatrix(0, 12) = "Codiproveidor"
   reixa.ColWidth(11) = 0
   
   
End Sub

Sub amplesform()
 If cadbl(llegir_ini("TamanyForm", "ample", iniconfigreixa)) > 0 Then
   formproveidorsqualitat.Tag = "canvianttamany"
   formproveidorsqualitat.Width = llegir_ini("TamanyForm", "ample", iniconfigreixa)
   formproveidorsqualitat.Height = llegir_ini("TamanyForm", "alt", iniconfigreixa)
   If cadbl(llegir_ini("PosicioForm", "Left", iniconfigreixa)) > 0 Then
     formproveidorsqualitat.Left = cadbl(llegir_ini("PosicioForm", "Left", iniconfigreixa))
     formproveidorsqualitat.Top = cadbl(llegir_ini("PosicioForm", "Top", iniconfigreixa))
   End If
   formproveidorsqualitat.Tag = ""
  End If
End Sub
Private Sub Form_Resize()

   If formproveidorsqualitat.Height - reixa.Top - 800 < 1 Then Exit Sub
   reixa.Width = formproveidorsqualitat.Width - 400
   reixa.Height = formproveidorsqualitat.Height - reixa.Top - 800
'   Fbotons.Left = formproveidorsqualitat.Width - Fbotons.Width - 300
 '  etregistres.Top = reixa.Height + reixa.Top
   If formproveidorsqualitat.Tag <> "canvianttamany" Then
       guardar_posicions_finestre
   End If
End Sub
Sub guardar_posicions_finestre()
    escriure_ini "TamanyForm", "ample", atrim(formproveidorsqualitat.Width), iniconfigreixa
    escriure_ini "TamanyForm", "alt", atrim(formproveidorsqualitat.Height), iniconfigreixa
    escriure_ini "PosicioForm", "Left", atrim(formproveidorsqualitat.Left), iniconfigreixa
    escriure_ini "PosicioForm", "Top", atrim(formproveidorsqualitat.Top), iniconfigreixa
    
End Sub

Sub guardar_amples_reixa()
Dim j As Integer
If iniconfigreixa <> "" Then
  For j = 0 To reixa.Cols - 1
   escriure_ini "AmplesReixaQualitat", UCase(reixa.TextMatrix(0, j)), atrim(Redondejar(reixa.ColWidth(j), 0)), iniconfigreixa
 Next j
End If
End Sub
Sub carregar_amples_reixa()
Dim j As Integer
Dim x As Double
If iniconfigreixa <> "" Then
  x = reixa.Left + 35
  If cadbl(llegir_ini("AmplesReixaQualitat", UCase(reixa.TextMatrix(0, j)), iniconfigreixa)) = 0 Then Exit Sub
  For j = 0 To reixa.Cols - 1
   vamplecol = cadbl(llegir_ini("AmplesReixaQualitat", UCase(reixa.TextMatrix(0, j)), iniconfigreixa))
   If vamplecol = 0 Then vamplecol = 500
   reixa.ColWidth(j) = vamplecol
   If j = filtre.Count Then Load filtre(j)
   filtre(j).Width = reixa.ColWidth(j)
   If filtre(j).Width = 150 Then
          filtre(j).Visible = False
            Else: filtre(j).Visible = True
   End If
   filtre(j).Left = x
   x = x + cadbl(reixa.ColWidth(j))
 Next j
 j = 15
End If
DoEvents
reixa.row = 0
reixa.col = 0
reixa.ColSel = 0
reixa.RowSel = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set controlcanviat = Nothing
guardar_amples_reixa
End Sub

Private Sub reixa_Click()
  If reixa.row = 1 And reixa.RowSel = reixa.Rows - 1 Then
      reixa.row = 0
      aplicar_filtre
  End If
End Sub

Private Sub reixa_DblClick()
  Dim v As String
  Dim vresp As Long
  If reixa.TextMatrix(0, reixa.col) = "PRL" Then
     vresp = MsgBox("Aquest proveidor té PRL?", vbExclamation + vbDefaultButton2 + vbYesNo, "PRL?")
     v = ""
     If vresp = vbYes Then v = "S"
     If vresp = vbNo Then v = "N"
     dbqualitat.Execute "update proveidors_qualitat set PRL='" + v + "' where codiproveidor=" + atrim(reixa.TextMatrix(reixa.row, reixa.Cols - 1))
     reixa.Text = v
  End If
  If reixa.TextMatrix(0, reixa.col) = "Tipus CQ" Then
     v = UCase(InputBox("Escriu " + vbNewLine + "[L]-Certificado por Lote." + vbNewLine + "[C]-Calidad concertada." + vbNewLine + "[R]-Sense.", "Tipus Control"))
     If v = "L" Then
        v = "Certificado por Lote"
         Else
            If v = "C" Then
                v = "Calidad concertada"
              Else:
                 If v = "R" Then
                      v = "R"
                       Else: v = ""
                 End If
            End If
     End If
     If v <> "" Then
           If v = "R" Then v = ""
           dbqualitat.Execute "update proveidors_qualitat set tipuscontrolCQ='" + v + "' where codiproveidor=" + atrim(reixa.TextMatrix(reixa.row, reixa.Cols - 1))
           reixa.Text = v
     End If
  End If
  If reixa.TextMatrix(0, reixa.col) = "Caducitat tipus CQ" Then
      If reixa.TextMatrix(reixa.row, reixa.col - 1) = "Calidad concertada" Then
          v = InputBox("Entra la data de caducitat de la concertada", "Data caducitat concertada.", reixa.Text)
          If IsDate(v) Then
             dbqualitat.Execute "update proveidors_qualitat set CQ_datacaducitat=#" + Format(v, "mm/dd/yy") + "# where codiproveidor=" + atrim(reixa.TextMatrix(reixa.row, reixa.Cols - 1))
             reixa.Text = v
          End If
           Else: MsgBox "No hi ha data de calidad concertada no "
      End If
  End If
  If reixa.TextMatrix(0, reixa.col) = "Tipus Productes" Then
      possar_llista_tipificacions "tipusproductes", reixa.Text
  End If
  If reixa.TextMatrix(0, reixa.col) = "GFSI" Then
      possar_llista_tipificacions "GFSI", reixa.Text, True
  End If
  If reixa.TextMatrix(0, reixa.col) = "NO GFSI" Then
      possar_llista_tipificacions "NOGFSI", reixa.Text
  End If

End Sub


Private Sub reixa_GotFocus()
 'Dim vcol As Double
 'vcol = reixa.col
 'reixa.col = 0: reixa.ColSel = reixa.Cols - 1
 
End Sub

Private Sub reixa_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim vdatacaducitat As String
   Dim rst As Recordset
   If y < reixa.RowHeight(0) Then
     ordenar_reixa
   End If
   If Button = 2 Then
     If Mid(reixa.TextMatrix(0, reixa.col), 1, 4) = "Rec." Then
            mSol.Tag = "": mRec.Tag = ""
            Me.PopupMenu menuRoS, , x, y + reixa.Top + reixa.RowHeight(reixa.row)
            If mSol.Tag <> "" Then
                Set rst = dbqualitat.OpenRecordset("select * from proveidors_qualitat")
                If InStr(1, reixa.Text, "Sol:") > 0 Then
                     If InStr(1, reixa.Text, mSol.Tag) = 0 Then reixa.Text = "Sol: " + mSol.Tag + Mid(reixa.Text, 5)
                      Else: reixa.Text = "Sol: " + mSol.Tag
                End If
                posar_color_casella reixa.col + 1, reixa.row, reixa.Text
                'If mSol.Tag = " " Then reixa.Text = ""
                dbqualitat.Execute "update proveidors_qualitat set " + atrim(rst.Fields(reixa.col + 1).Name) + "='" + reixa.Text + "' where codiproveidor=" + atrim(reixa.TextMatrix(reixa.row, reixa.Cols - 1))
            End If
            If mRec.Tag <> "" Then
                reixa.Text = mRec.Tag
                posar_color_casella reixa.col + 1, reixa.row, reixa.Text
                vdatacaducitat = InputBox("Entra la data de caducitat:", "Caducitat", Format(DateAdd("yyyy", 3, mRec.Tag), "dd/mm/yy"))
                If IsDate(vdatacaducitat) Then
                       reixa.TextMatrix(reixa.row, reixa.col + 1) = vdatacaducitat
                       posar_color_casella reixa.col + 2, reixa.row, vdatacaducitat
                End If
                Set rst = dbqualitat.OpenRecordset("select * from proveidors_qualitat")
                dbqualitat.Execute "update proveidors_qualitat set " + atrim(rst.Fields(reixa.col + 1).Name) + "='" + reixa.Text + "' where codiproveidor=" + atrim(reixa.TextMatrix(reixa.row, reixa.Cols - 1))
                dbqualitat.Execute "update proveidors_qualitat set " + atrim(rst.Fields(reixa.col + 2).Name) + "='" + reixa.TextMatrix(reixa.row, reixa.col + 1) + "' where codiproveidor=" + atrim(reixa.TextMatrix(reixa.row, reixa.Cols - 1))
            End If
      End If
   End If
   Set rst = Nothing
End Sub

Sub ordenar_reixa()
   vordrereixa = reixa.col
   reixa.Redraw = False
   aplicar_filtre
   If reixa.Rows > 0 Then reixa.TopRow = 1
   reixa.Redraw = True
End Sub
Private Sub reixa_RowColChange()
   cllistatipificacions.Visible = False
   bPDF.Visible = False
   If InStr(1, vcampsPDF, "[" + reixa.TextMatrix(0, reixa.col) + "]") > 0 Then
     bPDF.Visible = True
     bPDF.Left = reixa.CellLeft + reixa.CellWidth - bPDF.Width + reixa.Left
     bPDF.Top = reixa.CellTop + reixa.Top
     bPDF.Tag = UCase(reixa.TextMatrix(0, reixa.col))
   End If
End Sub
