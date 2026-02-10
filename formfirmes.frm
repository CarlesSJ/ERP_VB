VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formfirmes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Firmes"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   7515
   Icon            =   "formfirmes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Tag             =   "2"
   Begin VB.CommandButton bFitxaTecnica 
      BackColor       =   &H006BEBB1&
      Caption         =   "Fitxa Tècnica"
      Height          =   1230
      Left            =   6705
      OLEDropMode     =   1  'Manual
      Picture         =   "formfirmes.frx":048A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Fitxa Tècnica de la referencia. (Arrastrar PDF) (CTRL+CLICK Eliminar)"
      Top             =   90
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   5010
      TabIndex        =   14
      Top             =   15
      Width           =   1545
      Begin VB.Label etcomanda 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   15
         Top             =   -15
         Width           =   60
      End
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0000C000&
      Caption         =   "Plaç"
      Height          =   600
      Left            =   4245
      Picture         =   "formfirmes.frx":09D9
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Firmar packinglist"
      Top             =   750
      Width           =   825
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00008000&
      Caption         =   "FINAL"
      Height          =   600
      Left            =   5070
      Picture         =   "formfirmes.frx":0E63
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Firmar packinglist"
      Top             =   750
      Width           =   825
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H006BEBB1&
      Caption         =   "Tècnic"
      Height          =   600
      Left            =   3420
      Picture         =   "formfirmes.frx":12ED
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Firmar packinglist"
      Top             =   750
      Width           =   825
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF00&
      Caption         =   "PVP"
      Height          =   600
      Left            =   2595
      Picture         =   "formfirmes.frx":1777
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Firmar packinglist"
      Top             =   750
      Width           =   825
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pendent de firmar per mi"
      Height          =   705
      Left            =   135
      TabIndex        =   5
      Top             =   30
      Width           =   6525
      Begin VB.CommandButton bultimaconsulta 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         Picture         =   "formfirmes.frx":1C01
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ultima consulta feta"
         Top             =   255
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FDDECE&
         Caption         =   "Amb PVP"
         Height          =   420
         Left            =   4845
         TabIndex        =   9
         Top             =   210
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FDDECE&
         Caption         =   "Noves i Modif."
         Height          =   420
         Left            =   3375
         TabIndex        =   8
         Top             =   210
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FDDECE&
         Caption         =   "PK2"
         Height          =   420
         Left            =   1905
         TabIndex        =   7
         Top             =   210
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FDDECE&
         Caption         =   "Fulles Principals"
         Height          =   420
         Left            =   435
         TabIndex        =   6
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.CommandButton eliminar 
      Height          =   450
      Left            =   6180
      Picture         =   "formfirmes.frx":1D0C
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Borrar una firma"
      Top             =   1020
      Width           =   465
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Pack'Lst"
      Height          =   600
      Left            =   1770
      Picture         =   "formfirmes.frx":2296
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Firmar packinglist"
      Top             =   750
      Width           =   825
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "I.M.P."
      Height          =   600
      Left            =   945
      Picture         =   "formfirmes.frx":2720
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Firmar IMP"
      Top             =   750
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comanda"
      Height          =   600
      Left            =   120
      Picture         =   "formfirmes.frx":2BAA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Firmar comanda"
      Top             =   750
      Width           =   825
   End
   Begin VB.Data datafirmes 
      Caption         =   "datafirmes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1530
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "comandes_firmes"
      Top             =   5460
      Visible         =   0   'False
      Width           =   2865
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "formfirmes.frx":3034
      Height          =   5175
      Left            =   90
      OleObjectBlob   =   "formfirmes.frx":3049
      TabIndex        =   0
      Top             =   1485
      Width           =   7290
   End
   Begin VB.Menu mpendents 
      Caption         =   "Pendents de firmar"
      Visible         =   0   'False
      Begin VB.Menu mactivadessensefirma 
         Caption         =   "Activades sense la meva firma"
      End
      Begin VB.Menu mambpreuisenselamevafirma 
         Caption         =   "Amb preu i sense la meva firma"
      End
   End
   Begin VB.Menu mutils 
      Caption         =   "Utils"
      Begin VB.Menu mfirmamassivaardo 
         Caption         =   "Firma massiva ARDO"
      End
      Begin VB.Menu mboarrarfirmes 
         Caption         =   "Borrar firmes"
         Begin VB.Menu mborrartotes 
            Caption         =   "Borrar totes les firmes d'una comanda"
         End
      End
      Begin VB.Menu mfirmesborrades 
         Caption         =   "Veure firmes borrades"
      End
   End
End
Attribute VB_Name = "formfirmes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vprimercop As Boolean

Private Sub bFitxaTecnica_Click()
    Dim vnomfitxer As String
    vnomfitxer = NomFitxaTecnica(cadbl(etcomanda))
    If existeix(vnomfitxer) Then obrir_document vnomfitxer
End Sub

Private Sub bFitxaTecnica_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim vnomfitxer As String
  If Shift = 2 Then
     If MsgBox("Segur que vols eliminar aquesta fitxa tècnica?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
        vnomfitxer = NomFitxaTecnica(cadbl(etcomanda))
        If existeix(vnomfitxer) Then Kill vnomfitxer: wait 1
        refrescar_firmes
     End If
  End If
End Sub

Private Sub bFitxaTecnica_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim vnomfitxer As String
   Dim vnomfitxer_desti As String
   vnomfitxer_desti = NomFitxaTecnica(cadbl(etcomanda))
   vnomfitxer = UCase(Data.Files(1))
   assignar_FT vnomfitxer, vnomfitxer_desti  'NomFitxaTecnica(cadbl(etcomanda))
End Sub
Sub assignar_FT(vnomfitxer As String, vnomfitxer_desti As String)
   If InStr(1, vnomfitxer, ".PDF") = 0 Then MsgBox "El fitxer ha de ser PDF", vbCritical, "Error": Exit Sub
   If existeix(vnomfitxer) Then
       If existeix(vnomfitxer_desti) Then
            Kill vnomfitxer_desti
       End If
       Copiar_Fitxer vnomfitxer, vnomfitxer_desti
       wait 1
       refrescar_firmes
   End If
End Sub

Private Sub bultimaconsulta_Click()
   Dim vcom As String
   vcom = llegir_ini("Firmes", "ultimaconsultafirmes", "comandes.ini")
   If Len(vcom) > 4 Then activadessensefirma "", ""
End Sub

Private Sub Command1_Click()
   Dim vx As Double
   Dim vy As Double
 
   If comprovarsijaestafirmat("COM") Then Exit Sub
   datafirmes.Recordset.AddNew
   datafirmes.Recordset!comanda = cadbl(etcomanda)
   datafirmes.Recordset!usuari = nomordinador
   datafirmes.Recordset!Data = Now
   datafirmes.Recordset!tipus = "COM"
   datafirmes.Recordset.Update


   vx = formfirmes.Left: vy = formfirmes.Top
   formfirmes.Visible = False
   formcomandes.botoconsultar
   wait 1
   If cadbl(etcomanda) <> cadbl(formcomandes.Text1) Then
       formfirmes.Show
      formfirmes.Visible = True
      formfirmes.Left = vx: formfirmes.Top = vy
       Else: Unload formfirmes
   End If
End Sub
Function comprovarsijaestafirmat(vtipus As String) As Boolean
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from comandes_firmes where dataanulacio=null and comanda=" + atrim(etcomanda) + " and tipus='" + vtipus + "'")
   If rst.EOF Then
          comprovarsijaestafirmat = False
          
         Else:
           If vtipus = "PVP" Then
              rst.MoveLast
              If rst.RecordCount >= 3 Then
                MsgBox "El PVP ja l'han firmat tres usuaris no cal firmar mes.", vbCritical, "Error"
                comprovarsijaestafirmat = True
                 Else: comprovarsijaestafirmat = False
              End If
           End If
           If vtipus = "IM2" Then
              rst.MoveLast
              If rst.RecordCount >= 2 Then
                MsgBox "El IMP ja l'han firmat dos usuaris no cal firmar mes.", vbCritical, "Error"
                comprovarsijaestafirmat = True
                 Else: comprovarsijaestafirmat = False
              End If
           End If
            If vtipus <> "IM2" And vtipus <> "PVP" Then
                 comprovarsijaestafirmat = True
                 MsgBox "El " + vtipus + " ja l'han firmat no cal firmar mes.", vbCritical, "Error"
           End If
   End If
   Set rst = Nothing
End Function
Private Sub DBGrid1_Click()

End Sub

Function escullir_impanterior(vnumtreball As Double, vordre As Double, vclient As Double, vdirenvio As Double) As String
   Unload formseleccio
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "clixesnous.mdb"
  formseleccio.Data1.RecordSource = "select direnvio,NOMCLIENT,NOMDIRENVIO from CLIENTSVINCULATS where id_treball=" + atrim(vnumtreball) + " and ordremodificacio=" + atrim(vordre) + " order by nomclient "
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(1).Width = 3000
  formseleccio.DBGrid2.Columns(2).Width = 2000
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.Width = 8000
  If formseleccio.Data1.Recordset.EOF Then Exit Function
  formseleccio.Data1.Recordset.MoveLast: formseleccio.Data1.Recordset.MoveFirst
  If formseleccio.Data1.Recordset.RecordCount > 1 Then
        formseleccio.Show 1
          Else: seleccioret = 1
  End If
  If seleccioret = 1 Then
     vdirenvio = cadbl(formseleccio.Data1.Recordset!direnvio)
     escullir_impanterior = ruta_documentacio_clixes + "\" + Format(vnumtreball, "00000") + "\IMP" + Format(vnumtreball, "00000") + "-" + Format(vordre, "000") + "-" + Format(vclient, "000000") + "_" + atrim(vdirenvio) + ".doc"
     If Not existeix(escullir_impanterior) Then escullir_impanterior = escullir_impanterior + "x"
  End If
  Unload formseleccio
End Function

Private Sub Command10_Click()
  If comprovarsijaestafirmat("FIN") Then Exit Sub
  If MsgBox("Estàs a punt de firmar LA FINALITZACIO DE LES FIRMES de la comanda, es correcte?" + observacioPVP, vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
   datafirmes.Recordset.AddNew
   datafirmes.Recordset!comanda = cadbl(etcomanda)
   datafirmes.Recordset!usuari = nomordinador
   datafirmes.Recordset!Data = Now
   datafirmes.Recordset!tipus = "FIN"
   datafirmes.Recordset.Update
  End If
End Sub

Private Sub Command11_Click()
If comprovarsijaestafirmat("PLA") Then Exit Sub
  If MsgBox("Estàs a punt de firmar PLAÇ de la comanda, es correcte?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
   datafirmes.Recordset.AddNew
   datafirmes.Recordset!comanda = cadbl(etcomanda)
   datafirmes.Recordset!usuari = nomordinador
   datafirmes.Recordset!Data = Now
   datafirmes.Recordset!tipus = "PLA"
   datafirmes.Recordset.Update
  End If
End Sub

Private Sub Command2_Click()
  Dim vnomfitxer As String
  Dim vruta As String
  If comprovarsijaestafirmat("IM2") Then Exit Sub
  formcomandes.obrir_imp_treball cadbl(formcomandes.Data1.Recordset!numtreball), cadbl(formcomandes.Data1.Recordset!numordremodificacio), cadbl(formcomandes.Data1.Recordset!client), cadbl(formcomandes.Data1.Recordset!direnvio)
  If cadbl(formcomandes.Data1.Recordset!numordremodificacio) > 1 Then
     vnomfitxer = "noavisar"
     formcomandes.obrir_imp_treball cadbl(formcomandes.Data1.Recordset!numtreball), cadbl(formcomandes.Data1.Recordset!numordremodificacio) - 1, cadbl(formcomandes.Data1.Recordset!client), cadbl(formcomandes.Data1.Recordset!direnvio), vnomfitxer
     'si no existeix el imp anterior vol una llista de destins de la versio anterior per escullir
     If Not existeix(vnomfitxer) Then
        vnomfitxer = escullir_impanterior(cadbl(formcomandes.Data1.Recordset!numtreball), cadbl(formcomandes.Data1.Recordset!numordremodificacio) - 1, cadbl(formcomandes.Data1.Recordset!client), cadbl(formcomandes.Data1.Recordset!direnvio))
        If Not existeix(vnomfitxer) Then vnomfitxer = vnomfitxer + "x"
        If existeix(vnomfitxer) Then
           obrir_document vnomfitxer
        End If
     End If
  End If
  vruta = ruta_documentacio_clixes + "\" + Format(cadbl(formcomandes.Data1.Recordset!numtreball), "00000") + "\Arxiu_documentacio_relacionada" + "\v" + atrim(cadbl(formcomandes.Data1.Recordset!numordremodificacio))
  If existeix(vruta) Then idp = ShellExecute(Me.hwnd, "Open", "c:\windows\explorer.exe", " " + vruta, "", 1)
  If MsgBox("Estàs a punt de firmar el IMP del treball, es correcte?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
  
     datafirmes.Recordset.AddNew
   datafirmes.Recordset!comanda = cadbl(etcomanda)
   datafirmes.Recordset!usuari = nomordinador
   datafirmes.Recordset!Data = Now
   datafirmes.Recordset!tipus = "IM2"
   datafirmes.Recordset.Update
  End If
End Sub

Private Sub Command3_Click()
   Dim vnumc As String
   Do
     firmar_packinglist
     vnumc = InputBox("Vols firmar una altra comanda?" + vbNewLine + "ESCRIU LA COMANDA SI VOLS FER UNA ALTRA O RES PER SORTIR", "FIRMAR UNA ALTRA COMANDA")
     If cadbl(vnumc) > 0 Then carregar_comanda_i_firmes vnumc
   Loop Until cadbl(vnumc) = 0
End Sub
Sub firmar_packinglist()
 If comprovarsijaestafirmat("PK2") Then Exit Sub
   datafirmes.Recordset.FindFirst "tipus='PK1'"
   If datafirmes.Recordset.NoMatch Then MsgBox "No pots firmar el PK2 si encara no hi ha el PK1 assignat.", vbCritical, "Atenció": Exit Sub
   datafirmes.Recordset.FindFirst "tipus='PK1' and usuari='" + nomordinador + "'"
   If Not datafirmes.Recordset.NoMatch Then MsgBox "No pot firmar el packing-list dues vegades la mateixa persona.", vbCritical, "Error": Exit Sub
   escriure_ini "baixes", "imprimirpackinglist", "1", "comandes.ini"
   Shell rutadelfitxer(llegir_ini("General", "rutaprogbaixes", fitxerini)) + "palets.exe comandes.ini " + atrim(cadbl(etcomanda) * -1), vbNormalFocus

  If MsgBox("Estàs a punt de firmar el PACKINGLIST de la comanda, es correcte?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
   datafirmes.Recordset.AddNew
   datafirmes.Recordset!comanda = cadbl(etcomanda)
   datafirmes.Recordset!usuari = nomordinador
   datafirmes.Recordset!Data = Now
   datafirmes.Recordset!tipus = "PK2"
   datafirmes.Recordset.Update
  End If
  escriure_ini "baixes", "imprimirpackinglist", "0", "comandes.ini"
End Sub

Private Sub Command4_Click()
    escriure_ini "Firmes", "ultimaconsultafirmes", "", "comandes.ini"
    'activadessensefirma "(producte<>'PC' and producte <>'PC2' and producte<>'PCP') and", " tipus='INI' "
    activadessensefirma "(producte<>'PC' and producte <>'PC2' and producte<>'PCP') and", " tipus='INI' and comanda not in (select comanda from comandes_firmes_actives where tipus='COM')"
     
End Sub

Private Sub Command5_Click()
escriure_ini "Firmes", "ultimaconsultafirmes", "", "comandes.ini"
'activadessensefirma "(producte='PC' or producte ='PC2' or producte='PCP') and", " tipus='PK1' and comanda not in (select comanda from comandes_firmes where tipus='PK2') "
activadessensefirma " proximaseccio<>'T' AND ", " usuari<>'" + nomordinador + "' AND tipus='PK1'  and comanda not in (select comanda from comandes_firmes_actives where tipus='PK2') "
End Sub

Private Sub Command6_Click()
escriure_ini "Firmes", "ultimaconsultafirmes", "", "comandes.ini"
  activadessensefirma "(impressio='N' or impressio ='M') and", "tipus='IM1'"
End Sub

Private Sub Command7_Click()
escriure_ini "Firmes", "ultimaconsultafirmes", "", "comandes.ini"
  ambpreuisenselamevafirma
End Sub

Private Sub Command8_Click()
  If comprovarsijaestafirmat("PVP") Then Exit Sub
  If MsgBox("Estàs a punt de firmar que el PVP de la comanda està bé, es correcte?" + observacioPVP, vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
   datafirmes.Recordset.AddNew
   datafirmes.Recordset!comanda = cadbl(etcomanda)
   datafirmes.Recordset!usuari = nomordinador
   datafirmes.Recordset!Data = Now
   datafirmes.Recordset!tipus = "PVP"
   datafirmes.Recordset.Update
  End If
End Sub
Function observacioPVP() As String
    If formcomandes.Command26(6).ToolTipText <> "" Then
       observacioPVP = vbNewLine + "OBSERVACIÓ PVP:" + vbNewLine + formcomandes.Command26(6).ToolTipText
    End If
End Function
Private Sub Command9_Click()
If comprovarsijaestafirmat("TEC") Then Exit Sub
  If MsgBox("Estàs a punt de firmar TÈCNIC de la comanda, es correcte?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
   datafirmes.Recordset.AddNew
   datafirmes.Recordset!comanda = cadbl(etcomanda)
   datafirmes.Recordset!usuari = nomordinador
   datafirmes.Recordset!Data = Now
   datafirmes.Recordset!tipus = "TEC"
   datafirmes.Recordset.Update
  End If
End Sub

Private Sub eliminar_Click()
  If datafirmes.Recordset.EOF Then
      MsgBox "Has d'escullir una firma.", vbCritical, "error"
      Exit Sub
  End If
  If datafirmes.Recordset!anulada Then MsgBox "Ja està eliminada aquesta firma.", vbCritical, "Error": Exit Sub
  If nomordinador <> datafirmes.Recordset!usuari Then MsgBox "Només pot treure la firma la persona que l'ha posat.", vbCritical, "Error": Exit Sub
  If MsgBox("Segur que vols eliminar aquesta firma?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
      datafirmes.Recordset.Edit
      datafirmes.Recordset!anulada = True
      datafirmes.Recordset!dataanulacio = Now
      datafirmes.Recordset.Update
      datafirmes.Refresh
  End If
End Sub

Private Sub etcomanda_Change()
  Frame2.Width = etcomanda.Width + 50
End Sub

Sub Activar_FitxaTecnica(vestat As Boolean)
   If Not vestat Then
      formfirmes.BackColor = &H5C31DD
      bFitxaTecnica.BackColor = &H8080FF
   End If
   If vestat Then
      formfirmes.BackColor = &H8000000F
      bFitxaTecnica.BackColor = &H6BEBB1
   End If
End Sub
Private Sub Form_Activate()
  Dim vx As Double
  Dim vy As Double
  Dim vcom As String
  vcom = llegir_ini("Firmes", "ultimaconsultafirmes", "comandes.ini")
  bultimaconsulta.BackColor = &H8000000F
  If Len(vcom) > 4 Then bultimaconsulta.BackColor = QBColor(12)
  If vprimercop = False Then
    vx = cadbl(llegir_ini("PosicioFormFirmes", "Left", "comandes.ini"))
    vy = cadbl(llegir_ini("PosicioFormFirmes", "Top", "comandes.ini"))
    If vx > 0 And vy > 0 Then formfirmes.Left = vx: formfirmes.Top = vy
    vprimercop = True
'    formfirmes.Visible = True
  End If
  
  refrescar_firmes
 
End Sub

Private Sub Form_Load()
  
  refrescar_firmes
  
End Sub
Sub refrescar_firmes()
  If cadbl(formcomandes.Text1) = 0 Then Exit Sub
  etcomanda = formcomandes.Text1
  Command1.Visible = True
  Command2.Visible = True
  Command8.Visible = True
  Command9.Visible = True
  If Not formcomandes.Data1.Recordset.EOF Then
    If formcomandes.Data1.Recordset!producte = "PC" Or formcomandes.Data1.Recordset!producte = "PC2" Or formcomandes.Data1.Recordset!producte = "PCP" Then
        Command1.Visible = False
        Command2.Visible = False
        Command8.Visible = False
        Command8.Visible = False
    End If
  End If
  datafirmes.DatabaseName = cami
  datafirmes.RecordSource = "select * from comandes_firmes where anulada=false and comanda=" + atrim(etcomanda) + " order by data desc"
  datafirmes.Refresh
  If HiHaFitxaTecnica(cadbl(etcomanda)) Then
       Activar_FitxaTecnica True
         Else:
            Activar_FitxaTecnica False
           ' MirarSiHiHaPdfFirmat cadbl(etcomanda)  'PER ARA JA NO ES FA SERVIR
  End If
End Sub
Sub MirarSiHiHaPdfFirmat(vnumc As Double)
   Dim vruta As String
   Dim rst As Recordset
   Dim vdir As String
   
   Set rst = dbtmp.OpenRecordset("select * from comandesmesextres where comanda=" + atrim(vnumc))
   vruta = "\\ord_copies\DadesProduccio\Arxius Produccio\DadesGenerals\FitxesTecniquesRefInplacsa\FICHAS TECNICAS-ESPECIFICACIONES y NORMATIVAS\"
   vdir = Dir(vruta + Format(rst!client, "000000") + "*", vbDirectory)
   If vdir <> "" Then vruta = vruta + vdir + "\FIRMADA\"
   vdir = Dir(vruta + "*.pdf")
   While vdir <> ""
      If vdir <> "." And vdir <> ".." Then
            'vdir
            If InStr(1, vdir, rst!refinplacsa) > 0 Or InStr(1, vdir, rst!refclient) > 0 Then
                If MsgBox("He trobat una fitxa tècnica" + vbNewLine + "Vols veure-la?", vbInformation + vbYesNo, "Fitxa") = vbYes Then
                    obrir_document vruta + vdir
                    If MsgBox("Vols relacionar aquesta fitxa amb aquesta referencia?", vbYesNo + vbExclamation, "Atenció") = vbYes Then
                         assignar_FT vruta + vdir, NomFitxaTecnica(vnumc)
                    End If
                End If
            End If
      End If
      vdir = Dir
   Wend
   Set rst = Nothing
End Sub
Function HiHaFitxaTecnica(vnumc As Double) As Boolean
   Dim vnomfitxer As String
   vnomfitxer = NomFitxaTecnica(vnumc)
   If existeix(vnomfitxer) Then
          HiHaFitxaTecnica = True
   End If
End Function
Function NomFitxaTecnica(vnumc As Double) As String
   Dim vnomfitxer As String
   Dim rst As Recordset
   vnomfitxer = "\\ord_copies\DadesProduccio\Arxius Produccio\DadesGenerals\FitxesTecniquesRefInplacsa"
   Set rst = dbtmp.OpenRecordset("select refinplacsa from comandes_extres where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
      If atrim(rst!refinplacsa) <> "" Then
           vnomfitxer = vnomfitxer + "\FT-" + UCase(atrim(rst!refinplacsa) + ".pdf")
           NomFitxaTecnica = vnomfitxer
      End If
   End If
   Set rst = Nothing
End Function
Private Sub Form_Unload(Cancel As Integer)
 If cadbl(formfirmes.Left) > 0 Then
  escriure_ini "PosicioFormFirmes", "Left", atrim(formfirmes.Left), "comandes.ini"
  escriure_ini "PosicioFormFirmes", "Top", atrim(formfirmes.Top), "comandes.ini"
  vprimercop = False
 End If
End Sub

Sub activadessensefirma(vsubconsulta As String, vsubfirma As String)
  Dim vnumc As String
  Dim rst As Recordset
  Dim vcom As String
  Dim vnomordinador As String
  
  vcom = llegir_ini("Firmes", "ultimaconsultafirmes", "comandes.ini")
  Unload formseleccio
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.CommandXLS.Visible = True
  formseleccio.Data1.DatabaseName = cami
  If Len(vcom) > 4 Then GoTo carregar_comandes
  'poso ordinador_rg com a inici de comandes valides per mirar firmes
  vsql = "SELECT DISTINCT comanda From comandes_firmes_actives WHERE " + vsubfirma
  If InStr(1, LCase(vsubfirma), " in (") = 0 Then
    vsql = vsql + " AND comanda Not In (select distinct comanda from comandes_firmes_actives where "
    vnomordinador = nomordinador
     
    If vnomordinador = "ORDINADOR_LP" Or vnomordinador = "ORD_JOSEPM" Then
        vsql = vsql + " (usuari='ORDINADOR_LP' OR usuari='ORD_JOSEPM')"
         Else
           vsql = vsql + " usuari='" + vnomordinador + "')"
    End If
  End If
    'Clipboard.Clear
    'Clipboard.SetText vsql
'  MsgBox vsql
  Set rst = dbtmp.OpenRecordset(vsql)
  While Not rst.EOF
    vcom = vcom + IIf(vcom <> "", ",", "") + atrim(rst!comanda)
    rst.MoveNext
  Wend
 ' MsgBox vcom
  If vcom = "" Then vcom = "0"
  escriure_ini "Firmes", "ultimaconsultafirmes", vcom, "comandes.ini"
carregar_comandes:
  'Clipboard.Clear
  'Clipboard.SetText "SELECT comandes.comanda, clients.nom, comandes.refclient FROM comandes LEFT JOIN clients ON comandes.client = clients.codi where " + vsubconsulta + " comanda in (" + vcom + ") order by nom"
  formseleccio.Data1.RecordSource = "SELECT comandes.comanda, clients.nom, comandes.refclient FROM comandes LEFT JOIN clients ON comandes.client = clients.codi where " + vsubconsulta + " comanda in (" + vcom + ") order by nom"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 1800
  formseleccio.DBGrid2.Columns(1).Width = 4000
  formseleccio.DBGrid2.Columns(2).Width = 2000
  'formseleccio.DBGrid2.Columns(3).Width = 3000
  formseleccio.DBGrid2.Font.Size = 14
  formseleccio.Width = 10000
  If formseleccio.Data1.Recordset.EOF Then MsgBox "No hi cap de pendent", vbInformation, "Atenció": Exit Sub
  formseleccio.Show 1
  If seleccioret = 1 Then
    If Not formseleccio.Data1.Recordset.EOF Then
     vnumc = atrim(cadbl(formseleccio.Data1.Recordset!comanda))
     'If MsgBox("Vols carregar aquesta comanda?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
      carregar_comanda_i_firmes vnumc
      'Unload formfirmes
    ' End If
    End If
     Else: Unload formseleccio
  End If
  
End Sub
Sub carregar_comanda_i_firmes(vnumc As String)
  formcomandes.Data1.RecordSource = "select * from comandes where comanda=" + atrim(vnumc)
  formcomandes.Data1.Refresh
  Unload formseleccio
  refrescar_firmes
End Sub
Private Sub ambpreuisenselamevafirma()
  Dim vnumc As String
  Dim vsql As String
  Dim rstf As Recordset
  Dim vcont As Byte
  Dim vhihausr As Boolean
   Dim vcom As String
  vcom = llegir_ini("Firmes", "ultimaconsultafirmes", "comandes.ini")
  If Len(vcom) > 4 Then GoTo carregar_comandes
  
  Unload formseleccio
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  
  'poso ordinador_rg com a inici de comandes valides per mirar firmes
  ''vsql = "select comandes_firmes.comanda FROM comandes_firmes LEFT JOIN comandes ON comandes_firmes.comanda = comandes.comanda where comandeS_firmes.anulada=false and comandes.pvp>0 and (comandes_firmes.tipus='INI') and comandes_firmes.comanda Not In (select comanda from comandes_firmes where "
  vsql = "SELECT comandes_firmes.comanda, clients.grupdeclient FROM (comandes_firmes LEFT JOIN comandes ON comandes_firmes.comanda = comandes.comanda) LEFT JOIN clients ON comandes.client = clients.codi WHERE comandes.proximaseccio<>'T' and (((comandes_firmes.anulada)=False) AND ((comandes.pvp)>0) AND ((comandes_firmes.tipus)='INI') AND ((clients.grupdeclient)<>'ARDO'))"
  'vsql = vsql +" and comandes_firmes.comanda Not In (select comanda from comandes_firmes where tipus='PVP' and anulada=false)"
  Set rstf = dbtmp.OpenRecordset("select comanda,usuari from comandes_firmes where tipus='PVP' and anulada=false order by comanda")
  Set rst = dbtmp.OpenRecordset(vsql)
  vcol = ""
  While Not rst.EOF
    vhihausr = False
    vcont = 0
    rstf.FindFirst "comanda=" + atrim(rst!comanda)
    While Not rstf.NoMatch
       vcont = vcont + 1
       If UCase(rstf!usuari) = UCase(nomordinador) Then vhihausr = True
       rstf.FindNext "comanda=" + atrim(rst!comanda)
    Wend
    If vcont < 2 And Not vhihausr Then
        If InStr(1, vsql, Str(rst!comanda)) = 0 Then
            vcom = vcom + IIf(vcom <> "", ",", "") + atrim(rst!comanda)
        End If
    End If
    rst.MoveNext
  Wend
  If vcom = "" Then MsgBox "No hi ha pendents de revisar.", vbInformation, "Pendents PVP": GoTo fi
  escriure_ini "Firmes", "ultimaconsultafirmes", vcom, "comandes.ini"
carregar_comandes:
  formseleccio.Data1.RecordSource = "SELECT comandes.comanda, clients.nom, comandes.refclient FROM comandes LEFT JOIN clients ON comandes.client = clients.codi where comanda in (" + vcom + ") order by nom"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 1800
  formseleccio.DBGrid2.Columns(1).Width = 4000
  formseleccio.DBGrid2.Columns(2).Width = 2000
  'formseleccio.DBGrid2.Columns(3).Width = 3000
  formseleccio.DBGrid2.Font.Size = 14
  formseleccio.Width = 10000
  If formseleccio.Data1.Recordset.EOF Then MsgBox "No hi cap de pendent", vbInformation, "Atenció": Exit Sub
  formseleccio.Show 1
  If seleccioret = 1 Then
   If Not formseleccio.Data1.Recordset.EOF Then
    vnumc = atrim(cadbl(formseleccio.Data1.Recordset!comanda))
    'If MsgBox("Vols carregar aquesta comanda?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
     formcomandes.Data1.RecordSource = "select * from comandes where comanda=" + atrim(vnumc)
     formcomandes.Data1.Refresh
     Unload formseleccio
     refrescar_firmes
    ' Unload formfirmes
    'End If
   End If
     Else: Unload formseleccio
  End If
fi:
  Set rstf = Nothing
  Set rst = Nothing
End Sub

Private Sub mborrartotes_Click()
  If UCase(InputBox("Escriu la contrasenya per borrar totes les firmes de la comanda " + atrim(etcomanda), "Borrar firmes")) = "INPLACSA" Then
      dbtmp.Execute "delete * from comandes_firmes where comanda=" + atrim(etcomanda)
      datafirmes.Refresh
  End If
End Sub

Private Sub mfirmamassivaardo_Click()
  Dim vagrupacioardo As String
  Dim rst As Recordset
  Dim rstfirmes As Recordset
  Dim vcomandessenseARDO1 As String
  
  vagrupacioardo = InputBox("Entra el nom de l'agrupació ARDO que vols firmar massivament (PVP2): " + vbNewLine + "NOMES ES FIRMARAN LES QUE JA TENEN UNA FIRMA.", "AGRUPACIÓ ARDO")
  If vagrupacioardo = "" Then Exit Sub
  ratoli "espera"
  Set rstfirmes = dbtmp.OpenRecordset("select * from comandes_firmes where tipus='PVP' order by comanda")
  Set rst = dbtmp.OpenRecordset("select * from comandes where  numpressupost='" + atrim(vagrupacioardo) + "' and producte<>'PC' and producte<>'PC2' and producte<>'PCI3' and producte<>'PCP'")
  While Not rst.EOF
     rstfirmes.FindFirst "comanda=" + atrim(rst!comanda)
     If rstfirmes.NoMatch Then
           vcomandessenseARDO1 = vcomandessenseARDO1 + IIf(vcomandessenseARDO1 <> "", ",", "") + Str(rst!comanda)
         Else
           rstfirmes.MoveNext ' "comanda=" + atrim(rst!comanda)
           If Not rstfirmes.EOF Then
               If rstfirmes!comanda <> rst!comanda Then
                      possarfirmaARDO2 rst!comanda
               End If
                 Else: possarfirmaARDO2 rst!comanda
           End If
     End If
     rst.MoveNext
  Wend
  ratoli "normal"
  If vcomandessenseARDO1 <> "" Then
       MsgBox "Comandes que no tenen PVP1 ENTRAT: (Copiades al PORTAPAPERS) " + vbNewLine + vcomandessenseARDO1, vbExclamation + vbOKOnly, "SENSE PVP1"
       Clipboard.Clear
       Clipboard.SetText vcomandessenseARDO1
  End If
  Set rst = Nothing
End Sub
Sub possarfirmaARDO2(vnumc As Double)
   dbtmp.Execute "insert into comandes_firmes (comanda,usuari,tipus,data) values (" + atrim(vnumc) + ",'ARDO_PVP2','PVP',now)"
End Sub

Private Sub mfirmesborrades_Click()
  reixa.BackColor = QBColor(12)
  datafirmes.RecordSource = "select usuari,tipus,data, dataanulacio,anulada from comandes_firmes where anulada=true and comanda=" + atrim(etcomanda) + " order by data desc"
  datafirmes.Refresh
End Sub

