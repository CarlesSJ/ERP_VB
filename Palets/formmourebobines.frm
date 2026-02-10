VERSION 5.00
Begin VB.Form formmourebobines 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Moure bobines a la secció"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   Icon            =   "formmourebobines.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox checkamagarlam 
      Caption         =   "Amagar LAM"
      Height          =   210
      Left            =   3060
      TabIndex        =   4
      Top             =   825
      Width           =   1530
   End
   Begin VB.CommandButton Command1 
      Height          =   465
      Left            =   105
      Picture         =   "formmourebobines.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Passar les bobines sel.leccionades a Situació LAM"
      Top             =   270
      Width           =   3360
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Marcar/Desmarcar Tots"
      Height          =   255
      Left            =   435
      TabIndex        =   2
      Top             =   795
      Width           =   2205
   End
   Begin VB.CommandButton imprimiralbaraproveidor 
      Height          =   480
      Left            =   3705
      Picture         =   "formmourebobines.frx":076F
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir llista palets."
      Top             =   270
      Width           =   1050
   End
   Begin VB.ListBox llistabobines 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6330
      Left            =   420
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1065
      Width           =   3900
   End
   Begin VB.Label Label1 
      Caption         =   "Passar les bobines sel.leccionades a Situació LAM"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   60
      Width           =   4080
   End
End
Attribute VB_Name = "formmourebobines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
  Dim i As Integer
  For i = 0 To llistabobines.ListCount - 1
   llistabobines.Selected(i) = Check1.Value
  Next i
End Sub

Private Sub checkamagarlam_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   carregar_bobinespermoure "L"
End Sub

Private Sub Command1_Click()
  Dim i As Integer
  Dim vidpalet As Double
  Dim vidbobina As Double
  If UCase(InputBox("Totes les bobines sel.leccionades es passaran a seccio LAMINADORA." + Chr(10) + "SI ES CORRECTE ESCRIU [LAM]", "Passar bobines a LAM")) = "LAM" Then
      For i = 0 To llistabobines.ListCount - 1
         If llistabobines.Selected(i) Then
            vidpalet = cadbl(Mid(llistabobines.List(i), 1, InStr(1, llistabobines.List(i), "/") - 1))
            vidbobina = cadbl(Mid(llistabobines.List(i), InStr(1, llistabobines.List(i), "/") + 1, 4))
            If vidpalet > 0 And vidbobina > 0 Then
              dbtmp.Execute "update  bobines set sit='LAM' WHERE idpalet=" + atrim(vidpalet) + " and idbobina=" + atrim(vidbobina)
            End If
         End If
      Next i
      carregar_bobinespermoure "L"
  End If
End Sub

Private Sub Form_Load()

  carregar_bobinespermoure "L"
End Sub
Sub mirarquinescomandestenenanonim(vnumc As Double, ByRef vnumc1 As Double, ByRef vnumc2 As Double, ByRef vnumc3 As Double)
   Dim rstc As Recordset
   vnumc1 = 999999999
   vnumc2 = 999999999
   vnumc3 = 999999999
   Set rstc = dbcomandes.OpenRecordset("select linkcomanda1,linkcomanda2,proximaseccio,comanda from comandes where comanda=" + atrim(vnumc), , ReadOnly)
   If rstc.EOF Then Exit Sub
   Set rstc = dbcomandes.OpenRecordset("SELECT comandes.comanda, productes.ruta FROM comandes INNER JOIN productes ON comandes.producte = productes.codi where comanda=" + atrim(rstc!linkcomanda1) + " or comanda=" + atrim(rstc!linkcomanda2) + " or comanda=" + atrim(rstc!comanda))
   If Not rstc.EOF Then If InStr(1, rstc!ruta, "I") = 0 Then vnumc1 = rstc!comanda
   rstc.MoveNext
   If Not rstc.EOF Then If InStr(1, rstc!ruta, "I") = 0 Then vnumc2 = rstc!comanda
   rstc.MoveNext
   If Not rstc.EOF Then If InStr(1, rstc!ruta, "I") = 0 Then vnumc3 = rstc!comanda
   Set rstc = Nothing
End Sub
Sub preparataulatemporal()
   On Error Resume Next
   dbllistat.Execute "drop table llistatperpujar"
   dbtmp.Execute "SELECT Parcials.idpalet, Parcials.idbobina, Bobines.Sit AS ample, Bobines.disponible AS metres, Bobines.Sit, CDbl(parcials.comanda) AS comanda, 0 as ordre,' ' as nomproveidor INTO llistatperpujar IN '" + nomfitxertemporal + "' FROM (Bobines INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)) INNER JOIN comandes ON cdbl(Parcials.comanda) = comandes.comanda WHERE comandes.proximaseccio='I' and CDbl([parcials].[comanda])=0;"
   dbllistat.Execute "delete * from llistatperpujar"
   On Error GoTo 0
End Sub
Sub carregar_bobinespermoure(vseccio As String)
   Dim rst As Recordset
   Dim rstp As Recordset
   Dim rstpalet As Recordset
   Dim rstordreimpresio As Recordset
   Dim dbplanificacio As Database
   Dim vnumc1 As Double
   Dim vnumc2 As Double
   Dim vnumc3 As Double
   Dim vnumbob As String
   
   ratoli "espera"
   llistabobines.Clear
   Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificaciooperaris.mdb", , True)
   '' 19/07/22 trec el filtre de només baixar les bobines que a planificacio OPERARIS (PACO) ha
     ' donat ordre de planfificació per tan s'ha de preparar les bobinees... ara seran totes
     ' en MARC i en MIRALLES han donat l'ordre
     'nomes canvio la linia seguent
  ' Set rst = dbplanificacio.OpenRecordset("select * from planificaciolam where ordre>0 and ordre<900 order by ordre")
   Set rst = dbplanificacio.OpenRecordset("SELECT planificaciolam.*, comandes.proximaseccio FROM comandes RIGHT JOIN planificaciolam ON comandes.comanda = planificaciolam.comanda where proximaseccio='L' or proximaseccio='I'")
   preparataulatemporal
   While Not rst.EOF
      
      'Set rstordreimpresio = dbtmpb.OpenRecordset("select comanda from impresores_ordreimpresio where comanda=" + atrim(rst!comanda))
      'If rstordreimpresio.EOF Then GoTo proxim
      mirarquinescomandestenenanonim rst!comanda, vnumc1, vnumc2, vnumc3
      Set rstp = dbtmp.OpenRecordset("SELECT distinct Parcials.idpalet, Parcials.idbobina,bobines.disponible, Bobines.Sit, Parcials.comanda FROM Bobines INNER JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet) WHERE parcials.utilitzada=false and (((Parcials.comanda)='" + atrim(vnumc1) + "' or parcials.comanda='" + atrim(vnumc2) + "' or parcials.comanda='" + atrim(vnumc3) + "'))")
           
      While Not rstp.EOF
        Set rstpalet = dbtmp.OpenRecordset("select * from palets where idpalet=" + atrim(rstp!idpalet))
        vnumbob = atrim(rstp!idpalet) + "/" + atrim(rstp!idbobina)
        
        If checkamagarlam.Value = 1 Then    'si hi ha el check d'amagar lam amagarles
             If UCase(Mid(rstp!sit, 1, 3)) = "LAM" Then 'Or Len(atrim(rstp!sit)) < 3 Then
                vnumbob = ""
             End If
        End If
        
        If vnumbob <> "" Then
                llistabobines.AddItem vnumbob + Space(10 - Len(vnumbob)) + UCase(atrim(rstp!sit))
                dbllistat.Execute "insert into llistatperpujar (idpalet,idbobina,ample,metres,sit,comanda,ordre,nomproveidor) values (" + atrim(rstp!idpalet) + "," + atrim(rstp!idbobina) + ",'" + atrim(rstpalet!ample) + "'," + atrim(rstp!disponible) + ",'" + atrim(rstp!sit) + "'," + atrim(rstp!comanda) + "," + atrim(IIf(cadbl(rst!ordre) = 0, 999, passaradecimalpunt(cadbl(rst!ordre)))) + ",'" + nomproveidor(rstpalet!idpalet) + "')"
        End If
        rstp.MoveNext
      Wend
proxim:
      rst.MoveNext
   Wend
   Set rst = Nothing
   Set rstpalet = Nothing
   Set rstc = Nothing
   Set dbplanificacio = Nothing
   ratoli "normal"
End Sub

Private Sub imprimiralbaraproveidor_Click()
  Dim llistat As CrystalReport
  Set llistat = Form1.llistat
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistatbobinesperpujar.rpt"
 llistat.Destination = crptToPrinter
 llistat.CopiesToPrinter = 1
 llistat.DataFiles(0) = nomfitxertemporal
 llistat.SortFields(0) = "+{llistatperpujar.ordre}"
 llistat.SortFields(1) = "+{llistatperpujar.sit}"
 llistat.DiscardSavedData = True
 llistat.Formulas(1) = "titol='LListat de bobines per LAMINADORES.'"
 llistat.Formulas(0) = ""
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 llistat.Formulas(4) = ""
 llistat.Formulas(5) = ""
 llistat.Formulas(6) = ""
 llistat.Formulas(7) = ""
 llistat.Formulas(8) = ""
 llistat.Formulas(9) = ""
 llistat.Formulas(10) = ""
 llistat.Formulas(11) = ""
 DoEvents
 If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 If Form1.mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
End Sub
