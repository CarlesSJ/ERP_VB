VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form mantenimentbobina 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment de Bobina Anònima"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox checknoimprimirparcial 
      Caption         =   "no imprimir paper parcial"
      Height          =   225
      Left            =   90
      TabIndex        =   24
      Top             =   1710
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Historia de la Bobina"
      Height          =   2205
      Left            =   -15
      TabIndex        =   21
      Top             =   4125
      Width           =   5715
      Begin MSDBGrid.DBGrid reixaparcials 
         Bindings        =   "mantenimentbobina.frx":0000
         Height          =   1755
         Left            =   75
         OleObjectBlob   =   "mantenimentbobina.frx":0013
         TabIndex        =   22
         Top             =   210
         Width           =   5550
      End
   End
   Begin VB.Data parcials 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   195
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2490
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton avis 
      Height          =   540
      Left            =   3915
      Picture         =   "mantenimentbobina.frx":171C
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "S'enviarà incidència a Planificació."
      Top             =   3210
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox observacions 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   90
      MaxLength       =   100
      TabIndex        =   10
      Top             =   3360
      Width           =   3630
   End
   Begin VB.CommandButton ok 
      BackColor       =   &H00C0FFC0&
      Height          =   795
      Left            =   4620
      Picture         =   "mantenimentbobina.frx":1CA6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3210
      Width           =   765
   End
   Begin VB.TextBox mtrsrestants 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3825
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2400
      Width           =   1620
   End
   Begin VB.TextBox mtrsgastats 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3810
      TabIndex        =   5
      Top             =   1620
      Width           =   1620
   End
   Begin VB.TextBox mtrsinicials 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3810
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   1620
   End
   Begin VB.Label mtrsoriginals 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3060
      TabIndex        =   23
      Top             =   60
      Width           =   2565
   End
   Begin VB.Label grup 
      Height          =   240
      Left            =   15
      TabIndex        =   19
      Top             =   405
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label etutilitzada 
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
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   4065
      TabIndex        =   18
      Top             =   765
      Width           =   1545
   End
   Begin VB.Label etavis 
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2205
      TabIndex        =   17
      Top             =   2910
      Width           =   3270
   End
   Begin VB.Label comanda 
      Height          =   210
      Left            =   30
      TabIndex        =   16
      Top             =   165
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label bobina 
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1950
      TabIndex        =   15
      Top             =   90
      Width           =   1665
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1755
      TabIndex        =   14
      Top             =   75
      Width           =   270
   End
   Begin VB.Label etassignacio 
      BackStyle       =   0  'Transparent
      Caption         =   "Assignació:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   150
      TabIndex        =   13
      Top             =   675
      Width           =   1545
   End
   Begin VB.Label assignacio 
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1590
      TabIndex        =   12
      Top             =   660
      Width           =   3105
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Observacions:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   75
      TabIndex        =   11
      Top             =   3045
      Width           =   2085
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Metres restants:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1290
      TabIndex        =   8
      Top             =   2475
      Width           =   2490
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3570
      TabIndex        =   6
      Top             =   1380
      Width           =   285
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   3705
      X2              =   5580
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "(en aquesta comanda)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1335
      TabIndex        =   3
      Top             =   1920
      Width           =   2085
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Metres gastats:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1290
      TabIndex        =   2
      Top             =   1590
      Width           =   2205
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Metres reals:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1290
      TabIndex        =   1
      Top             =   1080
      Width           =   2205
   End
   Begin VB.Label palet 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99.999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   1665
   End
End
Attribute VB_Name = "mantenimentbobina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim shamodificatalgu As Boolean

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
 
End Sub
Sub calcularmtrsrestants()
  etavis = ""
     mtrsrestants = cadbl(mtrsinicials) - cadbl(mtrsgastats)
  
  If cadbl(mtrsrestants) < 500 Then etavis = "Menys de 500 mtrs ACABADA."
  If cadbl(mtrsrestants) < 1 Then etavis = "BOBINA ACABADA."
  If (cadbl(mtrsgastats)) >= (cadbl(assignacio.tag) + 100) And cadbl(etassignacio.tag) > 10000 Then
     avis.visible = True
       Else: avis.visible = False
  End If
  If cadbl(mtrsinicials) = 0 Then mtrsrestants = 0
End Sub

Private Sub Command3_Click()
  MsgBox "S'enviarà una incidencia a planificació per arreglar el descuadre d'aquesta bobina amb la planificació que hi havia per les pròximes comandes.", vbOKOnly + vbInformation, "Inicència"
End Sub
Sub obrestocks(Optional noobrirbd As Boolean)
 Dim camistocks As String
' Set ws = DBEngine.CreateWorkspace("", "admin", "")
 ' If estaobertstocks Then dbtemp.Execute "delete * from selecciobobentrada": Exit Sub
camistocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
'If camistocks = "{[}]" Then camistocks = "\\Ser2\documentos\Stock Reclamaciones\Estoc inplacsa.mdb"
'If Not existeix(camistocks) Then
'    MsgBox "Error obrint la la base de dades de Estocs (Palets) intentarem obrir la BD per defecte", vbCritical, "Error"
'    camistocks = "\\serverprodu\dades\progcomandes\dades\palets.mdb"
'End If

If camistocks = "{[}]" Then escriure_ini "General", "ruta_stocks", rutadelfitxer(cami) + "palets.mdb", "comandes.ini"
camistocks = llegir_ini("General", "ruta_stocks", "comandes.ini")
If Not noobrirbd Then
   Set dbstocks = OpenDatabase(camistocks)
 '  dbtemp.Execute "delete * from selecciobobentrada"
End If
  
End Sub

Sub comprovarnivellsdestoc()
   Dim rstg As Recordset
   Dim rstp As Recordset
   obrestocks
   Set rstg = dbstocks.OpenRecordset("select * from grupsdepalets")
   While Not rstg.EOF
      Set rstp = dbstocks.OpenRecordset("SELECT Sum(Parcials.metres) AS tmetres, Parcials.comanda From parcials GROUP BY Parcials.comanda HAVING (((Parcials.comanda)='" + atrim(rstg!numerogrup) + "'));")
      If cadbl(rstp!tmetres) < cadbl(rstg!estocminim) And cadbl(rstg!estocminim) > 0 Then
        passaravis 0, 0, Format(Now, "dd/mm/yy") + " - Estoc mínim superat en el grup " + atrim(rstg!numerogrup) + " - " + atrim(rstg!nomdelgrup) + " Estoc mínim: " + atrim(rstg!estocminim) + "Mtrs  Estoc actual: " + atrim(rstp!tmetres) + " Mtrs.", 0
      End If
      rstg.MoveNext
   Wend
   Set rstg = Nothing
   Set rstp = Nothing
   
End Sub

Private Sub Form_Activate()
   mtrsinicials = bobinesdentrada.calcular_mtrsdispreals(cadbl(palet), cadbl(bobina))
   assignacio = carregarassignacio(cadbl(palet), cadbl(bobina), atrim(comanda), atrim(grup))
   calcularmtrsrestants
   mtrsgastats.SetFocus
   If etutilitzada = "Utilitzada" Then
      If DateDiff("d", CVDate(etutilitzada.tag), CVDate(Now)) <> 0 Then
        mtrsgastats.Locked = True
      End If
      mtrsinicials = cadbl(mtrsinicials) + cadbl(mtrsgastats)
      calcularmtrsrestants
     Else: mtrsgastats.Locked = False
   End If
   ratoli "normal"
   guardar_packinglistoriginal cadbl(comanda)
End Sub
Function carregarassignacio(palet As Double, bobina As Double, numc As Double, grp As Double) As String
  Dim rstp As Recordset
  Dim rstb As Recordset
  mtrsgastats.tag = ""
  utilitzada = ""
  parcials.RecordSource = "select * from parcials where  idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " order by utilitzada"
  parcials.Refresh
  Set rstp = dbstocks.OpenRecordset("select * from parcials where cdbl(orcomassignacio)>1000 and idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and comanda='" + atrim(numc) + "'")
  Set rstb = dbstocks.OpenRecordset("select * from bobines where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina))
  If rstp.EOF Then Set rstp = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and comanda='" + atrim(grup) + "'"): numc = grp
  If Not rstb.EOF Then mtrsoriginals = atrim(Redondejar(cadbl(rstb!mts), 0)) + " Mtrs"
  comanda.tag = ""
  If Not rstp.EOF Then
    carregarassignacio = atrim(numc) + " <==> " + atrim(cadbl(rstp!metres)) + " Mtrs"
    assignacio.tag = atrim(cadbl(rstp!metres))
    etassignacio.tag = numc
    comanda.tag = rstp!id
    If rstp!utilitzada Then
       mtrsgastats = rstp!metres
       mtrsgastats.tag = "utilitzada"
       'mtrsinicials = cadbl(mtrsinicials) + cadbl(mtrsgastats)
       etutilitzada = "Utilitzada"
        If IsDate(rstp!Data) Then
          etutilitzada.tag = CVDate(rstp!Data)
           Else: etutilitzada.tag = Format(Now, "dd/mm/yy")
       End If
    End If
    observacions = atrim(rstp!observacions)
    colocarsealiddparcial rstp!id
  End If
  Set rstp = Nothing
  Set rstb = Nothing

End Function
Sub colocarsealiddparcial(id As Double)
  If parcials.Recordset.EOF And parcials.Recordset.BOF Then Exit Sub
   parcials.Recordset.MoveFirst
   While Not parcials.Recordset.EOF
      If cadbl(parcials.Recordset!id) = cadbl(id) Then GoTo surt
      parcials.Recordset.MoveNext
   Wend
surt:
End Sub
Private Sub Label8_Click()


End Sub

Private Sub Form_Click()
 ' passaravis cadbl(palet), cadbl(bobina), "Molta diferencia de metres gastats amb els assignats.", comanda, explicacio
  'assignacio = carregarassignacio(cadbl(palet), cadbl(bobina), atrim(comanda), atrim(grup))
  'bobinesdentrada.imprimir_bobinaparcial palet, bobina, , 2
  'comprovardifmetresassignats
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 110 Then KeyCode = 188
End Sub

Private Sub Form_Load()
  
  If llegir_ini("General", "ruta_stocks", "comandes.ini") = "{[}]" Then
     escriure_ini "General", "ruta_stocks", "\\serverprodu\dades\progcomandes\dades\palets.mdb", "comandes.ini"
  End If
  parcials.DatabaseName = llegir_ini("General", "ruta_stocks", "comandes.ini")
  shamodificatalgu = False
  
End Sub
Sub guardar_packinglistoriginal(numc As Double)
   Dim rsth As Recordset
   Set rsth = dbstocks.OpenRecordset("select * from historic_packinglist where comanda='" + atrim(numc) + "'")
   If rsth.EOF Then
      dbstocks.Execute "insert into historic_packinglist select * from parcials where orcomassignacio<>'500' and comanda='" + atrim(numc) + "'"
      Set rsth = dbstocks.OpenRecordset("select * from historic_packinglist where comanda='" + atrim(numc) + "'")
      If rsth.EOF Then dbstocks.Execute "insert into historic_packinglist (comanda) values ('" + atrim(numc) + "')"
   End If
   Set rsth = Nothing
End Sub
Private Sub mtrsgastats_Change()
   calcularmtrsrestants
   If mantenimentbobina.ActiveControl.Name = "mtrsgastats" Then shamodificatalgu = True
End Sub

Private Sub mtrsgastats_KeyDown(KeyCode As Integer, Shift As Integer)
If mtrsgastats.Locked Then MsgBox "Qualsevol modificació de metres d'una bobina previament omplerta s'ha de passar nota a oficines per comprovar-ho. Gràcies", vbInformation + vbOKOnly, "Rectificació de metres"
End Sub
Function metresajust(numc As Double, p As Double, b As Double) As Double
   Dim rstm As Recordset
   Set rstm = dbstocks.OpenRecordset("select sum(metres) as total from parcials where orcomassignacio='500' and comanda='" + atrim(numc) + "' and idpalet=" + atrim(cadbl(p)) + " and idbobina=" + atrim(cadbl(b)))
   metresajust = 0
   If Not rstm.EOF Then metresajust = cadbl(rstm!total)
   Set rstm = Nothing
End Function
Function comprovardifmetresassignats() As Boolean
   Dim metresassignats As Double
   Dim metresgastats As Double
   Dim mtrsajust As Double
   
   explicacio = atrim(observacions)
   If cadbl(etassignacio.tag) < 10000 Then comprovardifmetresassignats = True: Exit Function
   mtrsajust = metresajust(cadbl(etassignacio.tag), cadbl(palet), cadbl(bobina))
   metresgastats = cadbl(mtrsgastats) + mtrsajust
   metresassignats = cadbl(assignacio.tag)
   ratoli "normal"
   
   If metresassignats < 1 Then comprovardifmetresassignats = True: Exit Function
   If (metresassignats + 1000) < metresgastats Or (metresassignats - 1000) > metresgastats Then
         If MsgBox("Hi ha molta diferencia entre els metres assignats i els que has possat." + Chr(10) + Chr(13) + " Ès correcte?", vbExclamation + vbYesNo, "Atenció") = vbYes Then
            comprovardifmetresassignats = True
            'passaravis cadbl(palet), cadbl(bobina), "Molta diferencia de metres gastats amb els assignats.", comanda, explicacio, mtrsajust
           Else: comprovardifmetresassignats = False
         End If
       Else: comprovardifmetresassignats = True
   End If
   ratoli "espera"
   
End Function
Function comprovarsiespotdeixara0() As Boolean
  comprovarsiespotdeixara0 = True
  If noespota0 Then
     MsgBox "No pots deixar el camp de metres gastats sense valor." + Chr(10) + Chr(13) + " Si no l'has utilitzada posa-la a 0 metres si no ho saps segur possa almenys 1 metre per poder rectificar-ho despres.", vbExclamation + vbOKOnly, "Atenció"
     ratoli "normal"
     comprovarsiespotdeixara0 = False
  End If
End Function
Private Sub ok_Click()
  Dim modificar  As Boolean
  If Trim(mtrsgastats) = "" Then
    If Not comprovarsiespotdeixara0 Then Exit Sub
  End If
  ok.Enabled = False
  ratoli "espera"
  modificar = False
  explicacio = ""
  
 If shamodificatalgu Then
  If etutilitzada <> "Utilitzada" Then
    modificar = True
   Else
     If DateDiff("d", CVDate(etutilitzada.tag), CVDate(Now)) = 0 Then
        modificar = True
       Else:
         ratoli "normal"
         MsgBox "No es poden modificar valors que no siguin del mateix dia. Aviseu a oficines"
         ok.Enabled = True
     End If
  End If
  If cadbl(mtrsrestants) < -500 Then
     If MsgBox("Els metres restants son negatius. ES CORRECTE?", vbCritical + vbYesNo, "Atenció") = vbNo Then Exit Sub
  End If
  If Not comprovardifmetresassignats Then ok.Enabled = True: Exit Sub
     
  If modificar Then
    assignarmodificacionsalparcial
  End If
  Set parcials.Recordset = Nothing
  parcials.RecordSource = ""
  parcials.Refresh
  parcials.RecordSource = "select * from parcials where idpalet=" + atrim(cadbl(palet)) + " and idbobina=" + atrim(cadbl(bobina))
  parcials.Refresh
 End If
  'dbstocks.Close
  'obrestocks
  'wait 2
  ratoli "normal"
  ok.Enabled = True
  Unload mantenimentbobina
End Sub
Sub assignarmodificacionsalparcial()
  Dim rstp As Recordset
  Dim segur As Boolean
  Dim vresp As String
  
  colocarsealiddparcial cadbl(comanda.tag)
  If parcials.Recordset.EOF Then Exit Sub
  If cadbl(parcials.Recordset!comanda) > 1999 And cadbl(parcials.Recordset!comanda) < 3000 Then
'     If cadbl(rstp!metres) - cadbl(mtrsgastats) > 0 Then
       Set rstp = dbstocks.OpenRecordset("select * from parcials where id=" + atrim(parcials.Recordset!id))
       parcials.Recordset.AddNew
       For i = 0 To rstp.Fields.Count - 1
        If parcials.Recordset.Fields(i).Name <> "id" Then
         parcials.Recordset.Fields(i) = rstp.Fields(i)
        End If
       Next i
       rstp.Edit
       rstp!metres = cadbl(rstp!metres) - cadbl(mtrsgastats)
       rstp.Update
       rstp.Bookmark = rstp.LastModified
       rstp.Close
 '     End If
            Else: parcials.Recordset.Edit
  End If
  Set rstp = Nothing
  If parcials.Recordset!utilitzada Then
      parcials.Recordset!metres = cadbl(mtrsgastats)
      If atrim(observacions) <> "" Then parcials.Recordset!observacions = atrim(observacions)
     Else
        parcials.Recordset!metres = cadbl(mtrsgastats)
        parcials.Recordset!seccio = lletraseccio
        parcials.Recordset!Data = Now
        parcials.Recordset!operari = numop
        parcials.Recordset!utilitzada = True
        'If observacions = "" Then observacions = "."
        parcials.Recordset!observacions = atrim(observacions)
        parcials.Recordset!comanda = cadbl(comanda)
  End If
  parcials.Recordset.Update
  parcials.Refresh
  'rstp = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(palet) + " and idbobina=" + atrim(bobina) + " and comanda='" + atrim(cadbl(comanda.Tag)) + "'")
  'If Not rstp.EOF Then
  wait 2
 
 
 explicacio = ""
 If InStr(1, etavis, "ACABADA") = 0 Then
   ratoli "normal"
   vresp = ""
   If checknoimprimirparcial.Value = 0 Then vresp = UCase(InputBox("Bobina fisicament acabada?" + Chr(10) + Chr(13) + "Escriu SI per donar per acabada.", "Bobina acabada?"))
   If vresp = "SI" Then
       explicacio = InputBox("Vols entrar una explicació del perquè dones per acabada aquesta bobina?" + Chr(10) + Chr(13) + "SI HAS DONAT PER ACABADA ACCIDENTALMENT ESCRIU [ERROR] A LA CASELLA.", "Bobina acabada")
       ratoli "espera"
       If UCase(explicacio) <> "ERROR" Then passarbobinaaacavada cadbl(palet), cadbl(bobina)
       ratoli "normal"
      Else:
          ratoli "espera"
          actualitzarmetresgrupsestoc cadbl(palet), cadbl(bobina)
          If checknoimprimirparcial.Value <> 1 Then
            bobinesdentrada.imprimir_bobinaparcial palet, bobina, , 1
            wait (2)
          End If
   End If
     Else: passarbobinaaacavada cadbl(palet), cadbl(bobina)
 End If
 actualitzarmetresgrupsestoc cadbl(palet), cadbl(bobina)
 dbstocks.Execute "delete * from parcials where metres=0 and idpalet=" + atrim(cadbl(parcials.Recordset!idpalet)) + " and idbobina=" + atrim(parcials.Recordset!idbobina)
 'wait 1
End Sub
Sub actualitzarmetresgrupsestoc(p As Double, b As Double)
   Dim mtrsrestants As Double
   Dim rstp As Recordset
   Dim c As Byte
   Dim grup As Double
   c = 0
   mtrsrestants = bobinesdentrada.calcular_mtrsdispreals(p, b)
   Set rstp = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(p) + " and idbobina=" + atrim(b))
   While Not rstp.EOF
      If cadbl(rstp!comanda) > 1999 And cadbl(rstp!comanda) < 3000 Then
        c = c + 1
        grup = cadbl(rstp!comanda)
        If c > 1 Or mtrsrestants = 0 Then
          rstp.Delete
          If c = 1 Then c = 0
        End If
      End If
      rstp.MoveNext
   Wend
   Set rstp = Nothing
   wait 2
   If grup > 0 Then
    mtrsrestants = bobinesdentrada.calcular_mtrsdispreals(p, b)
    If mtrsrestants > 499 Then dbstocks.Execute "update parcials set metres=" + atrim(mtrsrestants) + "  where comanda='" + atrim(grup) + "' and idpalet=" + atrim(p) + " and idbobina=" + atrim(b)
    If mtrsrestants < 500 And mtrsrestants > 0 Then passarbobinaaacavada p, b
   End If
End Sub
Sub passarbobinaaacavada(p As Double, b As Double)
   Dim mtrsrestants As Double
   Dim vmsg As String
   Dim rstp As Recordset
   wait 3
   mtrsrestants = bobinesdentrada.calcular_mtrsdispreals(p, b)
   If mtrsrestants < 1 Then GoTo fi
   If mtrsrestants > 499 Then
     passaravis p, b, "Donada per acabada", comanda, explicacio
   End If
   dbstocks.Execute "insert into parcials (idpalet,idbobina,operari,metres,comanda,data,seccio,utilitzada,orcomassignacio) values (" + atrim(p) + "," + atrim(b) + "," + atrim(cadbl(numop)) + "," + atrim(cadbl(mtrsrestants)) + ",100,now,'" + lletraseccio + "',true,0)"
fi:
   Set rstp = dbstocks.OpenRecordset("select * from parcials where idpalet=" + atrim(p) + " and idbobina=" + atrim(b))
   While Not rstp.EOF
      If Not rstp!utilitzada And cadbl(rstp!comanda) > 3000 Then vmsg = vmsg + " " + atrim(rstp!comanda)
      If cadbl(rstp!comanda) > 1999 And cadbl(rstp!comanda) < 3000 Then rstp.Delete
      rstp.MoveNext
   Wend
   If vmsg <> "" Then passaravis p, b, "!!!BOBINA DONADA PER ACABADA AMB ASSIGNACIONS PER ALTRES COMANDES.", comanda, "S'Ha donat per acabada i la seguent comanda estava assignada a aquesta bobina." + Chr(13) + Chr(10) + "Comanda: " + vmsg
   'Set rstp = dbstocks.OpenRecordset("select * from parcials where not utilitzada and cdbl(comanda)>3000 and idpalet=" + atrim(p) + " and idbobina=" + atrim(b))
   'If Not rstp.EOF Then
   '  While Not rstp.EOF
   '    vmsg = vmsg + " " + atrim(rstp!comanda)
   '    rstp.MoveNext
   '  Wend
'     If vmsg <> "" Then passaravis p, b, "!!! DONADA PER ACABADA AMB ASSIGNACIONS PER ALTRES COMANDES.", comanda, "S'Ha donat per acabada i la seguent comanda estava assignada a aquesta bobina." + Chr(13) + Chr(10) + "Comanda: " + vmsg
   'End If
   Set rstp = Nothing
   'bobinesdentrada.imprimir_bobinaparcial p, b, , 1
End Sub
Sub passaravis(p As Double, b As Double, avis, Optional comanda As String, Optional explicacio As String, Optional mtrsajust As Double)
   Dim rutamdb As String
   Dim dbavisos As Database
   Dim rsta As Recordset
   rutamdb = rutadelfitxer(cami) + "avisosincidencies.mdb"
   Set dbavisos = DBEngine.OpenDatabase(rutamdb)
   explicacio = treure_apostruf(explicacio)
   avis = treure_apostruf(avis)
'   MsgBox "insert into avisos_baixes  (data,seccio,nomoperari,numoperari,palet,bobina,avis,comanda) values (now,'" + atrim(lletraseccio) + "','" + atrim(Form1.nomoperari) + "','" + atrim(numop) + "','" + atrim(palet) + "','" + atrim(bobina) + "','" + treure_apostruf(avis) + "','" + atrim(comanda) + "')"
   Set rsta = dbavisos.OpenRecordset("select * from avisos_baixes where seccio='" + atrim(lletraseccio) + "' and comanda='" + atrim(comanda) + "' and avis='" + treure_apostruf(avis) + "'")
   If rsta.EOF Then
    dbavisos.Execute ("insert into avisos_baixes  (data,seccio,nomoperari,numoperari,palet,bobina,avis,comanda,mtrsassignats,mtrsrestants,mtrsgastats,observacio) values (now,'" + atrim(lletraseccio) + "','" + atrim(form1.nomoperari) + "','" + atrim(numop) + "','" + atrim(p) + "','" + atrim(b) + "','" + treure_apostruf(avis) + "','" + atrim(comanda) + "'," + atrim(cadbl(assignacio.tag)) + "," + atrim(cadbl(mtrsrestants)) + "," + atrim(cadbl(mtrsgastats) + cadbl(mtrsajust)) + ",'" + treure_apostruf(explicacio) + "')")
   End If
   Set rsta = Nothing
   dbavisos.Close
   Set dbavisos = Nothing
End Sub
 Function rutadelfitxer(cam As String) As String
   Dim c As Byte
   c = 0
   While InStr(c + 1, cam, "\") <> 0
    c = InStr(c + 1, cam, "\")
   Wend
   If c = 0 Then c = Len(cam)
   rutadelfitxer = Mid(cam, 1, c)
End Function
