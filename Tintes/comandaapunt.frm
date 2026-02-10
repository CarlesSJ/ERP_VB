VERSION 5.00
Begin VB.Form comandapreparada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comanda Preparada"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5835
   Icon            =   "comandaapunt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameestatgestionat 
      Height          =   5355
      Left            =   4965
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   5790
      Begin VB.CommandButton Command5 
         BackColor       =   &H0000FFFF&
         Caption         =   "Màquina"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1350
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2865
         Width           =   3330
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0076B5E9&
         Caption         =   "Palet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1350
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   795
         Width           =   3330
      End
   End
   Begin VB.Frame framelectura 
      Height          =   5265
      Left            =   2385
      TabIndex        =   5
      Top             =   5190
      Visible         =   0   'False
      Width           =   5670
      Begin VB.Frame Frame2 
         Height          =   4500
         Left            =   1410
         TabIndex        =   6
         Top             =   150
         Width           =   2520
         Begin VB.CommandButton Command2 
            Caption         =   "Lectura acabada"
            Height          =   570
            Left            =   510
            TabIndex        =   9
            Top             =   3735
            Width           =   1425
         End
         Begin VB.TextBox lectura 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Left            =   255
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   390
            Width           =   2025
         End
         Begin VB.Label Label3 
            Caption         =   "Lectura codis de llauna"
            Height          =   255
            Left            =   465
            TabIndex        =   8
            Top             =   180
            Width           =   1770
         End
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   3510
      Picture         =   "comandaapunt.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1620
      Width           =   1830
   End
   Begin VB.TextBox observacions 
      Height          =   600
      Left            =   405
      MaxLength       =   255
      TabIndex        =   2
      Top             =   945
      Width           =   4950
   End
   Begin VB.ListBox llistadellaunes 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   390
      TabIndex        =   0
      Top             =   2145
      Width           =   5025
   End
   Begin VB.TextBox datapreparada 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   420
      TabIndex        =   11
      Top             =   330
      Width           =   1110
   End
   Begin VB.CommandButton Command3 
      Height          =   480
      Left            =   3525
      Picture         =   "comandaapunt.frx":0947
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4785
      Width           =   1830
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Copiar a una altra comanda de la reixa"
      Height          =   435
      Left            =   165
      TabIndex        =   17
      Top             =   4830
      Width           =   1635
   End
   Begin VB.Label nummaquina 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2595
      TabIndex        =   16
      Top             =   75
      Width           =   2130
   End
   Begin VB.Label Label4 
      Caption         =   "Data preparada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   555
      TabIndex        =   10
      Top             =   75
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   3555
      Picture         =   "comandaapunt.frx":0ED1
      Stretch         =   -1  'True
      Top             =   -375
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.Label Label2 
      Caption         =   "Observacions varies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   495
      TabIndex        =   3
      Top             =   690
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Llaunes assignades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   510
      TabIndex        =   1
      Top             =   1905
      Width           =   3660
   End
End
Attribute VB_Name = "comandapreparada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   framelectura.visible = True
   framelectura.Left = 30
   framelectura.Top = 30
   lectura = ""
   lectura.SetFocus
   While framelectura.visible
      DoEvents
   Wend
   actualitzarllaunesassignades
   'guardar_numerosdellaunes lectura
End Sub
Sub guardar_numerosdellaunes(lectura As TextBox)
 
End Sub
Private Sub Command2_Click()
  Dim rstll As Recordset
  Dim l As String
  Dim vnumc As Double
  lectura = atrim(lectura)
  vnumc = cadbl(formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, 0))
  If vnumc = 0 Then Exit Sub
  If llistadellaunes.ListCount > 0 And Len(lectura) = 0 Then
    If MsgBox("No has possat cap llauna." + Chr(10) + " Vols borrar totes les llaunes assignades?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then
        GoTo fi
         Else
           dbtintes.Execute "delete * from assignaciollaunesacomandes where comanda=" + atrim(vnumc)
           GoTo fi
    End If
      Else: If Len(lectura) = 0 Then GoTo fi
  End If
  If InStr(1, Mid(lectura, Len(lectura) - 1), vbCrLf) = 0 Then lectura = lectura + vbCrLf
  'trec els enters i posso comes
  l = "'" + Trim(substituir(lectura, vbCrLf, "','"))
  If Mid(l, Len(l), 1) = "'" Then l = Mid(l, 1, Len(l) - 2)
  'busco totes les llaunes
  Set rstll = dbtintes.OpenRecordset("select codi,numllauna,descripcio from dadesllaunes where dadesllaunes.numllauna in (" + l + ")")
  If Not tintessondelacomanda(rstll) Then MsgBox "Hi ha tintes que no son per aquesta comanda, sol.luciona-ho i torna a assignar-les.", vbCritical, "Error": GoTo fi
  If Not rstll.EOF Then dbtintes.Execute "delete * from assignaciollaunesacomandes where comanda=" + atrim(vnumc)
  While Not rstll.EOF
     dbtintes.Execute "insert into assignaciollaunesacomandes (comanda,coditinta,numllauna) values (" + atrim(vnumc) + "," + atrim(cadbl(rstll!codi)) + ",'" + atrim(rstll!numllauna) + "')"
     rstll.MoveNext
  Wend
fi:
  framelectura.visible = False
  Set rstll = Nothing
End Sub
Function tintessondelacomanda(rstllt As Recordset) As Boolean
  Dim vidtreball As Double
  Dim vordre As Double
  Dim rst As Recordset
  Dim vsql As String
  Dim vsql2 As String
  Dim rstll As Recordset
  Set rstll = rstllt.Clone
  tintessondelacomanda = True
  vidtreball = cadbl(formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, 7))
  formtintes.reixacomandes.col = 7
  If formtintes.reixacomandes.CellBackColor = 15971192 Then vidtreball = IIf(vidtreball < 0, vidtreball * -1, vidtreball)
  formtintes.reixacomandes.col = 0: formtintes.reixacomandes.ColSel = formtintes.reixacomandes.Cols - 1
  vordre = cadbl(formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, 8))
'  Set rst = dbclixes.OpenRecordset("select coditinta from tintes where id_treball=" + atrim(vidtreball) + " and ordremodificacio=" + atrim(vordre))
  vsql = "SELECT Tintes.coditinta,id_tinter From tintes WHERE id_tinter "
  vsql2 = " in(select id_tinter from tintes where Tintes.id_treball=" + atrim(vidtreball) + " and tintes.ordremodificacio=+" + atrim(vordre) + ")  or  id_tinter in(select tinterlinkambid_treball from tintes where tinterlinkambid_treball>0 and Tintes.id_treball=" + atrim(vidtreball) + " and tintes.ordremodificacio=+" + atrim(vordre) + ")"
  vsql = vsql + vsql2
   
  Set rst = dbclixes.OpenRecordset(vsql)
  
  While Not rstll.EOF
     rst.FindFirst "coditinta='" + atrim(rstll!codi) + "'"
     If rst.NoMatch Then
         If Not mirarsialternativa(vsql2, rstll!codi) Then
              MsgBox "La tinta " + rstll!descripcio + " no existeix a la comanda.", vbCritical, "Atenció": tintessondelacomanda = False
         End If
     End If
     rstll.MoveNext
  Wend
  Set rst = Nothing
End Function
Function mirarsialternativa(vsql As String, vcoditinta As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbclixes.OpenRecordset("select * from tintes_alternatives where id_tinter " + vsql)
   If Not rst.EOF Then
       rst.FindFirst "coditinta='" + atrim(vcoditinta) + "'"
       If Not rst.NoMatch Then mirarsialternativa = True
   End If
   Set rst = Nothing
End Function
Function demanaraquinestatdegestio() As String
   frameestatgestionat.tag = ""
   frameestatgestionat.visible = True
   frameestatgestionat.Left = 30
   frameestatgestionat.Top = 30
   While frameestatgestionat.visible
     DoEvents
   Wend
   demanaraquinestatdegestio = frameestatgestionat.tag
End Function
Sub canviarlestatdegestionat()
  Dim vnumc As Double
  Dim vnumtreball As Double
  Dim col As Double
  Dim vcanvisituacio As String
  vnumc = cadbl(formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, 0))
  vnumtreball = cadbl(formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, formtintes.numcol("NºTreball")))
  vcanvisituacio = demanaraquinestatdegestio
  If vcanvisituacio <> "M" And vcanvisituacio <> "P" Then Exit Sub
  For col = 0 To formtintes.reixacomandes.Cols - 1
          If formtintes.reixacomandes.TextMatrix(0, col) = "Gestionat?" Then
             formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, col) = vcanvisituacio
             If vcanvisituacio = "M" Then vcolor = QBColor(14)
             If vcanvisituacio = "P" Then vcolor = &H80C0FF
             formtintes.reixacomandes.col = col
             formtintes.reixacomandes.CellBackColor = vcolor
          End If
  Next col
  dbtintes.Execute "insert into comandesrevisadesatintes (comanda,numtreball,estatgestio) values (" + atrim(vnumc) + "," + atrim(cadbl(vnumtreball)) + ",'N')"
  dbtintes.Execute "update  comandesactives set gestionat='" + vcanvisituacio + "' where comanda=" + atrim(vnumc)
  dbtintes.Execute "update  comandesrevisadesatintes set estatgestio='" + vcanvisituacio + "' where comanda=" + atrim(vnumc)
  'formtintes.canviarestatcomanda_reixacomandes vnumc, vnumtreball, vcanvisituacio
End Sub

Private Sub Command3_Click()
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vllaunes As String
  Dim vnumc As Double
  Dim vprinter As Printer
  Set oapp = New CRAXDDRT.Application
  If llistadellaunes.ListCount = 0 Then Exit Sub
  guardarobservacioidata
  canviarlestatdegestionat
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "etiqueta_tintapreparada.rpt")
  'oreport.Database.Tables.Item(0).Location = ""
  'oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + UCase(atrim(numllauna)) + "'"
   'report.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
 ' oreport.Sections("D").ReportObjects.Item("serie2").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  oreport.DiscardSavedData
  carregarllaunes vllaunes
 
  vnumc = cadbl(formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, 0))
  oreport.FormulaFields.GetItemByName("lot").Text = "'" + Format(vnumc, "#,##0") + "'"
  oreport.FormulaFields.GetItemByName("llaunes").Text = """" + atrim(vllaunes) + """"
  oreport.FormulaFields.GetItemByName("data").Text = "'" + atrim(datapreparada) + "'"
  oreport.FormulaFields.GetItemByName("nummaquina").Text = "'" + atrim(nummaquina) + "'"
  oreport.FormulaFields.GetItemByName("observacions").Text = "'" + treure_apostruf(observacions) + "'"
  Set vprinter = triarimpresoratickets
  'MsgBox vprinter.DeviceName
  'If UCase(vprinter.DeviceName) = "TICKETS" Then
  If InStr(1, UCase(vprinter.DeviceName), "TICKETS") > 0 Then
     oreport.SelectPrinter vprinter.DriverName, vprinter.DeviceName, vprinter.Port
       Else: MsgBox "No s'ha trobat la impresora Tickets instal.lada al sistema", vbCritical, "Error": Exit Sub
  End If
  
  'MsgBox oreport.PrinterName
  oreport.PaperOrientation = crPortrait
  If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
    Else
      oreport.DisplayProgressDialog = False
      oreport.PrintOut False, 1
  End If
  Unload Me
End Sub
Function triarimpresoratickets() As Printer
  
  For Each triarimpresoratickets In Printers
    
    If InStr(1, UCase(triarimpresoratickets.DeviceName), "TICKETS") > 0 Then
       Exit Function
    End If
  Next
  Set triarimpresoratickets = Printer
End Function
Sub carregarllaunes(vllaunes As String)
  Dim rst As Recordset
  Dim rstdades As Recordset
  Dim vnumc As Double
  vnumc = cadbl(formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, 0))
  
  Set rstdades = dbtintes.OpenRecordset("select * from dadesllaunes", , ReadOnly)
  Set rst = dbtintes.OpenRecordset("SELECT * from assignaciollaunesacomandes where comanda=" + atrim(vnumc), , ReadOnly)
  While Not rst.EOF
     rstdades.FindFirst "numllauna='" + atrim(rst!numllauna) + "'"
     If Not rstdades.NoMatch Then
        vllaunes = vllaunes + atrim(rst!numllauna) + " " + atrim(rstdades!descripcio) + "¿"
        'poso les llaunes marcades com que estan a impresores
        dbtintes.Execute "update llaunes set aimpresores=true where numllauna='" + atrim(rst!numllauna) + "'"
     End If
     rst.MoveNext
  Wend
  
  Set rst = Nothing
  Set rstdades = Nothing
End Sub



Private Sub Command4_Click()
   frameestatgestionat.tag = "P"
   frameestatgestionat.visible = False
   
End Sub

Private Sub Command5_Click()

   frameestatgestionat.tag = "M"
   frameestatgestionat.visible = False
End Sub

Private Sub Command6_Click()
  Dim i As Integer
  Dim vnumc_origen As Double
  vnumc_origen = cadbl(formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, 0))
  vnumc = cadbl(InputBox("Entra el numero de comanda on vols copiar-ho.", "Copiar llaunes"))
  If vnumc = 0 Then Exit Sub
  For i = 1 To formtintes.reixacomandes.Rows - 1
      If formtintes.reixacomandes.TextMatrix(i, 0) = vnumc Then Exit For
  Next i
  If formtintes.reixacomandes.TextMatrix(i, 0) <> vnumc Then MsgBox "No he trobat aquesta comanda a la reixa", vbCritical, "Error": Exit Sub
    formtintes.reixacomandes.Row = i
    formtintes.reixacomandes.col = 0
    formtintes.reixacomandes.ColSel = formtintes.reixacomandes.Cols - 1
    formtintes.reixacomandes.RowSel = i
   framelectura.visible = True
   framelectura.Left = 30
   framelectura.Top = 30
   lectura = ""
   carregarllaunesassignadesalallistadellaunes vnumc_origen
End Sub

Private Sub Form_Load()
     
  actualitzarllaunesassignades
  If datapreparada = "" Then datapreparada = Format(Now, "dd/mm/yy")
  nummaquina = formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, formtintes.numcol("Nº_Maq"))
End Sub
Sub actualitzarllaunesassignades()
  Dim rst As Recordset
  Dim rstdades As Recordset
  Dim vnumc As Double
  vnumc = cadbl(formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, 0))
  Set rst = dbtintes.OpenRecordset("select * from comandesrevisadesatintes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then datapreparada = IIf(atrim(rst!datacomandapreparada) <> "", atrim(rst!datacomandapreparada), datapreparada): observacions = IIf(atrim(rst!observacio) <> "", atrim(rst!observacio), observacions)
  Set rstdades = dbtintes.OpenRecordset("select * from dadesllaunes", , ReadOnly)
  Set rst = dbtintes.OpenRecordset("SELECT * from assignaciollaunesacomandes where comanda=" + atrim(vnumc), , ReadOnly)
  llistadellaunes.Clear
  While Not rst.EOF
     rstdades.FindFirst "numllauna='" + atrim(rst!numllauna) + "'"
     llistadellaunes.AddItem formtintes.justificar(rst!numllauna, 10, "E") + formtintes.justificar(rstdades!descripcio, 30, "E")
     rst.MoveNext
  Wend
  Set rst = Nothing
  Set rstdades = Nothing
End Sub
Sub carregarllaunesassignadesalallistadellaunes(vnumc_origen As Double)

  Dim rst As Recordset
  Dim rstdades As Recordset
  Dim vnumc As Double
  vnumc = cadbl(formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, 0))
  Set rst = dbtintes.OpenRecordset("select * from comandesrevisadesatintes where comanda=" + atrim(vnumc_origen))
  If Not rst.EOF Then datapreparada = IIf(atrim(rst!datacomandapreparada) <> "", atrim(rst!datacomandapreparada), datapreparada): observacions = IIf(atrim(rst!observacio) <> "", atrim(rst!observacio), observacions)
  Set rstdades = dbtintes.OpenRecordset("select * from dadesllaunes", , ReadOnly)
  Set rst = dbtintes.OpenRecordset("SELECT * from assignaciollaunesacomandes where comanda=" + atrim(vnumc_origen), , ReadOnly)
  lectura = ""
  While Not rst.EOF
     rstdades.FindFirst "numllauna='" + atrim(rst!numllauna) + "'"
     lectura = lectura + vbCrLf + atrim(rst!numllauna)
     rst.MoveNext
  Wend
  Set rst = Nothing
  Set rstdades = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
  guardarobservacioidata
End Sub
Sub guardarobservacioidata()
   Dim vnumc As Double
   Dim vdata As String
   Dim rst As Recordset
   vnumc = cadbl(formtintes.reixacomandes.TextMatrix(formtintes.reixacomandes.Row, 0))
   If llistadellaunes.ListCount = 0 Then
     If MsgBox("Si ho hi ha llaunes la data i les observacions es borraran." + Chr(13) + "Vols borrar aquesta informació", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then
         Exit Sub
        Else: dbtintes.Execute "update comandesrevisadesatintes set observacio='',datacomandapreparada=null where comanda=" + atrim(vnumc): GoTo fi
     End If
   End If
   If Not IsDate(datapreparada) And atrim(datapreparada) <> "" Then MsgBox "La data no es correcte", vbCritical, "Atenció": Exit Sub
   vdata = IIf(atrim(datapreparada) <> "", datapreparada, Null)
   
   Set rst = dbtintes.OpenRecordset("select * from comandesrevisadesatintes where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
      rst.Edit
      rst!observacio = observacions
      rst!datacomandapreparada = vdata
      rst.Update
   End If
   
fi:
   Set rst = Nothing
End Sub

