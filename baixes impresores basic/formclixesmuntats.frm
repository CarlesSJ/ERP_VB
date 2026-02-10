VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form formclixesmuntats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clixes muntats entre dues dates"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18075
   Icon            =   "formclixesmuntats.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   18075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Data datamuntadora 
      Caption         =   "datamuntadora"
      Connect         =   "Access"
      DatabaseName    =   "W:\progcomandes\dades\baixes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"formclixesmuntats.frx":00D2
      Top             =   270
      Visible         =   0   'False
      Width           =   2805
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      Height          =   9960
      Left            =   9495
      TabIndex        =   7
      Top             =   900
      Width           =   8190
      _cx             =   5080
      _cy             =   5080
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "formclixesmuntats.frx":0237
      Height          =   10035
      Left            =   150
      OleObjectBlob   =   "formclixesmuntats.frx":024F
      TabIndex        =   6
      Top             =   945
      Width           =   9105
   End
   Begin VB.Frame Frame1 
      Height          =   630
      Left            =   150
      TabIndex        =   0
      Top             =   135
      Width           =   7785
      Begin VB.TextBox chorafi 
         Height          =   285
         Left            =   4665
         TabIndex        =   9
         Top             =   225
         Width           =   825
      End
      Begin VB.TextBox chorainici 
         Height          =   285
         Left            =   1980
         TabIndex        =   8
         Top             =   225
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   300
         Left            =   5730
         Picture         =   "formclixesmuntats.frx":0F9E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1860
      End
      Begin VB.TextBox cdatafi 
         Height          =   285
         Left            =   3555
         TabIndex        =   3
         Top             =   225
         Width           =   1065
      End
      Begin VB.TextBox cdatainici 
         Height          =   285
         Left            =   885
         TabIndex        =   1
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Data fi:"
         Height          =   225
         Left            =   3000
         TabIndex        =   4
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Data inici:"
         Height          =   225
         Left            =   135
         TabIndex        =   2
         Top             =   225
         Width           =   975
      End
   End
End
Attribute VB_Name = "formclixesmuntats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vcarpetadestiCACHE As String

Private Sub Command1_Click()
   vsql1 = "SELECT muntadoratot.comanda, First(muntadoratot.totalhores) AS PrimeroDetotalhores, Max([muntadores].[datafi] & ' ' & [muntadores].[horafi]) AS Dataihora_fi, First([muntadoratot].[firma] & ' - '+[muntadoratot].[nomfirma]) AS [Nom operari] "
   vsql1 = vsql1 + " FROM muntadoratot LEFT JOIN muntadores ON muntadoratot.comanda = muntadores.comanda "
   vsql2 = " GROUP BY muntadoratot.comanda "
   
   
   datamuntadora.RecordSource = vsql1 + vsql2 + " HAVING Max(muntadores.datafi) Is Not Null and (cvdate(Max(muntadores.datafi)&' '& Max(muntadores.horafi)) between #" + Format(cdatainici, "mm/dd/yy") + " " + chorainici + "# and #" + Format(cdatafi, "mm/dd/yy") + " " + chorafi + "#) ORDER BY CVDate(Max([muntadores].[datafi]) & ' ' & Max([muntadores].[horafi])) "
   datamuntadora.Refresh
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   Dim vnumc As Double
   Dim vrutaPDFcache As String
   Dim vrutaPDF As String
  ' AcroPDF1.LoadFile "a"
   vnumc = cadbl(DBGrid1.Columns(0))
   If vnumc = 0 Then Exit Sub
   vrutaPDF = "Les_" + atrim(atrim(Int(cadbl(vnumc) / 1000)) + "000")
   vrutaPDF = llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini") + "\" + vrutaPDF + "\" + atrim(numc)
   vrutaPDF = vrutaPDF + atrim(vnumc) + "\" + atrim(vnumc) + "_BaixaMuntadora.pdf"
   vrutaPDFcache = vcarpetadestiCACHE + "\Les_" + Trim(Int(vnumc / 1000)) + "000\" + atrim(vnumc) + "\" + atrim(vnumc) + "_BaixaMuntadora.pdf"

   If existeix(vrutaPDFcache) Then vrutaPDF = vrutaPDFcache
   If existeix(vrutaPDF) Then
     If existeix("c:\temp\pdftemp.pdf") Then Kill "c:\temp\pdftemp.pdf"
     FileCopy vrutaPDF, "c:\temp\pdftemp.pdf"
     'AcroPDF1.LoadFile "c:\temp\pdftemp.pdf"
     AcroPDF1.src = vrutaPDF
      AcroPDF1.setLayoutMode "SinglePage"
      AcroPDF1.setShowToolbar False
      AcroPDF1.setShowScrollbars False
      AcroPDF1.setView ("FitH")
  End If
End Sub

Private Sub Form_Load()
   Dim vsql1 As String
   Dim vsql2 As String
   
   datamuntadora.DatabaseName = rutadelfitxer(cami) + "baixes.mdb"
   vsql1 = "SELECT muntadoratot.comanda, First(muntadoratot.totalhores) AS PrimeroDetotalhores, Max([muntadores].[datafi] & ' ' & [muntadores].[horafi]) AS Dataihora_fi, First([muntadoratot].[firma] & ' - '+[muntadoratot].[nomfirma]) AS [Nom operari] "
   vsql1 = vsql1 + " FROM muntadoratot LEFT JOIN muntadores ON muntadoratot.comanda = muntadores.comanda "
   vsql2 = " GROUP BY muntadoratot.comanda;"
   datamuntadora.RecordSource = vsql1 + " where muntadoratot.comanda=-1 " + vsql2
   datamuntadora.Refresh
   cdatainici = Format(Now, "dd/mm/yy")
   chorainici = "06:00"
   cdatafi = Format(Now, "dd/mm/yy")
   chorafi = "14:00"
   vcarpetadestiCACHE = llegir_ini("ruta", "ruta_comandes_exportades", rutadelfitxer(cami) + "valorsprograma.ini")
   vcarpetadestiCACHE = vcarpetadestiCACHE + "\cache_Fabricacio"
   
End Sub
