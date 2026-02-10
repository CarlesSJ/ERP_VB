VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formrevisarentregues 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revisar entregues"
   ClientHeight    =   10470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10335
   Icon            =   "formrevisarentregues.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10470
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Framepassword 
      BackColor       =   &H00EAD9CE&
      Height          =   8595
      Left            =   330
      TabIndex        =   4
      Top             =   10260
      Visible         =   0   'False
      Width           =   7230
      Begin VB.TextBox cpassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   135
         TabIndex        =   18
         Top             =   7545
         Width           =   5505
      End
      Begin VB.CommandButton cbotonum 
         BackColor       =   &H00C0FFC0&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Index           =   10
         Left            =   5685
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   420
         Width           =   1365
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   8
         Left            =   3855
         TabIndex        =   16
         Top             =   4020
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   5
         Left            =   3855
         TabIndex        =   15
         Top             =   2220
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   2
         Left            =   3855
         TabIndex        =   14
         Top             =   420
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   5850
         Width           =   3630
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   7
         Left            =   1980
         TabIndex        =   12
         Top             =   4020
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   6
         Left            =   105
         TabIndex        =   11
         Top             =   4020
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   4
         Left            =   1980
         TabIndex        =   10
         Top             =   2220
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   3
         Left            =   105
         TabIndex        =   9
         Top             =   2220
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   1
         Left            =   1980
         TabIndex        =   8
         Top             =   420
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   420
         Width           =   1770
      End
      Begin VB.CommandButton cbotonum 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Index           =   11
         Left            =   3855
         TabIndex        =   6
         Top             =   5865
         Width           =   1770
      End
      Begin VB.CommandButton Command14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   5715
         Picture         =   "formrevisarentregues.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7560
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command7 
      Height          =   705
      Left            =   9345
      Picture         =   "formrevisarentregues.frx":145C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   165
      Width           =   705
   End
   Begin VB.Data databobines 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   525
      Left            =   7545
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from bobinesent"
      Top             =   1245
      Visible         =   0   'False
      Width           =   2115
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "formrevisarentregues.frx":19E6
      Height          =   5745
      Left            =   420
      OleObjectBlob   =   "formrevisarentregues.frx":19FC
      TabIndex        =   2
      Top             =   1875
      Width           =   9540
   End
   Begin VB.TextBox cfrontal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2970
      TabIndex        =   0
      Top             =   150
      Width           =   6285
   End
   Begin VB.ListBox llista 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   195
      TabIndex        =   19
      Top             =   7830
      Width           =   9855
   End
   Begin VB.Label Label3 
      Caption         =   "PROVA-HO A VEURE QUE TAL ARA "
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
      Height          =   330
      Left            =   3885
      TabIndex        =   21
      Top             =   7635
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.Label Label1 
      Caption         =   "Llista de pendents "
      Height          =   210
      Left            =   240
      TabIndex        =   20
      Top             =   7635
      Width           =   1965
   End
   Begin VB.Label Label2 
      Caption         =   "Frontal->"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   570
      TabIndex        =   1
      Top             =   555
      Width           =   2205
   End
End
Attribute VB_Name = "formrevisarentregues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbobina_Change()

End Sub
Sub actualitzar_reixa(vfrontal As String)
   Dim rst As Recordset
   Dim vnumc As Double
   Dim vpalet As Double
   
   If InStr(1, vfrontal, "/") > 0 Then vnumc = cadbl(Mid(vfrontal + "  ", 1, InStr(1, vfrontal, "/") - 1)): vpalet = cadbl(Mid(vfrontal, InStr(1, vfrontal, "/") + 1))
   If vpalet = 0 Then
       Set rst = dbcomandes.OpenRecordset("select * from bobinesent where comanda=" + atrim(cadbl(vfrontal)))
       If rst.EOF Then GoTo fi
       databobines.RecordSource = "SELECT DISTINCT bobinesent.comanda, First(bobinesent.numalbara) AS nalbara, bobinesent.numpalet, First(bobinesent.revisatTORERU) AS revisat From bobinesent Where numalbara = " + atrim(rst!numalbara) + " GROUP BY bobinesent.comanda, bobinesent.numpalet;"
       databobines.Refresh
       GoTo fi
   End If
   Set rst = dbcomandes.OpenRecordset("select * from bobinesent where comanda=" + atrim(vnumc) + " and numpalet=" + atrim(vpalet))
   If Not rst.EOF Then
      dbcomandes.Execute "update bobinesent set revisattoreru='S',modificat=true where comanda=" + atrim(vnumc) + " and numpalet=" + atrim(vpalet)
      
      databobines.RecordSource = "SELECT DISTINCT bobinesent.comanda, First(bobinesent.numalbara) AS nalbara, bobinesent.numpalet, First(bobinesent.revisatTORERU) AS revisat From bobinesent Where numalbara = " + atrim(rst!numalbara) + " GROUP BY bobinesent.comanda, bobinesent.numpalet;"
      databobines.Refresh
   End If
fi:
   Set rst = Nothing
End Sub

Private Sub cbotonum_Click(Index As Integer)
 If cbotonum(Index).Caption = "OK" Then Framepassword.Visible = False: GoTo fi
   cpassword.Tag = cpassword.Tag + cbotonum(Index).Caption
   If Framepassword.Tag = "password" Then
      cpassword = cpassword + "*"
       Else: cpassword = cpassword.Tag
   End If
   cpassword.SetFocus
fi:
End Sub

Private Sub cfrontal_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      acceptar_codis
   End If
End Sub
Sub acceptar_codis()
     cfrontal = substituirtot(cfrontal, "-", "/")
     cfrontal = substituirtot(cfrontal, ".", "")
     cfrontal = substituirtot(cfrontal, ",", "")
     actualitzar_reixa cfrontal
      cfrontal = ""
      cfrontal.SetFocus
     actualitzar_llista
End Sub

Private Sub Command14_Click()
If Len(cpassword.Tag) = 0 Then Exit Sub
  cpassword.Tag = Mid(cpassword.Tag, 1, Len(cpassword.Tag) - 1)
  If Framepassword.Tag = "password" Then
     cpassword = Mid(cpassword, 1, Len(cpassword) - 1)
      Else: cpassword = Mid(cpassword, 1, Len(cpassword) - 1)
  End If
End Sub

Private Sub Command7_Click()
       Framepassword.Visible = True
       Framepassword.Top = 1800
       Framepassword.Left = 2500
       Framepassword.Tag = ""
       cpassword.SetFocus
       While Framepassword.Visible
         DoEvents
       Wend
       cfrontal = cpassword.Tag
       acceptar_codis
End Sub

Private Sub cpassword_Change()
 If Framepassword.Tag = "password" Then
     KeyAscii = 0
      Else: If KeyAscii > 21 Then cpassword.Tag = cpassword + Chr(KeyAscii)
  End If
  If KeyAscii = 13 Then
      KeyAscii = 0
      Framepassword.Visible = False
  End If
End Sub
Sub actualitzar_llista_defets()
   Dim rst As Recordset
   Dim rst2 As Recordset
   
   Set rst = dbcomandes.OpenRecordset("select * from linies_expedicions")
   While Not rst.EOF
      Set rst2 = dbcomandes.OpenRecordset("SELECT bobinesent.revisatTORERU, linies_expedicions.albara FROM bobinesent RIGHT JOIN linies_expedicions ON bobinesent.numalbara = linies_expedicions.albara WHERE (((bobinesent.revisatTORERU)<>'S' or revisatTORERU is null) AND ((linies_expedicions.albara)=" + atrim(cadbl(rst!albara)) + "));")
      If rst2.EOF Then rst.Edit: rst!enviat = True: rst.Update
      If Not rst2.EOF Then rst.Edit: rst!enviat = False: rst.Update
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub actualitzar_llista()
   Dim rst As Recordset
   Dim vetfet As Boolean
   Dim vliniaactiva As Long
   llista.Clear
   vliniaactiva = -1
   actualitzar_llista_defets
   Set rst = dbcomandes.OpenRecordset("SELECT linies_expedicions.albara, First(linies_expedicions.nomclient) AS Pnomclient, First(linies_expedicions.enviat) AS fet From linies_expedicions GROUP BY linies_expedicions.albara order by First(linies_expedicions.enviat);")
   If rst.EOF Then Exit Sub
   'llista.AddItem "PENDENTS ----------"
   While Not rst.EOF
     'If rst!fet = True Then llista.AddItem "FETS ----------": vetfet = True
     llista.AddItem IIf(rst!fet, "[FET] ", "") + atrim(rst!albara) + " " + atrim(rst!pnomclient)
     llista.ItemData(llista.NewIndex) = rst!albara
     If Not databobines.Recordset.EOF Then
         If databobines.Recordset!nalbara = rst!albara Then vliniaactiva = llista.NewIndex
     End If
     rst.MoveNext
   Wend
   If vliniaactiva > -1 Then llista.ListIndex = vliniaactiva
   Set rst = Nothing
End Sub
Private Sub Form_Activate()
cfrontal.SetFocus
End Sub

Private Sub Form_Load()
    databobines.DatabaseName = App.Path + "\torerus.mdb"
    databobines.RecordSource = "select * from bobinesent where id=-1"
    databobines.Refresh
  actualitzar_llista
End Sub

Private Sub llista_Click()
  Dim rst As Recordset
  If llista.ItemData(llista.ListIndex) > 0 Then
    databobines.RecordSource = "SELECT DISTINCT bobinesent.comanda, First(bobinesent.numalbara) AS nalbara, bobinesent.numpalet, First(bobinesent.revisatTORERU) AS revisat From bobinesent Where numalbara = " + atrim(llista.ItemData(llista.ListIndex)) + " GROUP BY bobinesent.comanda, bobinesent.numpalet;"
     Else: databobines.RecordSource = "SELECT DISTINCT bobinesent.comanda, First(bobinesent.numalbara) AS nalbara, bobinesent.numpalet, First(bobinesent.revisatTORERU) AS revisat From bobinesent Where numalbara = -1 GROUP BY bobinesent.comanda, bobinesent.numpalet;"
  End If
  databobines.Refresh
fi:
  Set rst = Nothing
End Sub

Private Sub llista_GotFocus()
  cfrontal.SetFocus
End Sub

Private Sub reixa_DblClick()
   Dim vnumc As Double
   Dim vpalet As Double
   Dim vfrontal As String
   If databobines.Recordset.EOF Then Exit Sub
   vnumc = databobines.Recordset!comanda
   vpalet = databobines.Recordset!numpalet
   If databobines.Recordset!revisat = "S" Then
      If MsgBox("Vols marcar aquest palet com a no revisat?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
           dbcomandes.Execute "update bobinesent set revisattoreru='',modificat=true where comanda=" + atrim(vnumc) + " and numpalet=" + atrim(vpalet)
           databobines.Refresh
      End If
   End If
   actualitzar_llista
End Sub

Private Sub reixa_GotFocus()
cfrontal.SetFocus
End Sub
