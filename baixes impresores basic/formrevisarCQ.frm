VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form formrevisarCQ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revisar CQ"
   ClientHeight    =   12555
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   8295
   ControlBox      =   0   'False
   Icon            =   "formrevisarCQ.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12555
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton brevissioPCC 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Revisar PCC"
      Height          =   945
      Left            =   375
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5565
      Width           =   7155
   End
   Begin VB.Timer TimercontroltotOK 
      Interval        =   1000
      Left            =   3435
      Top             =   195
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   3450
      Top             =   825
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Timer TimerControlWordObert 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1530
      Top             =   90
   End
   Begin VB.CommandButton bOkvisualització 
      BackColor       =   &H006BEBB1&
      Caption         =   "Tancar Visualtizació del Document"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   8025
      Picture         =   "formrevisarCQ.frx":048A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5850
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H005C31DD&
      Caption         =   "Imprimir Etiqueta VQ"
      Height          =   1170
      Left            =   405
      Picture         =   "formrevisarCQ.frx":0760
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   11115
      Width           =   6900
   End
   Begin VB.CommandButton bverificarCB 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Valor Codi Barres: "
      Height          =   405
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2310
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Frame Frame4 
      Caption         =   "L E C T U R A   D E L S    D E L T E S"
      Height          =   4530
      Left            =   270
      TabIndex        =   10
      Top             =   6540
      Width           =   7335
      Begin VB.CommandButton bcolors 
         Caption         =   "Colors"
         Height          =   660
         Index           =   5
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3660
         Width           =   6900
      End
      Begin VB.CommandButton bcolors 
         Caption         =   "Colors"
         Height          =   660
         Index           =   4
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3000
         Width           =   6900
      End
      Begin VB.CommandButton bcolors 
         Caption         =   "Colors"
         Height          =   660
         Index           =   3
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2340
         Width           =   6900
      End
      Begin VB.CommandButton bcolors 
         Caption         =   "Colors"
         Height          =   660
         Index           =   2
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1680
         Width           =   6900
      End
      Begin VB.CommandButton bcolors 
         Caption         =   "Colors"
         Height          =   660
         Index           =   1
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1020
         Width           =   6900
      End
      Begin VB.CommandButton bcolors 
         Caption         =   "Colors"
         Height          =   660
         Index           =   0
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   6900
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Verificar codi de barres"
      Height          =   2760
      Left            =   240
      TabIndex        =   4
      Top             =   2775
      Width           =   7380
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   1650
      End
      Begin VB.CommandButton bmotius 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Motiu 4"
         Height          =   855
         Index           =   3
         Left            =   5490
         Picture         =   "formrevisarCQ.frx":0A9B
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   435
         Width           =   1740
      End
      Begin VB.CommandButton bmotius 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Motiu 3"
         Height          =   855
         Index           =   2
         Left            =   3690
         Picture         =   "formrevisarCQ.frx":0F25
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   435
         Width           =   1785
      End
      Begin VB.CommandButton bmotius 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Motiu 2"
         Height          =   855
         Index           =   1
         Left            =   1890
         Picture         =   "formrevisarCQ.frx":13AF
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   435
         Width           =   1785
      End
      Begin VB.CommandButton bmotius 
         BackColor       =   &H00F1B75F&
         Caption         =   "Motiu 1"
         Height          =   855
         Index           =   0
         Left            =   75
         Picture         =   "formrevisarCQ.frx":1839
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   435
         Width           =   1785
      End
      Begin VB.Label etgraulectura 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H005C31DD&
         Height          =   285
         Left            =   5475
         TabIndex        =   24
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label etnumcodidebarres 
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
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Top             =   150
         Width           =   3480
      End
      Begin VB.Label etresultatlecturacodibarres 
         Alignment       =   2  'Center
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
         ForeColor       =   &H005C31DD&
         Height          =   390
         Left            =   1155
         TabIndex        =   19
         Top             =   2235
         Width           =   5325
      End
      Begin VB.Label etestatuslectura 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Esperant la lectura del codi de barres..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F1B75F&
         Height          =   660
         Left            =   555
         TabIndex        =   9
         ToolTipText     =   "Esperant la lectura del codi de barres..."
         Top             =   1710
         Visible         =   0   'False
         Width           =   6420
      End
   End
   Begin VB.Frame Framepdf 
      Caption         =   "Veure el  PDF"
      Height          =   1845
      Left            =   4515
      TabIndex        =   1
      Top             =   435
      Width           =   2595
      Begin VB.CommandButton Command2 
         Height          =   1155
         Left            =   675
         Picture         =   "formrevisarCQ.frx":1CC3
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1185
      End
   End
   Begin VB.Frame Frameimp 
      BackColor       =   &H00C0C0FF&
      Caption         =   "R E V I S A R   IMP "
      Height          =   1845
      Left            =   300
      TabIndex        =   0
      Top             =   420
      Width           =   2595
      Begin VB.CommandButton Command1 
         Height          =   1155
         Left            =   735
         Picture         =   "formrevisarCQ.frx":23F0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   390
         Width           =   1185
      End
   End
   Begin VB.Frame FramePCC 
      BackColor       =   &H00FFFFFF&
      Height          =   12480
      Left            =   7830
      TabIndex        =   25
      Top             =   450
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton bPCC 
         BackColor       =   &H0025EFAD&
         Caption         =   "Revisat"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   11295
         Width           =   6450
      End
      Begin VB.Image Image1 
         Height          =   11025
         Left            =   465
         Picture         =   "formrevisarCQ.frx":2BEE
         Stretch         =   -1  'True
         Top             =   225
         Width           =   7425
      End
   End
   Begin VB.Label cnumbobina 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   4890
      TabIndex        =   18
      Top             =   30
      Width           =   2715
   End
   Begin VB.Label cComanda 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   2475
      TabIndex        =   17
      Top             =   15
      Width           =   2715
   End
   Begin VB.Menu mopcions 
      Caption         =   "Opcions"
      Begin VB.Menu mcrearcodibarres 
         Caption         =   "Crear Codi de barres"
      End
      Begin VB.Menu mbobinazero 
         Caption         =   "Revisió Bobina Zero"
      End
   End
End
Attribute VB_Name = "formrevisarCQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long


Private Const SW_MAXIMIZE = 3
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const HWND_NOTOPMOST = -2
Const PROCESS_TERMINATE = &H1
Const SYNCHRONIZE = &H100000
Const PROCESS_QUERY_INFORMATION = &H400
Const cColorVermell = &HC0C0FF
Const cColorVerd = &HC0FFC0
Const cColorBlau = &HF1B75F
Dim vrutafitxersescaner As String
Dim vultimahoralectura As String
Dim dbbaixesannex As Database
Dim objWord As Object
Dim objDoc As Object
Dim vidProcesPdf As Long
  

Private Sub bcolors_Click(Index As Integer)
    Dim vnumc As Double
    Dim vnumbob As Double
    vnumc = cadbl(cComanda.tag): vnumbob = cadbl(cnumbobina.tag)
    demanarvalorsdelta vnumc, vnumbob, cadbl(bcolors(Index).tag), Mid(bcolors(Index).caption, 1, InStr(1, bcolors(Index).caption, " [")), formrevisarCQ.Left + (formrevisarCQ.width / 2), formrevisarCQ.Top + 3400
    
    carregar_dades_bobina vnumc, vnumbob, True
    ratoli "normal"
End Sub
Sub demanarvalorsdelta(vnumc As Double, vnumbob As Double, vcoditinta As Double, vnomtinta As String, vX As Double, vY As Double)
    Dim rst As Recordset
    Dim v As Double
    Dim vresp As String
    Dim vi As Byte
    Dim vidtreball As Double
    Dim vordremodificacio As Double
    Dim vvalordelta As String
    Dim vdeltamaxim As Double
    Dim vsql As String
    Dim vdeltamaximTINTES As Double
    vi = 1
    vdeltamaxim = 2.5
    vdeltamaximTINTES = 2
    
demanardelta:
    v = form1.demanarvalordelta(vnomtinta, vdeltamaxim, vX, vY)
    If v = 9 Then GoTo sensedelta
    If v > vdeltamaxim Then
            vmsg = "Delta màxim superat  Màx:" + atrim(vdeltamaxim) + "  Llegit:" + atrim(v)
            If vmsg = "" Then GoTo demanardelta
    End If
    If v <= 0 Then vmsg = "Delta llegit valor zero."
    If vmsg <> "" Then vresp = InputBox("Escriu el motiu de: " + Chr(10) + vmsg + "SI T'HAS EQUIVOCAT POSSA [ERROR] SISPLAU")
    If vmsg <> "" Then enviaremailgeneric "missatgesgenericsimpresores", "ERROR Lectura Delta [" + nommaq + "] - " + form1.nomoperari + "  Comanda:" + atrim(vnumc), "Color: " + atrim(vnomtinta) + Chr(10) + treure_apostruf("Error: " + vmsg + Chr(10) + "Resposta operari: " + vresp)
    If v > vdeltamaximTINTES Then enviaremailgeneric "tintes@inplacsa.com", "ERROR Lectura Delta>2 [" + nommaq + "] - " + form1.nomoperari + "  Comanda:" + atrim(vnumc), "Color: " + atrim(vnomtinta) + Chr(10) + treure_apostruf("Error: " + vmsg + Chr(10) + "Resposta operari: " + vresp)
    guardarvalordelta vnumc, vnumbob, v, vcoditinta, vnomtinta, cadbl(form1.tmetres)
    dbbaixesannex.Execute "update RevisioCQ_Deltes set valordelta=" + passaradecimalpunt(atrim(v)) + " where comanda=" + atrim(vnumc) + " and numbobina=" + atrim(vnumbob) + " and coditinta=" + atrim(vcoditinta)
sensedelta:
          '"missatgesgenericsimpresores"
    vmsg = ""
    vresp = ""
    vvalordelta = atrim(Redondejar(v, 2))
    
fi:
   ' If cadbl(vvalordelta) = 9 Then vvalordelta = "N/S"
   '      llistat.Formulas(vcont) = "delta" + atrim(vi) + "='VD: " + vvalordelta + " - " + atrim(rst!Color) + "'"
    '      vcont = vcont + 1
     '     vi = vi + 1
   '   End If
      
End Sub
Sub PossarElsValorsDeltaalLlistat(vcont As Byte)
    Dim rst As Recordset
    Dim vi As Byte
    vi = 1
    Set rst = dbbaixesannex.OpenRecordset("select * from  RevisioCQ_Deltes where comanda=" + atrim(cComanda.tag) + " and numbobina=" + atrim(cnumbobina.tag))
    While Not rst.EOF
      vvalordelta = atrim(rst!valordelta)
      If cadbl(rst!valordelta) = 9 Then vvalordelta = "N/S"
      llistat.Formulas(vcont) = "delta" + atrim(vi) + "='VD: " + vvalordelta + " - " + atrim(rst!nomtinta) + "'"
      vcont = vcont + 1
      vi = vi + 1
      rst.MoveNext
    Wend
    Set rst = Nothing
      
End Sub
Sub guardarvalordelta(vnumc As Double, vnumbob As Double, vvalor As Double, vcoditinta As Double, vnomtinta As String, vmetres As Double)
  Dim vvalues As String
  vvalues = "(" + atrim(vnumbob) + "," + atrim(vnumc) + ",now," + atrim(vmetres) + "," + atrim(vcoditinta) + ",'" + treure_apostruf(vnomtinta) + "'," + atrim(numop) + "," + passaradecimalpunt(atrim(vvalor)) + ")"
  dbtmpb.Execute "insert into impresores_valorsdelta (numbobina,comanda,hora,metres,coditinta,nomdelatinta,operari,valordelta) values " + vvalues
End Sub


Private Sub bmotius_Click(Index As Integer)
   
   If nummaq = 7 Then motius_manualment: GoTo fi
   If etestatuslectura.visible Then etestatuslectura.visible = False: GoTo fi
   vultimahoralectura = Now
   llegir_motiu Index, cadbl(cComanda.tag), cadbl(cnumbobina.tag)
fi:
carregar_dades_bobina cadbl(cComanda.tag), cadbl(cnumbobina.tag)
ratoli "normal"
End Sub
Sub motius_manualment()
   Dim vvalorescaner As Double
   Dim vAvg As String
   SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
   vvalorescaner = form1.demanarvalorescaner(cadbl(cComanda.tag), formrevisarCQ.Left + (formrevisarCQ.width / 2), formrevisarCQ.Top + (formrevisarCQ.Height / 2))
   If vvalorescaner <> 0 Then
      vAvg = valorgrauLletra(vvalorescaner)
      dbbaixesannex.Execute "update RevisioCQ_codibarres set fitxerescaner='Manual',averagegrade='(" + atrim(vAvg) + ")' where comanda=" + atrim(cComanda.tag) + " and numbobina=" + cnumbobina.tag
   End If
   SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
Sub llegir_motiu(Index As Integer, vnumc As Double, vnumbobina As Double)
     bmotius(Index).BackColor = cColorBlau
     etresultatlecturacodibarres = ""
     etestatuslectura.visible = True: Timer1.Enabled = True: etestatuslectura.tag = ""
     dbbaixesannex.Execute "update RevisioCQ_CodiBarres SET FITXERESCANER='' where comanda=" + atrim(vnumc) + " and numbobina=" + atrim(vnumbobina) + " and nummotiu=" + atrim(Index + 1)
     dbbaixesannex.Execute "update RevisioCQ_CodiBarres SET AverageGrade='' where comanda=" + atrim(vnumc) + " and numbobina=" + atrim(vnumbobina) + " and nummotiu=" + atrim(Index + 1)
     While etestatuslectura.visible
        wait 1
     Wend
     If etestatuslectura.tag <> "" Then
         If InStr(1, etresultatlecturacodibarres.tag, "(F)") = 0 Then
             dbbaixesannex.Execute "update RevisioCQ_CodiBarres SET FITXERESCANER='" + atrim(etestatuslectura.tag) + "' where comanda=" + atrim(vnumc) + " and numbobina=" + atrim(vnumbobina) + " and nummotiu=" + atrim(Index + 1)
             dbbaixesannex.Execute "update RevisioCQ_CodiBarres SET AverageGrade='" + atrim(etresultatlecturacodibarres.tag) + "' where comanda=" + atrim(vnumc) + " and numbobina=" + atrim(vnumbobina) + " and nummotiu=" + atrim(Index + 1)
             etgraulectura = "Grau: " + atrim(valormigdelgrau)
         End If
     End If
    
     carregar_dades_bobina vnumc, vnumbobina
     
End Sub
Function valormigdelgrau() As String
   Dim rst As Recordset
   Dim vsum As Double
   Dim vcont As Double
   Set rst = dbbaixesannex.OpenRecordset("Select * from  RevisioCQ_CodiBarres where comanda=" + atrim(cComanda.tag) + " and numbobina=" + atrim(cnumbobina.tag))
   If Not rst.EOF Then rst.MoveLast: rst.MoveFirst: vcont = rst.RecordCount
   While Not rst.EOF
      vsum = vsum + valorgrau(lletragrau(atrim(rst!averagegrade)))
      rst.MoveNext
   Wend
   If vcont > 0 Then valormigdelgrau = valorgrauLletra(Redondejar(vsum / vcont, 0))
   
End Function
Function valorgrauLletra(v As Double) As String
   If v = 1 Then valorgrauLletra = "F"
   If v = 4 Then valorgrauLletra = "A"
   If v = 3 Then valorgrauLletra = "B"
   If v = 2 Then valorgrauLletra = "C"
   
End Function

Function valorgrau(v As String) As Double
   If v = "F" Or v = "G" Then valorgrau = 1
   If v = "A" Then valorgrau = 4
   If v = "B" Then valorgrau = 3
   If v = "C" Then valorgrau = 2
   If v = "D" Then valorgrau = 2
End Function
Function lletragrau(v As String) As String
   Dim vX As String
   vX = " "
   If InStr(1, v, "(") <> 0 Then
       vX = Mid(v, InStr(1, v, "(") + 1, 1)
   End If
   lletragrau = vX
End Function

Private Sub bOkvisualització_Click()
  Dim vnumc As Double
  Dim vi As Long
  Dim vespdf As Boolean
  Dim hProcess As Long
  vnumc = cadbl(cComanda.tag)
    If bOkvisualització.tag = "PDF" Then vespdf = True
  possar_boto_ok False
  If vespdf Then
       hProcess = OpenProcess(PROCESS_TERMINATE, 0, vidProcesPdf)
       vi = TerminateProcess(hProcess, 0): Exit Sub
  End If
  If atrim(objWord) <> "" Then
          objDoc.Close SaveChanges:=False
          objWord.Quit
  End If

  Set objDoc = Nothing
  Set objWord = Nothing
  dbbaixesannex.Execute "update RevisioCQ set imp_verificat=True where comanda=" + atrim(vnumc)
 
  carregar_imp_pdf
  ratoli "normal"

End Sub

Private Sub bPCC_Click()
FramePCC.visible = False
End Sub

Private Sub brevissioPCC_Click()
  FramePCC.Left = 0
  FramePCC.Top = 30
  FramePCC.visible = True
  FramePCC.ZOrder 0
  dbbaixesannex.Execute "update RevisioCQ_Extres set PCCrevisat=True where comanda=" + atrim(cadbl(cComanda.tag)) + " and numbobina=" + atrim(cadbl(cnumbobina.tag))
  carregar_imp_pdf
End Sub

Private Sub bverificarCB_Click()
   Dim vvalorescaner As Double
   Dim vnumc As Double
   vnumc = cadbl(cComanda.tag)
   vvalorescaner = form1.demanarvalorescaner(vnumc, formrevisarCQ.Left + (formrevisarCQ.width / 2), formrevisarCQ.Top + (formrevisarCQ.Height / 2))
   If vvalorescaner <> 0 Then dbbaixesannex.Execute "update RevisioCQ set valorCodiBarres=" + passaradecimal(atrim(vvalorescaner)) + " where comanda=" + atrim(vnumc)
   carregar_imp_pdf
End Sub
Sub possar_boto_ok(vActivar As Boolean, Optional vespdf As Boolean)
  Static vampleform As Double
  Static valtform As Double
  If cadbl(vampleform) = 0 Then vampleform = formrevisarCQ.width: valtform = formrevisarCQ.Height
  If vActivar Then
   If Not vespdf Then TimerControlWordObert.Enabled = True
   bOkvisualització.Top = 50
   bOkvisualització.Left = 50
   bOkvisualització.visible = True
   formrevisarCQ.width = bOkvisualització.width + 180
   formrevisarCQ.Height = bOkvisualització.Height + 530
   If vespdf Then bOkvisualització.tag = "PDF"
     Else
        bOkvisualització.visible = False
        formrevisarCQ.width = vampleform
        formrevisarCQ.Height = valtform
        bOkvisualització.tag = ""
  End If
End Sub
Function ObtenirhWndDocumentWord(objDoc As Object) As Long
    Dim hWndWord As Long
    Dim strTitle As String

    strTitle = objDoc.Windows(1).caption & " - Microsoft Word" 'Obtenim el titol de la finestra.
    hWndWord = FindWindow("OpusApp", strTitle) 'Busquem la finestra amb el titol.
    ObtenirhWndDocumentWord = hWndWord
End Function
Private Sub Command1_Click()
  Dim rstc As Recordset
  Dim vnumc As Double
  Dim vnomdocument As String
  Dim vcont As Byte
  Dim v As Long
  TancarTotsElsWords
  vnumc = cadbl(cComanda.tag)
  
  Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
  obrir_imp_treball cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), cadbl(rstc!client), cadbl(rstc!direnvio), vnomdocument, objDoc, objWord
  If vnomdocument <> "" Then
      'v = ObtenirhWndDocumentWord(objDoc)
      'ShowWindow v, SW_MAXIMIZE
      CanviarVistaDocumentWord objDoc
      possar_boto_ok True
     
     'If vcont >= 5 Then MsgBox "No he trobat el document " + atrim(vnomdocument), vbCritical, "Error"
  End If
  Set rstc = Nothing
  ratoli "normal"
  
End Sub
Public Sub TancarTotsElsWords()
    Dim objWordApp As Object
    Dim objDoc As Object
    Dim i As Integer

    On Error GoTo ErrorHandler

    ' Intentar obtenir una instància existent de Word
    Set objWordApp = GetObject(, "Word.Application")

    If Not objWordApp Is Nothing Then
        ' Tancar tots els documents sense guardar canvis (pots ajustar-ho)
        For Each objDoc In objWordApp.Documents
            objDoc.Close SaveChanges:=False ' Pots posar True o wdDoNotSaveChanges
        Next objDoc

        ' Tancar l'aplicació de Word
        objWordApp.Quit SaveChanges:=False ' Pots posar True o wdDoNotSaveChanges
    Else
        MsgBox "No s'ha trobat cap instància de Word oberta.", vbInformation
    End If

Exit Sub

ErrorHandler:
  
    Set objWordApp = Nothing
End Sub
Sub CanviarVistaDocumentWord(objDoc As Object)
    ' Amaga les barres d'eines
  '  objDoc.Application.DisplayToolbars = False

    ' Canvia la vista a "Print Layout" (Disseny d'impressió)
    objDoc.ActiveWindow.View.Type = 3 ' wdPrintView

    ' Maximitzar la finestra per ocupar tota l'amplada
    objDoc.ActiveWindow.WindowState = 1 ' wdWindowStateMaximize

    ' Canvia el zoom per adaptar-se a l'amplada de la pàgina
    objDoc.ActiveWindow.View.Zoom.PageFit = 2
End Sub
Sub obrir_imp_treball(treball As Double, modificacio As Double, codiclient As Double, direnvio As Double, vnomdocument As String, objDoc As Object, objWord As Object)
   Dim generarfitxer_imp As String
   If modificacio = 0 Then modificacio = 1
   generarfitxer_imp = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\IMP" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "-" + Format(codiclient, "000000") + "_" + atrim(direnvio) + ".doc"
   If Not existeix(generarfitxer_imp) Then generarfitxer_imp = generarfitxer_imp + "x"
   If existeix(generarfitxer_imp) Then
     'obrir_document generarfitxer_imp
      Set objWord = CreateObject("Word.Application")
      objWord.visible = True  ' Mostrar Word
      Set objDoc = objWord.Documents.Open(generarfitxer_imp)
      Sleep 2
     vnomdocument = generarfitxer_imp
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + generarfitxer_imp, vbCritical, "Error"
  End If
    
End Sub

Private Sub Command2_Click()
Dim rstc As Recordset
  Dim vnumc As Double
  Dim vnomdocument As String
  Dim vcont As Byte
  vnumc = cadbl(cComanda.tag)
  Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
  vidProcesPdf = 0
  possar_boto_ok True, True
  bOkvisualització.Enabled = False
  obrir_pdf_treball rstc!numtreball, rstc!numordremodificacio, vnomdocument
  Set rstc = Nothing
  vcont = 0
  While estaobertelPDF(vidProcesPdf)
    wait 1
    vcont = vcont + 1
    If formrevisarCQ.visible = False Then
       bOkvisualització.tag = "PDF"
       bOkvisualització_Click
       Unload formrevisarCQ
    End If
    If vcont < 3 Then bOkvisualització.Enabled = False Else bOkvisualització.Enabled = True
  Wend
  possar_boto_ok False, True
  ratoli "normal"
End Sub
Function estaobertelPDF(vidPdf As Long) As Boolean
  Dim hProcess As Long
  Dim lExitCode As Long
 hProcess = OpenProcess(SYNCHRONIZE + PROCESS_QUERY_INFORMATION + PROCESS_TERMINATE, 0, vidPdf)
  estaobertelPDF = True
  If hProcess <> 0 Then
            DoEvents ' Permet que altres esdeveniments es processin
            If GetExitCodeProcess(hProcess, lExitCode) Then
                If lExitCode <> 259 Then ' 259 significa que el procés encara s'està executant (STILL_ACTIVE)
                    estaobertelPDF = False
                    Else:
                     estaobertelPDF = True
                End If

            End If
              Else: estaobertelPDF = False
  End If
End Function

Sub obrir_pdf_treball(treball As Double, modificacio As Double, Optional vnomfitxerpdf As String)
   Dim generarfitxer_pdf As String
   If modificacio = 0 Then modificacio = 1
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "_SC.pdf"
   If Not existeix(generarfitxer_pdf) Then generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + ".pdf"
   If existeix(generarfitxer_pdf) Then
     If vnomfitxerpdf <> "noobrir" Then
           ' obrir_document generarfitxer_pdf
           If Not existeix(generarfitxer_pdf) Then MsgBox "No hi ha el PDF"
           vidProcesPdf = ShellExecute(Screen.ActiveForm.hwnd, "Open", generarfitxer_pdf, "", "", SW_MAXIMIZE)
           wait 2
           vidProcesPdf = buscar_finestra(substituir(generarfitxer_pdf, rutadelfitxer(generarfitxer_pdf), ""))
           v = GetWindowThreadProcessId(vidProcesPdf, vidProcesPdf)
           
     End If
     vnomfitxerpdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + ".pdf"
    'Else: MsgBox "No he trobat el fitxer" + Chr(10) + generarfitxer_pdf + Chr(10) + " i tampoc el de separació de colors.", vbCritical, "Error"
  End If
End Sub

Private Sub Command3_Click()
  If Command3.BackColor = &H5C31DD Then Exit Sub
impresio_etiqueta
wait 2
Frame3.Enabled = False
Frame4.Enabled = False
Command3.Enabled = False
ratoli "normal"
'Unload formrevisarCQ
End Sub
Sub impresio_etiqueta()
  Dim ultimalinia As String
   Dim vvalorescaner As String
   Dim vindexformules As Byte
   Dim i As Byte

   ultimalinia = "Op: " + atrim(numop) + "    NºBob.Salida: " + atrim(cnumbobina.tag) + "   Fecha: " + Format(Now, "dd/mm/yy")
   For i = 0 To 100
     llistat.Formulas(i) = ""
   Next i
   llistat.Formulas(0) = "lot=" + atrim(cComanda.tag)
   llistat.Formulas(1) = "ultimalinia='" + atrim(ultimalinia) + "'"
   llistat.Formulas(2) = "nommaquina='" + atrim(nummaq) + "-" + atrim(nommaq) + "'"
   llistat.Formulas(3) = "nummaq='" + atrim(nummaq) + "'"
   
   vvalorescaner = "[" + atrim(valormigdelgrau) + "]"
   'If vvalorescaner = "" Or vvalorescaner = "0" Then Exit Sub
   llistat.Formulas(4) = "valorescaner='" + atrim(vvalorescaner) + "'"
   If vvalidaciocodidebarres <> "-" Then
       llistat.Formulas(5) = "codidebarres='CB: " + atrim(etnumcodidebarres) + "'"
        Else: llistat.Formulas(5) = ""
   End If
'   If vdigimarc Then
'        llistat.Formulas(6) = "digimarc='Digimarc OK'"
'          Else: llistat.Formulas(6) = "digimarc=''"
'   End If
   
   If (vavispeu <> "") And numbob = 0 Then
      MsgBox "Peu/Data: [" + vavispeu + "] VERIFICA'L." + Chr(10) + "Fes [OK] PER CONTINUAR.", vbExclamation + vbOKOnly, "PEU IMPRENTA"
   End If
   llistat.Formulas(7) = "peuimprenta='" + atrim(vavispreu) + "'"
'   llistat.Formulas(7) = "peuimprenta='prova 123'"
   vindexformules = 8
   PossarElsValorsDeltaalLlistat vindexformules
   form1.calcularvalorsreducciocilindre cadbl(cComanda.tag), nummaq, vindexformules, llistat
   llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "verificacioqualitatimpresores.rpt"
   llistat.Destination = crptToPrinter
    llistat.CopiesToPrinter = 1
   llistat.DataFiles(0) = ""
   llistat.DiscardSavedData = True
' llistat.PrinterName = llegir_ini("Impressores", "nomfulla", "baixesimpressora.ini")
' llistat.PrinterPort = llegir_ini("Impressores", "portfulla", "baixesimpressora.ini")
' llistat.PrinterDriver = llegir_ini("Impressores", "driverfulla", "baixesimpressora.ini")
   form1.escullir_impresora_tickets
   DoEvents
   If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
   llistat.Action = 1
   llistat.PrinterDriver = ""
   llistat.PrinterName = ""
   llistat.PrinterPort = ""
   'MsgBox "ATENCIÓ CONTROL DE VERIFICACIÓ DE QUALITAT." + Chr(10) + "VERIFICA LA IMPRESIÓ AMB L'ETIQUETA IMPRESA", vbInformation, "VERIFICACIÓ QUALITAT"
End Sub

Private Sub Form_Activate()
  colocarFormalasegonapantalla
  
End Sub
Sub colocarFormalasegonapantalla()
  formrevisarCQ.Left = cadbl(llegir_ini("Baixes", "Impresores_FinestraetCQ_Left", "comandes.ini"))
  formrevisarCQ.Top = cadbl(llegir_ini("Baixes", "Impresores_FinestraetCQ_Top", "comandes.ini"))
 
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Load()
  Set dbbaixesannex = OpenDatabase(rutadelfitxer(cami) + "baixes_annex.mdb")
  vrutafitxersescaner = "\\serverprodu\Dades\progcomandes\dades\Lectures_Codisdebarres_Impresores" + atrim(nummaq) + "\"
  'carregar_dades 223993, 0
End Sub
Sub carregar_dades(vnumc As Double, vnumbobina As Double)
   Frame3.Enabled = True
   Frame4.Enabled = True
   Command3.Enabled = True

   etnumcodidebarres = "Sense codi de barres"
   cComanda = Format(vnumc, "#,000")
   cComanda.tag = atrim(vnumc)
   cnumbobina = "Bobina: " + atrim(vnumbobina)
   cnumbobina.tag = vnumbobina
   possar_totvermell
   crear_dades vnumc, vnumbobina
   etgraulectura = "Grau: " + atrim(valormigdelgrau)
   carregar_imp_pdf
   carregar_dades_bobina vnumc, vnumbobina
End Sub
Sub carregar_dades_bobina(vnumc As Double, vnumbobina As Double, Optional vnomesDeltes As Boolean)
Dim rst As Recordset
  Dim rstc As Recordset
  Dim rstmotius As Recordset
  Dim vnummotius As Integer
  
  Dim i As Byte
  '  dbbaixesannex.Execute "insert into RevisioCQ_CodiBarres (comanda,numbobina,nummotiu) values (" + atrim(cComanda.tag) + "," + atrim(cnumbobina.tag) + "," + atrim(i) + ")"
  '  dbbaixesannex.Execute "insert into RevisioCQ_Deltes (comanda,numbobina,coditinta,nomtinta) values (" + atrim(cComanda.tag) + "," + atrim(cnumbobina.tag) + "," + atrim(rst!coditinta) + ",'" + atrim(rst!color) + "')"
  Set rstc = dbbaixesannex.OpenRecordset("select * from RevisioCQ_CodiBarres where comanda=" + atrim(vnumc) + " and numbobina=" + atrim(vnumbobina) + " order by nummotiu")
  For i = 0 To 3
    If Not rstc.EOF Then
        bmotius(i).visible = True
        bmotius(i).BackColor = IIf(atrim(rstc!fitxerescaner) = "", cColorVermell, cColorVerd)
        bmotius(i).tag = atrim(rstc!fitxerescaner)
         Else: bmotius(i).visible = False
    End If
    If Not rstc.EOF Then rstc.MoveNext
  Next
  
  Set rstc = dbbaixesannex.OpenRecordset("select * from RevisioCQ_Deltes where comanda=" + atrim(vnumc) + " and numbobina=" + atrim(vnumbobina))
  For i = 0 To 5
    If Not rstc.EOF Then
        bcolors(i).visible = True
        bcolors(i).caption = rstc!nomtinta + " [DeltaE: " + atrim(rstc!valordelta) + "]"
        bcolors(i).BackColor = IIf(cadbl(atrim(rstc!valordelta)) = 0, cColorVermell, cColorVerd)
        bcolors(i).tag = atrim(rstc!coditinta)
         Else: bcolors(i).visible = False
    End If
    If Not rstc.EOF Then rstc.MoveNext
  Next
  
fi:
  Set rst = Nothing
  Set rstc = Nothing
End Sub
Sub crear_dades(vnumc As Double, vnumbobina As Double)
  Dim rst As Recordset
  Dim rstc As Recordset
  Dim rstmotius As Recordset
  Dim vnummotius As Integer
  Dim vnumcodibarres As String
  
  Dim i As Byte
  Dim vbandes As Double
  
  Set rstc = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
  If rstc.EOF Then GoTo fi
'    Clipboard.Clear
'  Clipboard.SetText "SELECT id_treball,nummodificacio, numerodemotius, midacilindre,midadesarroll FROM repasclixes LEFT JOIN repasdadestintes ON repasclixes.id_repas = repasdadestintes.id_repas where id_treball=" + atrim(rstc!numtreball) + " and nummodificacio=" + atrim(rstc!numordremodificacio) + " ORDER BY repasdadestintes.midadesarroll desc , repasdadestintes.numerodemotius DESC;"
  Set rst = dbclixes.OpenRecordset("SELECT clixes.codidebarres, Modificacions.ordre, Modificacions.bandes FROM clixes INNER JOIN Modificacions ON clixes.id_treball = Modificacions.id_treball WHERE clixes.id_treball=" + atrim(rstc!numtreball) + " and Modificacions.ordre=" + atrim(rstc!numordremodificacio))
  vbandes = 1
  If Not rst.EOF Then vbandes = cadbl(rst!bandes): vnumcodibarres = atrim(rst!codidebarres): etnumcodidebarres = vnumcodibarres
  Set rstc = dbclixes.OpenRecordset("SELECT id_treball,nummodificacio, numerodemotius, midacilindre,midadesarroll FROM repasclixes LEFT JOIN repasdadestintes ON repasclixes.id_repas = repasdadestintes.id_repas where id_treball=" + atrim(rstc!numtreball) + " and nummodificacio=" + atrim(rstc!numordremodificacio) + " ORDER BY repasdadestintes.midadesarroll desc , repasdadestintes.numerodemotius DESC;")
  
  If Not rstc.EOF Then
       If cadbl(rstc!midadesarroll) > 0 Then
           vnummotius = vbandes * (cadbl(rstc!midacilindre) / cadbl(rstc!midadesarroll))
       End If
  End If
  If vnummotius = 0 Then vnummotius = 1
  If vnumcodibarres = "" Then vnummotius = 0
  
  Set rst = dbbaixesannex.OpenRecordset("select * from RevisioCQ where comanda=" + atrim(vnumc))
  If rst.EOF Then
       dbbaixesannex.Execute "insert into RevisioCQ (comanda,imp_verificat) values (" + atrim(vnumc) + ",False)"
  End If
  Set rst = dbbaixesannex.OpenRecordset("select * from RevisioCQ_extres where comanda=" + atrim(cComanda.tag) + " and numbobina=" + atrim(cnumbobina.tag))
  If rst.EOF Then
       dbbaixesannex.Execute "insert into RevisioCQ_extres (comanda,numbobina) values (" + atrim(cComanda.tag) + "," + atrim(cnumbobina.tag) + ")"
  End If
  etgraulectura = "Grau: "
  If vnummotius > 4 Then vnummotius = 4
  For i = 1 To vnummotius
     Set rstmotius = dbbaixesannex.OpenRecordset("select * from RevisioCQ_Codibarres where comanda=" + atrim(cComanda.tag) + " and numbobina=" + atrim(cnumbobina.tag) + " and nummotiu=" + atrim(i))
     If rstmotius.EOF Then
         dbbaixesannex.Execute "insert into RevisioCQ_CodiBarres (comanda,numbobina,nummotiu) values (" + atrim(cComanda.tag) + "," + atrim(cnumbobina.tag) + "," + atrim(i) + ")"
     End If
  Next i
  vsql = "select * from tintes where (id_treball=" + atrim(cadbl(rstc!id_treball)) + " and ordremodificacio=" + atrim(cadbl(rstc!nummodificacio)) + " and tinterlinkambid_treball<1) or id_tinter in (select tinterlinkambid_treball  from tintes where id_treball=" + atrim(cadbl(rstc!id_treball)) + " and ordremodificacio=" + atrim(cadbl(rstc!nummodificacio)) + " and tinterlinkambid_treball>0)"
'  Clipboard.Clear
'  Clipboard.SetText vsql
  Set rst = dbclixes.OpenRecordset(vsql)
  While Not rst.EOF
      If InStr(1, atrim(rst!color), "P-") > 0 And InStr(1, atrim(rst!color), "PRIMAR") = 0 Then
          Set rstmotius = dbbaixesannex.OpenRecordset("select * from RevisioCQ_Deltes where comanda=" + atrim(cComanda.tag) + " and numbobina=" + atrim(cnumbobina.tag) + " and coditinta=" + atrim(rst!coditinta))
          If rstmotius.EOF Then
               dbbaixesannex.Execute "insert into RevisioCQ_Deltes (comanda,numbobina,coditinta,nomtinta) values (" + atrim(cComanda.tag) + "," + atrim(cnumbobina.tag) + "," + atrim(rst!coditinta) + ",'" + atrim(rst!color) + "')"
          End If
      End If
      rst.MoveNext
  Wend
  
fi:
  Set rst = Nothing
  Set rstc = Nothing
End Sub
Sub carregar_imp_pdf()
  Dim rst As Recordset
  Set rst = dbbaixesannex.OpenRecordset("select * from RevisioCQ where comanda=" + atrim(cComanda.tag))
  If Not rst.EOF Then
      If rst!imp_verificat Then Frameimp.BackColor = cColorVerd
      'If rst!pdf_verificat Then Framepdf.BackColor = cColorVerd
'      If cadbl(rst!valorcodibarres) <> 0 Then bverificarCB.BackColor = cColorVerd: bverificarCB.caption = "Valor Codi Barres: " + atrim(cadbl(rst!valorcodibarres))
  End If
  Set rst = dbbaixesannex.OpenRecordset("select * from RevisioCQ_extres where comanda=" + atrim(cComanda.tag) + " and numbobina=" + atrim(cadbl(cnumbobina.tag)))
  If Not rst.EOF Then If rst!PCCrevisat Then brevissioPCC.BackColor = cColorVerd
  Set rst = Nothing
End Sub
Sub possar_totvermell()
  Dim i As Byte
  brevissioPCC.BackColor = cColorVermell
  Frameimp.BackColor = cColorVermell
  'Framepdf.BackColor = cColorVermell
  bverificarCB.BackColor = cColorVermell
  bverificarCB.caption = "Valor Codi Barres: "
  For i = 0 To 3
    bmotius(i).BackColor = cColorVermell
  Next i
  For i = 0 To 5
    bcolors(i).BackColor = cColorVermell
  Next i
End Sub

Private Sub Form_Paint()
   
  ' escriure_ini "Baixes", "Impresores_FinestraetCQ_Left", formrevisarCQ.Left, "comandes.ini"
  ' escriure_ini "Baixes", "Impresores_FinestraetCQ_top", formrevisarCQ.Top, "comandes.ini"
  ' colocarFormalasegonapantalla
End Sub

Private Sub Form_Unload(Cancel As Integer)
   escriure_ini "Baixes", "Impresores_FinestraetCQ_Left", formrevisarCQ.Left, "comandes.ini"
   escriure_ini "Baixes", "Impresores_FinestraetCQ_top", formrevisarCQ.Top, "comandes.ini"
   Set dbbaixesannex = Nothing
End Sub

Private Sub Framepdf_DblClick()
  Unload formrevisarCQ
End Sub

Private Sub mbobinazero_Click()
  form1.imprimir_controlqualitat cadbl(form1.comanda), numop, 0
End Sub

Private Sub mcrearcodibarres_Click()
   Dim vlink As String
   Dim vcodi As String
   MsgBox "Al generar aquest codi de barres s'enviarà un correu al encarregat per avisar que s'ha generat un codi alternatiu.", vbCritical, "Codi de barres"
   vlink = "https://barcode.tec-it.com/barcode.ashx?data=ElCodiEAN13&code=EAN13&translate-esc=on"
   vcodi = InputBox("Entra el codi de barres que vols crear EAN13" + vbNewLine + "S'OBRIRÀ UNA PAGINA WEB A L'ALTRA PANTALLA FES IMPRIMIR CTRL+P PER IMPRIMIR-LA.", "Codi de barres", etnumcodidebarres)
   If atrim(vcodi) <> "" Then
      ' Shell "cmd.exe /c """ + substituir(vlink, "ElCodiEAN13", vcodi) + """"
        Call ShellExecute(0, "open", substituir(vlink, "ElCodiEAN13", vcodi), "", "", SW_SHOWNORMAL)
        enviaremailgeneric "impresores@inplacsa.com", "Codi de Barres generat desde [" + nommaq + "] - " + form1.nomoperari + "  Comanda:" + atrim(form1.comanda), "S'ha generat un codi de barres manualment per poder saltar la lectura del motiu."
   End If
   
End Sub

Private Sub Timer1_Timer()
  Static vpos As Integer
  Dim rst As Recordset
  Static vcont As Byte
  Dim vintents As Byte
  Dim vnomfitxerLectures As String
  vnomfitxerLectures = rutadelfitxer(cami) + "Lectures_Codisdebarres_Impresores\" + atrim(nummaq) + "\guardado automático.csv"
  If vpos = 0 Then vpos = 1
  'If etestatuslectura.visible = False Then vpos = 1: etestatuslectura.visible = True
  etestatuslectura = Mid(etestatuslectura.ToolTipText, vpos) + "   " + Mid(etestatuslectura.ToolTipText, 1, vpos)
  vpos = vpos + 1
  If vpos = Len(etestatuslectura.ToolTipText) Then vpos = 1
  vcont = vcont + 1
  vintents = 0
  If vcont = 10 Then
    vcont = 0
    If fitxerobert(vnomfitxerLectures) Then GoTo fi
    If existeix("c:\temp\LecturaCodideBarres_" + atrim(nummaq) + ".csv") Then Kill "c:\temp\LecturaCodideBarres_" + atrim(nummaq) + ".csv"
    On Error GoTo errordll
access_fitxer:
    Copiar_Fitxer vnomfitxerLectures, "c:\temp\LecturaCodideBarres_" + atrim(nummaq) + ".csv"
    wait 1
    Set dbbaixesannex = OpenDatabase(rutadelfitxer(cami) + "baixes_annex.mdb")
    Set rst = dbbaixesannex.OpenRecordset("SELECT [Average Grade],Codibarres_9.Code, Codibarres_9.Date, Codibarres_9.Time, Codibarres_9.Filename From Codibarres_9 ;", dbOpenSnapshot, dbAppendOnly )
    On Error GoTo 0
    If Not rst.EOF Then
        rst.MoveLast
        'If vultimahoralectura = "" Then vultimahoralectura = atrim(rst!Date) + " " + atrim(rst!Time)
        If DateDiff("s", vultimahoralectura, atrim(rst!Date) + " " + atrim(rst!Time)) > 0 Then
           If cadbl(treureespais(atrim(rst!code))) <> cadbl(treureespais(atrim(etnumcodidebarres))) Then MsgBox "Aquest codi de barres escanejat no coincideix amb el de la comanda." + vbNewLine + "Treball:" + atrim(cadbl(treureespais(atrim(etnumcodidebarres)))) + "  ->  Escanejat: " + atrim(cadbl(treureespais(atrim(rst!code)))), vbCritical, "Error": Set rst = Nothing: vcont = 0: Exit Sub
          ' vultimahoralectura = atrim(rst!Date) + " " + atrim(rst!Time)
           etestatuslectura.tag = substituir(rst!filename, vrutafitxersescaner, "")
           etresultatlecturacodibarres.tag = rst![Average Grade]
           etresultatlecturacodibarres = "Avg_Grade: " + rst![Average Grade]
           Timer1.Enabled = False
           etestatuslectura.visible = False
        End If
    End If
    Set rst = Nothing
  End If
fi:
  Exit Sub
errordll:
   vintents = vintents + 1
   If vintents < 4 Then wait 1: GoTo access_fitxer
    MsgBox "Error accedint a les dades del escaner." + vbNewLine + "Assegura que el arxiu mstext35.dll estigui copiat a syswow64 i registrat.", vbCritical, "Error"
End Sub
Function treureespais(TextOriginal As String) As String
   Dim i As Long
    Dim CaracterActual As String
    Dim CadenaNeta As String
    CadenaNeta = ""
    For i = 1 To Len(TextOriginal)
        CaracterActual = Mid(TextOriginal, i, 1)
        If CaracterActual <> " " Then
            CadenaNeta = CadenaNeta & CaracterActual
        End If
    Next i
    treureespais = CadenaNeta
End Function

Function fitxerobert(vnomfitxer As String) As Boolean
  On Error GoTo Error
  Open vnomfitxer For Input Shared As #1
  Close #1
  Exit Function
Error:
  fitxerobert = True
End Function
Sub comprovar_sitotrevisat()
  Dim rst As Recordset
  Command3.BackColor = &H5C31DD
  vtotok = True
  Set rst = dbbaixesannex.OpenRecordset("select * from RevisioCQ where comanda=" + atrim(cComanda.tag))
  If Not rst.EOF Then If Not rst!imp_verificat Then vtotok = False
  Set rst = dbbaixesannex.OpenRecordset("select * from RevisioCQ_extres where comanda=" + atrim(cComanda.tag) + " and numbobina=" + atrim(cadbl(cnumbobina.tag)))
  If Not rst.EOF Then If Not rst!PCCrevisat Then vtotok = False
  If vtotok = True Then
    Set rst = dbbaixesannex.OpenRecordset("select * from RevisioCQ_CodiBarres where comanda=" + atrim(cComanda.tag) + " and numbobina=" + atrim(cnumbobina.tag))
    While Not rst.EOF
      If valorgrau(lletragrau(atrim(rst!averagegrade))) < 2 Then vtotok = False
      rst.MoveNext
    Wend
  End If
  If vtotok = True Then
      Set rst = dbbaixesannex.OpenRecordset("select * from RevisioCQ_Deltes where comanda=" + atrim(cComanda.tag) + " and numbobina=" + atrim(cnumbobina.tag) + " and valordelta=0")
      If Not rst.EOF Then vtotok = False
  End If
  If vtotok Then Command3.BackColor = &H6BEBB1
End Sub

Private Sub Timer2_Timer()
  
End Sub

Private Sub TimercontroltotOK_Timer()
comprovar_sitotrevisat
End Sub

Private Sub TimerControlWordObert_Timer()
  
  If atrim(objDoc) = "" Then TimerControlWordObert.Enabled = False: possar_boto_ok False
End Sub
Function buscar_finestra(partialWindowName As String)
    Dim hwnd As Long
    Dim currentTitle As String
    Dim currentTitleLength As Long
    Dim maxTitleLength As Long

    ' Inicialitza el handle a la primera finestra.
    hwnd = FindWindowEx(0, 0, vbNullString, vbNullString)

    ' Itera a través de totes les finestres.
    Do While hwnd <> 0
        ' Obté la longitud màxima del títol de la finestra.
        maxTitleLength = 255 ' Un valor raonable.

        ' Inicialitza la variable per emmagatzemar el títol.
        currentTitle = String$(maxTitleLength, 0)

        ' Obté el títol de la finestra.
        currentTitleLength = GetWindowText(hwnd, currentTitle, maxTitleLength)

        ' Si s'ha obtingut el títol i conté el títol parcial, retorna el handle.
        If currentTitleLength > 0 Then
            currentTitle = Left$(currentTitle, currentTitleLength)
            If InStr(1, LCase$(currentTitle), LCase$(partialWindowName)) > 0 Then
                'SetWindowPos hwnd, 0, x, y, cx, cy, 0 ' &H1
                buscar_finestra = hwnd
                Exit Function
            End If
        End If

        ' Obté el següent handle de finestra.
        hwnd = FindWindowEx(0, hwnd, vbNullString, vbNullString)
    Loop

    ' Si no troba la finestra, retorna 0.
    buscar_finestra = 0
End Function


