VERSION 5.00
Begin VB.Form formretorntintes 
   BackColor       =   &H00FDDECE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retorn de tintes"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7755
   Icon            =   "Retorn de tintes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bconsultarecarreges 
      Height          =   360
      Left            =   5835
      Picture         =   "Retorn de tintes.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Llista de recarregues pendents"
      Top             =   2295
      Width           =   345
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "RECARREGAR Llauna"
      Height          =   675
      Left            =   60
      Picture         =   "Retorn de tintes.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Recarregar una llauna amb una producció nova."
      Top             =   2310
      Width           =   5760
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1455
      Top             =   915
   End
   Begin VB.CommandButton bcancel 
      Caption         =   "Tancar Finestre"
      Height          =   675
      Left            =   6300
      Picture         =   "Retorn de tintes.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancelar el retorn"
      Top             =   2325
      Width           =   1320
   End
   Begin VB.CommandButton bok 
      Caption         =   " Retorn Llauna"
      Height          =   390
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Fer un retorn de tinta"
      Top             =   1275
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEE4D7&
      Height          =   1230
      Left            =   75
      TabIndex        =   2
      Top             =   15
      Width           =   7575
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FDDECE&
         Height          =   315
         Left            =   2880
         Picture         =   "Retorn de tintes.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Guardar la Tara per defecte."
         Top             =   195
         Width           =   345
      End
      Begin VB.TextBox pesnet 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4560
         TabIndex        =   4
         Top             =   510
         Width           =   1110
      End
      Begin VB.TextBox ctara 
         BackColor       =   &H00EEE4D7&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2115
         TabIndex        =   3
         Top             =   510
         Width           =   1110
      End
      Begin VB.Label pesdelatinta 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   285
         TabIndex        =   12
         Top             =   555
         Width           =   1410
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5715
         TabIndex        =   11
         Top             =   615
         Width           =   540
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3270
         TabIndex        =   10
         Top             =   570
         Width           =   540
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tara"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2310
         TabIndex        =   9
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Pes Net"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4785
         TabIndex        =   8
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3930
         TabIndex        =   7
         Top             =   495
         Width           =   390
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1845
         TabIndex        =   6
         Top             =   390
         Width           =   390
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Pes Bascula"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   135
         TabIndex        =   5
         Top             =   210
         Width           =   1785
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEE4D7&
      Height          =   1140
      Left            =   75
      TabIndex        =   13
      Top             =   1155
      Width           =   7605
      Begin VB.CommandButton Command5 
         Caption         =   " Retorn a Magatzem"
         Height          =   570
         Left            =   4650
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Fer un retorn de tinta"
         Top             =   525
         Width           =   1395
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Canvi de Situació"
         Height          =   765
         Left            =   6150
         Picture         =   "Retorn de tintes.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fer un retorn de tinta"
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton bcomposicio 
         BackColor       =   &H00F3B378&
         Caption         =   "Composició Bases Inkmaker"
         Height          =   570
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   270
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton Command3 
         Height          =   465
         Left            =   4050
         Picture         =   "Retorn de tintes.frx":2C6C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Imprimir etiqueta llauna"
         Top             =   345
         Width           =   480
      End
      Begin VB.TextBox nllauna 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1905
         TabIndex        =   15
         Top             =   345
         Width           =   2115
      End
      Begin VB.ComboBox colordelatinta 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4800
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "(Ex: A5000-A5009)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3120
         TabIndex        =   23
         Top             =   105
         Width           =   1470
      End
      Begin VB.Label etpesteoric 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   585
         TabIndex        =   19
         Top             =   840
         Width           =   5520
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº llauna"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   17
         Top             =   75
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Color de la Tinta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5250
         TabIndex        =   16
         Top             =   90
         Visible         =   0   'False
         Width           =   1140
      End
   End
End
Attribute VB_Name = "formretorntintes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, _
                                ByVal bInvert As Long) As Long
Const WM_USER = &H400
Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" _
                Alias "SendMessageA" _
                (ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
                
Private Sub bcomposicio_Click()
  ensenyar_formulacioxrXkg
End Sub
Sub ensenyar_formulacioxrXkg()
  Dim rst As Recordset
  Dim vmsg As String
  Dim vformula As String
  Dim vkg As Double
  Dim vcalcultanx100 As Double
  Dim vtitolmsg As String
  Dim nlinies As Double
  Dim vtxtprova As TextBox
  Dim hinici As Date
  Unload avis
  vformula = bcomposicio.tag
  
  vsql = "SELECT dbo.tblFormula.Code, dbo.tblFormula.Description, dbo.tblComponenti.DescComponente, [Quantity]/10 AS [%decomponent]"
  vsql = vsql + " FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente"
  vsql = vsql + " WHERE ((dbo.tblFormulaDetail.Formulation=0 and (dbo.tblFormula.Code)='" + vformula + "'));"
  Set rst = conODBC.OpenRecordset(vsql)
  
'  Set rst = dbtintes.OpenRecordset("SELECT Formules.codiformula, Formules.descripcioformula, Formules.series, Componentsbase.nomcomponent, DetallFormules.[%decomponent] FROM Componentsbase RIGHT JOIN (Formules LEFT JOIN DetallFormules ON Formules.idformula = DetallFormules.IDFormula) ON Componentsbase.idcomponent = DetallFormules.IdComponente where codiformula='" + vformula + "';")
  If rst.EOF Then GoTo fi
  vtitolmsg = atrim(rst!code) + " - " + atrim(rst!Description)
  vkg = cadbl(InputBox("Entra els Kg de formulació que vols calcular.", "Atenció"))
  nllauna = ""
  While Not rst.EOF
    If InStr(1, UCase(atrim(rst!DescComponente)), "BASE ") > 0 Then
            vcalcultanx100 = Redondejar((vkg * cadbl(rst![%decomponent])) / 100, 3)
            vmsg = vmsg + atrim(vcalcultanx100) + " Kg --- " + atrim(rst!DescComponente) + Chr(10)
            nlinies = nlinies + 1
    End If
    rst.MoveNext
  Wend
  If vmsg <> "" Then
    Load avis
    avis.caption = vtitolmsg
    avis.missatge = vmsg
    avis.missatge.Alignment = 0
    avis.missatge.Height = 250 * nlinies
    avis.missatge.FontBold = True
    avis.Height = 400 + ((180 * nlinies) * 4)
    avis.Show
    avis.Top = 1
    avis.Left = 1
    hinici = Now
    While IsFormLoaded(avis) And DateDiff("s", hinici, Now) < 90
       SetForegroundWindow avis.hwnd
       If avis.botoxinxeta.BackColor = &H8000000F Then
          avis.caption = vtitolmsg + "     (" + atrim(90 - DateDiff("s", hinici, Now)) + ")"
       End If
       wait 1
    Wend
    Unload avis
  End If
fi:
  Set rst = Nothing
End Sub

Private Sub bconsultarecarreges_Click()
   ensenyar_recarregues_pendents
End Sub
Sub ensenyar_recarregues_pendents()
   Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  
  sql = "SELECT Recarregarllaunes.numllauna, tintes.descripcio, Recarregarllaunes.data FROM Recarregarllaunes INNER JOIN tintes ON Recarregarllaunes.idtinta = tintes.idtinta order by data"
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
  formseleccio.Data1.RecordSource = sql
  formseleccio.width = 9000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(0).width = 1000
  formseleccio.DBGrid2.Columns(1).width = 4000
  formseleccio.DBGrid2.Columns(2).width = 2000
  formseleccio.Show 1
  If seleccioret = 1 Then
    If MsgBox("Vols eliminar aquesta recarrega pendent?", vbYesNo, "Atenció") = vbYes Then
       dbtintes.Execute "delete * from recarregarllaunes where numllauna='" + atrim(formseleccio.Data1.Recordset!numllauna) + "'"
    End If
  End If
  Unload formseleccio
End Sub
Private Sub colordelatinta_Click()
  ' colorllaunaiescullitsoniguals
End Sub

Private Sub Command1_Click()
   escriure_ini "Tintes", "pestara", ctara, "comandes.ini"
End Sub

Private Sub bcancel_Click()
  Form1.botoretorn.tag = "1"
  Unload formretorntintes
End Sub

Private Sub bok_Click()
   boto_retorn_llaunes
End Sub
Sub boto_retorn_llaunes(Optional vretornamagatzem As Boolean)
 Dim rst As Recordset
   Dim situacioactual As String
   If Not comprovar_pes_basculainferiora26kg Then Exit Sub
   If nllauna = "" Then MsgBox "No hi ha el numero de llauna", vbCritical, "Error": Exit Sub
   If nllauna.tag <> "1" Then
     vverificaciollauna = InputBox("Repeteix el numero de llauna per comfirmar", "Verificar Llauna")
     If UCase(vverificaciollauna) <> UCase(nllauna) Then MsgBox "Llaunes diferents", vbCritical, "Error": GoTo fi
  End If
   'If cadbl(pesnet) <= 1 Then If MsgBox("Aquesta llauna pesa molt poc " + atrim(cadbl(pesnet)) + "Kg), es correcte que vols fer un retorn?" + Chr(10) + "POTSER VOLIES FER UNA RECARREGA", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   ''If etpesteoric.tag = "igual" Then
  ''  If MsgBox("Els pes teoric i real son molt semblants, vols fer retorn igualments?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
  '' End If
   'If Not colorllaunaiescullitsoniguals Then Exit Sub
   'nllauna.ForeColor = QBColor(15)
    Set rst = dbtintes.OpenRecordset("select situacio from llaunes where numllauna='" + atrim(nllauna) + "'")
   If Not rst.EOF Then situacioactual = atrim(rst!situacio)
   'If UCase(InputBox("Entra un altra cop el Nº de llauna per assegurar que fas bé el retorn.", "Retorn de llaunes")) <> UCase(nllauna) Then
   '   MsgBox "Els numeros de llauna no coincideixen", vbCritical, "Error"
    '  nllauna.ForeColor = QBColor(0)
   '   Exit Sub
   'End If
   nllauna.ForeColor = QBColor(0)
   If existeix("c:\ordprog.ini") Then pesnet = atrim(cadbl(InputBox("Entra el pes de tinta que vols retornar:", "Nomes programació")))
   If cadbl(pesnet) < 1 Then
       ferelretorndetinta nllauna, cadbl(pesnet), True
       situacioactual = ""
         Else: ferelretorndetinta nllauna, cadbl(pesnet), True ' IIf(situacioactual = "IMP" False, False)
   End If
   If situacioactual = "IMP" Then
      canviardesituacio nllauna, situacioactual
   End If
   If cadbl(pesnet) < 1 Then dbtintes.Execute "update llaunes set activa=false where numllauna='" + atrim(nllauna) + "'"
   If vretornamagatzem Then
      dbtintes.Execute "update llaunes set aimpresores=False where numllauna='" + atrim(nllauna) + "'"
   End If
fi:
    nllauna = ""
    nllauna.tag = ""
    nllauna.SetFocus
End Sub
Sub canviardesituacio(numllauna As String, situacioactual As String)
    Load formsituacio
    formsituacio.llistadellaunes.AddItem numllauna
    formsituacio.Show 1
End Sub
Function colorllaunaiescullitsoniguals() As Boolean
Dim rstcolor As Recordset
colorllaunaiescullitsoniguals = False
  If atrim(nllauna) = "" Then Exit Function
   If colordelatinta.ListIndex < 0 Then MsgBox "Primer escull el color de la tinta de la llauna.": Exit Function
   Set rstcolor = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, familiescolors.codi FROM familiescolors RIGHT JOIN (tintes LEFT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta) ON familiescolors.codi = tintes.idfamcolor where llaunes.numllauna='" + atrim(nllauna) + "';")
   
   If Not rstcolor.EOF Then
         If rstcolor!codi <> colordelatinta.ItemData(colordelatinta.ListIndex) Then MsgBox "El color escullit no coincideix amb el de la llauna": Exit Function
       Else: MsgBox "No he trobat el color de la tinta relacionada amb aquesta llauna": Exit Function
   End If
   colorllaunaiescullitsoniguals = True
End Function

Sub emplenarcombocolortinta()
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("SELECT * FROM familiescolors order by descripcio;")
  colordelatinta.Clear
  While Not rst.EOF
    colordelatinta.AddItem atrim(rst!descripcio)
    colordelatinta.ItemData(colordelatinta.NewIndex) = rst!codi
    rst.MoveNext
  Wend
  Set rst = Nothing
End Sub
Function comprovar_pes_basculainferiora26kg() As Boolean

  If cadbl(pesnet) > 26 Then
     MsgBox "El pes de la llauna no pot ser superior a 26Kg", vbCritical, "Error"
     comprovar_pes_basculainferiora26kg = False
      Else: comprovar_pes_basculainferiora26kg = True
  End If
  If cadbl(pesnet) <= 0 Then MsgBox "EL PES DE LA BASCULA ES ZERO O INFERIOR AIXÍ NO ES POT FER UN RETORN.", vbCritical, "REVISA LA BASCULA": comprovar_pes_basculainferiora26kg = False
End Function
Private Sub Command2_Click()
  Dim rst As Recordset
  Dim vverificaciollauna As String
  If nllauna = "" Then MsgBox "No hi ha el numero de llauna", vbCritical, "Error": Exit Sub
  If Not comprovar_pes_basculainferiora26kg Then Exit Sub
  comprovar_llauna_entrada
  If comprovarsillaunahihaValhistoria(nllauna) Then Exit Sub
  Set rst = dbtintes.OpenRecordset("select numllauna from recarregarllaunes where numllauna='" + atrim(nllauna) + "'")
     If Not rst.EOF Then
        MsgBox "Aquesta llauna ja esta apunt per fer la recarrega, no cal tornar a fer-ho", vbCritical, "Error"
        Set rst = Nothing
        Exit Sub
     End If
  
'  If etpesteoric.tag = "igual" Then
'    If MsgBox("Els pes teoric i real son molt semblants, vols fer retorn igualment?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
'  End If
  ferelretorndetinta nllauna, cadbl(pesnet), True
  Set rst = dbtintes.OpenRecordset("select idtinta from llaunes where numllauna='" + atrim(nllauna) + "'")
  If rst.EOF Then MsgBox "Tinta no trovada.": Exit Sub
  dbtintes.Execute "insert into recarregarllaunes (numllauna,idtinta,data) values ('" + nllauna + "'," + atrim(rst!idtinta) + ",now)"
  Set rst = Nothing
  MsgBox "Llauna " + nllauna + " pendent de càrrega." + Chr(10) + "ARA FES LA TINTA AMB INKMAKER I ET DEMANARÉ QUE LA RELACIONIS.", vbInformation, "Recarrega de llauna"
  nllauna = ""
  nllauna.SetFocus
End Sub

Private Sub Command3_Click()
 cridar_imprimir_etiqueta nllauna
End Sub
Sub cridar_imprimir_etiqueta(nllauna As String)
 Dim vinici As Double
 Dim vfi As Double
 Dim rst As Recordset
 If atrim(nllauna) <> "" Then
        If InStr(1, nllauna, "-") > 0 Then imprimiriniciifidellauna nllauna: GoTo fi
        Set rst = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(nllauna) + "'")
        If Not rst.EOF Then
           'If rst!capacitatactual < 1 Then MsgBox "Aquesta llauna no te tinta no s'imprimirà cap etiqueta", vbCritical, "Error": GoTo fi
           imprimir_etiqueta nllauna
             Else: MsgBox "No he trobat aquest numero de llauna    " + nllauna, vbCritical, "Error"
        End If
   End If
fi:
   Set rst = Nothing
End Sub
Sub imprimiriniciifidellauna(nllauna As String)
  Dim rst As Recordset
  Dim vinici As Double
  Dim vfi As Double
  Dim vnumllauna As String
  Dim vcont As Byte
  Dim i As Double
  nllauna = atrim(UCase(nllauna))
  On Error Resume Next
  vinici = cadbl(Mid(nllauna, 2, InStr(1, nllauna, "-") - 2))
  vfi = cadbl(Mid(nllauna, InStr(3, nllauna, "A") + 1))
  On Error GoTo 0
  vcont = 0
  If vfi - vinici <= 10 And vinici < vfi Then
      For i = vinici To vfi
        vnumllauna = "A" + atrim(i)
        Set rst = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(vnumllauna) + "' and activa")
        If Not rst.EOF Then
           If rst!capacitatactual < 1 Then MsgBox "La llauna " + vnumllauna + " no te tinta no s'imprimirà cap etiqueta", vbCritical, "Error": GoTo cont
           'If MsgBox("Ara s'imprimira l'etiqueta " + vnumllauna, vbInformation + vbOKCancel, "Impresió de " + nllauna) = vbOK Then
              imprimir_etiqueta vnumllauna
            '  Else: GoTo fi
           'End If
           vcont = vcont + 1
           If vcont > 10 Then GoTo fi
        End If
cont:
             'Else: MsgBox "No he trobat aquest numero de llauna    " + vnumllauna, vbCritical, "Error"
      Next i
        Else: MsgBox "El diferencial entre inici i fi de llauna no pot ser mes de 10", vbCritical, "Error"
  End If
fi:
  Set rst = Nothing
End Sub

Private Sub Command4_Click()
   formsituacio.Show 1
End Sub

Private Sub Command5_Click()
   If comprovarsilallaunasutilitzaperunaaltracomanda Then
       If MsgBox("Vols continuar amb el retorn?", vbInformation + vbYesNo, "Atenció") = vbNo Then GoTo fi
   End If
   boto_retorn_llaunes True
fi:
End Sub
Function comprovarsilallaunasutilitzaperunaaltracomanda() As Boolean
   Dim rst As Recordset
   Dim rstll As Recordset
   Dim vnumc As Double
   Set rstll = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, tintes.codi FROM Llaunes INNER JOIN tintes ON Llaunes.idtinta = tintes.idtinta where numllauna='" + atrim(nllauna) + "'")
   If rstll.EOF Then MsgBox "No he trobat la llauna " + atrim(nllauna): Exit Function
   dbtintes.Execute "DELETE comandes.proximaseccio, assignaciollaunesacomandes.* FROM assignaciollaunesacomandes INNER JOIN comandes ON assignaciollaunesacomandes.comanda = comandes.comanda WHERE (((comandes.proximaseccio)<>'E' And (comandes.proximaseccio)<>'I'));"
   Set rst = dbtintes.OpenRecordset("select comanda,coditinta from assignaciollaunesacomandes where coditinta=" + atrim(rstll!codi))
   If Not rst.EOF Then
       vnumc = cadbl(rst!comanda)
       Set rst = dbtintes.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
       If rst.EOF Then
           dbtintes.Execute "delete * from assignaciollaunesacomandes where comanda=" + atrim(vnumc)
            Else
                MsgBox "Aquesta tinta està assignada a almenys una altra comanda. Comanda: " + atrim(vnumc)
                comprovarsilallaunasutilitzaperunaaltracomanda = True
       End If
   End If
   Set rst = Nothing
End Function


Private Sub ctara_Change()
  On Error Resume Next
  If Screen.ActiveControl.Name = "ctara" Then pesnet = cadbl(pesdelatinta) - cadbl(ctara)
End Sub

Private Sub Form_Activate()
  'SetForegroundWindow formretorntintes.hwnd
  On Error Resume Next
  nllauna.SetFocus
  
End Sub

Private Sub Form_Click()
'dbtintes.Execute "insert into recarregarllaunes (numllauna,idtinta,data) values ('A1234',2456,now)"
'dbtintes.Execute "delete * from recarregarllaunes where data<dateadd('n',-30,now)"

End Sub

Private Sub Form_DblClick()
   'Form1.etpesbascula = InputBox("Pes?", "Atenció")
   
End Sub

Private Sub Form_Load()
   Form1.botoretorn.tag = "1"
   emplenarcombocolortinta
   ctara = llegir_ini("Tintes", "pestara", "comandes.ini")
   If cadbl(ctara) = 0 Then ctara = atrim(1 + 1 / 3)
    ColocarEnTop formretorntintes, True
    ColocarEnTop formretorntintes, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Form1.botoretorn.tag = ""
   Form1.visible = False
End Sub

Sub comprovar_llauna_entrada()
Dim vpesteoric As Double
   Dim vformula As String
   etpesteoric = ""
   etpesteoric.tag = ""
   nllauna.BackColor = QBColor(15)
   calcularkgdisponiblesllauna nllauna, vpesteoric
   If vpesteoric > 0 Then etpesteoric = "Pes Teòric: " + atrim(vpesteoric) + " Kg"
   If cadbl(pesnet) < (vpesteoric + 0.3) And cadbl(pesnet) > (vpesteoric - 0.3) Then
       etpesteoric = etpesteoric + "   PESA IGUAL, ES CORRECTE?"
       etpesteoric.tag = "igual"
   End If
   bcomposicio.visible = False
   bcomposicio.tag = ""
   If Len(nllauna) > 4 Then
        If comprovarsillaunahihaValhistoria(nllauna) Then
            nllauna.BackColor = QBColor(12)
            MsgBox "La llauna " + atrim(nllauna) + " ja s'ha buidat i no es pot utilitzar", vbCritical, "Error"
        End If
        vformula = comprovarsillaunahihaformulaalhistoria(nllauna)
        If vformula <> "" Then
           bcomposicio.visible = True
           bcomposicio.tag = vformula
        End If
   End If

End Sub
Function comprovarsillaunahihaformulaalhistoria(vllauna As String) As String
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, historiallauna.tipusmoviment,historiallauna.formula FROM Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna WHERE (((Llaunes.numllauna)='" + treure_apostruf(nllauna) + "') AND ((historiallauna.tipusmoviment)='C'));")
   If Not rst.EOF Then
      comprovarsillaunahihaformulaalhistoria = atrim(rst!formula)
   End If
   Set rst = Nothing
End Function

Function comprovarsillaunahihaValhistoria(vllauna As String) As Boolean
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, historiallauna.tipusmoviment,historiallauna.formula FROM Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna WHERE (((Llaunes.numllauna)='" + treure_apostruf(nllauna) + "') AND ((historiallauna.tipusmoviment)='V'));")
   If Not rst.EOF Then
      comprovarsillaunahihaValhistoria = True
    
   End If
   Set rst = Nothing
End Function

Private Sub nllauna_GotFocus()
 nllauna.tag = ""
End Sub

Private Sub nllauna_KeyPress(KeyAscii As Integer)
  Static vhoraprimerapulsacio As Date
  Static vcont As Integer
  
  If nllauna = "" Then
    vhoraprimerapulsacio = Now
  End If
  'nllauna.tag = Str(KeyAscii)
  If Len(nllauna) > 4 Then
    If DateDiff("s", vhoraprimerapulsacio, Now) < 1 Then
       nllauna.tag = "1"
         Else: nllauna.tag = "": vhoraprimerapulsacio = "0:00:00"
    End If
  End If
  If KeyAscii = 13 Then comprovar_llauna_entrada
End Sub

Private Sub pesnet_Change()
  On Error Resume Next
  If Screen.ActiveControl.Name = "pesnet" Then ctara = cadbl(pesdelatinta) - cadbl(pesnet)
End Sub

Private Sub Timer1_Timer()
  ' Label11.caption = nllauna.tag
   FlashWindow formretorntintes.hwnd, 1
   FlashWindow Form1.hwnd, 1
   Form1.possarpesbascula
   pesdelatinta = Form1.etpesbascula
   pesnet = cadbl(pesdelatinta) - cadbl(ctara)
   If cadbl(pesdelatinta) < 1 And nllauna = "" And Not existeix("c:\ordprog.ini") Then
      Form1.tag = ""
      Unload formretorntintes
      Exit Sub
   End If
   'Timer1.tag = cadbl(Timer1.tag) + 1
   'If cadbl(Timer1.tag) > 29 Then
   '   If Not IsFormLoaded(formseleccio) Then Unload formretorntintes
   '     Else: formretorntintes.caption = "Gestió retorn i cargues de tintes (" + atrim(30 - cadbl(Timer1.tag)) + ")"
   'End If
End Sub

Private Sub Timer2_Timer()
  
  
End Sub
