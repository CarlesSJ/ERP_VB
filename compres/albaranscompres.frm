VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form albaranscompres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manteniment d'albarans de compres."
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10980
   Icon            =   "albaranscompres.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   8520
      Picture         =   "albaranscompres.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Taspàs definitiu a SAP"
      Top             =   45
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Height          =   390
      Left            =   9510
      Picture         =   "albaranscompres.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Comptabilitzar a Bip"
      Top             =   30
      Width           =   525
   End
   Begin VB.CommandButton sortir 
      Height          =   390
      Left            =   10485
      Picture         =   "albaranscompres.frx":13DE
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Sortir"
      Top             =   30
      Width           =   390
   End
   Begin VB.CommandButton imprimiralbaraproveidor 
      Height          =   390
      Left            =   10065
      Picture         =   "albaranscompres.frx":1968
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Imprimir el nostra albarà de  proveïdor."
      Top             =   30
      Width           =   390
   End
   Begin VB.Data albarans 
      Caption         =   "Albarans de Compres"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   3090
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "albaransbip"
      Top             =   45
      Width           =   3060
   End
   Begin VB.Frame Frame1 
      Caption         =   "Manteniment d'albarans de compres"
      Height          =   4605
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   10830
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Linies d'albarà"
         Height          =   2430
         Left            =   135
         TabIndex        =   2
         Top             =   1905
         Width           =   10620
         Begin VB.CommandButton Command6 
            Height          =   255
            Left            =   4950
            Picture         =   "albaranscompres.frx":1EF2
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Buscar l'Albarà."
            Top             =   285
            Width           =   270
         End
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   10260
            Picture         =   "albaranscompres.frx":247C
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Moure linia a un altra albarà"
            Top             =   780
            Width           =   300
         End
         Begin VB.CommandButton Command4 
            Height          =   330
            Left            =   10260
            Picture         =   "albaranscompres.frx":2A06
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Eliminar linia seleccionada."
            Top             =   420
            Width           =   300
         End
         Begin VB.Data liniesalbara 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   1440
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   -45
            Visible         =   0   'False
            Width           =   1410
         End
         Begin MSDBGrid.DBGrid reixa 
            Bindings        =   "albaranscompres.frx":2F90
            Height          =   1980
            Left            =   45
            OleObjectBlob   =   "albaranscompres.frx":2FA7
            TabIndex        =   3
            Top             =   270
            Width           =   10170
         End
         Begin VB.Label etpostit 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Height          =   195
            Left            =   6540
            TabIndex        =   29
            Top             =   45
            Visible         =   0   'False
            Width           =   45
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00D29F7D&
         Caption         =   "Capçalera d'albarà"
         Height          =   1440
         Left            =   135
         TabIndex        =   1
         Top             =   300
         Width           =   10575
         Begin VB.CommandButton Command3 
            Height          =   300
            Left            =   1725
            Picture         =   "albaranscompres.frx":439C
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Buscar l'Albarà."
            Top             =   900
            Width           =   300
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   1500
            Picture         =   "albaranscompres.frx":4926
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "Buscar l'Albarà."
            Top             =   390
            Width           =   300
         End
         Begin VB.CommandButton consultar 
            Height          =   300
            Left            =   4950
            Picture         =   "albaranscompres.frx":4EB0
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Buscar l'Albarà."
            Top             =   900
            Width           =   300
         End
         Begin VB.TextBox Text5 
            DataField       =   "nalbprov"
            DataSource      =   "albarans"
            Height          =   285
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "Dos clics per canviar-lo"
            Top             =   900
            Width           =   1155
         End
         Begin VB.TextBox Text4 
            DataField       =   "ncodiprov"
            DataSource      =   "albarans"
            Height          =   285
            Left            =   2175
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   900
            Width           =   1155
         End
         Begin VB.TextBox Text3 
            DataField       =   "ndata"
            DataSource      =   "albarans"
            Height          =   285
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   7
            ToolTipText     =   "Dos clics per canviar la data"
            Top             =   390
            Width           =   1155
         End
         Begin VB.TextBox Text2 
            DataField       =   "nnumcomanda"
            DataSource      =   "albarans"
            Height          =   285
            Left            =   540
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   900
            Width           =   1155
         End
         Begin VB.TextBox Text1 
            DataField       =   "nempresa"
            DataSource      =   "albarans"
            Height          =   285
            Left            =   2205
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   390
            Width           =   1155
         End
         Begin VB.TextBox albara 
            DataField       =   "numalbara"
            DataSource      =   "albarans"
            Height          =   285
            Left            =   585
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   390
            Width           =   885
         End
         Begin VB.Label nomproveidor 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   2175
            TabIndex        =   24
            Top             =   1200
            Width           =   3600
         End
         Begin VB.Label comptabilitzat 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Comptabilitzat (Passat a SAP)"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   540
            Left            =   5655
            TabIndex        =   18
            Top             =   180
            Width           =   1995
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Albarà Proveïdor"
            Height          =   180
            Left            =   3630
            TabIndex        =   15
            Top             =   690
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Codi Proveidor"
            Height          =   180
            Left            =   2205
            TabIndex        =   14
            Top             =   690
            Width           =   1290
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Num. Comanda"
            Height          =   180
            Left            =   555
            TabIndex        =   13
            Top             =   690
            Width           =   1260
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
            Height          =   225
            Left            =   4230
            TabIndex        =   12
            Top             =   195
            Width           =   600
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa"
            Height          =   180
            Left            =   2445
            TabIndex        =   11
            Top             =   195
            Width           =   960
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Albarà"
            Height          =   180
            Left            =   690
            TabIndex        =   10
            Top             =   195
            Width           =   960
         End
      End
   End
   Begin VB.Label msgnoenviatsasap 
      BackStyle       =   0  'Transparent
      Caption         =   "No enviats a SAP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   75
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Menu mllistats 
      Caption         =   "Llistats"
      Begin VB.Menu mllistatcompresentredates 
         Caption         =   "Llistat d'albarans entre dates"
      End
   End
   Begin VB.Menu malbarantspendentsdenviarasap 
      Caption         =   "Albarans pendents d'enviar a SAP"
   End
   Begin VB.Menu gCSVenvasos 
      Caption         =   "Generar CSV Impost envasos"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "albaranscompres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function nomprov(codi As Double) As String
  Dim rst As Recordset
  Set rst = dbtmpb.OpenRecordset("select nom from proveidors_comercial where codicomptable='" + atrim(codi) + "'")
  If Not rst.EOF Then nomprov = rst!nom
End Function

Private Sub albarans_Reposition()
   If Not albarans.Recordset.EOF Then
       liniesalbara.RecordSource = "select * from albaransbip where numalbara=" + atrim(cadbl(albarans.Recordset!numalbara))
       nomproveidor = nomprov(albarans.Recordset!ncodiprov)
      Else:
        liniesalbara.RecordSource = "select * from albaransbip where numalbara=-1"
        nomproveidor = ""
   End If
   liniesalbara.Refresh
   If Not albarans.Recordset.EOF Then
      If albarans.Recordset!menviat Then
         comptabilitzat.Visible = True
        Else: comptabilitzat.Visible = False
      End If
   End If
   
End Sub

Private Sub Command1_Click()
  If MsgBox("Segur que vols enviar aquest albarà a comptabilitat?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
      If DateDiff("d", CVDate("15/01/" + atrim(Year(Now))), CVDate(Now)) < 0 Then
        If Year(albarans.Recordset!ndata) <> Year(Now) Then
             If UCase(InputBox("Aquest albarà no es de l'any encurs... per evitar errors no es permet pujar-lo al SAP." + vbNewLine + "SI VOLS CONTINUAR IGUALMENT ESCRIU [CONTINUAR]?", "Error D'ANY")) <> "CONTINUAR" Then Exit Sub
        End If
      End If
      If albarans.Recordset!menviat Then MsgBox "OJU que ja està enviat aquest albarà si ja l'havien importat l'han d'actualitzar a comptabilitat."
     ' comprespalets.generar_fitxer_bip Text5
      assignardecimalipunt
      If CDbl("1,234") > 2 Then MsgBox "El simbol decimal està malament, tanca tots els programes i torna a intentar-ho", vbCritical, "Error"
      
      comprespalets.generar_fitxer_sap Text5, Text4
      MsgBox "Albarà enviat correctament."
  End If
End Sub

Private Sub Command2_Click()
   Dim numalb As Double
  albarans.RecordSource = "select distinct numalbara,first(enviat) as menviat,first(nomfitxer) as mnomfitxer,first(numcomanda) as nnumc,first(empresa) as nempresa,first(numcomanda) as nnumcomanda,first(data) as ndata,first(codiproveidorcomercial) as ncodiprov, first(numalbaraprov) as nalbprov from albaransbip GROUP BY albaransbip.numalbara order by numalbara DESC"
  albarans.Refresh
  numalb = cadbl(InputBox("Entra el Nº d'albarà NOSTRA que vols buscar.", "Atenció"))
  albarans.Recordset.FindFirst "numalbara=" + atrim(numalb)
  If albarans.Recordset.NoMatch Then MsgBox "No he trobat aquest Nº d'albarà."
End Sub

Private Sub Command3_Click()
Dim numcomanda As Double
  numcomanda = cadbl(InputBox("Entra el Nº de comanda de compra que vols buscar.", "Atenció"))
  If numcomanda = 0 Then
     albarans.RecordSource = "select distinct numalbara,first(enviat) as menviat,first(nomfitxer) as mnomfitxer,first(numcomanda) as nnumc,first(empresa) as nempresa,first(numcomanda) as nnumcomanda,first(data) as ndata,first(codiproveidorcomercial) as ncodiprov, first(numalbaraprov) as nalbprov from albaransbip GROUP BY albaransbip.numalbara order by numalbara DESC"
     albarans.Refresh
       Else
         albarans.RecordSource = "select distinct numalbara,first(enviat) as menviat,first(nomfitxer) as mnomfitxer,first(numcomanda) as nnumc,first(empresa) as nempresa,first(numcomanda) as nnumcomanda,first(data) as ndata,first(codiproveidorcomercial) as ncodiprov, first(numalbaraprov) as nalbprov from albaransbip where numcomanda=" + atrim(numcomanda) + " GROUP BY albaransbip.numalbara order by numalbara DESC"
         albarans.Refresh
  End If
  
  If albarans.Recordset.EOF Then
     MsgBox "No he trobat aquest Nº de comanda."
     albarans.RecordSource = "select distinct numalbara,first(enviat) as menviat,first(nomfitxer) as mnomfitxer,first(numcomanda) as nnumc,first(empresa) as nempresa,first(numcomanda) as nnumcomanda,first(data) as ndata,first(codiproveidorcomercial) as ncodiprov, first(numalbaraprov) as nalbprov from albaransbip GROUP BY albaransbip.numalbara order by numalbara DESC"
     albarans.Refresh
  End If

End Sub

Private Sub Command4_Click()
   If liniesalbara.Recordset.EOF Then Exit Sub
   If InputBox("Segur que vols eliminar aquesta linia d'albarà?" + Chr(10) + "Escriu [eliminar] per eliminar-la", "Atenció") = "eliminar" Then
       liniesalbara.Recordset.Delete
       liniesalbara.Refresh
   End If
End Sub

Private Sub Command5_Click()
   canviaralbaradelalinia
End Sub
Sub canviaralbaradelalinia()
   Dim v As String
   v = InputBox("Entra el número el nou nº d'albarà:" + Chr(10) + "Si es albarà nou escriu [NOU]" + Chr(10) + "ATENCIÓ UTILITZA AQUESTA FUNCIÓ AMB COMPTE.", "Canvi d'albarà")
   If atrim(v) = "" Then Exit Sub
   If cadbl(v) > noualbara Then MsgBox "El nº que has entrat es mes gran que l'ultim numero registrat, si vols crear un de nou has d'escriure NOU a la pregunta anterior.", vbCritical, "Error": Exit Sub
   If UCase(v) = "NOU" Then v = atrim(noualbara + 1)
   liniesalbara.Recordset.Edit
   liniesalbara.Recordset!numalbara = cadbl(v)
   liniesalbara.Recordset.Update
   albarans.Refresh
   albarans.Recordset.FindFirst "numalbara=" + atrim(v)
End Sub
Function noualbara() As Double
   Dim rst As Recordset
   noualbara = 1
   Set rst = dbcompres.OpenRecordset("select max(numalbara) as gran from albaransbip")
   If Not rst.EOF Then noualbara = rst!gran
   Set rst = Nothing
End Function

Private Sub Command6_Click()
  Dim numalb As String
  Dim rst As Recordset
  numalb = InputBox("Entra el Nº de LOT de PROVEIDOR que vols buscar." + Chr(10) + "Pots utilitzar * per buscar semblants.", "Atenció")
  Set rst = albarans.Database.OpenRecordset("select numalbaraprov from albaransbip where numlotproveidor like '" + atrim(numalb) + "'")
  If rst.EOF Then MsgBox "No he trobat aquest Lot a cap albarà.", vbCritical, "Error": GoTo fi
  rst.MoveLast
  If rst.RecordCount = 1 Then numalb = rst!numalbaraprov: GoTo nomes1
  albarans.RecordSource = "select distinct numalbara,first(enviat) as menviat,first(nomfitxer) as mnomfitxer,first(numcomanda) as nnumc,first(empresa) as nempresa,first(numcomanda) as nnumcomanda,first(data) as ndata,first(codiproveidorcomercial) as ncodiprov, first(numalbaraprov) as nalbprov from albaransbip GROUP BY albaransbip.numalbara order by numalbara DESC"
  albarans.Refresh
  numalb = escullir_albara_lot("select distinct numalbaraprov from albaransbip where numlotproveidor like '" + atrim(numalb) + "'")
nomes1:
  albarans.Recordset.FindFirst "nalbprov='" + atrim(numalb) + "'"
  If albarans.Recordset.NoMatch Then MsgBox "No he trobat l'albarà de proveidor " + atrim(numalb)
fi:
  Set rst = Nothing
End Sub

Function simboldecimal() As String
   vsimboldecimal = "."
   vsimbolmiler = ","
   If InStr(1, Trim(CDbl(1 / 2)), ",") Then vsimbolmiler = ".": vsimboldecimal = ","
   simboldecimal = vsimboldecimal
End Function

Private Sub Command7_Click()
 Dim vhihainplacsa As Boolean
  Dim vhihaplasel As Boolean
 ' If simboldecimal = "," Then MsgBox "El simbol decimal es la coma no es poden fer importacions fins que sigui el punt.", vbCritical, "Error"
  
  comprovarsihihafitxerspendentsdimportarasap vhihainplacsa, vhihaplasel, True, True
  If vhihaplasel Then
    MsgBox "Hi ha albarans de Plasel", vbInformation, "Atenció"
    ShellAndWait "\\servidorsap\seidor_COMUNICADOR\PROGRAMA\Compres\Plasel\SEI_Importacions.exe"
    mirar_resultat_importacio "\\servidorsap\seidor_COMUNICADOR\LOGcOMPRES\Plasel"
  End If
  If vhihainplacsa Then
    MsgBox "Hi ha albarans de Inplacsa", vbInformation, "Atenció"
    ShellAndWait "\\servidorsap\seidor_COMUNICADOR\PROGRAMA\Compres\Inplacsa\SEI_Importacions.exe"
    mirar_resultat_importacio "\\servidorsap\seidor_COMUNICADOR\LOGCOMPRES\Inplacsa"
  End If
End Sub
Sub mirar_resultat_importacio(vdir As String)
  Dim v As String
  Dim vlinia As String
  Dim vmesactual As String
  v = Dir(vdir + "\Log_" + format(Now, "yyyymmdd") + "*.txt")
  While v <> ""
    If cadbl(Mid(v, Len(v) - 9, 6)) > vgran Then
      vgran = cadbl(Mid(v, Len(v) - 9, 6))
      vmesactual = v
    End If
    'v = Dir(vdir + "\Log_" + Format(Now, "yyyymmdd") + "*.txt")
    v = Dir
  Wend
  If vmesactual <> "" Then
    Open vdir + "\" + vmesactual For Input As #1
    While Not EOF(1)
       Input #1, vlinia
       If InStr(1, UCase(vlinia), "ERROR") > 0 Then obrir_document vdir + "\" + vmesactual:  GoTo fi
    Wend
fi:
    Close #1
  End If

End Sub

Private Sub consultar_Click()
  Dim numalb As String
  If consultar.BackColor = QBColor(12) Then
     consultar.BackColor = Command2.BackColor:
     albarans.Recordset.Filter = ""
     albarans.RecordSource = "select distinct numalbara,first(enviat) as menviat,first(nomfitxer) as mnomfitxer,first(numcomanda) as nnumc,first(empresa) as nempresa,first(numcomanda) as nnumcomanda,first(data) as ndata,first(codiproveidorcomercial) as ncodiprov, first(numalbaraprov) as nalbprov from albaransbip GROUP BY albaransbip.numalbara order by numalbara DESC"
     albarans.Refresh
     Exit Sub
  End If
  albarans.RecordSource = "select distinct numalbara,first(enviat) as menviat,first(nomfitxer) as mnomfitxer,first(numcomanda) as nnumc,first(empresa) as nempresa,first(numcomanda) as nnumcomanda,first(data) as ndata,first(codiproveidorcomercial) as ncodiprov, first(numalbaraprov) as nalbprov from albaransbip GROUP BY albaransbip.numalbara order by numalbara DESC"
  albarans.Refresh
  numalb = InputBox("Entra el Nº d'albarà de PROVEIDOR que vols buscar." + Chr(10) + "Pots utilitzar * per buscar semblants.", "Atenció")
  albarans.RecordSource = "select distinct numalbara,first(enviat) as menviat,first(nomfitxer) as mnomfitxer,first(numcomanda) as nnumc,first(empresa) as nempresa,first(numcomanda) as nnumcomanda,first(data) as ndata,first(codiproveidorcomercial) as ncodiprov, first(numalbaraprov) as nalbprov from albaransbip where numalbaraprov like '" + atrim(numalb) + "' GROUP BY albaransbip.numalbara order by numalbara DESC"
  'albarans.Recordset.Filter = "nalbprov like '" + atrim(numalb) + "'"
  albarans.Refresh
  consultar.BackColor = QBColor(12)
  
  
  'albarans.Recordset.FindFirst "nalbprov ='" + atrim(numalb) + "'"
  'If albarans.Recordset.NoMatch Then MsgBox "No he trobat aquest Nº d'albarà de proveidor."
End Sub

Private Sub Form_Click()
'  comprespalets.comprovarsijaexisteix 881, "\\servidorsap\SGI_COMUNICADOR\ENTALBCOMPRAS\INPLACSA\A-Articles.csv"
End Sub

Private Sub Form_Load()
   Dim vhihainplacsa As Boolean
   Dim vhihaplasel As Boolean
   assignardecimalipunt
   albarans.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
   liniesalbara.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
   albarans.RecordSource = "select distinct numalbara,first(enviat) as menviat,first(nomfitxer) as mnomfitxer,first(numcomanda) as nnumc,first(empresa) as nempresa,first(numcomanda) as nnumcomanda,first(data) as ndata,first(codiproveidorcomercial) as ncodiprov, first(numalbaraprov) as nalbprov from albaransbip GROUP BY albaransbip.numalbara order by numalbara DESC"
   albarans.Refresh
'   comprovarsihihafitxerspendentsdimportarasap vhihainplacsa, vhihaplasel, , True
   
End Sub
Sub comprovarsihihafitxerspendentsdimportarasap(Optional vhihainplacsa As Boolean, Optional vhihaplasel As Boolean, Optional vnoavisar As Boolean, Optional vseidor As Boolean)
   Dim vruta As String
   Dim vcontador As Integer
   Dim vdir As String
   Dim vmsg As String
   On Error GoTo errorsap
   'vruta = llegir_ini("Vendes", "rutasap_INPLACSA", "comandes.ini")
   If llegir_ini("Compres", "rutaSapSeidor_INPLACSA", "comandes.ini") = "{[}]" Then
     escriure_ini "Compres", "rutaSapSeidor_INPLACSA", "\\servidorsap\seidor_COMUNICADOR\ENTALBCOMPRAS\INPLACSA", "comandes.ini"
     escriure_ini "Compres", "rutaSapSeidor_PLASEL", "\\servidorsap\seidor_COMUNICADOR\ENTALBCOMPRAS\PLASEL", "comandes.ini"
   End If
   If vseidor Then vruta = llegir_ini("Compres", "rutaSapSeidor_INPLACSA", "comandes.ini")
   mirarsihihafitxers vruta + "\C-*.csv", vcontador
   If vcontador > 0 Then vhihainplacsa = True: vmsg = "  Hi ha fitxers de INPLACSA pendents d'importar a SAP."
   'vruta = llegir_ini("Vendes", "rutasap_PLASEL", "comandes.ini")
   If vseidor Then vruta = llegir_ini("Compres", "rutaSapSeidor_PLASEL", "comandes.ini")
   mirarsihihafitxers vruta + "\C-*.csv", vcontador
   If vcontador > 0 Then vhihaplasel = True: vmsg = vmsg + Chr(10) + "  Hi ha fitxers de PLASEL pendents d'importar a SAP."
   If vmsg <> "" And Not vnoavisar Then MsgBox vmsg, vbCritical, "Atenció"
   
   Exit Sub
errorsap:
     MsgBox "L'usuari " + Environ("username") + " no te acces al servidor de SAP no es podran enviar albarans."
     'Command2.Enabled = False
End Sub
Sub mirarsihihafitxers(vruta As String, vcontador As Integer)
   vdir = Dir(vruta)
   vcontador = 0
   While vdir <> ""
     If vdir <> "." And vdir <> ".." Then vcontador = vcontador + 1
     vdir = Dir
   Wend
End Sub

Private Sub Image1_Click()

End Sub

Private Sub gCSVenvasos_Click()
   Dim vdatainici As String
   Dim vdatafi As String
   Dim rst As Recordset
   Dim rstprov As Recordset
   Dim vlinia As String
   Dim vkgtotal As String
   Dim vkgNOreciclats As String
   Dim vTIPUSNIF_NIF_RAOSOCIAL As String
   
   vdatainici = InputBox("Entra la data d'inici de l'exportació.", "Inici")
   If StrPtr(vdatainici) = 0 Then Exit Sub
   vdatafi = InputBox("Entra la data de fi de l'exportació.", "Fi")
   If StrPtr(vdatafi) = 0 Then Exit Sub
   Set rst = dbtmp.OpenRecordset("select * from albaransbip where kgimpostenvasos>0 and data>=#" + format(vdatainici, "mm/dd/yy") + "# and data<=#" + format(vdatafi, "mm/dd/yy") + "#")
   Set rstprov = dbtmp.OpenRecordset("SELECT proveidors.tipusproveidorIMPOST, proveidors_comercial.codicomptable FROM proveidors_comercial LEFT JOIN proveidors ON proveidors_comercial.codi = proveidors.codi")
   Open "c:\temp\CSV_Impost_Envasos.csv" For Output As 1
   Print #1, "Número Asiento;Fecha Hecho Contabilizado;Concepto;Clave Producto;Descripción Producto;Régimen Fiscal;Justificante;Prov./Dest.: Tipo Documento;Prov./Dest.: Nº documento;Prov./Dest.: Razón social;Kilogramos;Kilogramos No Reciclados;Observaciones"
   While Not rst.EOF
     rstprov.FindFirst "codicomptable=" + atrim(cadbl(rst!codiproveidorcomercial))
     If Not rstprov.NoMatch Then
         If rstprov!tipusproveidorIMPOST = "INTRA" Then
           vkgtotal = format(rst!quantitat, "###0.000")
           vkgNOreciclats = format(rst!kgimpostenvasos, "###0.000")
             'miro si el decimal m'ha posat el punt i el canvio per , EL CSV HA DE SER ,
           vkgtotal = substituir(vkgtotal, ".", ",")
           vkgNOreciclats = substituir(vkgNOreciclats, ".", ",")
           vTIPUSNIF_NIF_RAOSOCIAL = infoproveidorseparatperpuntsicomes(rstprov!codicomptable)
           
           vlinia = atrim(rst!id)  'ASSENTAMENT
           vlinia = vlinia + ";" + atrim(rst!data) 'DATA RECEPCIO
           vlinia = vlinia + ";1" 'concepto -> (1 Adquisición intracomunitaria)
           vlinia = vlinia + ";B" 'clave producto (En miralles ha dit que sempre seria la B
           vlinia = vlinia + ";PRODUCTE ENVAS PLASTIC" 'Descripció   (pels adquirientes no es NECESSARI)
           vlinia = vlinia + ";A" 'REGIM FISCAL
           vlinia = vlinia + ";"  'JUSTIFICANTE
           vlinia = vlinia + ";" + vTIPUSNIF_NIF_RAOSOCIAL 'tipusNIF NIF RAOSOCIAL ORIGEN PROVEIDOR
           vlinia = vlinima + ";" + vkgtotal 'KG TOTALS
           vlinia = vlinia + ";" + vkgNOreciclats 'KG no reciclats
           
         End If
     End If
     If vlinia <> "" Then Print #1, vlinia
     rst.MoveNext
   Wend
   Close 1
   Set rst = Nothing
End Sub
Function infoproveidorseparatperpuntsicomes(vcodicomptable)

End Function

Private Sub imprimiralbaraproveidor_Click()
   comprespalets.impresiodalbara Text5
End Sub

Private Sub malbarantspendentsdenviarasap_Click()
   Dim rst As Recordset
   albarans.RecordSource = "select distinct numalbara,first(enviat) as menviat,first(nomfitxer) as mnomfitxer,first(numcomanda) as nnumc,first(empresa) as nempresa,first(numcomanda) as nnumcomanda,first(data) as ndata,first(codiproveidorcomercial) as ncodiprov, first(numalbaraprov) as nalbprov from albaransbip where not enviat GROUP BY albaransbip.numalbara order by numalbara DESC"
   albarans.Refresh
   msgnoenviatsasap.Visible = True
End Sub

Private Sub mllistatcompresentredates_Click()
'llistat_albaranscompresentredates
 llistatalbaranscompresentredates
End Sub

Function escullir_albara_lot(vsql As String) As Double
      Dim vtipusmaterial As String
      Load formseleccio
      formseleccio.Caption = "Escull l'albarà "
      formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
      formseleccio.Data1.RecordSource = vsql
      formseleccio.refrescar
      formseleccio.DBGrid2.Columns(0).Width = 1900
      'formseleccio.DBGrid2.Columns(1).Width = 5000
      formseleccio.Width = 5200
      'formseleccio.DBGrid2.Columns(0).Width = 1000
      'formseleccio.DBGrid2.Columns(1).Width = 3000
      formseleccio.Command2.Tag = "0"
      formseleccio.Caption = "Escullir albarà"
      formseleccio.Show 1
      If seleccioret = 1 Then
           escullir_albara_lot = cadbl(formseleccio.Data1.Recordset!numalbaraprov)
      End If
      Unload formseleccio
     
End Function
Sub llistatalbaranscompresentredates()
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim datainici As String
  Dim datafi As String
  Dim vcodimaterial As Double
  datainici = InputBox("Entra la data d'inici del llistat", "Inici")
  If Not IsDate(datainici) Then MsgBox "Data no valida", vbCritical, "Error": Exit Sub
  datafi = InputBox("Entra la data fi del llistat", "Inici")
  If Not IsDate(datafi) Then MsgBox "Data no valida", vbCritical, "Error": Exit Sub
  vcodimaterial = cadbl(escullirmaterial)
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + "llistat_albaranscompresentredates.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "compres.mdb"
  oreport.FormulaFields.GetItemByName("dates").Text = "'Inici: " + format(datainici, "dd/mm/yy") + " -> Fi: " + format(datafi, "dd/mm/yy") + "'"
  oreport.RecordSelectionFormula = "{albaransbip.data}>=#" + format(datainici, "mm/dd/yy") + "# and {albaransbip.data}<=#" + format(datafi, "mm/dd/yy") + "#" + IIf(vcodimaterial > 0, " and {albaransbip.article}='" + atrim(vcodimaterial) + "'", "")
  oreport.DiscardSavedData
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport

   veurereport.CRViewer.ViewReport
   veurereport.Show 1
End Sub

Private Sub reixa_AfterColUpdate(ByVal ColIndex As Integer)
   activar_desactivar_impostpalet cadbl(reixa.Columns("Kg_ImpEnv")), cadbl(reixa.Columns("Palet Creat"))
End Sub
Sub activar_desactivar_impostpalet(vimpost As Double, vpalet As Double)
   If vpalet > 0 Then dbtmp.Execute "update palets set teimpost=" + IIf(vimpost > 0, "True", "False") + " where idpalet=" + atrim(vpalet)
End Sub

Private Sub reixa_Change()
  reixa.Columns("Import") = cadbl(reixa.Columns("preu")) * cadbl(reixa.Columns("quantitat"))
  
End Sub

Private Sub reixa_DblClick()
   Dim vnumpalet As Double
   Dim vBaseImpEnv As Double
   Dim vImpEnv As Double
   Dim v As String
   
   If reixa.Columns(reixa.col).Caption = "Palet Creat" Then
     
     'tancar i buscar el palet
     vnumpalet = cadbl(reixa.Text)
     Form1.palets.RecordSource = "select * from palets"
     Form1.palets.Refresh
     Form1.palets.Recordset.FindFirst "idpalet=" + atrim(vnumpalet)
     albaranscompres.Hide
     
   End If
   If UCase(reixa.Columns(reixa.col).DataField) = "KGIMPOSTENVASOS" Then
        v = InputBox("Escriu el valor de Kg de BASE IMPOSABLE de l'impost d'envasos." + vbNewLine + "Escriu [treure] per treure l'impost d'aquesta compra.", "Base imposable impost")
        If UCase(v) = "TREURE" Then GoTo guardarcanvi
        vBaseImpEnv = cadbl(v)
        If vBaseImpEnv = 0 Then Exit Sub
        vImpEnv = cadbl(InputBox("Escriu el valor de Kg on s'aplica l'impost d'envasos.", "Kg impost", vBaseImpEnv))
        If vImpEnv = 0 Then Exit Sub
guardarcanvi:
        If liniesalbara.Recordset.EditMode = 0 Then liniesalbara.Recordset.Edit
        liniesalbara.Recordset!kgbaseimposableimpostenvasos = vBaseImpEnv
        liniesalbara.Recordset!kgimpostenvasos = vImpEnv
        liniesalbara.Recordset.Update
        activar_desactivar_impostpalet cadbl(reixa.Columns("Kg_ImpEnv")), cadbl(reixa.Columns("Palet Creat"))
   End If
   
   
End Sub

Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If UCase(reixa.Columns(reixa.col).DataField) = "KGIMPOSTENVASOS" Then
         etpostit.Visible = True: etpostit.Caption = "Doble Clic per canviar l'IMPOST ENVASOS"
       Else: etpostit.Visible = False
   End If
End Sub

Private Sub sortir_Click()
  Unload albaranscompres
End Sub

Private Sub Text3_DblClick()
   Dim numalb As String
   Dim albaraactual As Double
   numalb = InputBox("Entra la data d'albarà de proveidor.", "Canvi de data", Text3)
   albaraactual = cadbl(albara)
   If IsDate(numalb) Then
       albarans.Database.Execute "update albaransbip set data=#" + format(numalb, "mm/dd/yy") + "# where numalbara=" + atrim(albara)
       albarans.Refresh
       albarans.Recordset.FindFirst "numalbara=" + atrim(albaraactual)
   End If
End Sub


Function ajuntaralbaransduplicats(numalbprov As String, numalb As String) As Boolean
   Dim rst As Recordset
   Dim nomfitxer As String
   Set rst = albarans.Database.OpenRecordset("select * from albaransbip where numalbaraprov='" + numalbprov + "' and numalbara<>" + atrim(cadbl(numalb)))
   ajuntaralbaransduplicats = True
   If Not rst.EOF Then
       If MsgBox("Hi ha l'albarà Nº:" + atrim(rst!numalbara) + " que també te l'albarà de proveidor Nº: " + atrim(numalbprov) + Chr(10) + "VOLS AJUNTAR EL " + atrim(numalb) + " amb el " + atrim(rst!numalbara) + "?", vbExclamation + vbYesNo, "Albarans repetits") = vbYes Then
       
       'elimino l'albarà antic de comptabilitat
           nomfitxer = llegir_ini("Compres", "rutabip", "comandes.ini") + "\empre" + format(cadbl(llegir_ini("Compres", "numempresabip_" + atrim(rst!empresa), "comandes.ini")), "000") + "\" + albarans.Recordset!mnomfitxer
           If existeix(nomfitxer) Then
            Kill (nomfitxer)
           End If
       
           ajuntaralbaransduplicats = True
           albarans.Database.Execute "update albaransbip set numalbara=" + atrim(rst!numalbara) + " where numalbara=" + atrim(cadbl(numalb))
           albarans.Recordset.FindFirst "numalbara=" + atrim(rst!numalbara)
           
          Else: ajuntaralbaransduplicats = False
       End If
   End If
End Function

Sub triar_proveidor(vcodicomptable As String, vnomproveidor As String)
  Load formseleccio
  formseleccio.sortirs.Tag = "filtre"
  'formseleccio.Data1.DatabaseName = cami
  Set formseleccio.Data1.Recordset = dbcomandes.OpenRecordset("SELECT proveidors.codi, proveidors.nom, proveidors_comercial.codicomptable, proveidors_comercial.nom as Nom_Empresafacturadora FROM proveidors LEFT JOIN proveidors_comercial ON proveidors.codi = proveidors_comercial.codiproduccio")
  'formseleccio.Data1.RecordSource = "select * from proveidors"
  formseleccio.refrescar
  formseleccio.Width = 12000
  formseleccio.Show 1
  If seleccioret = 1 Then
   vcodicomptable = atrim(cadbl(formseleccio.Data1.Recordset!codicomptable))
   vnomproveidor = atrim(formseleccio.Data1.Recordset!Nom_Empresafacturadora)
   Unload formseleccio
  End If
  Unload formseleccio
End Sub

Private Sub Text4_DblClick()
 Dim numalb As String
   Dim albaraactual As Double
   Dim vnomproveidor As String
   Dim vcodicomptable As String
    albaraactual = cadbl(albara)
    triar_proveidor vcodicomptable, vnomproveidor
    If vcodicomptable <> "" Then
     albarans.Database.Execute "update albaransbip set codiproveidorcomercial='" + atrim(vcodicomptable) + "' where numalbara=" + atrim(albara)
     albarans.Database.Execute "update albaransbip set nomproveidorcomercial='" + atrim(vnomproveidor) + "' where numalbara=" + atrim(albara)
     albarans.Refresh
     albarans.Recordset.FindFirst "numalbara=" + atrim(albaraactual)
    End If
End Sub

Private Sub Text5_DblClick()
   Dim numalb As String
   Dim albaraactual As Double
   Dim rsta As Recordset
   Dim nomfitxer As String
   numalb = InputBox("Entra el numero d'albarà de proveidor.", "Canvi d'albarà", Text5)
   albaraactual = cadbl(albara)
   If Len(numalb) > 0 Then
       ajuntaralbaransduplicats numalb, albara
       albarans.Database.Execute "update albaransbip set numalbaraprov='" + numalb + "' where numalbara=" + atrim(albara)
       albaraactual = albarans.Recordset!numalbara
       
       'elimino  l'albarà enviat abip per colocarhi el nou
       Set rsta = albarans.Database.OpenRecordset("select *from albaransbip where numalbara=" + atrim(albaraactual))
       If Not rsta.EOF Then
         nomfitxer = llegir_ini("Compres", "rutabip", "comandes.ini") + "\empre" + format(cadbl(llegir_ini("Compres", "numempresabip_" + atrim(rsta!empresa), "comandes.ini")), "000") + "\" + rsta!nomfitxer
         If existeix(nomfitxer) Then
            Kill (nomfitxer)
            comprespalets.generar_fitxer_bip numalb
         End If
       End If
       nomfitxer = generarfitxeralbaratxt(rsta)
       albarans.Database.Execute "update albaransbip set nomfitxer='" + nomfitxer + "' where numalbara=" + atrim(albara)
       albarans.Refresh
       albarans.Recordset.FindFirst "numalbara=" + atrim(albaraactual)
   End If
End Sub
