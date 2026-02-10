VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formcontrolclixesentrats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control clixes entrats"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19335
   Icon            =   "formcontrolclixesentrats.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   19335
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton copcionsvisualitzacio 
      Caption         =   "Tots"
      Height          =   225
      Index           =   2
      Left            =   225
      TabIndex        =   8
      Top             =   1560
      Width           =   780
   End
   Begin VB.OptionButton copcionsvisualitzacio 
      Caption         =   "Veure només els pendents REVISAR"
      Height          =   225
      Index           =   1
      Left            =   4065
      TabIndex        =   7
      Top             =   1560
      Width           =   3645
   End
   Begin VB.OptionButton copcionsvisualitzacio 
      Caption         =   "Veure només els pendents MARCAR"
      Height          =   225
      Index           =   0
      Left            =   1020
      TabIndex        =   6
      Top             =   1560
      Value           =   -1  'True
      Width           =   3165
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Revisar Tintes clixes nous o modificats"
      Height          =   465
      Left            =   195
      TabIndex        =   5
      Top             =   75
      Width           =   3645
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificacions"
      Height          =   705
      Left            =   4230
      Picture         =   "formcontrolclixesentrats.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Veure PDF"
      Top             =   705
      Width           =   1980
   End
   Begin VB.CommandButton Command3 
      Height          =   705
      Left            =   2205
      Picture         =   "formcontrolclixesentrats.frx":0BB6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Veure PDF"
      Top             =   690
      Width           =   1980
   End
   Begin VB.CommandButton Command2 
      Height          =   705
      Left            =   180
      Picture         =   "formcontrolclixesentrats.frx":0EC0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Veure IMP"
      Top             =   690
      Width           =   1980
   End
   Begin VB.CommandButton Command1 
      Height          =   705
      Left            =   16995
      Picture         =   "formcontrolclixesentrats.frx":178A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Clixé revisat"
      Top             =   810
      Width           =   1980
   End
   Begin VB.Data dataclixesentrats 
      Caption         =   "dataclixesentrats"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   6810
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   105
      Visible         =   0   'False
      Width           =   2730
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "formcontrolclixesentrats.frx":1A60
      Height          =   6075
      Left            =   120
      OleObjectBlob   =   "formcontrolclixesentrats.frx":1A7C
      TabIndex        =   0
      Top             =   1860
      Width           =   18945
   End
End
Attribute VB_Name = "formcontrolclixesentrats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vordre As String
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long



Private Sub Command1_Click()
   If IsNull(dataclixesentrats.Recordset!datafet) Then
      If MsgBox("Vols marcar aquesta linia com a FET?", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbYes Then
        possarobservacions
        dataclixesentrats.Recordset.Edit
        dataclixesentrats.Recordset!datafet = Now
        dataclixesentrats.Recordset.Update
      End If
   End If
End Sub

Private Sub Command2_Click()
  Dim rstc As Recordset
  Set rstc = dbtmpb.OpenRecordset("select id_treball,ordremodificacio,direnvio,codiclient from clientsvinculats where id_treball=" + atrim(dataclixesentrats.Recordset!numtreball) + " and ordremodificacio=" + atrim(dataclixesentrats.Recordset!versio))
  If Not rstc.EOF Then
     obrir_imp_treball cadbl(rstc!id_treball), cadbl(rstc!ordremodificacio), cadbl(rstc!codiclient), cadbl(rstc!direnvio)
      Else: MsgBox "No he trobat cap IMP amb aquest treball, no puc ensenyar el IMP.", vbCritical, "ERROR"
  End If
End Sub
Sub obrir_imp_treball(treball As Double, modificacio As Double, codiclient As Double, direnvio As Double)
   Dim generarfitxer_imp As String
   If modificacio = 0 Then modificacio = 1
   generarfitxer_imp = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\IMP" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "-" + Format(codiclient, "000000") + "_" + atrim(direnvio) + ".doc"
   If existeix(generarfitxer_imp) Then
     obrir_document generarfitxer_imp
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + generarfitxer_imp, vbCritical, "Error"
  End If
End Sub

Private Sub Command3_Click()
   
   obrir_pdf_treball cadbl(dataclixesentrats.Recordset!numtreball), cadbl(dataclixesentrats.Recordset!versio)
End Sub

Private Sub Command4_Click()
  Dim vfitxer As String
  If dataclixesentrats.Recordset.EOF Then Exit Sub
  vfitxer = rutamodifispdftreball(dataclixesentrats.Recordset!numtreball, dataclixesentrats.Recordset!versio)
  If existeix(vfitxer) Then
     obrir_document (vfitxer)
       Else: MsgBox "No trobo el fitxer de modificacions per aquest treball.", vbCritical, "Error"
  End If
End Sub
Function rutamodifispdftreball(vidtreball As Double, vordre As Double) As String
   'On Error Resume Next
   'MkDir ruta_documentacio_clixes + "\" + Format(vidtreball, "00000")
   rutamodifispdftreball = ruta_documentacio_clixes + "\" + Format(vidtreball, "00000") + "\MODIFI" + Format(vidtreball, "00000") + "-" + Format(vordre, "000") + ".pdf"
   'If existeix(rutamodifispdftreball) Then Kill rutamodifispdftreball
End Function

Private Sub Command5_Click()
  verificaciotintestreballsnousomodificats
End Sub
Sub verificaciotintestreballsnousomodificats()
  Load formseleccionou
  formseleccionou.caption = "Treballs per revisar"
  formseleccionou.Data1.DatabaseName = rutadelfitxer(cami) + "CLIXESNOUS.MDB"
  formseleccionou.Data1.RecordSource = "SELECT Modificacions.id_treball, Modificacions.ordre, [marca] & ' - ' & [linia] AS Marcailinia, Modificacions.estatrevisiotintes FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball WHERE (((InStr(1,[estatrevisiotintes],'DISSENY'))>0) AND ((InStr(1,[estatrevisiotintes],'+IMP'))=0));"
  formseleccionou.refrescar
  formseleccionou.DBGrid2.Columns(0).width = 1000
  formseleccionou.DBGrid2.Columns(1).width = 500
  formseleccionou.DBGrid2.Columns(2).width = 5500
  formseleccionou.DBGrid2.Columns(3).width = 3500
  formseleccionou.width = 12000
  formseleccionou.Show 1
  If seleccioret = 1 Then
   ShellAndWait "\\serverprodu\dades\progcomandes\aplicacio\clixesnous.exe " + "comandes.ini ''  modificartintes " + atrim(formseleccionou.Data1.Recordset!id_treball) + " " + atrim(formseleccionou.Data1.Recordset!ordre) + " +IMP", vbNormalFocus
'   wait 2
'   AppActivate "Manteniment de les Tintes    " + atrim(formseleccionou.Data1.Recordset!id_treball) + "/" + atrim(formseleccionou.Data1.Recordset!ordre)

  End If
  
  Unload formseleccionou

End Sub

Private Sub copcionsvisualitzacio_Click(Index As Integer)
   carregar_controlclixes
   
End Sub

Private Sub Form_Load()
  vordre = "marcailinia"
  dataclixesentrats.DatabaseName = rutadelfitxer(cami) + "baixes.mdb"
  carregar_controlclixes
End Sub
Sub carregar_controlclixes()
   Dim vsql As String
   actualitzarcampsbuits
   If copcionsvisualitzacio(0).Value = True Then vsql = "datafet=null and datarepas<>null "
   If copcionsvisualitzacio(2).Value = True Then vsql = "1=1"
   If copcionsvisualitzacio(1).Value = True Then vsql = "datarepas=null"
   vsql = vsql + IIf(vordre <> "", " order by " + vordre, "")
   dataclixesentrats.RecordSource = "select * from clixesentrats_control where " + vsql
   dataclixesentrats.Refresh
End Sub

   
Sub actualitzarcampsbuits()
   dbtmpb.Execute "UPDATE Clixes INNER JOIN clixesentrats_control ON Clixes.id_treball = clixesentrats_control.numtreball SET clixesentrats_control.marcailinia = [clixes].[marca] & ' - ' & [clixes].[linia], clixesentrats_control.modificada = IIf([clixesentrats_control].[versio]=1,'N','M') where (clixesentrats_control.marcailinia=null or clixesentrats_control.modificada=null) or  (clixesentrats_control.marcailinia='' or clixesentrats_control.modificada='')"
End Sub
Private Sub reixa_DblClick()
   If reixa.Columns(reixa.col).DataField = "observacions" Then
      possarobservacions
      Exit Sub
   End If
    If reixa.Columns(reixa.col).DataField = "numtaula" Then
      possarnumerodetaula
      Exit Sub
   End If
   
   If Not dataclixesentrats.Recordset.EOF Then
     If IsNull(dataclixesentrats.Recordset!datarepas) Then
        obrir_repas_clixes cadbl(dataclixesentrats.Recordset!numtreball), cadbl(dataclixesentrats.Recordset!versio)
     End If
   End If
End Sub
Sub obrir_repas_clixes(vidtreball As Double, vordre As Double)
   escriure_ini "clixes", "repasidtreball", atrim(vidtreball), "comandes.ini"
   escriure_ini "clixes", "repasordre", atrim(vordre), "comandes.ini"
   If FindWindow(vbNullString, "Repàs de clixes") = 0 Then
      Shell llegir_ini("General", "rutallistats", "comandes.ini") + "repas de clixes.exe", vbNormalFocus
       Else: AppActivate "Repàs de clixes"
   End If
End Sub

Sub possarnumerodetaula()
     Dim v As String
     Dim veliminar As Boolean
      v = InputBox("Escriu la TAULA que vols. T-" + vbNewLine + "ESCRIU [CAP] per deixar-la sense taula.", "Canvi de taula")
      If StrPtr(v) = 0 Then Exit Sub
      If UCase(v) = "CAP" Then veliminar = True
      If Not veliminar And (cadbl(v) < 1 Or cadbl(v) > 4) Then MsgBox "Aquest número de taula no es vàlid.", vbCritical, "ERROR": GoTo fi
      dataclixesentrats.Recordset.Edit
      dataclixesentrats.Recordset!numtaula = IIf(veliminar = False, "T-" + atrim(v), "")
      dataclixesentrats.Recordset.Update
      dataclixesentrats.Recordset.Move 0
fi:
End Sub
Sub possarobservacions()
     Dim v As String
      v = InputBox("Escriu la observació que vols.", "Canvi d'observació", atrim(dataclixesentrats.Recordset!observacions))
      If StrPtr(v) = 0 Then Exit Sub
      dataclixesentrats.Recordset.Edit
      dataclixesentrats.Recordset!observacions = v
      dataclixesentrats.Recordset.Update
      dataclixesentrats.Recordset.Move 0

End Sub

Private Sub reixa_HeadClick(ByVal ColIndex As Integer)
   If vordre = reixa.Columns(ColIndex).DataField Then
         vordre = vordre + " DESC"
       Else: vordre = reixa.Columns(ColIndex).DataField
   End If
   carregar_controlclixes
End Sub

Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If dataclixesentrats.Recordset.EOF Then Command4.visible = False: Exit Sub
  If IsNull(dataclixesentrats.Recordset!datarepas) Then
      Command1.Enabled = False
        Else: Command1.Enabled = True
  End If
  If dataclixesentrats.Recordset!versio > 1 Then
      Command4.visible = True
        Else: Command4.visible = False
  End If
End Sub


