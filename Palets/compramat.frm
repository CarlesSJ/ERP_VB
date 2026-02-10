VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form compramat 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Compra de material"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   Icon            =   "compramat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8445
      Picture         =   "compramat.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Comprar"
      Top             =   6435
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      Caption         =   "Afegir Comanda/Client"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8430
      TabIndex        =   19
      Top             =   5160
      Width           =   1725
      Begin VB.CommandButton Command3 
         Caption         =   "Client"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   645
         Width           =   1500
      End
      Begin VB.CommandButton afegircomanda 
         Caption         =   "Comanda"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Comprar"
      Top             =   5175
      Width           =   1170
   End
   Begin VB.TextBox convkilos 
      BackColor       =   &H00C0C0FF&
      Height          =   300
      Left            =   1350
      TabIndex        =   16
      Top             =   1305
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009AA6FA&
      Height          =   1005
      Left            =   330
      TabIndex        =   12
      Top             =   -60
      Width           =   6945
      Begin VB.ComboBox descmat 
         BackColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   780
         TabIndex        =   13
         Top             =   345
         Width           =   5895
      End
      Begin VB.Label infomaterial 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   690
         Width           =   6705
      End
      Begin VB.Label codimat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   60
         TabIndex        =   15
         Top             =   345
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Material per comprar"
         Height          =   255
         Left            =   945
         TabIndex        =   14
         Top             =   135
         Width           =   3030
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5415
      Picture         =   "compramat.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Comprar"
      Top             =   1050
      Width           =   1170
   End
   Begin VB.Data percomandaoclient 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1545
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4965
      Visible         =   0   'False
      Width           =   1500
   End
   Begin MSDBGrid.DBGrid reixaxcomocli 
      Bindings        =   "compramat.frx":109E
      Height          =   2730
      Left            =   120
      OleObjectBlob   =   "compramat.frx":10BA
      TabIndex        =   10
      Top             =   5145
      Width           =   8265
   End
   Begin VB.TextBox amplecompra 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   3195
      TabIndex        =   8
      Top             =   1020
      Width           =   855
   End
   Begin VB.TextBox numdecompra 
      Height          =   360
      Left            =   7575
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton toteslescompres 
      Caption         =   "Ensenyar Compres"
      Height          =   555
      Left            =   9195
      TabIndex        =   5
      Top             =   90
      Width           =   870
   End
   Begin VB.CommandButton botocomprar 
      Caption         =   "Compar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4215
      Picture         =   "compramat.frx":215B
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Comprar"
      Top             =   1050
      Width           =   1170
   End
   Begin VB.TextBox quantitatcomprar 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   1365
      TabIndex        =   1
      Top             =   1005
      Width           =   1185
   End
   Begin VB.Data compres 
      Caption         =   "compres"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3735
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "compresmaterial"
      Top             =   1050
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSDBGrid.DBGrid reixacompres 
      Bindings        =   "compramat.frx":26E5
      Height          =   3405
      Left            =   90
      OleObjectBlob   =   "compramat.frx":26F7
      TabIndex        =   0
      Top             =   1680
      Width           =   10110
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6900
      TabIndex        =   17
      Top             =   1185
      Width           =   1725
   End
   Begin VB.Label etamplecompra 
      BackStyle       =   0  'Transparent
      Caption         =   "Ample:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2580
      TabIndex        =   9
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Compra:"
      Height          =   240
      Left            =   7830
      TabIndex        =   7
      Top             =   15
      Width           =   1155
   End
   Begin VB.Label numreserva 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mtrs Comprar:                       Conv. Kilos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   105
      TabIndex        =   2
      Top             =   1005
      Width           =   1290
   End
End
Attribute VB_Name = "compramat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub escullir_material()
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   'formseleccio.Data1.RecordSource = "select codi,descripcio,refproducte,proveidor,* from materials where codi>499" + IIf(descmat.Tag <> "", " and " + descmat.Tag, "")
   formseleccio.Data1.RecordSource = "SELECT materials.codi as [Codi], materials.descripcio as [Descripcio], materials.refproducte as [Ref_Producte], proveidors.nom as [Proveidor] FROM materials LEFT JOIN proveidors ON materials.proveidor = proveidors.codi WHERE (((materials.codi)>499)" + IIf(descmat.Tag <> "", " And " + descmat.Tag, "") + ")"
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns(0).Width = 500
   formseleccio.DBGrid2.Columns(1).Width = 2500
   formseleccio.DBGrid2.Columns(2).Width = 1000
   formseleccio.DBGrid2.Columns(3).Width = 1500
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.Width = formseleccio.Width + (formseleccio.Width / 3)
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           descmat = formseleccio.DBGrid2.Columns("descripcio") + " | " + formseleccio.DBGrid2.Columns("ref_producte")
           codimat = formseleccio.DBGrid2.Columns("codi")
           infomaterial = "Proveidor: " + formseleccio.DBGrid2.Columns("Proveidor")
        End If
   End If
   If seleccioret = 9 Then
      descmat = ""
      codimat = ""
      infomaterial = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub

Private Sub afegircomanda_Click()
 Dim comanda As String
 Dim metres As String
 Dim resp As Byte
 comanda = InputBox("Entra la comanda que vols afegir.", "Atenció")
 resp = comparasielmaterialcorrespont(comanda)
 If resp = 0 Then MsgBox "Aquesta comanda no existeix": Exit Sub
 If resp = 2 Then MsgBox "El material d'aquesta comanda no es el mateix que'l comprat": Exit Sub
 If resp = 3 Then MsgBox "Les carecteristiques del material de la comanda no es el mateix que'l comprat": Exit Sub
 If resp = 4 Then MsgBox "El material entrat en aquesta comanda no te el codi>500 i no es pot comparar amb la compra feta": Exit Sub
 If comanda <> "" And resp = 1 Then
    metres = InputBox("Entra els metres per aquesta pre-reserva.", "Atenció")
    afegirpercomandaoclientxreserva cadbl(comanda), "", "", cadbl(metres)
 End If
End Sub
Sub actualitzar_material_comanda(numcom As String)
    Dim rstinfo As Recordset
    Set rstinfo = dbtmpb.OpenRecordset("select * from comandes where comanda=" + numcom)
    MsgBox "Aquesta comanda te el material asignat inferior al codi 500. S'ha de canviar, ESCULL UN MATERIAL NOU"
    assignarmat.demanar_nou_material numcom, cadbl(rstinfo!materialex), cadbl(rstinfo!colorex), cadbl(rstinfo!aditiuex)
End Sub
Function comparasielmaterialcorrespont(comanda As String) As Byte
   Dim rstcom As Recordset
   Dim rstreseva As Recordset
   Dim rstmaterial As Recordset
   Dim resp As Byte
   Dim micres As Double
   Set rstcom = dbtmpb.OpenRecordset("select * from comandes where comanda=" + comanda)
   If Not rstcom.EOF Then
      If rstcom!materialex < 500 Then
         actualitzar_material_comanda comanda
      End If
      Set rstcom = dbtmpb.OpenRecordset("select * from comandes where comanda=" + comanda)
      If Not rstcom.EOF Then
        If rstcom!materialex < 500 Then comparasielmaterialcorrespont = 4: Exit Function
      End If
   End If
   resp = 1
   If Not rstcom.EOF Then
      Set rstreserva = dbtmp.OpenRecordset("select * from reserves where idreserva=" + atrim(numreserva))
      Set rstmaterial = dbtmp.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstcom!materialex)))
      If Not rstreserva.EOF Then
          resp = 2
          If cadbl(rstreserva!familia) = cadbl(rstmaterial!familia) Then
             If cadbl(rstreserva!subfamilia) = cadbl(rstmaterial!subfamilia) Then
               If cadbl(rstreserva!familiacol) = cadbl(rstmaterial!familiacol) Then
                 If cadbl(rstreserva!subfamiliacol) = cadbl(rstmaterial!subfamiliacol) Then
                   If cadbl(rstreserva!familiaad) = cadbl(rstmaterial!familiaad) Then
                     If cadbl(rstreserva!subfamiliaad) = cadbl(rstmaterial!subfamiliaad) Then
                          resp = 1
                     End If
                   End If
                 End If
               End If
             End If
          End If
          micres = assignarmat.micresmaterial(cadbl(rstcom!mesuraesp), rstcom!espessor, atrim(rstcom!tubolam))
          If resp = 1 Then
             resp = 3
             If cadbl(rstreserva!ample) >= (cadbl(rstcom!ampleesq) - 1) Then
                 If cadbl(rstreserva!plegat) = cadbl(rstcom!plegatesq) Then
                   If cadbl(rstreserva!solapa) = cadbl(rstcom!solapa) Then
                     If cadbl(rstreserva!espesor) = micres Then
                       'If assignarmat.aatrim(rstreserva!carestractat) = assignarmat.aatrim(rstcom!tractatex) Then
                         If assignarmat.aatrim(rstreserva!obert) = assignarmat.aatrim(rstcom!oberturaex) Then
                           If cabool(rstreserva!microperforat) = cabool(rstcom!micropex) Then
                             If atrim(rstreserva!semielaborat) = atrim(rstcom!tubolam) Then
                               resp = 1
                             End If
                           End If
                         End If
                       'End If
                     End If
                   End If
                 End If
              End If
          End If
      End If
     Else: resp = 0
   End If
   comparasielmaterialcorrespont = resp
End Function
Private Sub amplecompra_LostFocus()
calcular_kilos
End Sub

Private Sub botocomprar_Click()
  Dim rstcomocli As Recordset
  Dim idcompra As Long
  Dim numreservainici As String
  numreservainici = numreserva
 'If cadbl(numreserva) = 0 Then
 '  If MsgBox("Si fas aquesta compra es reservaran " + assignarmat.mtrsnecessaris + " mtrs ", vbInformation + vbYesNo, "Atenció") = vbNo Then Exit Sub
 'End If
 If assignarmat.codiclient <> "" And cadbl(assignarmat.comanda) = 0 Then
    If MsgBox("Aquesta compra s'assignarà al client " + Chr(10) + Chr(13) + assignarmat.codiclient + " - " + assignarmat.nomclient, vbInformation + vbYesNo, "Comprar material") = vbNo Then Exit Sub
   Else: If MsgBox("Aquesta compra s'assignarà a la comanda " + assignarmat.comanda, vbInformation + vbYesNo, "Comprar material") = vbNo Then Exit Sub
 End If
 If cadbl(quantitatcomprar) = 0 Or descmat = "" Then MsgBox "Falten camps per emprlenar": Exit Sub
 If cadbl(assignarmat.iespesor) = 0 Then MsgBox "L'espesor sel.leccionat no es correcte.": Exit Sub
  mirarampleareserva
  compres.Recordset.AddNew
    compres.Recordset!codimat = cadbl(codimat)
    compres.Recordset!descmat = atrim(Mid(descmat, 1, 30))
    compres.Recordset!idreserva = cadbl(numreserva)
    compres.Recordset!data = Now
    compres.Recordset!metres = cadbl(quantitatcomprar)
    compres.Recordset!kgpendents = cadbl(convkilos)
    compres.Recordset!kilos = cadbl(convkilos)
    compres.Recordset!micres = cadbl(assignarmat.iespesor)
    compres.Recordset!ample = amplecompra
    idcompra = compres.Recordset!idcomandacompra
  compres.Recordset.Update
  compres.Recordset.Bookmark = compres.Recordset.LastModified
  percomandaoclient.RecordSource = "select * from percomandaoclient where idcompra=" + atrim(cadbl(compres.Recordset!idcomandacompra))
  percomandaoclient.Refresh
  percomandaoclient.Recordset.AddNew
  percomandaoclient.Recordset!idreserva = numreserva
  percomandaoclient.Recordset!metres = cadbl(assignarmat.mtrsnecessaris) - cadbl(quantitatcomprar.Tag)
   If assignarmat.codiclient <> "" And cadbl(assignarmat.comanda) = 0 Then
       percomandaoclient.Recordset!numclient = assignarmat.codiclient
       percomandaoclient.Recordset!nomclient = assignarmat.nomclient
     Else: percomandaoclient.Recordset!numcomanda = cadbl(assignarmat.comanda)
   End If
  percomandaoclient.Recordset!idcompra = idcompra
  percomandaoclient.Recordset.Update
  
  percomandaoclient.Refresh
  percomandaoclient.Recordset.Bookmark = percomandaoclient.Recordset.LastModified
  'If cadbl(percomandaoclient.Recordset!numcomanda) > 0 Then dbtmp.Execute "update  reserves set metresreservats=metresreservats+" + atrim(cadbl(assignarmat.mtrsnecessaris)) + " where idreserva=" + atrim(numreserva)
  If cadbl(numreservainici) = 0 Then assignarmat.filtrar_materials
End Sub
Sub mirarampleareserva(Optional ampleres As Double)
 Dim i As Integer
  Dim trobat As Boolean
  Dim rample As Double
  rample = cadbl(amplecompra)
  If ampleres > 0 Then rample = ampleres
  If assignarmat.reixa.Cols > 2 And rample > 0 Then
      i = 1
      While i < assignarmat.reixa.Rows And Not trobat
         If cadbl(assignarmat.reixa.TextMatrix(i, assignarmat.columnadelcamp("ample"))) = rample And cadbl(assignarmat.reixa.TextMatrix(i, assignarmat.columnadelcamp("estareservat"))) <> 1 Then
            trobat = True
           Else: i = i + 1
         End If
      Wend
      If trobat Then
        If cadbl(assignarmat.reixa.TextMatrix(i, assignarmat.columnadelcamp("idreserva"))) <> 0 Then
          numreserva = assignarmat.reixa.TextMatrix(i, assignarmat.columnadelcamp("idreserva"))
        End If
      End If
      If cadbl(numreserva) = 0 Then MsgBox "Aquest ample amb aquest material no existeix, crearé una reserva nova.": numreserva = atrim(assignarmat.crear_novareserva(rample))
'      assignarmat.filtrar_materials
      If trobat Then assignarmat.reixa.row = i
  End If
End Sub

Private Sub Command1_Click()
  Dim rstcom As Recordset
  Dim msg As String
  If compres.Recordset.EOF Then MsgBox "No hi ha cap compra sel.leccionada": Exit Sub
  Set rstcom = dbtmp.OpenRecordset("select * from percomandaoclient where idcompra=" + atrim(cadbl(compres.Recordset!idcomandacompra)))
  While Not rstcom.EOF
     If rstcom!numcomanda <> 0 Then
        msg = msg + " Comanda: " & atrim(rstcom!numcomanda) + Chr(10) + Chr(13)
       Else: If cadbl(rstcom!numclient) <> 0 Then msg = msg + " Client: " & atrim(rstcom!numcomanda) + " - " + atrim(rstcom!nomclient) + Chr(10) + Chr(13)
     End If
     rstcom.MoveNext
  Wend
  If InputBox("Eliminar afectarà als següents moviments:" + Chr(10) + Chr(13) + Chr(10) + Chr(13) + msg + Chr(10) + Chr(13) + "ESCRIU [ELIMINAR] PER ELIMINAR LA COMPRA", "Eliminar compra") <> "ELIMINAR" Then Set rstcom = Nothing: Exit Sub
  Set rstcom = Nothing
  dbtmp.Execute "delete * from percomandaoclient where idcompra=" + atrim(cadbl(compres.Recordset!idcomandacompra))
  compres.Recordset.Delete
  compres.Refresh
  comprovar_reserves_negatives
End Sub
Public Sub comprovar_reserves_negatives()
  Dim subconsulta As String
  subconsulta = "select reserves.idreserva FROM Reserves LEFT JOIN percomandaoclient ON Reserves.idreserva = percomandaoclient.idreserva WHERE (((Reserves.metresreservats)<=0) and  ((percomandaoclient.numcomanda)=0 Or (percomandaoclient.numcomanda) Is Null) AND ((percomandaoclient.numclient)=0 Or (percomandaoclient.numclient) Is Null) AND ((percomandaoclient.idcompra)=0 Or (percomandaoclient.idcompra) Is Null));"
  dbtmp.Execute ("delete * from reserves where idreserva in (" + subconsulta + ")")
End Sub

Private Sub Command3_Click()
  Dim nomclient As String
  Dim codiclient As String
  Dim metres As String
  Load formseleccio
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select * from clients order by nom"
  formseleccio.DBGrid2.AllowDelete = False
  formseleccio.refrescar
  formseleccio.Width = formseleccio.Width + (formseleccio.Width / 3)
  formseleccio.Show 1
  If seleccioret = 9 Then
     codiclient = ""
     nomclient = ""
  End If
  If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           nomclient = formseleccio.DBGrid2.Columns("nom")
           codiclient = formseleccio.DBGrid2.Columns("codi")
        End If
  End If
  formseleccio.Data1.RecordSource = ""
  formseleccio.Data1.Refresh
  Unload formseleccio
  If codiclient <> "" Then
     metres = InputBox("Entra els metres per aquesta pre-reserva.", "Atenció")
     If cadbl(metres) > 0 Then afegirpercomandaoclientxreserva 0, codiclient, nomclient, cadbl(metres)
  End If
End Sub
Sub afegirpercomandaoclientxreserva(comanda As Double, codiclient As String, nomclient As String, metres As Double)
  percomandaoclient.Recordset.AddNew
  percomandaoclient.Recordset!idreserva = numreserva
  percomandaoclient.Recordset!metres = metres
   If codiclient <> "" And comanda = 0 Then
       percomandaoclient.Recordset!numclient = codiclient
       percomandaoclient.Recordset!nomclient = nomclient
     Else: percomandaoclient.Recordset!numcomanda = comanda
   End If
  percomandaoclient.Recordset!idcompra = cadbl(compres.Recordset!idcomandacompra)
  percomandaoclient.Recordset.Update
  percomandaoclient.Refresh
End Sub

Private Sub Command4_Click()
   If Not percomandaoclient.Recordset.EOF Then
       If InputBox("Escriu [ELIMINAR] per eliminar aquesta Pre-Reserva.", "Atenció") = "ELIMINAR" Then
           percomandaoclient.Recordset.Delete
           percomandaoclient.Refresh
       End If
      Else: MsgBox "Primer sel.lecciona una pre-reserva."
   End If
End Sub

Private Sub compres_Reposition()
 If cadbl(compres.Recordset!idcomandacompra) > 0 Then
    percomandaoclient.RecordSource = "select * from percomandaoclient where idcompra=" + atrim(cadbl(compres.Recordset!idcomandacompra))
     Else: percomandaoclient.RecordSource = ""
  End If
 percomandaoclient.Refresh
End Sub

Sub calcular_kilos()
  convkilos = format(conversiokilos(cadbl(codimat), cadbl(amplecompra), cadbl(quantitatcomprar), cadbl(assignarmat.iespesor), assignarmat.itl, cadbl(assignarmat.isolapa)), "#,##0.00")
  
End Sub

Private Sub convkilos_LostFocus()
   If cadbl(assignarmat.iespesor) = 0 Then MsgBox "L'espesor no pot estar a zero": Exit Sub
   If codimat <> "" Then quantitatcomprar = format(conversiokilos(cadbl(codimat), cadbl(amplecompra), cadbl(convkilos) * -1, cadbl(assignarmat.iespesor), assignarmat.itl, cadbl(assignarmat.isolapa)), "#,##0")
End Sub

Private Sub descmat_DropDown()
 
 escullir_material
 SendKeys "{TAB}"
 calcular_kilos
End Sub
Function conversiokilos(codimat As Long, amplemat As Double, metres As Double, espesor As Double, semielaborat As String, solapa As Double) As Double
  Dim kilos As Double
  Dim ample As Double
  Dim rstmaterial As Recordset
  kilos = 0
  'If metres < 0 Then Stop
  If espesor < 0 Then
      kilos = (espesor * -1) / 1000: ample = amplemat / 100: GoTo jatincgrmsm2
  End If
  Set rstmaterial = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(codimat)), dbOpenSnapshot, dbReadOnly)
  If Not rstmaterial.EOF Then
   ample = amplemat
   kilos = demetresakilos(ample, cadbl(rstmaterial!grmcm3), espesor, semielaborat, solapa)
jatincgrmsm2:
   'si els metres son negatius son kilos
   If metres > 0 Then
      kilos = kilos * ample * metres
     Else:
       'multiplico per -1 per treure el negatiu dels metres, k son kilos
       If (kilos * ample) > 0 Then kilos = (metres * -1) / (kilos * ample)
       'kilos = (((metres * -1) * 1000) / cadbl(rstmaterial!grmcm3)) / ample
   End If
  End If
  conversiokilos = kilos
  Set rstmaterial = Nothing
End Function
Private Sub Form_Load()
  percomandaoclient.DatabaseName = camistock
End Sub

Private Sub numdecompra_LostFocus()
  If numdecompra = "0" Then numdecompra = "": Exit Sub
  If cadbl(numdecompra) < 10000 Then numdecompra = format(Now, "yy") + "0000000" + format(numdecompra, "0000")
End Sub

Private Sub quantitatcomprar_LostFocus()
   If cadbl(assignarmat.iespesor) = 0 Then MsgBox "L'espesor no pot estar a zero": Exit Sub
   If codimat <> "" Then calcular_kilos
End Sub

Private Sub reixacompres_Click()
  numreserva = compres.Recordset!idreserva
End Sub

Private Sub reixacompres_DblClick()
  Dim r As String
  If reixacompres.Columns(reixacompres.col).DataField = "numcompra" Then
      r = InputBox("Entra el numero de compra.", "Atenció", format(Now, "yy") + "000000")
      If r <> "" Then
         reixacompres.Columns("numcompra") = r
         If compres.Recordset.EditMode = 0 Then compres.Recordset.Edit
         compres.Recordset.Update
      End If
  End If
  If reixacompres.Columns(reixacompres.col).DataField = "entregada" Then
       reixacompres.Columns("entregada") = Not cabool(reixacompres.Columns("entregada"))
       If compres.Recordset.EditMode = 0 Then compres.Recordset.Edit
       compres.Recordset.Update
  End If
  If reixacompres.Columns(reixacompres.col).DataField = "metres" Then
      r = InputBox("Entra els metres de compra.", "Atenció")
      If r <> "" Then
         reixacompres.Columns("metres") = r
         reixacompres.Columns("kilos") = format(conversiokilos(reixacompres.Columns("codimat"), cadbl(amplecompra), cadbl(r), cadbl(assignarmat.iespesor), assignarmat.itl, cadbl(assignarmat.isolapa)), "#,##0.00")
         If compres.Recordset.EditMode = 0 Then compres.Recordset.Edit
         compres.Recordset.Update
      End If
  End If
  If reixacompres.Columns(reixacompres.col).DataField = "Kgpendents" Then
      r = InputBox("Entra els kilos pendents d'entrega.", "Atenció")
      If r <> "" Then
         reixacompres.Columns("Kgpendents") = r
         If compres.Recordset.EditMode = 0 Then compres.Recordset.Edit
         compres.Recordset.Update
      End If
  End If
   If reixacompres.Columns(reixacompres.col).DataField = "kilos" Then
      r = InputBox("Entra els kilos de compra.", "Atenció")
      If r <> "" Then
         reixacompres.Columns("kilos") = r
         reixacompres.Columns("metres") = format(conversiokilos(reixacompres.Columns("codimat"), cadbl(amplecompra), cadbl(r) * -1, cadbl(assignarmat.iespesor), assignarmat.itl, cadbl(assignarmat.isolapa)), "#,##0")
         If compres.Recordset.EditMode = 0 Then compres.Recordset.Edit
         compres.Recordset.Update
      End If
  End If
End Sub

Private Sub Text1_Change()
 
End Sub

Private Sub reixacompres_HeadClick(ByVal ColIndex As Integer)
  Static direccio As String
  Static camp As String
  If camp = reixacompres.Columns(ColIndex).DataField Then
    direccio = IIf(direccio = " ASC ", " DESC ", " ASC ")
     Else: direccio = " ASC "
  End If
  camp = reixacompres.Columns(ColIndex).DataField
  compres.RecordSource = compres.Tag + " order by " + camp + direccio
  compres.Refresh

End Sub

Private Sub toteslescompres_Click()
Dim r As String
 If atrim(numdecompra) <> "" Then
    r = " where numcompra='" + atrim(numdecompra) + "'"
      Else:
        If MsgBox("Vols veure les NO ENTREGADES?", vbExclamation + vbYesNo, "Triar") = vbYes Then
           r = " where not entregada "
          Else: r = " where entregada"
        End If
 End If
 compres.RecordSource = "select * from compresmaterial  " + r + " order by data Desc"
 compres.Refresh

End Sub
