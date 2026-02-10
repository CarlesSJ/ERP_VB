VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formrevisatescanejats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revisar escanejats desde expedicions"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16560
   Icon            =   "revisio_escanajats_expedicions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   16560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bveurelots 
      Caption         =   "Filtrar Lots"
      Height          =   285
      Left            =   0
      TabIndex        =   15
      ToolTipText     =   "Filtrar lots CQ d'aquest albarà"
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00EEE4D7&
      Caption         =   "Veure Tot"
      Height          =   285
      Left            =   2085
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   570
      Width           =   1005
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00EEE4D7&
      Caption         =   "Veure CQ"
      Height          =   285
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   570
      Width           =   1005
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00EEE4D7&
      Caption         =   "Veure Alb"
      Height          =   285
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   570
      Width           =   1005
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FDDECE&
      Caption         =   "Llista de Lots sense CQ        (Últim any)"
      Height          =   525
      Left            =   10470
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   135
      Width           =   2490
   End
   Begin VB.CommandButton bcanvidenom 
      Height          =   285
      Left            =   8490
      Picture         =   "revisio_escanajats_expedicions.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Canvi de nom del document"
      Top             =   1020
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame Frame1 
      Caption         =   "Afegir arrastrant"
      Height          =   810
      Left            =   7800
      TabIndex        =   7
      Top             =   -15
      Width           =   2520
      Begin VB.CommandButton Command6 
         Caption         =   "CQ Lots"
         Height          =   435
         Left            =   1290
         TabIndex        =   9
         Top             =   255
         Width           =   1110
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Albarans"
         Height          =   435
         Left            =   135
         TabIndex        =   8
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   330
      Left            =   5850
      Picture         =   "revisio_escanajats_expedicions.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Filtrar"
      Top             =   510
      Width           =   375
   End
   Begin VB.TextBox vfiltre 
      Height          =   285
      Left            =   3135
      TabIndex        =   5
      Top             =   555
      Width           =   2685
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H006BEBB1&
      Caption         =   "Marcar Revisat"
      Height          =   390
      Left            =   14790
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   330
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tots"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "No Revisats"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   1065
   End
   Begin VB.Data Dataescanejats 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from registre_escanejades_expedicions where not revisat order by data desc"
      Top             =   165
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "revisio_escanajats_expedicions.frx":109E
      Height          =   7605
      Left            =   60
      OleObjectBlob   =   "revisio_escanajats_expedicions.frx":10B7
      TabIndex        =   0
      Top             =   855
      Width           =   16350
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No revisats"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F1B75F&
      Height          =   435
      Left            =   2535
      TabIndex        =   3
      Top             =   135
      Width           =   3615
   End
End
Attribute VB_Name = "formrevisatescanejats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bcanvidenom_Click()
    Dim vnom As String
    Dim vrutalocal As String
    Dim vrutadrive As String
    Dim vnumcmr As String
   
    If Dataescanejats.Recordset!tipus = "CMR" Then
      vrutalocal = llegir_ini("ruta", "rutaAlbaransSAPLOCAL", rutadelfitxer(cami) + "valorsprograma.ini") + "CMRs\"
      vrutadrive = llegir_ini("ruta", "rutaAlbaransSAPDRIVE", rutadelfitxer(cami) + "valorsprograma.ini") + "CMRs\"
      vnumcmr = substituir(Dataescanejats.Recordset!nomfitxer, "CMR_", "")
      vnumcmr = substituir(LCase(vnumcmr), ".pdf", "")
    End If
    If Dataescanejats.Recordset!tipus = "SAP" Then
      vrutalocal = llegir_ini("ruta", "rutaAlbaransSAPLOCAL", rutadelfitxer(cami) + "valorsprograma.ini")
      vrutadrive = llegir_ini("ruta", "rutaAlbaransSAPDRIVE", rutadelfitxer(cami) + "valorsprograma.ini")
    End If
    If Dataescanejats.Recordset!tipus = "CQ" Then
      vrutalocal = llegir_ini("ruta", "rutaCQlotsLOCAL", rutadelfitxer(cami) + "valorsprograma.ini")
      vrutadrive = llegir_ini("ruta", "rutaCQlotsDRIVE", rutadelfitxer(cami) + "valorsprograma.ini")
    End If
    If Dataescanejats.Recordset!tipus = "ALB" Then
      vrutalocal = llegir_ini("ruta", "rutaAlbaransProveidorsLOCAL", rutadelfitxer(cami) + "valorsprograma.ini")
      vrutadrive = llegir_ini("ruta", "rutaAlbaransProveidorsDRIVE", rutadelfitxer(cami) + "valorsprograma.ini") + "EscanejatDesdeExpedicions\" + atrim(Year(Now)) + "\"
    End If
    If existeix(vrutalocal + DBGrid1.Text) And existeix(vrutadrive + DBGrid1.Text) Then
       vnom = atrim(Dataescanejats.Recordset!nomfitxer)
       vnom = InputBox("Escriu el nou nom pel fitxer... ATENCIÓ DE RESPECTAR L'ESTRUCTURA." + vbNewLine + "PER ELIMINAR L'ESCANEJADA ESCRIU [ELIMINAR].", "CANVI DE NOM", vnom)
       If StrPtr(vnom) = 0 Then Exit Sub
       If UCase(vnom) = "ELIMINAR" Then
           If existeix(vrutalocal + atrim(Dataescanejats.Recordset!nomfitxer)) Then Kill vrutalocal + atrim(Dataescanejats.Recordset!nomfitxer)
           If existeix(vrutadrive + atrim(Dataescanejats.Recordset!nomfitxer)) Then Kill vrutadrive + atrim(Dataescanejats.Recordset!nomfitxer)
           Dataescanejats.Recordset.Delete
           Dataescanejats.Refresh
           If cadbl(vnumcmr) > 0 Then dbtmp.Execute "update transportistes_avisos set escanejat=false where numeroavis='" + atrim(vnumcmr) + "'"
          ' Dataescanejats.Recordset.Move 0
           MsgBox "Fitxers eliminats... " + vbNewLine + " ASSEGURA QUE ELS CQ I ALBARANS QUE TINGUIN ALGUNA RELACIÓ ESTIGUIN CORRECTES.", vbCritical, "ATENCIÓ"
           GoTo fi
       End If
       'canvio el nom de la LOCAL
       FileCopy vrutalocal + atrim(Dataescanejats.Recordset!nomfitxer), vrutalocal + "\" + vnom
       wait 1
       If existeix(vrutalocal + "\" + vnom) Then Kill vrutalocal + atrim(Dataescanejats.Recordset!nomfitxer)
       'canvio el nom del DRIVE
       FileCopy vrutadrive + atrim(Dataescanejats.Recordset!nomfitxer), vrutadrive + "\" + vnom
       wait 1
       If existeix(vrutadrive + "\" + vnom) Then Kill vrutadrive + atrim(Dataescanejats.Recordset!nomfitxer)
       Dataescanejats.Recordset.Edit
       Dataescanejats.Recordset!nomfitxer = vnom
       Dataescanejats.Recordset.Update
       Dataescanejats.Recordset.Move 0
         Else:
           If Not existeix(vrutalocal) Or Not existeix(vrutadrive) Then
            MsgBox "Error al accedir a un dels fitxers no es pot canviar el nom." + vbNewLine + "Local-> " + vrutalocal + DBGrid1.Text + vbNewLine + "Drive-> " + vrutadrive + DBGrid1.Text, vbCritical, "Error"
              Else
                 If MsgBox("Aquest fitxer no existeix al servidor possiblement hi ha hagut un error d'escaneig." + vbNewLine + "Vols eliminar aquesta entrada i tornar a escanejar-lo?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
                    Dataescanejats.Recordset.Delete
                    Dataescanejats.Refresh
'                    Dataescanejats.Recordset.Move 0
                 End If
           End If
    End If
fi:
End Sub
'
Function buscaralbaransdelCQ(vCQ As String) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select numalbaraprov from albaransbip where numlotproveidor='" + atrim(vCQ) + "'")
  While Not rst.EOF
     buscaralbaransdelCQ = buscaralbaransdelCQ + IIf(buscaralbaransdelCQ <> "", " or ", "") + " nomfitxer like'" + treuresimbolsnovalidsnomfitxer(atrim(rst!numalbaraprov)) + " *' "
     rst.MoveNext
  Wend
  Set rst = Nothing
End Function
Function buscarCQdelalbara(valbara As String) As String
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select numlotproveidor from albaransbip where numalbaraprov='" + atrim(valbara) + "'")
  While Not rst.EOF
     buscarCQdelalbara = buscarCQdelalbara + IIf(buscarCQdelalbara <> "", " or ", "") + " nomfitxer like'CQ_" + atrim(rst!numlotproveidor) + " *' "
     rst.MoveNext
  Wend
  Set rst = Nothing
End Function

Private Sub bveurelots_Click()
  If bveurelots.Caption = "Filtra Alb." Then
         filtrar_Albarans
          Else
           filtrar_CQs
  End If
           
End Sub
Sub filtrar_Albarans()
  Dim vllistaalbarans As String
  vllistaalbarans = buscaralbaransdelCQ(Mid(Dataescanejats.Recordset!nomfitxer, 4, InStr(1, Dataescanejats.Recordset!nomfitxer, " ") - 4))
  
  If vllistaalbarans <> "" Then
    Dataescanejats.RecordSource = "select * from registre_escanejades_expedicions where " + vllistaalbarans + " order by data desc"
    Dataescanejats.Refresh
    Label1.Caption = ""
  End If
End Sub
Sub filtrar_CQs()
  Dim vllistacqs As String
  Dim vnumalbara As String
  vnumalbara = Mid(Dataescanejats.Recordset!nomfitxer, 1, InStr(1, Dataescanejats.Recordset!nomfitxer, " "))
  vnumalbara = substituir(vnumalbara, "_", "/")
  vllistacqs = buscarCQdelalbara(vnumalbara)
  If vllistacqs <> "" Then
    Dataescanejats.RecordSource = "select * from registre_escanejades_expedicions where " + vllistacqs + " order by data desc"
    Dataescanejats.Refresh
    Label1.Caption = ""
  End If
End Sub

Private Sub Command10_Click()
Dataescanejats.RecordSource = "select * from registre_escanejades_expedicions where not revisat order by data desc"
Dataescanejats.Refresh
End Sub

Private Sub Command2_Click()
  Dataescanejats.RecordSource = "select * from registre_escanejades_expedicions order by data desc"
  Dataescanejats.Refresh
  Label1.Caption = "Tots"
End Sub

Private Sub Command3_Click()
    If MsgBox("Estas segur que vols canviar l'estat de REVISAT?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
    If DBGrid1.SelBookmarks.Count = 0 Then DBGrid1.SelBookmarks.Add DBGrid1.Bookmark
    For i = 0 To DBGrid1.SelBookmarks.Count - 1
      Dataescanejats.Recordset.Bookmark = DBGrid1.SelBookmarks(i)
      Dataescanejats.Recordset.Edit
      Dataescanejats.Recordset!revisat = Not Dataescanejats.Recordset!revisat
      If Dataescanejats.Recordset!revisat Then
           Dataescanejats.Recordset!datarev = Now
           Dataescanejats.Recordset!operari = nomordinador
            Else
              Dataescanejats.Recordset!datarev = Null
              Dataescanejats.Recordset!operari = ""
      End If
      Dataescanejats.Recordset.Update
    Next i
  
End Sub

Private Sub Command4_Click()
  If Label1.Caption = "Tots" Then
     Dataescanejats.RecordSource = "select * from registre_escanejades_expedicions WHERE NOMFITXER like '*" + treure_apostrof(vfiltre) + "*' order by data desc"
     Dataescanejats.Refresh
    Else
        Dataescanejats.RecordSource = "select * from registre_escanejades_expedicions where not revisat and NOMFITXER like '*" + treure_apostrof(vfiltre) + "*' order by data desc"
        Dataescanejats.Refresh
  End If
End Sub

Private Sub Command5_Click()
   Load formescanejaralbaransproveidor
    
   formescanejaralbaransproveidor.ettipusescaneig.Caption = "Esperant Albarans del Proveïdor..."
   formescanejaralbaransproveidor.Tag = "albarans"
   formescanejaralbaransproveidor.Caption = formescanejaralbaransproveidor.ettipusescaneig.Caption
   formescanejaralbaransproveidor.vcarpeta = "c:\temp\escaneigdocumentacio\"
   formescanejaralbaransproveidor.eliminar_fitxersdelacarpetaescaner
   DoEvents
   formescanejaralbaransproveidor.Show 1
End Sub

Private Sub Command6_Click()
 Load formescanejaralbaransproveidor
    
   formescanejaralbaransproveidor.ettipusescaneig.Caption = "Esperant CQ de lots..."
   formescanejaralbaransproveidor.Tag = "certificats"
   formescanejaralbaransproveidor.Caption = formescanejaralbaransproveidor.ettipusescaneig.Caption
   formescanejaralbaransproveidor.vcarpeta = "c:\temp\escaneigdocumentacio\"
   formescanejaralbaransproveidor.eliminar_fitxersdelacarpetaescaner
   DoEvents
   formescanejaralbaransproveidor.Show 1
End Sub

Private Sub Command7_Click()
  Dim rst As Recordset
   Load formseleccio
   formseleccio.sortirs.Tag = "filtre"
   formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
   formseleccio.Data1.RecordSource = "SELECT distinct numlotproveidor,data,nomproveidorcomercial,article,descripcio FROM albaransbip where numlotproveidor<>'-' and cal_CQ_lot=true and lotescanejat=false and ((albaransbip.data)>#11/1/2022#)"
  
  formseleccio.refrescar
  formseleccio.Width = 15000
  formseleccio.DBGrid2.Columns(0).Width = 2000
  formseleccio.DBGrid2.Columns(1).Width = 1000
  formseleccio.DBGrid2.Columns(2).Width = 5000
  formseleccio.DBGrid2.Columns(3).Width = 800
  formseleccio.DBGrid2.Columns(4).Width = 5000
  
  
  If formseleccio.Data1.Recordset.EOF Then MsgBox "No hi ha cap lot sense CQ", vbInformation, "Atenció": GoTo fi
  formseleccio.Show 1
  If seleccioret = 1 Then
     vfiltre = formseleccio.Data1.Recordset!numlotproveidor
     Dataescanejats.RecordSource = "select * from registre_escanejades_expedicions WHERE NOMFITXER like '*" + treure_apostrof(vfiltre) + "*' order by data desc"
     Dataescanejats.Refresh
  End If
fi:
  Unload formseleccio
End Sub

Private Sub Command8_Click()
   Dataescanejats.RecordSource = "select * from registre_escanejades_expedicions where not revisat AND tipus='ALB' order by data desc"
   Dataescanejats.Refresh
End Sub

Private Sub Command9_Click()
Dataescanejats.RecordSource = "select * from registre_escanejades_expedicions where not revisat AND tipus='CQ' order by data desc"
Dataescanejats.Refresh
End Sub

Private Sub DBGrid1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
   colocar_boto_edicio
End Sub

Private Sub DBGrid1_DblClick()
  Dim vruta As String
  Dim vresp As String
  
  
  If DBGrid1.Columns(DBGrid1.col).Caption = "Observacions" Then
      vresp = DBGrid1.Text
      vresp = InputBox("Escriu una observació.", "Observació", vresp)
      If StrPtr(vresp) = 0 Then Exit Sub
      Dataescanejats.Recordset.Edit
      Dataescanejats.Recordset!observacions = vresp
      Dataescanejats.Recordset.Update
  End If
  If DBGrid1.Columns(DBGrid1.col).Caption = "Nom fitxer" Then
    If Dataescanejats.Recordset!tipus = "SAP" Then
      vruta = llegir_ini("ruta", "rutaAlbaransSAPLOCAL", rutadelfitxer(cami) + "valorsprograma.ini")
    End If
    If Dataescanejats.Recordset!tipus = "CQ" Then
      vruta = llegir_ini("ruta", "rutaCQlotsLOCAL", rutadelfitxer(cami) + "valorsprograma.ini")
    End If
    If Dataescanejats.Recordset!tipus = "ALB" Then
      vruta = llegir_ini("ruta", "rutaAlbaransProveidorsLOCAL", rutadelfitxer(cami) + "valorsprograma.ini")
    End If
    If Dataescanejats.Recordset!tipus = "CMR" Then
      vruta = llegir_ini("ruta", "rutaAlbaransSAPLOCAL", rutadelfitxer(cami) + "valorsprograma.ini") + "CMRs\"
    End If
    If existeix(vruta + DBGrid1.Text) Then
          obrir_document vruta + DBGrid1.Text
           Else: MsgBox "No trobo o no es pot accedir al fitxer " + vruta + DBGrid1.Text
    End If
  End If
End Sub

Private Sub DBGrid1_LostFocus()
   If ActiveControl.Name <> "bcanvidenom" Then bcanvidenom.Visible = False
   If ActiveControl.Name <> "bveurelots" Then bveurelots.Visible = False
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    colocar_boto_edicio
End Sub
Sub colocar_boto_edicio()
    If DBGrid1.Columns(DBGrid1.col).Caption = "Nom fitxer" Then
       bcanvidenom.Left = DBGrid1.Columns(DBGrid1.col).Left + DBGrid1.Columns(DBGrid1.col).Width '- bcanvidenom.Width
       bcanvidenom.Top = DBGrid1.Top + (DBGrid1.RowHeight * (DBGrid1.row + 1) - 10)
       bcanvidenom.Visible = True
         Else: bcanvidenom.Visible = False
    End If
    If Not Dataescanejats.Recordset.EOF Then
        If Dataescanejats.Recordset!tipus = "ALB" Or Dataescanejats.Recordset!tipus = "CQ" Then
           bveurelots.Left = DBGrid1.Columns(2).Left + DBGrid1.Columns(2).Width - bveurelots.Width - bcanvidenom.Width - 10
           bveurelots.Top = DBGrid1.Top + (DBGrid1.RowHeight * (DBGrid1.row + 1) - 10)
           bveurelots.Caption = "Filtra Lots"
           If Dataescanejats.Recordset!tipus = "CQ" Then bveurelots.Caption = "Filtra Alb."
           bveurelots.Visible = True
             Else: bveurelots.Visible = False
        End If
           Else: bveurelots.Visible = False
    End If
End Sub
Private Sub Form_Load()
   Shell "c:\windows\system32\net use \\ord_josepm /user:josepm josepm", vbHide
   Dataescanejats.DatabaseName = cami
   Dataescanejats.RecordSource = "select * from registre_escanejades_expedicions where not revisat AND tipus='ALB' order by data desc"
End Sub

Private Sub vfiltre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Command4_Click: KeyCode = 0
End Sub

Private Sub vfiltre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then KeyAscii = 0
End Sub
