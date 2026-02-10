VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formabonosclients 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniment d'abonaments a Clients"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16770
   Icon            =   "formabonosclients.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   16770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Filtrar Lot"
      Height          =   285
      Left            =   2775
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Assignar Factura abono i data."
      Top             =   1665
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Assignar"
      Height          =   285
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Assignar Factura abono i data."
      Top             =   1665
      Width           =   1035
   End
   Begin VB.CommandButton beliminar 
      Height          =   285
      Left            =   660
      Picture         =   "formabonosclients.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Eliminar linia d'abonament"
      Top             =   1665
      Width           =   1035
   End
   Begin VB.CheckBox Checktramitats 
      Caption         =   "Tramitats"
      Height          =   195
      Left            =   10755
      TabIndex        =   17
      Top             =   1695
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Afegir Factura"
      Height          =   420
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   150
      Width           =   1350
   End
   Begin VB.Data Dataabonos 
      Caption         =   "dataabonos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6765
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "abonosclients"
      Top             =   120
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Frame Framedades 
      Caption         =   "Dades abonament"
      Enabled         =   0   'False
      Height          =   990
      Left            =   120
      TabIndex        =   7
      Top             =   690
      Width           =   11820
      Begin VB.TextBox cpais 
         BackColor       =   &H00E0E0E0&
         DataField       =   "pais"
         DataSource      =   "Dataabonos"
         Height          =   285
         Left            =   10365
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   165
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox ccolorreciclat 
         BackColor       =   &H00E0E0E0&
         DataField       =   "colorreciclat"
         DataSource      =   "Dataabonos"
         Height          =   285
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   480
         Width           =   1020
      End
      Begin VB.TextBox clotinplacsa 
         DataField       =   "lotinplacsa"
         DataSource      =   "Dataabonos"
         Height          =   285
         Left            =   3435
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox cdatafraabonament 
         DataField       =   "datafacturaabono"
         DataSource      =   "Dataabonos"
         Height          =   285
         Left            =   225
         TabIndex        =   1
         Top             =   465
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Acceptar"
         Height          =   450
         Left            =   10710
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   345
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox cimpost 
         BackColor       =   &H00E0E0E0&
         DataField       =   "totaimpost"
         DataSource      =   "Dataabonos"
         Height          =   285
         Left            =   9135
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   1035
      End
      Begin VB.TextBox cfacturaoriginal 
         BackColor       =   &H00E0E0E0&
         DataField       =   "numfacturaoriginal"
         DataSource      =   "Dataabonos"
         Height          =   285
         Left            =   4785
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox ccodicomptable 
         BackColor       =   &H00E0E0E0&
         DataField       =   "codiclient"
         DataSource      =   "Dataabonos"
         Height          =   285
         Left            =   6435
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox cnumfactura 
         DataField       =   "numfacturaabono"
         DataSource      =   "Dataabonos"
         Height          =   285
         Left            =   2010
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Color Reciclat"
         Height          =   225
         Left            =   8025
         TabIndex        =   18
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label etlotabonat 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot Inplacsa"
         Height          =   180
         Left            =   3600
         TabIndex        =   16
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Fra abonament"
         Height          =   180
         Left            =   165
         TabIndex        =   14
         Top             =   225
         Width           =   1530
      End
      Begin VB.Label etnomclient 
         BackStyle       =   0  'Transparent
         DataField       =   "nomclient"
         DataSource      =   "Dataabonos"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6390
         TabIndex        =   12
         Top             =   765
         Width           =   4125
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Impost"
         Height          =   225
         Left            =   9135
         TabIndex        =   11
         Top             =   210
         Width           =   1260
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Factura original client"
         Height          =   225
         Left            =   4770
         TabIndex        =   10
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi comptable client"
         Height          =   225
         Left            =   6375
         TabIndex        =   9
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Factura abonament"
         Height          =   180
         Left            =   1875
         TabIndex        =   8
         Top             =   240
         Width           =   1905
      End
   End
   Begin MSDBGrid.DBGrid reixa 
      Bindings        =   "formabonosclients.frx":0B14
      Height          =   4425
      Left            =   120
      OleObjectBlob   =   "formabonosclients.frx":0B29
      TabIndex        =   0
      Top             =   1935
      Width           =   16410
   End
End
Attribute VB_Name = "formabonosclients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function escullir_factura(vnumlot As Double) As String
  Load formseleccio
  formseleccio.sortirs.Tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "ImpostEnvasos.mdb"
  formseleccio.Data1.RecordSource = "select numfact as Num_Fra,datafactura as Data_Fra,U_GSP_INFABLOTE as Lot_Inplacsa from Importada_LiniesFacturesSAP_Inplacsa where itemcode<>'IMP_ENV' and itemcode<>'PLATES' and U_GSP_INFABLOTE='" + atrim(vnumlot) + "'"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 1500
  formseleccio.DBGrid2.Columns(1).Width = 1500
  formseleccio.DBGrid2.Columns(2).Width = 1000
  formseleccio.DBGrid2.Font.Size = 14
  formseleccio.Width = 6000
  formseleccio.Left = formseleccio.Left - 3000
  formseleccio.Show 1
  If seleccioret = 1 Then
           escullir_factura = cadbl(formseleccio.DBGrid2.Columns("Num_Fra"))
   End If
   Unload formseleccio
End Function

Private Sub beliminar_Click()
   If Dataabonos.Recordset.EOF Or Dataabonos.Recordset.BOF Then Exit Sub
   If MsgBox("Segur que vols eliminar aquest abonament?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
       Dataabonos.Recordset.Delete
       Dataabonos.Refresh
   End If
End Sub

Private Sub cdatafraabonament_LostFocus()
   If Not IsDate(cdatafraabonament) Then MsgBox "La data no es vàlida", vbCritical, "Error": cdatafraabonament = ""
End Sub

Private Sub Checktramitats_Click()
   If Checktramitats.Value = 1 Then
        Dataabonos.RecordSource = "select * from abonosclients where id not in(select id from abonosclients where (numremesa_592=0 or numremesa_592=null) or (numremesadestruccio_592=0 or numremesadestruccio_592=null) or (numremesa_A22=0 or numremesa_A22=null) or (numremesadestruccio_A22=0 or numremesadestruccio_A22=null))"
        Dataabonos.Refresh
          Else
            Dataabonos.RecordSource = "select * from abonosclients where (numremesa_592=0 or numremesa_592=null) or (numremesadestruccio_592=0 or numremesadestruccio_592=null) or (numremesa_A22=0 or numremesa_A22=null) or (numremesadestruccio_A22=0 or numremesadestruccio_A22=null)"
            Dataabonos.Refresh
   End If
End Sub

Private Sub clotinplacsa_LostFocus()
  Dim rst As Recordset
  Dim vimpost As Double
  Dim rst2 As Recordset
  If cadbl(clotinplacsa) = 0 Then Exit Sub
  Set rst = dbtmp.OpenRecordset("select * from Importada_LiniesFacturesSAP_Inplacsa where itemcode<>'IMP_ENV' AND itemcode<>'PLATES' and U_GSP_INFABLOTE='" + atrim(cadbl(clotinplacsa)) + "'")
  If rst.EOF Then
      MsgBox "No trobo la factura d'aquest lot a la base de dades.", vbCritical, "Error"
      cfacturaoriginal = ""
      ccodicomptable = ""
      etnomclient = ""
      clotinplacsa = ""
      cpais = ""
       Else
          rst.MoveLast: rst.MoveFirst
          ccodicomptable = atrim(rst!codicomptable)
          etnomclient = atrim(rst!nomclient)
          Dataabonos.Recordset!nif = atrim(rst!nif)
          If rst.RecordCount = 1 Then
            cfacturaoriginal = atrim(rst!numfact)
             Else
               cfacturaoriginal = escullir_factura(cadbl(clotinplacsa))
               If cadbl(cfacturaoriginal) = 0 Then
                  ccodicomptable = ""
                  etnomclient = ""
                  clotinplacsa = ""
                  cfacturaoriginal = ""
                  Dataabonos.Recordset!nif = ""
               End If
          End If
          cpais = buscar_pais_envio_comanda(cadbl(clotinplacsa))
          Set rst2 = dbtmp.OpenRecordset("Select * from capcaleraalbara where numfacturasap=" + atrim(cadbl(cfacturaoriginal)))
          If rst2.EOF Then MsgBox "No he localitzat l'albarà original d'expedició.", vbCritical, "Atenció": GoTo fi
          Set rst2 = dbtmp.OpenRecordset("select * from liniesalbara where numalbara=" + atrim(rst2!numalbara) + " and lotinplacsa=" + atrim(clotinplacsa))
          If rst2.EOF Then MsgBox "No he localitzat la linia d'albarà d'aquesta expedició.", vbCritical, "Atenció": GoTo fi
          If cadbl(rst2!KgImpostEnvasos) = 0 Then MsgBox "Aquest abonament no porta IMPOST no cal entrar-lo aquí", vbInformation, "Atenció": Dataabonos.Recordset.CancelUpdate: Exit Sub
          vimpost = cadbl(InputBox("Escriu la quantitat abonada d'aquest LOT " + atrim(clotinplacsa) + vbNewLine + "Quantitat de Kg o peces o metres.", "Import quantitat abonada"))
          vimpost = (cadbl(rst2!KgImpostEnvasos) * cadbl(vimpost)) / cadbl(rst2!quantitat)
          cimpost = Redondejar(vimpost, 2)
          vcolorreciclat = ""
          vcolorreciclat = UCase(colormaterialdelacomanda(clotinplacsa))
          If vcolorreciclat = "" Then
                While InStr(1, " [B] [V] [VR]", "[" + vcolorreciclat + "]") = 0
                    vcolorreciclat = UCase(InputBox("Escriu el color del tipus de material per reciclar." + vbNewLine + " [B]Blau   [V]Vermell   [VR]Verd", "Color reciclat"))
                    If StrPtr(vcolorreciclat) = 0 Then Dataabonos.Recordset.CancelUpdate: GoTo fi
                Wend
                If vcolorreciclat = "B" Then vcolorreciclat = "BLAU"
                If vcolorreciclat = "v" Then vcolorreciclat = "VERMELL"
                If vcolorreciclat = "VR" Then vcolorreciclat = "VERD"
          End If
          ccolorreciclat = vcolorreciclat
          wait 1
          Command1_Click
  End If
fi:
  Set rst = Nothing
  Set rst2 = Nothing
End Sub
Function buscar_pais_envio_comanda(vnumc As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("SELECT comandes.comanda, Clients_envios.pais FROM Clients_envios RIGHT JOIN comandes ON Clients_envios.id = comandes.direnvio WHERE (((comandes.comanda)=" + atrim(vnumc) + "));")
   If rst.EOF Then GoTo fi
   buscar_pais_envio_comanda = atrim(rst!pais)
fi:
   Set rst = Nothing
End Function

Private Sub Command1_Click()
    Dim vid As Long
    Framedades.Enabled = False
    vid = Dataabonos.Recordset.id
    If cadbl(clotinplacsa) = 0 Then
          If Dataabonos.Recordset.EditMode > 0 Then Dataabonos.Recordset.CancelUpdate
        Else:
            If Dataabonos.Recordset.EditMode > 0 Then Dataabonos.Recordset.Update
            Dataabonos.Recordset.FindFirst "id=" + atrim(vid)
            If Dataabonos.Recordset!pais = "ES" Then
                Dataabonos.Database.Execute "update abonosclients set numremesa_a22=-1, numremesa_592=-1 where id=" + atrim(Dataabonos.Recordset!id)
                  Else: Dataabonos.Database.Execute "update abonosclients set numremesa_a22=0 where id=" + atrim(Dataabonos.Recordset!id)
            End If
            If Not hihamaterialIMPOST("A22", Dataabonos.Recordset!lotinplacsa) Then
                        Dataabonos.Database.Execute "update abonosclients set numremesadestruccio_a22=-1 where id=" + atrim(Dataabonos.Recordset!id)
                          Else: Dataabonos.Database.Execute "update abonosclients set numremesadestruccio_a22=0 where id=" + atrim(Dataabonos.Recordset!id)
            End If
            If Not hihamaterialIMPOST("592", Dataabonos.Recordset!lotinplacsa) Then
                Dataabonos.Database.Execute "update abonosclients set numremesa_592=-1 where id=" + atrim(Dataabonos.Recordset!id)
                Dataabonos.Database.Execute "update abonosclients set numremesadestruccio_592=-1 where id=" + atrim(Dataabonos.Recordset!id)
                  Else
                    If Dataabonos.Recordset!pais = "ES" Then
                        Dataabonos.Database.Execute "update abonosclients set numremesa_592=-1 where id=" + atrim(Dataabonos.Recordset!id)
                         Else: Dataabonos.Database.Execute "update abonosclients set numremesa_592=0 where id=" + atrim(Dataabonos.Recordset!id)
                    End If
                    Dataabonos.Database.Execute "update abonosclients set numremesadestruccio_592=0 where id=" + atrim(Dataabonos.Recordset!id)
            End If
    End If
End Sub
Function hihamaterialIMPOST(vtipus As String, vnumc As Double) As Boolean
   Dim rst As Recordset
   Dim rsti As Recordset
   Set rst = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(vnumc))
   If rst.EOF Then GoTo fi
   Set rsti = dbtmp.OpenRecordset("select sum(kgventaespanya+kgventaimp_mes_esp) as T_A22,sum(kgventaad_intracom) as T_592 from impostenvasos where comanda=" + atrim(rst!comanda) + " or comanda=" + atrim(cadbl(rst!linkcomanda1)) + " or comanda=" + atrim(cadbl(rst!linkcomanda2)))
   If rsti.EOF Then GoTo fi
   If vtipus = "A22" Then If cadbl(rsti!T_A22) > 0 Then hihamaterialIMPOST = True
   If vtipus = "592" Then If cadbl(rsti!T_592) > 0 Then hihamaterialIMPOST = True
fi:
   Set rst = Nothing
   Set rsti = Nothing
End Function

Private Sub Command2_Click()
  Dim vlot As Double
  Dim rst As Recordset
  Dim vid As Long
  If Dataabonos.Recordset.EditMode > 0 Then Dataabonos.Recordset.CancelUpdate
  vlot = cadbl(InputBox("Escriu el Lot d'inplacsa que s'ha d'abonar.", "Atenció"))
  If vlot = 0 Then Exit Sub
  Set rst = dbtmp.OpenRecordset("Select * from abonosclients where (numremesa_592=0 or numremesa_592=null) and (numremesadestruccio_592=0 or numremesadestruccio_592=null) AND (numremesa_A22=0 or numremesa_A22=null) and (numremesadestruccio_A22=0 or numremesadestruccio_A22=null) and lotinplacsa=" + atrim(vlot))
  If Not rst.EOF Then If MsgBox("Ja hi ha entrat una comanda a l'espera de tramitar l'abonament amb aquest numero de lot." + vbNewLine + "ASSEGURA QUE AQUEST LOT TINGUI DOS ABONAMENTS I SIGUI CORRECTE." + vbNewLine + "VOLS CONTINUAR?", vbCritical + vbDefaultButton2 + vbYesNo, "Error") = vbNo Then Exit Sub
  'Framedades.Enabled = True
  'cdatafraabonament.SetFocus
  Dataabonos.Recordset.AddNew
  vid = cadbl(Dataabonos.Recordset.id)
  clotinplacsa = atrim(vlot)
  clotinplacsa_LostFocus
  Dataabonos.Refresh
  Dataabonos.Recordset.FindFirst "id=" + atrim(vid)
  
End Sub

Private Sub Command3_Click()
   Dim vnumc As Double
   Dim vdata As String
   Dim vnumfact As String
   Dim vimport As Double
   vnumc = cadbl(InputBox("Entra el LOT que vols assignar data i factura d'abonament.", "Entrada factura"))
   If vnumc = 0 Then Exit Sub
   Dataabonos.RecordSource = "select * from abonosclients where (numremesa_592=0 or numremesa_592=null) or (numremesadestruccio_592=0 or numremesadestruccio_592=null) or (numremesa_A22=0 or numremesa_A22=null) or (numremesadestruccio_A22=0 or numremesadestruccio_A22=null)"
   Dataabonos.Refresh
   Dataabonos.Recordset.FindFirst "lotinplacsa=" + atrim(cadbl(vnumc)) + " and (numfacturaabono='0' or numfacturaabono=null)"
   If Dataabonos.Recordset.NoMatch Then MsgBox "No he trobat aquest lot o no està pendent d'assignar numero d'abonament.", vbCritical, "Error": Exit Sub
   Dataabonos.Recordset.FindNext "lotinplacsa=" + atrim(cadbl(vnumc)) + " and (numfacturaabono='0' or numfacturaabono=null)"
   If Not Dataabonos.Recordset.NoMatch Then
       vimport = cadbl(InputBox("Aquest lot té mes d'una entrada." + vbNewLine + "ESCRIU L'IMPORT DE LA DEVOLUCIÓ.", "MES D'UN LOT"))
       Dataabonos.Recordset.FindFirst "totaimpost=" + passaradecimalpunt(cadbl(vimport)) + " and lotinplacsa=" + atrim(cadbl(vnumc)) + " and (numfacturaabono='0' or numfacturaabono=null)"
       If Dataabonos.Recordset.NoMatch Then MsgBox "No he trobat aquest lot o no està pendent d'assignar numero d'abonament.", vbCritical, "Error": Exit Sub
   End If
   vdata = InputBox("Escriu la data de la factura d'abonament.", "Data")
   If Not IsDate(vdata) Then MsgBox "Data no vàlida": Exit Sub
   vnumfact = InputBox("Escriu el numero de factura d'abonament.", "Num factura")
   Dataabonos.Recordset.Edit
   Dataabonos.Recordset!datafacturaabono = vdata
   Dataabonos.Recordset!numfacturaabono = vnumfact
   Dataabonos.Recordset.Update
   Command1_Click
   Dataabonos.Recordset.Move 0
   
End Sub

Private Sub Command4_Click()
 Dim vnumlot As String
  vnumlot = InputBox("Escriu el numero de LOT que vols filtrar.", "Atenció")
  If StrPtr(vnumlot) = 0 Then vnumlot = "1"
  vnumlot = atrim(cadbl(vnumlot))
  If vnumlot = 0 Then Exit Sub
If Checktramitats.Value = 1 Then
        Dataabonos.RecordSource = "select * from abonosclients where lotinplacsa" + IIf(vnumlot = 1, ">", "=") + atrim(vnumlot) + " AND numremesa_592=0 AND numremesa_592=null and numremesadestruccio_592=0 AND numremesadestruccio_592=null AND numremesa_A22=0 AND numremesa_A22=null and numremesadestruccio_A22=0 AND numremesadestruccio_A22=null"
        Dataabonos.Refresh
          Else
            Dataabonos.RecordSource = "select * from abonosclients where lotinplacsa" + IIf(vnumlot = 1, ">", "=") + atrim(vnumlot) + " and ((numremesa_592=0 or numremesa_592=null) or (numremesadestruccio_592=0 or numremesadestruccio_592=null) or (numremesa_A22=0 or numremesa_A22=null) or (numremesadestruccio_A22=0 or numremesadestruccio_A22=null))"
            Dataabonos.Refresh
   End If
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Form_DblClick()
 'Command1_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then KeyAscii = 0: Sendkeys "{TAB}"
End Sub

Private Sub Form_Load()

 Dataabonos.DatabaseName = rutadelfitxer(cami) + "ImpostEnvasos.mdb"
 
 'Dataabonos.RecordSource = "select * from abonosclients where (numremesa_592=0 or numremesa_592=null) or (numremesadestruccio_592=0 or numremesadestruccio_592=null) or (numremesa_A22=0 or numremesa_A22=null) or (numremesadestruccio_A22=0 or numremesadestruccio_A22=null)"
 Dataabonos.RecordSource = "select * from abonosclients order by datafacturaabono desc"
 Dataabonos.Refresh
 Checktramitats_Click
End Sub

Private Sub reixa_DblClick()
  Dim v As String
  If reixa.Columns(reixa.Col).DataField = "numfacturaabono" Then
        v = InputBox("Escriu el nou numero de factura.", "Num factura")
        Dataabonos.Recordset.Edit
        Dataabonos.Recordset!numfacturaabono = v
        Dataabonos.Recordset.Update
        Dataabonos.Recordset.Move 0
  End If
  
  If reixa.Columns(reixa.Col).DataField = "datafacturaabono" Then
        v = InputBox("Escriu la nova data de factura.", "Data factura")
        If Not IsDate(v) Then MsgBox "Data no vàlida.": Exit Sub
        Dataabonos.Recordset.Edit
        Dataabonos.Recordset!datafacturaabono = v
        Dataabonos.Recordset.Update
        Dataabonos.Recordset.Move 0
  End If
  Command2.SetFocus
End Sub

Private Sub reixa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If Dataabonos.Recordset.EOF Then beliminar.Enabled = False: Exit Sub
   If cadbl(Dataabonos.Recordset!numremesa_A22) > 0 Or cadbl(Dataabonos.Recordset!numremesadestruccio_A22) > 0 Or cadbl(Dataabonos.Recordset!numremesa_592) > 0 Or cadbl(Dataabonos.Recordset!numremesadestruccio_592) > 0 Then
        beliminar.Enabled = False
          Else: beliminar.Enabled = True
   End If
End Sub

Function colormaterialdelacomanda(vnumc As Double) As String
    Dim v1 As Double
    Dim v2 As Double
    Dim v3 As Double
    Dim c1 As Double
    Dim c2 As Double
    Dim vvalormesgran As Byte
    Dim v As String
   
    Dim vcolor As Double
    Dim rst As Recordset
    Set rst = dbtmp.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(vnumc))
    If rst.EOF Then Exit Function
    v1 = NUMEROCOLORmaterialdelacomanda(atrim(rst!comanda))
    v2 = NUMEROCOLORmaterialdelacomanda(atrim(rst!linkcomanda1))
    v3 = NUMEROCOLORmaterialdelacomanda(atrim(rst!linkcomanda2))
    vvalormesgran = IIf((IIf(v1 > v2, v1, v2)) > v3, (IIf(v1 > v2, v1, v2)), v3)
    colormaterialdelacomanda = IIf(vvalormesgran = 1, "Verd", IIf(vvalormesgran = 2, "Blau", "Vermell"))
End Function
Public Function NUMEROCOLORmaterialdelacomanda(vnumc As Double) As Double
    Dim rst As Recordset
    NUMEROCOLORmaterialdelacomanda = 0
    If vnumc = 0 Then Exit Function
    Set rst = dbtmp.OpenRecordset("SELECT materials.colorreciclatge FROM comandes INNER JOIN materials ON comandes.materialex = materials.codi where comanda=" + atrim(cadbl(vnumc)))
    If Not rst.EOF Then NUMEROCOLORmaterialdelacomanda = IIf(atrim(rst!colorreciclatge) = "Verd", 1, IIf(atrim(rst!colorreciclatge) = "Blau", 2, 3))
    Set rst = Nothing
End Function
