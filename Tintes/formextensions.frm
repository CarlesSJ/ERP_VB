VERSION 5.00
Begin VB.Form formextensions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extensions"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7455
   Icon            =   "formextensions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botoafegirrelacio 
      Height          =   390
      Left            =   7065
      Picture         =   "formextensions.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Alta  Registres"
      Top             =   1800
      Width           =   390
   End
   Begin VB.CommandButton botoeliminarrelacio 
      Height          =   390
      Left            =   7065
      Picture         =   "formextensions.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Eliminar relació d'aquesta extensió amb el treball escullit."
      Top             =   2190
      Width           =   390
   End
   Begin VB.TextBox observacions 
      Height          =   525
      Left            =   270
      MaxLength       =   100
      TabIndex        =   14
      Top             =   5160
      Width           =   6975
   End
   Begin VB.CommandButton Command52 
      Height          =   420
      Left            =   6900
      Picture         =   "formextensions.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Veure la comanda"
      Top             =   105
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   5730
      Picture         =   "formextensions.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Guardar la fulla de control de l'extensió i sortir."
      Top             =   5760
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   4185
      Picture         =   "formextensions.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprimir fulla de control de l'extensió."
      Top             =   5760
      Width           =   1515
   End
   Begin VB.TextBox cnumquadrat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   9
      Top             =   3960
      Width           =   1035
   End
   Begin VB.TextBox cvolum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2040
      TabIndex        =   8
      Top             =   3510
      Width           =   1035
   End
   Begin VB.ComboBox comboanilox 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "formextensions.frx":213C
      Left            =   2070
      List            =   "formextensions.frx":216A
      TabIndex        =   4
      Top             =   2385
      Width           =   1620
   End
   Begin VB.ListBox llistatreballs 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   5790
      TabIndex        =   2
      Top             =   1770
      Width           =   1245
   End
   Begin VB.Label etmarca 
      BackStyle       =   0  'Transparent
      Caption         =   "Marca"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F3B378&
      Height          =   540
      Left            =   180
      TabIndex        =   16
      Top             =   1815
      Width           =   5505
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Observacions:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   255
      TabIndex        =   15
      Top             =   4890
      Width           =   3375
   End
   Begin VB.Label ettreballactual 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   180
      TabIndex        =   12
      Top             =   135
      Width           =   2745
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Quadrat: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   7
      Top             =   4035
      Width           =   1560
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Volum:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   255
      TabIndex        =   6
      Top             =   3570
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Anilox utilitzat:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   225
      TabIndex        =   5
      Top             =   2490
      Width           =   1980
   End
   Begin VB.Label Label1 
      Caption         =   "Treballs afectats:"
      Height          =   225
      Left            =   5820
      TabIndex        =   3
      Top             =   1560
      Width           =   1320
   End
   Begin VB.Label etnumextensio 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Ext: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4320
      TabIndex        =   1
      Top             =   75
      Width           =   2550
   End
   Begin VB.Label etnomtinta 
      BackStyle       =   0  'Transparent
      Caption         =   "NOMDELATINTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1320
      Left            =   240
      TabIndex        =   0
      Top             =   375
      Width           =   5910
   End
End
Attribute VB_Name = "formextensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub netejar_pantalla_extensions()
    etnumextensio = ""
    comboanilox = ""
    base = ""
    atenuant = ""
    disolvent = ""
    etnumextensio.tag = ""
    observacions = ""
    llistatreballs.Clear
End Sub
Sub carregar_extensio(vnumextensio As String)
    Dim rst As Recordset
    netejar_pantalla_extensions
    Set rst = dbtintes.OpenRecordset("select * from extensions where codiextensio='" + atrim(vnumextensio) + "'")
    If Not rst.EOF Then
      etnumextensio = "Nº Ext: " + atrim(rst!codiextensio)
      etnumextensio.tag = atrim(rst!codiextensio)
      comboanilox = atrim(cadbl(rst!anilox))
      'base = atrim(cadbl(rst!base))
      'atenuant = atrim(cadbl(rst!atenuant))
      'disolvent = atrim(cadbl(rst!disolvent))
      'kilos = atrim(cadbl(rst!kilos))
      cvolum = atrim(cadbl(rst!volum))
      cnumquadrat = atrim(cadbl(rst!numquadrat))
      observacions = atrim(rst!observacions)
      llistatreballs.Clear
      Set rst = dbtintes.OpenRecordset("select * from extensions_treballsrelacionats where  codiextensio='" + atrim((etnumextensio.tag)) + "'")
      While Not rst.EOF
         llistatreballs.AddItem atrim(rst!numtreball) + "/" + atrim(rst!numordremodificacio)
         rst.MoveNext
      Wend
      possar_lamarca cadbl(ettreballactual.tag)
      
    End If
    Set rst = Nothing
End Sub
Sub possar_lamarca(vnumtreball As Double)
   Dim rst As Recordset
   etmarca = ""
   Set rst = dbclixes.OpenRecordset("select marca from clixes where id_Treball=" + atrim(vnumtreball))
   If Not rst.EOF Then etmarca = atrim(rst!marca)
   Set rst = Nothing
End Sub
Function buscar_laextensiomesgran(vcoditinta As String) As Long
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("select * from extensions where coditinta=" + atrim(cadbl(cadbl(etnomtinta.tag))) + " order by ordre desc")
  If Not rst.EOF Then
     buscar_laextensiomesgran = cadbl(rst!ordre) + 1
      Else: buscar_laextensiomesgran = 1
  End If
  Set rst = Nothing
End Function

Sub gravar_extensio_actual()
    Dim rst As Recordset
    Dim rstclixes As Recordset
    Dim vnommarca As String
    Set rst = dbtintes.OpenRecordset("select * from extensions where codiextensio='" + atrim(etnumextensio.tag) + "'")
    If Not rst.EOF Then
      rst.Edit
        Else
          rst.AddNew
          rst!coditinta = cadbl(etnomtinta.tag)
          rst!ordre = buscar_laextensiomesgran(cadbl(etnomtinta.tag))
          etnumextensio.tag = atrim(rst!coditinta) + "-" + atrim(rst!ordre)
    End If
    Set rstclixes = dbclixes.OpenRecordset("select marca from clixes where id_treball=" + atrim(cadbl(ettreballactual.tag)))
    If Not rstclixes.EOF Then vnommarca = atrim(rstclixes!marca)
    rst!codiextensio = etnumextensio.tag
    rst!anilox = cadbl(comboanilox)
   ' rst!base = cadbl(base)
   ' rst!atenuant = cadbl(atenuant)
   ' rst!disolvent = cadbl(disolvent)
   ' rst!kilos = cadbl(kilos)
    rst!volum = cadbl(cvolum)
    rst!numquadrat = cadbl(cnumquadrat)
    rst!observacions = atrim(observacions)
    rst!marca = vnommarca
    rst.Update
    guardar_treballsdelextensio etnumextensio.tag
    actualitzar_dadesextensioalstreballs etnumextensio.tag, cadbl(comboanilox), cadbl(cvolum)
    carregar_extensio etnumextensio.tag
    Set rst = Nothing
End Sub
Sub guardar_treballsdelextensio(vnumextensio As String, Optional vtreball As Double, Optional vordre As Double)
    Dim rst As Recordset
    Dim vnumtreballescullit As Double
    Dim vordreescullit As Double
    vnumtreballescullit = IIf(cadbl(vtreball) > 0, vtreball, cadbl(ettreballactual.tag))
    vordreescullit = IIf(cadbl(vordre) > 0, vordre, cadbl(ettreballactual.WhatsThisHelpID))
    Set rst = dbtintes.OpenRecordset("select * from extensions_treballsrelacionats where numtreball=" + atrim(vnumtreballescullit) + " and numordremodificacio=" + atrim(vordreescullit) + " and coditinta=" + atrim(cadbl(etnomtinta.tag)))
    If rst.EOF Then
       rst.AddNew
       rst!codiextensio = vnumextensio
       rst!numtreball = vnumtreballescullit
       rst!numordremodificacio = vordreescullit
       rst!coditinta = cadbl(etnomtinta.tag)
       rst.Update
    End If
End Sub
Sub actualitzar_dadesextensioalstreballs(vnumextensio As String, vanilox As Double, vvolum As Double)
    Dim rst As Recordset
    If vvolum <= 0 Then Exit Sub
    Set rst = dbtintes.OpenRecordset("select * from extensions_treballsrelacionats where codiextensio='" + atrim(vnumextensio) + "'")
    While Not rst.EOF
      modificareltreball cadbl(rst!numtreball), cadbl(rst!numordremodificacio), cadbl(rst!coditinta), vanilox, vvolum
      rst.MoveNext
    Wend
    Set rst = Nothing
End Sub
Sub modificareltreball(vtreball As Double, vmodificacio As Double, vcoditinta As Double, vanilox As Double, vvolum As Double)
   'asseguro que hi hagi tots els valors abans de modificar el treball
   'l´asaú va dir que modifiques directament el treball
   If vtreball > 0 And vmodificacio > 0 And vcoditinta > 0 And vanilox > 0 And vvolum > 0 Then
      dbclixes.Execute "update tintes set volum=" + passaradecimalpunt(Trim(vvolum)) + ",anilox=" + passaradecimalpunt(Trim(vanilox)) + " where coditinta='" + atrim(vcoditinta) + "' and id_treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vmodificacio)
   End If
End Sub

Private Sub base_Change()

End Sub

Private Sub botoafegirrelacio_Click()
  Dim vtreball As Double
  Dim vordre As Double
  Dim vcoditinta As String
  Dim rst As Recordset
  vtreball = cadbl(InputBox("Entra el numero de treball que vols relacionar.", "Treball"))
  If vtreball = 0 Then Exit Sub
  vordre = cadbl(InputBox("Entra el numero de versió que vols relacionar.", "Versió"))
  If vordre = 0 Then Exit Sub
  vcoditinta = atrim(cadbl(etnomtinta.tag))
  Set rst = dbclixes.OpenRecordset("select coditinta from tintes where coditinta='" + atrim(vcoditinta) + "' and id_treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vordre))
  
  If rst.EOF Then
    MsgBox "Aquest treball no utilitza aquesta tinta.", vbCritical, "Error"
     Else
       guardar_treballsdelextensio etnumextensio.tag, vtreball, vordre
        carregar_extensio etnumextensio.tag
  End If
  Set rst = Nothing
End Sub

Private Sub botoeliminarrelacio_Click()
  Dim vtreball As Double
  Dim vversio As Double
  If llistatreballs.ListIndex = -1 Then MsgBox "Primer has d'escullir un treball de la llista.", vbCritical, "Atenció": GoTo fi
  If MsgBox("Segur que vols eliminar aquesta relació del treball amb l'extensió?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then GoTo fi
  vtreball = cadbl(Mid(llistatreballs, 1, InStr(1, llistatreballs, "/") - 1))
  vversio = cadbl(Mid(llistatreballs, InStr(1, llistatreballs, "/") + 1))
  dbtintes.Execute "delete * from extensions_treballsrelacionats where numtreball=" + atrim(vtreball) + " and numordremodificacio=" + atrim(vversio) + " and coditinta=" + atrim(cadbl(etnomtinta.tag))
  carregar_extensio etnumextensio.tag
  If llistatreballs.ListCount = 0 Then
     If MsgBox("No queda cap treball relacionat amb aquesta extensió." + Chr(10) + "VOLS ELIMINAR L'EXTENSIÓ?", vbInformation + vbYesNo, "Atenció") = vbYes Then
         dbtintes.Execute "delete * from extensions where codiextensio='" + atrim(etnumextensio.tag) + "'"
     End If
     carregar_extensio etnumextensio.tag
  End If
fi:
End Sub

Private Sub Command1_Click()
' Dim rst As Recordset
  
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim camp As TextObject
  Dim f  As OLEObject
  Dim rstf As Recordset
  Dim rstt As Recordset
  gravar_extensio_actual
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "etiqueta_extensio_a5.rpt", 1)
  oreport.FormulaFields.GetItemByName("numextensio").Text = "'" + atrim(etnumextensio.tag) + "'"
  oreport.FormulaFields.GetItemByName("nomtinta").Text = "'" + atrim(etnomtinta) + "'"
  oreport.FormulaFields.GetItemByName("marca").Text = "'" + atrim(etmarca) + "'"
  oreport.FormulaFields.GetItemByName("anilox").Text = "'" + atrim(comboanilox) + "'"
  oreport.FormulaFields.GetItemByName("base").Text = "'" + atrim(base) + "'"
  oreport.FormulaFields.GetItemByName("atenuant").Text = "'" + atrim(atenuant) + "'"
  oreport.FormulaFields.GetItemByName("disolvent").Text = "'" + atrim(disolvent) + "'"
  oreport.FormulaFields.GetItemByName("kilos").Text = "'" + atrim(kilos) + "'"
  oreport.FormulaFields.GetItemByName("volum").Text = "'" + atrim(cvolum) + "'"
  oreport.FormulaFields.GetItemByName("numquadrat").Text = "'" + atrim(cnumquadrat) + "'"
  oreport.FormulaFields.GetItemByName("observacions").Text = "'" + treure_apostruf(observacions) + "'"
  oreport.FormulaFields.GetItemByName("treballs").Text = "'" + atrim(concatenartreballs) + "'"
  oreport.DiscardSavedData
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
  Set rstt = Nothing
  Set rstf = Nothing
End Sub
Function concatenartreballs() As String
  Dim i As Byte
  For i = 0 To llistatreballs.ListCount - 1
     concatenartreballs = concatenartreballs + " " + llistatreballs.List(i)
  Next i
  If concatenartreballs <> "" Then concatenartreballs = "NT: " + concatenartreballs
End Function
Private Sub Command2_Click()
   gravar_extensio_actual
   
   Unload Me
End Sub
Function triar_extensiosemblant() As String
  Load formseleccio
  formseleccio.caption = "Selecciona una extensió"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "SELECT extensions.codiextensio,EXTENSIONS.marca, extensions.anilox, extensions.base, extensions.atenuant, extensions.disolvent, extensions.observacions FROM extensions where coditinta=" + atrim(cadbl(etnomtinta.tag))
  formseleccio.refrescar
'  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(0).width = 1000
  formseleccio.DBGrid2.Columns(1).width = 2000
  formseleccio.DBGrid2.Columns(2).width = 500
  formseleccio.DBGrid2.Columns(3).width = 500
  formseleccio.DBGrid2.Columns(4).width = 500
  formseleccio.DBGrid2.Columns(5).width = 500
  formseleccio.DBGrid2.Columns(6).width = 4500
  formseleccio.cmissatge.tag = "1"
  formseleccio.Show 1
  If seleccioret = 1 Then
     triar_extensiosemblant = formseleccio.Data1.Recordset!codiextensio
  End If
  Unload formseleccio
    
End Function
Private Sub Command52_Click()
  Dim vnumext As String
  vnumext = triar_extensiosemblant
  If vnumext = "" Then GoTo fi
  If MsgBox("Segur que vols assignar aquesta extensió a aquesta tinta d'aquest treball?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo fi
  guardar_treballsdelextensio vnumext
  etnumextensio.tag = vnumext
  carregar_extensio etnumextensio.tag
fi:
End Sub

