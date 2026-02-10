VERSION 5.00
Begin VB.Form formCHAT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chat"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16755
   Icon            =   "formCHAT.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   16755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bseloperari 
      BackColor       =   &H00F1B75F&
      Caption         =   "Operaris"
      Height          =   420
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   180
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Frame Frame2 
      Height          =   9150
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   5595
      Begin VB.CommandButton bborrarconversa 
         Height          =   360
         Left            =   60
         Picture         =   "formCHAT.frx":048A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Eliminar tota la conversa seleccionada."
         Top             =   8370
         Width           =   360
      End
      Begin VB.CommandButton Command1 
         Height          =   540
         Left            =   4740
         Picture         =   "formCHAT.frx":0576
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Afegir conversa"
         Top             =   8400
         Width           =   660
      End
      Begin VB.ListBox llistadeconverses 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   8040
         Left            =   105
         MultiSelect     =   1  'Simple
         TabIndex        =   2
         Top             =   285
         Width           =   5385
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chat de la conversa"
      Height          =   9150
      Left            =   5700
      TabIndex        =   0
      Top             =   660
      Width           =   11010
      Begin VB.CommandButton bborrarmissatge 
         Height          =   390
         Left            =   135
         Picture         =   "formCHAT.frx":069F
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Eliminar el missatge seleccionat."
         Top             =   8415
         Width           =   405
      End
      Begin VB.CommandButton benvio 
         Height          =   540
         Left            =   10215
         Picture         =   "formCHAT.frx":078B
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8415
         Width           =   660
      End
      Begin VB.TextBox vmissatge 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   570
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   8385
         Width           =   9600
      End
      Begin VB.ListBox llistachat 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   7980
         ItemData        =   "formCHAT.frx":0800
         Left            =   180
         List            =   "formCHAT.frx":0802
         TabIndex        =   3
         Top             =   360
         Width           =   10785
      End
   End
   Begin VB.Label etoperari 
      Caption         =   "---------------------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   855
      TabIndex        =   7
      Top             =   150
      Width           =   12600
   End
End
Attribute VB_Name = "formCHAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vseccio As String
Dim voperari As Integer

Dim vremitent As String

Private Sub bborrarconversa_Click()
  Dim vdata As Date
   Dim rst As Recordset
   Dim vnummissatge As Long
   Dim i As Long
   vdata = trobar_data_conversa_seleccionat
   If IsDate(vdata) And vdata <> "0:00:00" Then
        vnummissatge = llistadeconverses.ItemData(llistadeconverses.ListIndex)
        Set rst = dbmissatges.OpenRecordset("select * from converses_assumpte where idmissatge=" + atrim(vnummissatge))
        If rst.EOF Then GoTo fi
        If rst!remitent <> vremitent Then MsgBox "AQUESTA CONVERSA NO L'HAS INICIAT TU, NO POTS ELIMINAR-LA.", vbCritical, "ERROR": GoTo fi
        If UCase(InputBox("ESTAS A PUNT D'ELIMINAR TOTA AQUESTA CONVERSA I EL CHAT RELACIONAT." + vbNewLine + "ESCRIU [ESBORRAR] PER CONFIRMAR QUE ESTAS ELIMINANT-LO.", "ESBORRAR CONVERSA")) = "ESBORRAR" Then
          i = llistadeconverses.ListIndex
          While atrim(llistadeconverses.List(i)) <> ""
               llistadeconverses.RemoveItem i
          Wend
          dbmissatges.Execute "delete * from converses_chat where idconversa=" + atrim(vnummissatge)
          dbmissatges.Execute "delete * from converses_assumpte where idmissatge=" + atrim(vnummissatge)
          carregar_missatges_operari "I", cadbl(etoperari.tag)
        End If
   End If
fi:
   Set rst = Nothing
End Sub

Private Sub bborrarmissatge_Click()
   Dim vdata As Date
   Dim i As Long
   Dim rst As Recordset
   vdata = trobar_data_chat_seleccionat
   If IsDate(vdata) Then
          Set rst = dbmissatges.OpenRecordset("select * from converses_assumpte where idmissatge=" + atrim(cadbl(llistadeconverses.ItemData(llistadeconverses.ListIndex))))
          If Not rst.EOF Then If rst!datalectura > vdata Then MsgBox "AQUEST MISSATGE NO ES POT BORRAR PERQUÈ JA ESTÀ LLEGIT PEL DESTINATARI.", vbCritical, "MISSATGE LLEGIT": GoTo fi
          i = llistachat.ListIndex
          If Mid(atrim(llistachat.List(i)), 1, 2) <> "[" + vremitent Then MsgBox "No pots borrar un missatge que no ès teu.", vbCritical, "Error": GoTo fi
          While atrim(llistachat.List(i)) <> ""
               llistachat.RemoveItem i
          Wend
          dbmissatges.Execute "delete * from converses_chat where idconversa=" + atrim(cadbl(llistadeconverses.ItemData(llistadeconverses.ListIndex))) + " and data=#" + Format(vdata, "mm/dd/yy hh:nn:ss") + "#"
   End If
fi:
  Set rst = Nothing
End Sub

Function trobar_data_conversa_seleccionat() As Date
    Dim i As Long
    Dim v As String
    Dim vpos As Long
    Dim vposfi As Long
    i = llistadeconverses.ListIndex
    If Trim(llistadeconverses.List(i)) = "" Then Exit Function
    i = i - 1
    If i < 1 Then i = 1
    While atrim(llistadeconverses.List(i)) <> ""
       i = i - 1
    Wend
    If atrim(llistadeconverses.List(i)) = "" Then
        i = i + 1
        vpos = InStr(1, atrim(llistadeconverses.List(i)), "]")
        vposfi = InStr(vpos + 1, atrim(llistadeconverses.List(i)), "[")
        If vposfi > vpos Then
            v = Mid(atrim(llistadeconverses.List(i)), vpos + 1, vposfi - vpos - 1)
            If IsDate(v) Then
                trobar_data_conversa_seleccionat = v
                llistadeconverses.ListIndex = i
            End If
        End If
    End If
End Function

Function trobar_data_chat_seleccionat() As Date
    Dim i As Long
    Dim v As String
    Dim vpos As Long
    i = llistachat.ListIndex
    If Trim(llistachat.List(i)) = "" Then Exit Function
    i = i - 1
    While atrim(llistachat.List(i)) <> ""
       i = i - 1
    Wend
    If atrim(llistachat.List(i)) = "" Then
        i = i + 1
        vpos = InStr(1, atrim(llistachat.List(i)), "]")
        v = Mid(atrim(llistachat.List(i)), vpos + 1)
        If IsDate(v) Then
            trobar_data_chat_seleccionat = v
            llistachat.ListIndex = i
        End If
    End If
End Function
Private Sub benvio_Click()
    If atrim(vmissatge) = "" Or voperari = 0 Then Exit Sub
    possar_conversa_alchat vremitent, atrim(vmissatge), Now
    grava_conversa_delchat vremitent, atrim(vmissatge), Now, llistadeconverses.ItemData(llistadeconverses.ListIndex)
    vmissatge = ""
    vmissatge.SetFocus
End Sub
Sub grava_conversa_delchat(vremitent As String, vmissatge As String, vdata As Date, vidmissatge As Long)
   Dim vvalues As String
   vmissatge = substituir(vmissatge, "'", "´")
   vvalues = "(" + atrim(vidmissatge) + ",'" + atrim(vmissatge) + "',#" + Format(vdata, "mm/dd/yy hh:nn:ss") + "#,'" + atrim(vremitent) + "')"
   dbmissatges.Execute "insert into converses_chat (idconversa,missatge,data,remitent) values " + vvalues
   dbmissatges.Execute "update converses_assumpte set dataultimcanvi=now,datalectura=null,operariultimcanvi='" + vremitent + "' where idmissatge=" + atrim(vidmissatge)
End Sub
Sub carregar_missatges_operari(vsec As String, vop As Integer)
  Dim rst As Recordset
  llistadeconverses.Clear
  llistachat.Clear
  vseccio = vsec
  voperari = vop
  vremitent = IIf(voperari = 0, "E", "T")
  Set rst = dbtmp.OpenRecordset("select * from operaris where maquina='" + vseccio + "' and codi=" + atrim(voperari))
  If Not rst.EOF Then etoperari = atrim(voperari) + "-" + atrim(rst!descripcio)
  Set rst = dbmissatges.OpenRecordset("select * from converses_assumpte where operari=" + atrim(vop) + " and seccio='" + atrim(vsec) + "' order by dataultimcanvi desc")
  
  While Not rst.EOF
     possar_conversa_llistadeconverses rst!remitent, atrim(rst!missatge), rst!idmissatge, rst!Data, rst!datalectura
     llistadeconverses.AddItem " "
     rst.MoveNext
  Wend
  Set rst = Nothing
End Sub

Private Sub bseloperari_Click()
  Dim vsec As String
  Dim rst As Recordset
  Dim vordre As Byte
  Dim vafegit As Boolean
  Dim rst2 As Recordset
  vsec = "I"
  dbmissatges.Execute "delete * from operaris_ordreCHAT"
  Set rst = dbmissatges.OpenRecordset("select operari,datalectura,operariultimcanvi from converses_assumpte where seccio='" + atrim(vsec) + "' and datalectura=null order by dataultimcanvi")
  vordre = 1
AFEGIR_REGISTRES:
  If rst.EOF Then GoTo afegir_PROXIMS
  While Not rst.NoMatch
   Set rst2 = dbmissatges.OpenRecordset("select * from operaris_ordrechat where  codi=" + atrim(rst!operari))
   If rst2.EOF Then
    dbmissatges.Execute "insert into operaris_ordrechat select * from operaris where maquina='" + vsec + "' and codi=" + atrim(rst!operari)
    dbmissatges.Execute "update operaris_ordrechat set ordre=" + atrim(vordre) + " where codi=" + atrim(rst!operari)
   End If
   If IsNull(rst!datalectura) Then
        If Trim(rst!operariultimcanvi) <> vremitent Then
             dbmissatges.Execute "update operaris_ordrechat set descripcio='['+[descripcio]+']' where codi=" + atrim(rst!operari)
               Else: dbmissatges.Execute "update operaris_ordrechat set descripcio='*'+[descripcio]+'*' where codi=" + atrim(rst!operari)
        End If
   End If
   vordre = vordre + 1
   rst.FindNext "operari<>" + atrim(rst!operari)
  Wend
afegir_PROXIMS:
  If Not vafegit Then
        Set rst = dbmissatges.OpenRecordset("select operari,datalectura,operariultimcanvi from converses_assumpte where seccio='" + atrim(vsec) + "' and datalectura<>null order by datalectura desc")
        vafegit = True
        GoTo AFEGIR_REGISTRES
  End If

  dbmissatges.Execute "insert into Operaris_OrdreCHAT select * from operaris where codi not in(select codi from operaris_ordrechat) and maquina='" + vsec + "' and actiu<>0 and codi>0"
  dbmissatges.Execute "update operaris_ordrechat set ordre=99 where ordre=0"
  
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "avisosincidencies.mdb"
  formseleccio.Data1.RecordSource = "select codi,descripcio from operaris_ordreCHAT where maquina='" + vsec + "' and actiu<>0 order by ordre,DESCRIPCIO "
  formseleccio.caption = "Selecció d'Operari"
  formseleccio.refrescar
   formseleccio.Height = form1.Height
   formseleccio.Top = 0
   formseleccio.DBGrid2.Font.Size = 16
   formseleccio.DBGrid2.RowHeight = 440
   formseleccio.DBGrid2.Height = form1.Height - formseleccio.DBGrid2.Top - 600
  formseleccio.Show 1
  If seleccioret = 1 Then
   etoperari.tag = cadbl(formseleccio.Data1.Recordset!codi)
   etoperari = etoperari.tag + "-" + atrim(formseleccio.Data1.Recordset!descripcio)
   carregar_missatges_operari vsec, etoperari.tag
   vremitent = "E"
  End If
End Sub

Private Sub Command1_Click()
   Dim vidmissatge As Long
   Dim vtitolconversa As String
   If voperari = 0 Then Exit Sub
   vtitolconversa = atrim(InputBox("Escriu el motiu de la conversa.", "Conversa"))
   If atrim(vtitolconversa) = "" Then Exit Sub
   vidmissatge = grava_conversa_assumpte(vremitent, vseccio, voperari, vtitolconversa, Now)
   possar_conversa_llistadeconverses vremitent, vtitolconversa, vidmissatge, Now, Now
   
End Sub
Function grava_conversa_assumpte(vremitent As String, vseccio As String, voperari As Integer, vmissatge As String, vdata As Date) As Long
   Dim vvalues As String
   Dim rst As Recordset
   vmissatge = substituir(vmissatge, "'", "´")
   vvalues = "('" + atrim(vseccio) + "'," + atrim(voperari) + ",#" + Format(vdata, "mm/dd/yy hh:nn:ss") + "#,'" + atrim(vmissatge) + "','" + vremitent + "', null)" '#" + Format(vdata, "mm/dd/yy hh:nn:ss") + "#')"
   dbmissatges.Execute "insert into converses_assumpte (seccio,operari,data,missatge,remitent,datalectura) values " + vvalues
   Set rst = dbmissatges.OpenRecordset("select * from converses_assumpte where operari=" + atrim(voperari) + " and seccio='" + vseccio + "' and data=#" + Format(vdata, "mm/dd/yy hh:mm:ss") + "#")
   If Not rst.EOF Then grava_conversa_assumpte = rst!idmissatge
   Set rst = Nothing
   
End Function

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
  If vremitent = "E" Then bseloperari.visible = True Else bseloperari.visible = False
End Sub

Private Sub Form_Load()
  
  
  
'  possar_conversa_alchat "E", "Te escribimos para recordarte los próximos cambios en la API de Vault. Si no usas la API de Vault ni ninguna herramienta de terceros basada en dicha API, puedes dejar de leer este mensaje. De no ser así, es importante que estés al tanto de algunos cambios que van a tener lugar en la API de Vault.", Now
'  possar_conversa_alchat "T", "Te escribimos para recordarte los próximos cambios en la API de Vault. Si no usas la API de Vault ni ninguna herramienta de terceros basada en dicha API, puedes dejar de leer este mensaje. De no ser así, es importante que estés al tanto de algunos cambios que van a tener lugar en la API de Vault.", Now
End Sub
Sub possar_conversa_llistadeconverses(vremitent As String, vmissatge As String, vidmissatge As Long, vdata As Date, vdatalectura As Variant)
  Dim vespaisllista As Integer
  Dim vtalls As String
  Dim vpos As Double
  Dim vpendent As String
  Dim vlistindex As Long
  vespaisllista = (llistadeconverses.width / (llistadeconverses.FontSize * Screen.TwipsPerPixelX)) * 1.3
  vpos = 1
  vpendent = IIf(IsNull(vdatalectura), "[NoL] ", "[L] ")
  llistadeconverses.AddItem vpendent + justificar(Format(vdata, "dd/mm/yy hh:nn:ss") + " " + IIf(vremitent = "E", "[Encarregat]", "[Operari]"), vespaisllista, "E")
  llistadeconverses.ItemData(llistadeconverses.NewIndex) = vidmissatge
  vlistindex = llistadeconverses.NewIndex
  Do
        vtalls = Mid(vmissatge, vpos, vespaisllista - 5)
        If InStr(1, vtalls, Chr(10)) > 0 Then
             vtalls = Mid(vtalls, 1, InStr(1, vtalls, Chr(10)))
             vpos = vpos + InStr(1, vtalls, Chr(10))
              Else
                vpos = vpos + (vespaisllista - 5)
        End If
        If vtalls <> "" Then
             llistadeconverses.AddItem justificar(atrim(vtalls), vespaisllista, "E")
             llistadeconverses.ItemData(llistadeconverses.NewIndex) = vidmissatge
        End If
  Loop Until vtalls = ""
  llistadeconverses.ListIndex = vlistindex
  llistadeconverses_Click
End Sub
Function justificar(v As String, longitut As Integer, DoE As String) As String
    v = Mid(v, 1, longitut)
    If DoE = "E" Then
       v = v + Space(longitut - Len(v))
      Else: v = Space(longitut - Len(v)) + v
    End If
    justificar = v
End Function

Sub possar_conversa_alchat(vTipus As String, vmissatge As String, vdata As Date)
  Dim vespaisllista As Integer
  Dim vtalls As String
  Dim vpos As Double
  vespaisllista = (llistachat.width / (llistachat.FontSize * Screen.TwipsPerPixelX)) * 1.21
  vpos = 1
  If vTipus = "E" Then
      llistachat.AddItem justificar(" ", vespaisllista, "E")
      llistachat.AddItem justificar("[ENCARREGAT] " + Format(vdata, "dd/mm/yy hh:nn:ss"), vespaisllista, "E")
      Do
        vtalls = Mid(vmissatge, vpos, vespaisllista - 5)
        If InStr(1, vtalls, Chr(10)) > 0 Then
             vtalls = Mid(vtalls, 1, InStr(1, vtalls, Chr(10)))
             vpos = vpos + InStr(1, vtalls, Chr(10))
              Else
                vpos = vpos + (vespaisllista - 5)
        End If
        If vtalls <> "" Then
              llistachat.AddItem justificar(atrim(vtalls), vespaisllista, "E")
        End If
      Loop Until vtalls = ""
  End If
  If vTipus = "T" Then
      llistachat.AddItem justificar(" ", vespaisllista, "D")
      llistachat.AddItem justificar("[Tu] " + Format(vdata, "dd/mm/yy hh:nn:ss"), vespaisllista, "D")
      Do
        vtalls = Mid(vmissatge, vpos, vespaisllista - 5)
        If InStr(1, vtalls, Chr(10)) > 0 Then
             vtalls = Mid(vtalls, 1, InStr(1, vtalls, Chr(10)))
             vpos = vpos + InStr(1, vtalls, Chr(10))
              Else
                vpos = vpos + (vespaisllista - 5)
        End If
        If vtalls <> "" Then
           llistachat.AddItem justificar(atrim(vtalls), vespaisllista, "D")
        End If
      Loop Until vtalls = ""
      
  End If
End Sub

Private Sub llistadeconverses_Click()
  Dim i As Integer
  Dim vidmissatge As Long
  Dim vprimeralinia As Byte
  Static vsocdins As Boolean
  If vsocdins Then Exit Sub
  vsocdins = True
  llistachat.Clear
  For i = 0 To llistadeconverses.ListCount - 1: llistadeconverses.Selected(i) = False: Next i
  i = llistadeconverses.ListIndex
  vidmissatge = llistadeconverses.ItemData(i)
  While llistadeconverses.ItemData(i) = vidmissatge And i > 0
    i = i - 1
  Wend
  If i > 0 Then i = i + 1
  While llistadeconverses.ItemData(i) = vidmissatge
    llistadeconverses.Selected(i) = True
    If Mid(llistadeconverses.List(i), 1, 1) = "[" Then vprimeralinia = i
    i = i + 1
    If i = llistadeconverses.ListCount Then GoTo cont
  Wend
  
cont:
   llistadeconverses.ListIndex = vprimeralinia
  carregar_chatconversa vidmissatge
  possar_conversa_a_llegida vidmissatge

  vsocdins = False
  
End Sub
Sub possar_conversa_a_llegida(vidmissatge As Long)
   Dim rst As Recordset
   Set rst = dbmissatges.OpenRecordset("select * from converses_assumpte where idmissatge=" + atrim(vidmissatge))
   If rst.EOF Then GoTo fi
   If rst!operariultimcanvi <> vremitent Then
       dbmissatges.Execute "update converses_assumpte set datalectura=now where idmissatge=" + atrim(vidmissatge) + " and datalectura=null"
       llistadeconverses.List(llistadeconverses.ListIndex) = "[L]" + Mid(llistadeconverses.List(llistadeconverses.ListIndex), InStr(1, llistadeconverses.List(llistadeconverses.ListIndex), " "))
   End If
fi:
   Set rst = Nothing
End Sub
Sub carregar_chatconversa(vidmissatge As Long)
   Dim rst As Recordset
   llistachat.Clear
   Set rst = dbmissatges.OpenRecordset("select * from converses_chat where idconversa=" + atrim(vidmissatge) + " order by data")
   While Not rst.EOF
      possar_conversa_alchat rst!remitent, atrim(rst!missatge), rst!Data
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub vmissatge_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("""") Then KeyAscii = 0
End Sub
