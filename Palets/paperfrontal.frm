VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{8D1418DD-FB6E-4C6F-A1DC-13E914E39989}#1.0#0"; "TBarCode11.ocx"
Begin VB.Form paperfrontal 
   Caption         =   "Paper Frontal"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4740
   Icon            =   "paperfrontal.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   4515
      Begin Crystal.CrystalReport llistat 
         Left            =   2085
         Top             =   1020
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Command2 
         Height          =   465
         Left            =   3090
         Picture         =   "paperfrontal.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprimir Paper Palet"
         Top             =   2820
         Width           =   945
      End
      Begin VB.ListBox llista 
         Height          =   1815
         Left            =   105
         TabIndex        =   4
         Top             =   1695
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2160
         Picture         =   "paperfrontal.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   870
      End
      Begin VB.TextBox comanda 
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label etcopies 
         BackStyle       =   0  'Transparent
         Caption         =   "2 Copies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   3195
         TabIndex        =   14
         Top             =   3285
         Width           =   825
      End
      Begin VB.Label estiletiqueta 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   690
         TabIndex        =   12
         Top             =   1170
         Width           =   3060
      End
      Begin VB.Label refclient 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   675
         TabIndex        =   9
         Top             =   915
         Width           =   3675
      End
      Begin VB.Label nomclient 
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
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   675
         TabIndex        =   8
         Top             =   660
         Width           =   3600
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   300
         Left            =   150
         TabIndex        =   7
         Top             =   660
         Width           =   525
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Palets Disponibles"
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   1485
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Comanda:"
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   900
      End
   End
   Begin TBarCode11LibCtl.TBarCode11 codidebarres 
      Height          =   690
      Left            =   5220
      TabIndex        =   13
      Top             =   3015
      Visible         =   0   'False
      Width           =   1965
      _cx             =   3466
      _cy             =   1217
      BackColor       =   15189445
      BackStyle       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   ""
      TextAlignment   =   0
      BarCode         =   20
      CDMethod        =   1
      CountCheckDigits=   0
      EscapeSequences =   0   'False
      Format          =   ""
      BearerBarWidth  =   -1
      BearerBarType   =   0
      ModuleWidth     =   "339"
      Orientation     =   0
      PrintDataText   =   0   'False
      PrintTextAbove  =   0   'False
      Ratio           =   ""
      RatioHint       =   "1B:2B:3B:4B:1S:2S:3S:4S"
      RatioDefault    =   "1:2:3:4:1:2:3:4"
      TextColor       =   0
      LastError       =   "Error: No barcode data!"
      LastErrorNo     =   -2147020590
      MustFit         =   0   'False
      TextDistance    =   0
      NotchHeight     =   -1
      CountModules    =   0
      DrawStatus      =   0
      SuppressErrorMsg=   0   'False
      CountRows       =   0
      EncodingMode    =   0
      OptResolution   =   0   'False
      DisplayText     =   ""
      BarWidthReduction=   0
      BarWidthReductionUnit=   0
      Quality         =   0
      CompositeComponent=   0
      RSS_SegmPerRow  =   -1
      TrimSpaces      =   0
      DefaultSet      =   0
      QuietZoneUnit   =   0
      QuietZoneLeft   =   0
      QuietZoneRight  =   0
      QuietZoneTop    =   0
      QuietZoneBottom =   0
      DefaultColorForQuietZoneLeft=   -1  'True
      DefaultColorForQuietZoneRight=   -1  'True
      DefaultColorForQuietZoneTop=   -1  'True
      DefaultColorForQuietZoneBottom=   -1  'True
      QuietZoneColorLeft=   16777215
      QuietZoneColorRight=   16777215
      QuietZoneColorTop=   16777215
      QuietZoneColorBottom=   16777215
      Compression     =   0
      SizeMode        =   0
      Dpi             =   600
      Decoder         =   1
      DrawMode        =   3
      CodePage        =   1
      CodePageCustom  =   0
      PropertyInternal=   ""
      MaximumTextIndex=   5
      ActiveTextIndex =   0
      TextPositionLeft=   0
      TextPositionTop =   0
      TextBlockWidth  =   0
      TextBlockHeight =   0
      TextClipping    =   -1  'True
      WordWrappingEnabled=   -1  'True
      TextRotation    =   0
      BarShape        =   0
      BarShapeImageFile=   ""
      Options         =   ""
      CBF_Rows        =   -1
      CBF_Columns     =   -1
      CBF_RowHeight   =   -1
      CBF_RowSeparatorHeight=   -1
      CBF_Format      =   0
      DM_Size         =   0
      DM_Rectangular  =   0   'False
      DM_Format       =   0
      DM_EnforceBinary=   0   'False
      DM_AppendIndex  =   -1
      DM_AppendCount  =   -1
      DM_AppendFileID =   -1
      Aztec_Size      =   0
      Aztec_EnforceBinary=   0   'False
      Aztec_ErrorCorrection=   -1
      Aztec_Runes     =   0   'False
      Aztec_Format    =   0
      Aztec_FormatSpecifier=   ""
      Aztec_AppendActive=   0   'False
      Aztec_AppendIndex=   65
      Aztec_AppendTotal=   65
      Aztec_AppendMessageID=   ""
      DotCode_SizeMode=   -1
      DotCode_Size    =   ""
      DotCode_PrintDirection=   0
      DotCode_Format  =   0
      DotCode_FormatSpecifier=   ""
      DotCode_EnforceBinary=   0   'False
      DotCode_Mask    =   -1
      DotCode_AppendActive=   0   'False
      DotCode_AppendIndex=   1
      DotCode_AppendTotal=   1
      HanXin_Size     =   0
      HanXin_EnforceBinary=   0   'False
      HanXin_ECLevel  =   0
      HanXin_Mask     =   -1
      MAXI_Mode       =   4
      MAXI_AppendIndex=   -1
      MAXI_AppendCount=   -1
      MAXI_Undercut   =   -1
      MAXI_Preamble   =   0   'False
      MAXI_PostalCode =   ""
      MAXI_CountryCode=   ""
      MAXI_ServiceClass=   ""
      MAXI_Date       =   "96"
      PDF417_Rows     =   -1
      PDF417_Columns  =   -1
      PDF417_ECLevel  =   -1
      PDF417_EncodationMode=   0
      PDF417_RowHeight=   -1
      PDF417_FileName =   ""
      PDF417_SegmentCount=   -1
      PDF417_TimeStamp=   -1
      PDF417_Sender   =   ""
      PDF417_Addressee=   ""
      PDF417_FileSize =   -1
      PDF417_CheckSum =   -1
      PDF417_RatioRowCol=   ""
      PDF417_SegmentIndex=   -1
      PDF417_FileID   =   ""
      PDF417_LastSegment=   0   'False
      MicroPDF_Mode   =   0
      MicroPDF_Version=   0
      QR_Version      =   0
      MQR_Version     =   0
      QR_Format       =   0
      QR_FmtAppIndicator=   ""
      QR_ECLevel      =   1
      QR_Mask         =   -1
      MQR_Mask        =   -1
      QR_AppendIndex  =   -1
      QR_AppendCount  =   -1
      QR_AppendParity =   -1
      QR_KanjiChineseCompaction=   -1
   End
   Begin TBarCode11LibCtl.TBarCode11 codidebarres2 
      Height          =   960
      Left            =   5145
      TabIndex        =   11
      Top             =   2190
      Width           =   2655
      _cx             =   4683
      _cy             =   1693
      BackColor       =   14149612
      BackStyle       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   "18004023200000004"
      TextAlignment   =   3
      BarCode         =   16
      CDMethod        =   2
      CountCheckDigits=   1
      EscapeSequences =   -1  'True
      Format          =   ""
      BearerBarWidth  =   -1
      BearerBarType   =   0
      ModuleWidth     =   "339"
      Orientation     =   0
      PrintDataText   =   -1  'True
      PrintTextAbove  =   0   'False
      Ratio           =   ""
      RatioHint       =   "1B:2B:3B:4B:1S:2S:3S:4S"
      RatioDefault    =   "1:2:3:4:1:2:3:4"
      TextColor       =   0
      LastError       =   "La operación se completó correctamente. "
      LastErrorNo     =   0
      MustFit         =   0   'False
      TextDistance    =   0
      NotchHeight     =   -1
      CountModules    =   145
      DrawStatus      =   0
      SuppressErrorMsg=   0   'False
      CountRows       =   1
      EncodingMode    =   0
      OptResolution   =   0   'False
      DisplayText     =   ""
      BarWidthReduction=   0
      BarWidthReductionUnit=   0
      Quality         =   65535
      CompositeComponent=   0
      RSS_SegmPerRow  =   -1
      TrimSpaces      =   0
      DefaultSet      =   0
      QuietZoneUnit   =   0
      QuietZoneLeft   =   0
      QuietZoneRight  =   0
      QuietZoneTop    =   0
      QuietZoneBottom =   0
      DefaultColorForQuietZoneLeft=   -1  'True
      DefaultColorForQuietZoneRight=   -1  'True
      DefaultColorForQuietZoneTop=   -1  'True
      DefaultColorForQuietZoneBottom=   -1  'True
      QuietZoneColorLeft=   16777215
      QuietZoneColorRight=   16777215
      QuietZoneColorTop=   16777215
      QuietZoneColorBottom=   16777215
      Compression     =   0
      SizeMode        =   0
      Dpi             =   600
      Decoder         =   1
      DrawMode        =   0
      CodePage        =   1
      CodePageCustom  =   0
      PropertyInternal=   ""
      MaximumTextIndex=   5
      ActiveTextIndex =   0
      TextPositionLeft=   0
      TextPositionTop =   0
      TextBlockWidth  =   0
      TextBlockHeight =   0
      TextClipping    =   -1  'True
      WordWrappingEnabled=   -1  'True
      TextRotation    =   0
      BarShape        =   0
      BarShapeImageFile=   ""
      Options         =   ""
      CBF_Rows        =   -1
      CBF_Columns     =   -1
      CBF_RowHeight   =   -1
      CBF_RowSeparatorHeight=   -1
      CBF_Format      =   0
      DM_Size         =   0
      DM_Rectangular  =   0   'False
      DM_Format       =   0
      DM_EnforceBinary=   0   'False
      DM_AppendIndex  =   -1
      DM_AppendCount  =   -1
      DM_AppendFileID =   -1
      Aztec_Size      =   0
      Aztec_EnforceBinary=   0   'False
      Aztec_ErrorCorrection=   -1
      Aztec_Runes     =   0   'False
      Aztec_Format    =   0
      Aztec_FormatSpecifier=   ""
      Aztec_AppendActive=   0   'False
      Aztec_AppendIndex=   65
      Aztec_AppendTotal=   65
      Aztec_AppendMessageID=   ""
      DotCode_SizeMode=   -1
      DotCode_Size    =   ""
      DotCode_PrintDirection=   0
      DotCode_Format  =   0
      DotCode_FormatSpecifier=   ""
      DotCode_EnforceBinary=   0   'False
      DotCode_Mask    =   -1
      DotCode_AppendActive=   0   'False
      DotCode_AppendIndex=   1
      DotCode_AppendTotal=   1
      HanXin_Size     =   0
      HanXin_EnforceBinary=   0   'False
      HanXin_ECLevel  =   0
      HanXin_Mask     =   -1
      MAXI_Mode       =   4
      MAXI_AppendIndex=   -1
      MAXI_AppendCount=   -1
      MAXI_Undercut   =   -1
      MAXI_Preamble   =   0   'False
      MAXI_PostalCode =   ""
      MAXI_CountryCode=   ""
      MAXI_ServiceClass=   ""
      MAXI_Date       =   "96"
      PDF417_Rows     =   -1
      PDF417_Columns  =   -1
      PDF417_ECLevel  =   -1
      PDF417_EncodationMode=   0
      PDF417_RowHeight=   -1
      PDF417_FileName =   ""
      PDF417_SegmentCount=   -1
      PDF417_TimeStamp=   -1
      PDF417_Sender   =   ""
      PDF417_Addressee=   ""
      PDF417_FileSize =   -1
      PDF417_CheckSum =   -1
      PDF417_RatioRowCol=   ""
      PDF417_SegmentIndex=   -1
      PDF417_FileID   =   ""
      PDF417_LastSegment=   0   'False
      MicroPDF_Mode   =   0
      MicroPDF_Version=   0
      QR_Version      =   0
      MQR_Version     =   0
      QR_Format       =   0
      QR_FmtAppIndicator=   ""
      QR_ECLevel      =   1
      QR_Mask         =   -1
      MQR_Mask        =   -1
      QR_AppendIndex  =   -1
      QR_AppendCount  =   -1
      QR_AppendParity =   -1
      QR_KanjiChineseCompaction=   -1
   End
   Begin TBarCode11LibCtl.TBarCode11 codidebarres1 
      Height          =   1245
      Left            =   4725
      TabIndex        =   10
      Top             =   375
      Width           =   5265
      _cx             =   9287
      _cy             =   2196
      BackColor       =   14149612
      BackStyle       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   "693536400100"
      TextAlignment   =   3
      BarCode         =   13
      CDMethod        =   14
      CountCheckDigits=   1
      EscapeSequences =   -1  'True
      Format          =   ""
      BearerBarWidth  =   -1
      BearerBarType   =   0
      ModuleWidth     =   "339"
      Orientation     =   0
      PrintDataText   =   -1  'True
      PrintTextAbove  =   0   'False
      Ratio           =   ""
      RatioHint       =   "1B:2B:3B:4B:1S:2S:3S:4S"
      RatioDefault    =   "1:2:3:4:1:2:3:4"
      TextColor       =   0
      LastError       =   "La operación se completó correctamente. "
      LastErrorNo     =   0
      MustFit         =   0   'False
      TextDistance    =   0
      NotchHeight     =   -1
      CountModules    =   113
      DrawStatus      =   0
      SuppressErrorMsg=   0   'False
      CountRows       =   1
      EncodingMode    =   0
      OptResolution   =   0   'False
      DisplayText     =   ""
      BarWidthReduction=   0
      BarWidthReductionUnit=   0
      Quality         =   16973827
      CompositeComponent=   0
      RSS_SegmPerRow  =   -1
      TrimSpaces      =   0
      DefaultSet      =   0
      QuietZoneUnit   =   0
      QuietZoneLeft   =   0
      QuietZoneRight  =   0
      QuietZoneTop    =   0
      QuietZoneBottom =   0
      DefaultColorForQuietZoneLeft=   -1  'True
      DefaultColorForQuietZoneRight=   -1  'True
      DefaultColorForQuietZoneTop=   -1  'True
      DefaultColorForQuietZoneBottom=   -1  'True
      QuietZoneColorLeft=   16777215
      QuietZoneColorRight=   16777215
      QuietZoneColorTop=   16777215
      QuietZoneColorBottom=   16777215
      Compression     =   0
      SizeMode        =   0
      Dpi             =   600
      Decoder         =   1
      DrawMode        =   0
      CodePage        =   1
      CodePageCustom  =   0
      PropertyInternal=   ""
      MaximumTextIndex=   5
      ActiveTextIndex =   0
      TextPositionLeft=   0
      TextPositionTop =   0
      TextBlockWidth  =   0
      TextBlockHeight =   0
      TextClipping    =   -1  'True
      WordWrappingEnabled=   -1  'True
      TextRotation    =   0
      BarShape        =   0
      BarShapeImageFile=   ""
      Options         =   ""
      CBF_Rows        =   -1
      CBF_Columns     =   -1
      CBF_RowHeight   =   -1
      CBF_RowSeparatorHeight=   -1
      CBF_Format      =   0
      DM_Size         =   0
      DM_Rectangular  =   0   'False
      DM_Format       =   0
      DM_EnforceBinary=   0   'False
      DM_AppendIndex  =   -1
      DM_AppendCount  =   -1
      DM_AppendFileID =   -1
      Aztec_Size      =   0
      Aztec_EnforceBinary=   0   'False
      Aztec_ErrorCorrection=   -1
      Aztec_Runes     =   0   'False
      Aztec_Format    =   0
      Aztec_FormatSpecifier=   ""
      Aztec_AppendActive=   0   'False
      Aztec_AppendIndex=   65
      Aztec_AppendTotal=   65
      Aztec_AppendMessageID=   ""
      DotCode_SizeMode=   -1
      DotCode_Size    =   ""
      DotCode_PrintDirection=   0
      DotCode_Format  =   0
      DotCode_FormatSpecifier=   ""
      DotCode_EnforceBinary=   0   'False
      DotCode_Mask    =   -1
      DotCode_AppendActive=   0   'False
      DotCode_AppendIndex=   1
      DotCode_AppendTotal=   1
      HanXin_Size     =   0
      HanXin_EnforceBinary=   0   'False
      HanXin_ECLevel  =   0
      HanXin_Mask     =   -1
      MAXI_Mode       =   4
      MAXI_AppendIndex=   -1
      MAXI_AppendCount=   -1
      MAXI_Undercut   =   -1
      MAXI_Preamble   =   0   'False
      MAXI_PostalCode =   ""
      MAXI_CountryCode=   ""
      MAXI_ServiceClass=   ""
      MAXI_Date       =   "96"
      PDF417_Rows     =   -1
      PDF417_Columns  =   -1
      PDF417_ECLevel  =   -1
      PDF417_EncodationMode=   0
      PDF417_RowHeight=   -1
      PDF417_FileName =   ""
      PDF417_SegmentCount=   -1
      PDF417_TimeStamp=   -1
      PDF417_Sender   =   ""
      PDF417_Addressee=   ""
      PDF417_FileSize =   -1
      PDF417_CheckSum =   -1
      PDF417_RatioRowCol=   ""
      PDF417_SegmentIndex=   -1
      PDF417_FileID   =   ""
      PDF417_LastSegment=   0   'False
      MicroPDF_Mode   =   0
      MicroPDF_Version=   0
      QR_Version      =   0
      MQR_Version     =   0
      QR_Format       =   0
      QR_FmtAppIndicator=   ""
      QR_ECLevel      =   1
      QR_Mask         =   -1
      MQR_Mask        =   -1
      QR_AppendIndex  =   -1
      QR_AppendCount  =   -1
      QR_AppendParity =   -1
      QR_KanjiChineseCompaction=   -1
   End
End
Attribute VB_Name = "paperfrontal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vpaletreferenciainplacsa As Boolean
Private Sub Command1_Click()
    carregar_elspalets
End Sub
Sub carregar_elspalets()
 Dim rstcli As Recordset
   Dim rstenvio As Recordset
   Dim numarticle As String
   Dim vrefdeclient As String
   If cadbl(comanda.Text) = 0 Then MsgBox "Primer apunta el numero de comanda. Espera que busqui els palets que hi han fets i despres escull el palet que vols que imprimeixi, tot seguit apreta el botó de impresora.": GoTo fi
   ratoli "espera"
   
   nomclient = ""
   refclient = ""
   nomclient.tag = ""
   refclient.tag = ""
   llista.tag = ""
   llista.Clear
   Set rstcli = dbtmpb.OpenRecordset("SELECT comandes.comanda ,comandes.refclient, comandes.refclientdeclient,comandes.client, clients.ultimcodiarticle ,clients.nom as nomclient FROM comandes INNER JOIN clients ON comandes.client = clients.codi WHERE comandes.comanda=" + atrim(cadbl(comanda)) + ";")
   If Not rstcli.EOF Then
     vrefdeclient = IIf(atrim(rstcli!refclientdeclient) <> "", atrim(rstcli!refclientdeclient), atrim(rstcli!refclient))
     If Len(vrefdeclient) < 2 Then MsgBox "Aquesta comanda no te Ref.Client entrada, abans s'ha d'entrar una referencia.", vbCritical + vbOKOnly, "Atenció": GoTo fi
     Set rstenvio = dbtmpb.OpenRecordset("SELECT comandes.comanda,clients_envios.nome, clients_envios.paletreferenciainplacsa,Clients_envios.estilfrontal,clients_envios.copiespaperfrontal FROM comandes INNER JOIN Clients_envios ON comandes.direnvio = Clients_envios.id WHERE (((comandes.comanda)=" + atrim(cadbl(comanda)) + "));")
     If Not rstenvio.EOF Then
        vpaletreferenciainplacsa = cabool(rstenvio!paletreferenciainplacsa)
        estiletiqueta = atrim(rstenvio!estilfrontal)
        If cadbl(rstenvio!copiespaperfrontal) > 1 Then
           etcopies.tag = cadbl(rstenvio!copiespaperfrontal): etcopies = atrim(rstenvio!copiespaperfrontal) + " Copies"
            Else: etcopies.tag = "1": etcopies = ""
        End If
     End If
     
     nomclient = atrim(rstenvio!nome) 'rstcli!nomclient + "(" + atrim(rstenvio!nome) + ")"
     nomclient.tag = rstcli!client
     'Set rsttmp = dbtmp.OpenRecordset("select codiarticle from papersfrontals where refclient='" + atrim(vrefdeclient) + "'")
     'If Not rsttmp.EOF Then
     '    numarticle = rsttmp!codiarticle
     '    refclient.tag = cadbl(numarticle) * -1
     '   Else:
     '        numarticle = cadbl(rstcli!ultimcodiarticle) + 1
     '        refclient.tag = numarticle
     'End If
     Set rsttmp = dbcomandes.OpenRecordset("select gtin14 from comandes_extres where comanda=" + atrim(cadbl(comanda)))
     If Not rsttmp.EOF Then numarticle = cadbl(rsttmp!gtin14): refclient.tag = atrim(numarticle)
     refclient = vrefdeclient + " -> CodiInplacsa: " + atrim(Format(numarticle, "00000"))
   End If
   DoEvents
   possarnumpalets
fi:
   Set rstcli = Nothing
   Set rstenvio = Nothing
   ratoli "normal"
End Sub
Sub possarnumpalets()
   Dim rstpal As Recordset
   Dim np As Double
   Dim rst As Recordset
   Dim vdatafabricacio As Date
   Dim vd As String
   Dim vnumalb As String
   'si no trobo bobines d'entrega buscaré bobines de rebobinadora
   vnumalb = cadbl(formvendes.datacapcalera.Recordset!numalbara)
   Set rstpal = dbbaixes.OpenRecordset("SELECT bobinesent.numpalet as palet,bobinesent.data as datafab FROM bobinesent WHERE ((comanda=" + atrim(cadbl(comanda)) + ") and numalbara=" + atrim(vnumalb) + ") order by numpalet;", , dbReadOnly)
   'Clipboard.Clear
   'Clipboard.SetText "SELECT bobinesent.numpalet as palet,bobinesent.data as datafab FROM bobinesent WHERE ((comanda=" + atrim(cadbl(comanda)) + ") and numalbara=" + atrim(vnumalb) + ") order by numpalet;"
   If rstpal.EOF Then
     Set rstpal = dbbaixes.OpenRecordset("SELECT bobinesreb.palet,bobinesreb.datafab,rebobinadores.datafi as datafabfi FROM bobinesreb INNER JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id WHERE (((rebobinadores.comanda)=" + atrim(cadbl(comanda)) + ")) order by palet;", , dbReadOnly)
     If Not rstpal.EOF Then
         vdatafabricacio = rstpal!datafabfi
          Else: vdatafabricacio = Now
     End If
       Else
         Set rst = dbbaixes.OpenRecordset("select max(datafi) as datafab from soldadores where comanda=" + atrim(comanda))
         If IsNull(rst!datafab) Then Set rst = dbbaixes.OpenRecordset("select max(datafi) as datafab from rebobinadores where comanda=" + atrim(comanda))
         If IsNull(rst!datafab) Then Set rst = dbbaixes.OpenRecordset("select max(datafi) as datafab from laminadores where comanda=" + atrim(comanda))
         If IsNull(rst!datafab) Then Set rst = dbbaixes.OpenRecordset("select max(datafi) as datafab from impressores where comanda=" + atrim(comanda))
         If Not rst.EOF Then
           If Not IsDate(rst!datafab) Then
             MsgBox "No hi ha la data de fabricació a soldadores", vbCritical, "Error"
             vd = InputBox("Entra la data de fabricació. dd/mm/yy", "Error")
             If Not IsDate(vd) Then
                Exit Sub
                 Else: vdatafabricacio = vd
             End If
            Else
             vdatafabricacio = atrim(rst!datafab)
           End If
         End If
           
   End If
   
   While Not rstpal.EOF
     llista.tag = vdatafabricacio
     If np <> rstpal!palet Then
        np = rstpal!palet
        llista.AddItem "Palet " + atrim(np)
        llista.ItemData(llista.NewIndex) = np
     End If
     rstpal.MoveNext
   Wend
   Set rstpal = Nothing
   Set rst = Nothing
End Sub

Private Sub Command2_Click()
   Dim rstpf As Recordset
   Dim rstb As Recordset
   Dim nomllistat As String
   If llista.ListIndex = -1 Then MsgBox "Primer escull un palet.", vbCritical, "Atenció": Exit Sub
   Set rstpf = dbtmp.OpenRecordset("select * from papersfrontals where numlotinplacsa=" + atrim(cadbl(comanda)) + " and numpalet=" + atrim(llista.ItemData(llista.ListIndex)))
   ratoli "espera"
   If rstpf.EOF Then
       crearnoupf rstpf
         Else
             'posso el rstpf a editar i aixií refrescare tots els valors del registre
             rstpf.Edit
             crearnoupf rstpf
   End If
   'imprimir pf
   Set rstpf = dbtmp.OpenRecordset("select * from papersfrontals where numlotinplacsa=" + atrim(cadbl(comanda)) + " and numpalet=" + atrim(llista.ItemData(llista.ListIndex)))
   If Not rstpf.EOF Then
     
     If Not existeix("c:\temp") Then MkDir "c:\temp"
     nomllistat = "frontalspaletsean13.rpt"
     If InStr(1, estiletiqueta, "128") > 0 Then
       codidebarres1.Text = rstpf!codidebarres1
       codidebarres2.Text = rstpf!codidebarres2
       codidebarres1.FontSize = 6
       If Len(rstpf!codidebarres2) > 25 Then
          codidebarres2.FontSize = 4
           Else: codidebarres2.FontSize = 12
       End If
       codidebarres1.EscapeSequences = True
       codidebarres2.EscapeSequences = True
       codidebarres1.BarCode = eBC_GS1_128 ' eBC_EAN128
       codidebarres2.BarCode = eBC_GS1_128 ' eBC_EAN128
       codidebarres1.CDMethod = eCDEAN128
       codidebarres2.CDMethod = eCDEAN128
       codidebarres2.CDMethod = eCDMod10
       'codidebarres1.CDMethod = eCDMod10
       codidebarres1.CDMethod = 0
       codidebarres2.CDMethod = 0
       codidebarres2.Format = ""
       codidebarres1.Format = ""
       
       
       'codidebarres2.Format = ""
       codidebarres1.SaveImage "c:\temp\codidebarres1", eIMBmp, 2300, 500, 600, 600
       If rstpf!codidebarres2 > 0 Then
          codidebarres2.SaveImage "c:\temp\codidebarres2", eIMBmp, 1200, 500, 600, 600
           Else: If existeix("c:\temp\codidebarres2.bmp") Then Kill "c:\temp\codidebarres2.bmp"
       End If
       If atrim(rstpf!formatetiqueta) = "ACTYS" Then
          codidebarres2.Format = ""
          codidebarres2.CDMethod = 0
          codidebarres2.SaveImage "c:\temp\codidebarres2", eIMBmp, 3000, 500, 600, 600
       End If
       nomllistat = "frontalspaletsean128" + atrim(rstpf!formatetiqueta) + "v9.rpt"
     End If
     If InStr(1, estiletiqueta, "EAN13") > 0 Then
       codidebarres1.Text = emplena12zeros(rstpf!codidebarres1)
       codidebarres2.Text = emplena12zeros(rstpf!codidebarres2)
       codidebarres1.BarCode = eBC_EAN13
       codidebarres2.BarCode = eBC_EAN13
       codidebarres1.CDMethod = eCDEAN13
       codidebarres2.CDMethod = eCDEAN13
       codidebarres1.Refresh
       codidebarres1.SaveImage "c:\temp\codidebarres1", eIMBmp, 2500, 500, 600, 600
       codidebarres2.SaveImage "c:\temp\codidebarres2", eIMBmp, 1200, 500, 600, 600
       nomllistat = "frontalspaletsean13.rpt"
     End If
     If InStr(1, estiletiqueta, "EAN-13") > 0 Then MsgBox "Aquesta etiqueta l'has de fer desde Papers Frontals genèrics.", vbInformation, "Atenció": ratoli "normal": Exit Sub
     '  codidebarres1.Text = emplena12zeros(rstpf!codidebarres1)
     '  codidebarres2.Text = emplena12zeros(rstpf!codidebarres2)
     '  codidebarres1.BarCode = eBC_EAN13
     '  codidebarres2.BarCode = eBC_EAN13
     '  codidebarres1.CDMethod = eCDEAN13
     '  codidebarres2.CDMethod = eCDEAN13
     '  codidebarres1.Refresh
     '  codidebarres1.SaveImage "c:\temp\codidebarres1", eIMBmp, 2500, 500, 600, 600
     '  codidebarres2.SaveImage "c:\temp\codidebarres2", eIMBmp, 1200, 500, 600, 600
     '  nomllistat = "frontalspaletsean-13CBarres.rpt"
     'End If
     
     rstpf.Edit
       copiafoto "c:\temp\codidebarres1.bmp", rstpf!codi1
       copiafoto "c:\temp\codidebarres2.bmp", rstpf!codi2
       'nomllistat = "frontalspaletsean13.rpt"
     rstpf.Update
     wait 3
     imprimir_frontal_v9 rstpf, nomllistat
     'If cadbl(refclient.tag) > 0 And InStr(1, estiletiqueta, "128") > 0 Then
     '  dbtmpb.Execute "update clients set ultimcodiarticle=" + atrim(cadbl(refclient.tag)) + " where codi=" + atrim(cadbl(nomclient.tag))
     'End If
     formvendes.datacapcalera.Database.Execute "update capcaleraalbara set papersfrontalsimpresos=True where numalbara=" + atrim(formvendes.datacapcalera.Recordset!numalbara)
     formvendes.possarcampscapçalera
   End If
   ratoli "normal"
End Sub
Function copiafoto(foto As String, fldTO As Field)

'This function takes the source field image and copies it
'into the destination field.
'The function first saves the image in the source field to a
'temp file on disc. Then reads this temp file into
'the destination field.
'The temp file is then deleted
'On Error Resume Next

Dim iFieldSize  As Long
Dim varChunk    As Variant
Dim baData()    As Byte
Dim iOffset     As Long
Dim sFName      As String
Dim iFileNum    As Long
Dim cnt         As Long
Dim z()         As Byte

Const CONCHUNKSIZE As Long = 16384

Dim iChunks As Long
Dim iFragmentSize As Long
    
    'Get a unique random filename
    If Not existeix(foto) Then foto = llegir_ini("General", "rutallistats", "comandes.ini") + "\blanc.jpg"
       ' Exit Function
    sFName = foto
    
    Open sFName For Binary Access Read As #1
    ReDim z(FileLen(sFName))
    Get #1, , z()
     fldTO.AppendChunk z
    Close #1
    
    'Delete the file
    'Kill (sFName)
    
End Function



Function emplena12zeros(codi As String) As String
  emplena12zeros = codi
  If Len(codi) < 13 Then emplena12zeros = String(12 - Len(codi), "0") + codi
End Function
Sub imprimir_frontal_v9(rstpf As Recordset, nomllistat As String)
Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", "comandes.ini") + nomllistat, 1)
 ' oreport.SQLQueryString = ""

  oreport.RecordSelectionFormula = "{papersfrontals.id}=" + atrim(rstpf!ID)
 
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "vendes.mdb"
  oreport.DiscardSavedData
'  oreport.VerifyOnEveryPrint = False
  
'  oreport.PrintOut False
   Load veurereport
   veurereport.width = 15000
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   ratoli "normal"
   veurereport.Show 1, Me
 
End Sub
Sub imprimir_frontal(rstpf As Recordset, nomllistat As String)
 llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + nomllistat
 llistat.Destination = crptToWindow
' llistat.Destination = crptToPrinter
 wait 2
  llistat.PrinterSelect
 'llistat.PrinterSelect
 llistat.CopiesToPrinter = cadbl(etcopies.tag)
 llistat.DataFiles(0) = rutadelfitxer(cami) + "vendes.mdb"
 llistat.SelectionFormula = "{papersfrontals.id}=" + atrim(rstpf!ID)
 llistat.DiscardSavedData = True
 llistat.Formulas(1) = ""
 llistat.Formulas(0) = ""
 llistat.Formulas(2) = ""
 llistat.Formulas(3) = ""
 llistat.Formulas(4) = ""
 'llistat.PrinterDriver = X.DriverName
 'llistat.PrinterName = X.DeviceName
 'llistat.PrinterPort = X.Port
 
 DoEvents
 'If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
 'If mllistaperpantalla.Checked Then llistat.Destination = crptToWindow
 llistat.Action = 1
 'Set dbllistat = Nothing
 'Set rstllistat = Nothing
End Sub
Function buscar_codidebarresdeltreball(ntreball As Double) As String
   Dim dbclixes As Database
   Dim rstclixes As Recordset
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
   Set rstclixes = dbclixes.OpenRecordset("select * from clixes where id_Treball=" + atrim(ntreball))
   If rstclixes.EOF Then Exit Function
   buscar_codidebarresdeltreball = atrim(rstclixes!codidebarres)
   Set rstclixes = Nothing
   Set dbclixes = Nothing
End Function
Function mirardesarrollxrbandes(ntreball As Double, nordre As Double) As Double
   Dim dbclixes As Database
   Dim rstclixes As Recordset
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
   Set rstclixes = dbclixes.OpenRecordset("select * from modificacions where id_Treball=" + atrim(ntreball) + " and ordre=" + atrim(nordre))
   If rstclixes.EOF Then Exit Function
   mirardesarrollxrbandes = cadbl(rstclixes!desarroll) / 1000
   Set rstclixes = Nothing
   Set dbclixes = Nothing
End Function
Sub crearnoupf(rstpf As Recordset)
   Dim datafab As Date
   Dim np As Byte
   Dim numpaletv As Integer
   Dim rstc As Recordset
   Dim rstb As Recordset
   Dim rstcli As Recordset
   Dim rstmat As Recordset
   Dim rstmat2 As Recordset
   Dim pespalet As Double
   Dim pesnet As Double
   Dim pesbrut As Double
   Dim metres As Double
   Dim bobines As Double
   Dim nomcli As String
   Dim rstpp As Recordset
   Dim rstc2 As Recordset
   Dim esactys As Boolean
   Dim v11o17 As Double
   Dim pecesxrmetre As Double
   Dim vmesoscaducitat As Long
   
   vmesoscaducitat = 9
   v11o17 = 11
   numpaletv = llista.ItemData(llista.ListIndex)
   np = llista.ListIndex + 1
   Set rstc = dbtmpb.OpenRecordset("SELECT comandes.*, InStr(1,[ruta],'I') AS hihaimpresores FROM comandes LEFT JOIN productes ON comandes.producte = productes.codi Where comanda = " + atrim(comanda))
   If Not rstc.EOF Then Set rstc2 = dbtmpb.OpenRecordset("select materialex from comandes where comanda=" + atrim(rstc!linkcomanda1))
   Set rstb = dbbaixes.OpenRecordset("SELECT rebobinadores.comanda, count(*) as tbobines,Sum(bobinesreb.pesnet) AS pesnet,Sum(bobinesreb.kilos) AS pesbrut,sum(bobinesreb.metres) as tmetres, bobinesreb.palet FROM bobinesreb INNER JOIN rebobinadores ON bobinesreb.controlid = rebobinadores.Id GROUP BY rebobinadores.comanda, bobinesreb.palet HAVING (((rebobinadores.comanda)=" + atrim(comanda) + ") AND ((bobinesreb.palet)=" + atrim(numpaletv) + "));")
   Set rstb = dbbaixes.OpenRecordset("SELECT comanda, count(*) as tbobines,Sum(kilosnets) AS pesnet,Sum(kilosiunitats) AS pesbrut,sum(metresisacs) as tmetres, numpalet as palet FROM bobinesent gROUP BY comanda, numpalet HAVING (((comanda)=" + atrim(comanda) + ") AND ((numpalet)=" + atrim(numpaletv) + "));")
   Set rstpp = dbbaixes.OpenRecordset("SELECT reb_pespalets.pespalet From reb_pespalets WHERE (((reb_pespalets.comanda)=" + comanda + ") AND ((reb_pespalets.numpalet)=" + atrim(numpaletv) + "));")
   If rstc!hihaimpresores > 0 Then
      pecesxrmetre = mirardesarrollxrbandes(rstc!numtreball, rstc!numordremodificacio)
      'If pecesxrmetre > 0 Then pecesxrmetre = 1 / pecesxrmetre
   End If
   If Not rstb.EOF Then
     pespalet = IIf(Not rstpp.EOF, rstpp!pespalet, 21)
     pesnet = IIf(rstb!pesnet = 0, rstb!pesbrut, rstb!pesnet)
     pesbrut = rstb!pesbrut
     metres = rstb!tmetres
     bobines = rstb!tbobines
     Set rstcli = dbtmpb.OpenRecordset("select nom from clients where codi=" + atrim(rstc!client))
     If Not rstcli.EOF Then
        nomcli = rstcli!nom
        If InStr(1, nomcli, "ACTYS") > 0 Then esactys = True
        If InStr(1, nomcli, "SCHWEPPES") > 0 Then v11o17 = 17
        If (InStr(1, nomcli, "DARTA ") > 0) Or (InStr(1, nomcli, "DARTA,") > 0) Or rstc!client = 6393 Then v11o17 = 31
        If InStr(1, nomcli, "FRIT ") > 0 Then v11o17 = 3110
        'If InStr(1, nomcli, "BENIMODO") > 0 Then v11o17 = 35
        If InStr(1, nomcli, "BENIMODO") > 0 Then v11o17 = 3102
        If InStr(1, nomcli, "ORKLA") > 0 Then v11o17 = 37: vmesoscaducitat = 6
     End If
   End If
   datafab = CVDate(llista.tag)
   If rstc.EOF Then Exit Sub
   Set rstmat = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstc!materialex)))
   Set rstmat2 = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstc2!materialex)))
   If rstpf.EditMode = 0 Then rstpf.AddNew
    rstpf!codiempresa = IIf(esactys, 3264960, 8402320) 'CODI ACTYS I INPLACSA
    rstpf!codiarticle = IIf(cadbl(refclient.tag) < 0, cadbl(refclient.tag) * -1, cadbl(refclient.tag))
    rstpf!codiclient = cadbl(nomclient.tag)
    rstpf!nomclient = nomclient
    rstpf!refclient = IIf(atrim(rstc!refclientdeclient) <> "", atrim(rstc!refclientdeclient), atrim(rstc!refclient))
    
    rstpf!material = descripciomaterial(rstmat) + " + " + descripciomaterial(rstmat2)
    'rstpf!material2 = descripciomaterial(rstmat2)
    If Len(rstpf!material2) < 3 Then rstpf!material2 = ""
    rstpf!ean13 = rstc!codibarras
    rstpf!pedidoclient = rstc!comandaclient
    rstpf!texteimp = Mid(atrim(rstc!marcailinia) + " ", 1, 50)
    rstpf!numlotinplacsa = cadbl(comanda)
    rstpf!datafabricacio = datafab
    rstpf!datacaducitat = DateAdd("m", vmesoscaducitat, datafab)
    rstpf!pesnet = pesnet
    rstpf!pesbrut = pesbrut + pespalet
    rstpf!metres = metres
    rstpf!bobines = bobines
    rstpf!amplada = cadbl(formvendes.datalinies.Recordset!ampladamaterial)
    rstpf!espesor = cadbl(formvendes.datalinies.Recordset!espesor)
    rstpf!desarroll = pecesxrmetre * 1000
    If pecesxrmetre > 0 Then rstpf!peces = Redondejar(metres / pecesxrmetre, 0)
    If rstpf!peces = 0 Then rstpf!peces = bobines
    rstpf!numpaletvisual = np
    rstpf!numpalet = numpaletv
    rstpf!totalpalets = llista.ListCount
    rstpf!gtin14 = generargtin14(rstpf)
    rstpf!scc = IIf(esactys, generarsccACTYS(rstpf), generarscc(rstpf))
    If InStr(1, estiletiqueta, "128") > 0 Then
      If Not esactys Then
           rstpf!codidebarres1 = generarcodi1(rstpf, v11o17)
           rstpf!codidebarres2 = generarcodi2(rstpf, v11o17)
           If v11o17 = 31 Then rstpf!codidebarres2 = 0
            Else
               rstpf!codidebarres1 = generarcodi1ACTYS(rstpf)
               rstpf!codidebarres2 = generarcodi2ACTYS(rstpf)
      End If
    End If
    If InStr(1, estiletiqueta, "EAN13") > 0 Then
       rstpf!codidebarres1 = rstpf!numlotinplacsa
       rstpf!codidebarres2 = treurecaracters(rstpf!refclient)
    End If
    If InStr(1, estiletiqueta, "EAN-13") > 0 Then
       rstpf!codidebarres1 = buscar_codidebarresdeltreball(rstc!numtreball)
       rstpf!codidebarres2 = 0
    End If
    If esactys Then rstpf!formatetiqueta = "ACTYS"
    rstpf!detallbobinesxpalet = formvendes.generarliniadepackinglist(cadbl(numpaletv), cadbl(formvendes.cnumalbara), cadbl(comanda))
     
   rstpf.Update
   
   wait 5
End Sub
Function treurecaracters(refclient As String) As String
   Dim ref As String
   ref = refclient
   For i = 1 To Len(refclient)
     If Not IsNumeric(Mid(refclient, i, 1)) Then substituir ref, Mid(refclient, i, 1), ""
   Next i
   treurecaracters = ref
End Function
Function generargtin14(rstpf As Recordset) As String
   generargtin14 = "9" + Format(rstpf!codiempresa, "0000000") + Format(rstpf!codiarticle, "00000")
   generargtin14 = Mid(generargtin14 + atrim(codidebarres1.CalculateCheckdigits(eCDEAN14, generargtin14)), 1, 14)
   'MsgBox Len(generargtin14)
End Function
Function generarscc(rstpf As Recordset) As String
   generarscc = "1" + Format(rstpf!codiempresa, "0000000") + Format(rstpf!ID, "000000000")
   'MsgBox Len(generarscc)
End Function
Function buscarlultimid(rstpf As Recordset) As Integer
   Dim rst As Recordset
   buscarlultimid = 0
   'Set dbstocks = OpenDatabase(camistock)
   Set rst = dbtmp.OpenRecordset("select cdbl(mid(scc,14)) as idpalet from papersfrontals where numlotinplacsa=" + atrim(rstpf!numlotinplacsa) + " and numpalet=" + atrim(rstpf!numpalet))
   If rst.EOF Then
      Set rst = dbtmp.OpenRecordset("select cdbl(mid(scc,14)) as idpalet from papersfrontals where mid(scc,11,3)='571' order by cdbl(mid(scc,14)) desc")
      If rst.EOF Then Exit Function
      buscarlultimid = rst!idpalet
         Else
         buscarlultimid = rst!idpalet - 1
   End If
   
   Set rst = Nothing
End Function
Function generarsccACTYS(rstpf As Recordset) As String
   Dim proximid As Integer
   proximid = buscarlultimid(rstpf) + 1
   generarsccACTYS = "3" + Format(rstpf!codiempresa, "000000") + "000571" + Format(proximid, "0000")
End Function
Function generarcodi1(rstpf As Recordset, Optional v11o17 As Double) As String
   Dim vvalor11o17 As String
   If v11o17 = 11 Then generarcodi1 = "02" + atrim(rstpf!gtin14) + "11" + Format(rstpf!datafabricacio, "yymmdd") + "15" + Format(rstpf!datacaducitat, "yymmdd") + "3101" + treurecoma(rstpf!pesnet) + "10" + Format(rstpf!numlotinplacsa, "00000#")
   If v11o17 = 17 Then generarcodi1 = "02" + atrim(rstpf!gtin14) + "17" + Format(DateAdd("m", 9, rstpf!datafabricacio), "yymmdd") + "37" + Format(rstpf!peces, "00000000#")
   If v11o17 = 31 Then generarcodi1 = "(10)" + Mid(rstpf!refclient, 1, 20) + "(3100)" + Format(Redondejar(rstpf!pesnet, 0), "000000")
   If v11o17 = 37 Then generarcodi1 = "02" + atrim(rstpf!gtin14) + "\F10" + Mid(rstpf!refclient, 1, 20) + "\F15" + Format(rstpf!datacaducitat, "yymmdd") + "37" + Format(rstpf!peces, "00000000#")
   
   If v11o17 = 3102 Then
      generarcodi1 = "02" + atrim(rstpf!gtin14) + "11" + Format(rstpf!datafabricacio, "yymmdd") + "3102" + treurecoma(rstpf!pesnet, 2)
   End If
   If v11o17 = 3110 Then
      generarcodi1 = "02" + atrim(rstpf!gtin14) + "17" + Format(rstpf!datacaducitat, "yymmdd") + "10" + Format(rstpf!numlotinplacsa, "00000#")
      'generarcodi1 = "(10)" + Mid(rstpf!refclient, 1, 20) + "(3100)" + Format(Redondejar(rstpf!pesnet, 0), "000000")
   End If
   If v11o17 = 35 Then generarcodi1 = "02" + atrim(rstpf!gtin14) + "15" + Format(rstpf!datacaducitat, "yymmdd") + "3100" + treurecoma(rstpf!pesbrut) + "10" + Format(rstpf!numlotinplacsa, "00000#") + "\F400" + Format(rstpf!pedidoclient, "00000#")
   codidebarres1.EscapeSequences = True
   
End Function
Function generarcodi2(rstpf As Recordset, v11o17 As Double) As String
   codidebarres1.Format = "00#################^"

   generarcodi2 = "00" + atrim(rstpf!scc) + Mid(codidebarres1.CalculateCheckdigits(eCDMod10, rstpf!scc), 1, 1)
   codidebarres1.Format = ""
   If v11o17 = 3102 Then
     generarcodi2 = generarcodi2 + "10" + Format(rstpf!numlotinplacsa, "00000#")
   End If
   If v11o17 = 3110 Then
     generarcodi2 = generarcodi2 + "3102" + treurecoma(rstpf!pesnet, 2) + "3110" + treurecoma(rstpf!metres, 0) + "37" + Format(rstpf!bobines, "0000000#")
   End If
End Function
Function generarcodi1ACTYS(rstpf As Recordset) As String
   Dim valor02 As String
   valor02 = IIf(atrim(rstpf!ean13) <> "", atrim(rstpf!ean13), atrim(rstpf!refclient))
   valor02 = Format(cadbl(valor02), "00000000000000")
   generarcodi1ACTYS = "02" + valor02 + "11" + Format(rstpf!datafabricacio, "yymmdd") + "37" + Format(rstpf!peces, "00000000#")
End Function
Function generarcodi2ACTYS(rstpf As Recordset) As String
   codidebarres1.Format = "00#################^"

   generarcodi2ACTYS = "00" + atrim(rstpf!scc) + Mid(codidebarres1.CalculateCheckdigits(eCDMod10, rstpf!scc), 1, 1)
   codidebarres1.Format = ""
   generarcodi2ACTYS = generarcodi2ACTYS + "10" + rstpf!pedidoclient
End Function
Function treurecoma(v, Optional vnumdecimals As Byte) As String
  Dim r As String
  Dim r2 As Double
  If vnumdecimals > 0 Then
   r = Format(v, "###.0" + String(vnumdecimals - 1, "0"))
   r = Mid(r, 1, InStr(1, r, ",") - 1) + Mid(r, InStr(1, r, ",") + 1)
    Else: r = v
  End If
  r2 = cadbl(r)
  treurecoma = Format(r2, "000000")
End Function

Private Sub Form_Load()
  etcopies.tag = "1"
  etcopies = ""
Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini))
codidebarres1.Enabled = True
codidebarres2.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'dbbaixes.Close
'   Set dbbaixes = Nothing
End Sub


