VERSION 5.00
Object = "{8D1418DD-FB6E-4C6F-A1DC-13E914E39989}#1.0#0"; "TBarCode11.ocx"
Begin VB.Form Form1 
   Caption         =   "Generar codi de barres"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin TBarCode11LibCtl.TBarCode11 tbarcode 
      Height          =   1215
      Left            =   285
      TabIndex        =   0
      Top             =   900
      Width           =   3705
      _cx             =   6535
      _cy             =   2143
      BackColor       =   15724527
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
      Text            =   "Adjust Properties"
      TextAlignment   =   0
      BarCode         =   62
      CDMethod        =   1
      CountCheckDigits=   0
      EscapeSequences =   0   'False
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
      CountModules    =   316
      DrawStatus      =   0
      SuppressErrorMsg=   0   'False
      CountRows       =   1
      EncodingMode    =   0
      OptResolution   =   0   'False
      DisplayText     =   ""
      BarWidthReduction=   0
      BarWidthReductionUnit=   0
      Quality         =   82
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
      Dpi             =   300
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function apiGetPrivateProfileString Lib "kernel32" _
       Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
       As String, ByVal lpKeyName As Any, ByVal lpDefault As _
       String, ByVal lpReturnedString As String, ByVal nSize As _
       Long, ByVal lpFileName As String) As Long
Private Declare Function apiWritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
        As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
        ByVal lpFileName As String) As Long


Private Sub Form_Load()
   Dim vfitxer As String
   Dim vpixelsample As Double
   Dim vpixelsalt As Double
   
 '  escriure_ini "Tbarcode", "nomfitxer", "c:\temp\prova1.bmp", "generartbarcode.ini"
 '  escriure_ini "Tbarcode", "pixelsample", "1000", "generartbarcode.ini"
 '  escriure_ini "Tbarcode", "pixelsalt", "800", "generartbarcode.ini"
 '  escriure_ini "Tbarcode", "text", "xl-935", "generartbarcode.ini"
 '  escriure_ini "Tbarcode", "printdatatext", "0", "generartbarcode.ini"
 '  escriure_ini "Tbarcode", "tipusbarcode", "62", "generartbarcode.ini"
   '62 es full asci
   '13 as ean 13
   
   
   
   vfitxer = llegir_ini("Tbarcode", "nomfitxer", "generartbarcode.ini")
   vpixelsample = cadbl(llegir_ini("Tbarcode", "pixelsample", "generartbarcode.ini"))
   vpixelsalt = cadbl(llegir_ini("Tbarcode", "pixelsalt", "generartbarcode.ini"))
   tbarcode.Text = llegir_ini("Tbarcode", "text", "generartbarcode.ini")
   tbarcode.BarCode = cadbl(llegir_ini("Tbarcode", "tipusbarcode", "generartbarcode.ini"))
   tbarcode.PrintDataText = cadbl(llegir_ini("Tbarcode", "printdatatext", "generartbarcode.ini"))
   tbarcode.Enabled = True
   On Error Resume Next
   Kill vfitxer
   On Error Resume Next
   tbarcode.SaveImage vfitxer, eIMBmp, vpixelsample, vpixelsalt, 600, 600
   End
End Sub
Function llegir_ini(ByVal Ap As String, ByVal cl As String, ByVal ini As String) As String
  Dim va As String
  Dim r As Integer
  cl = Trim(cl)
  va = Space$(255)
  r = apiGetPrivateProfileString(Ap, cl, "{[}]", va, 255, ini)
  If Mid(va, 1, 4) <> "{[}]" Then
     va = Mid(va, 1, Len(Trim(va)) - 1)
   Else: va = "{[}]"
  End If
  llegir_ini = va
End Function
Sub escriure_ini(Ap As String, cl As String, tex As String, ini As String)

  Dim r As Integer
  cl = Trim(cl)
  r = apiWritePrivateProfileString(Ap, cl, tex, ini)
End Sub
Function atrim(valo As Variant) As String
  On Error Resume Next
  If IsNull(valo) Then valo = ""
  atrim = Trim(valo)
End Function
Function cadbl(ByVal valo As Variant) As Double
  If Not IsNumeric(valo) Then valo = 0
  cadbl = CDbl(valo)
End Function

