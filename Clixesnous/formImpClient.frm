VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form formImpClient 
   Caption         =   "Imp del client"
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13560
   Icon            =   "formImpClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Framesortida 
      BackColor       =   &H00F1B75F&
      Caption         =   "Sortida que vol el client"
      Height          =   10215
      Left            =   -15
      TabIndex        =   0
      Top             =   60
      Width           =   13485
      Begin VB.Frame Frame1 
         Height          =   2775
         Left            =   225
         TabIndex        =   20
         Top             =   3240
         Width           =   1620
         Begin VB.TextBox cmargeesquerra 
            Height          =   330
            Left            =   420
            TabIndex        =   22
            Top             =   390
            Width           =   615
         End
         Begin VB.TextBox cobs1 
            Height          =   1650
            Left            =   105
            MaxLength       =   100
            TabIndex        =   21
            Top             =   1005
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Marge esquerra:"
            Height          =   300
            Left            =   150
            TabIndex        =   25
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Cm."
            Height          =   300
            Left            =   1095
            TabIndex        =   24
            Top             =   450
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "Observació esq.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   180
            TabIndex        =   23
            Top             =   780
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2790
         Left            =   8985
         TabIndex        =   14
         Top             =   3240
         Width           =   1620
         Begin VB.TextBox cobs2 
            Height          =   1650
            Left            =   90
            MaxLength       =   100
            TabIndex        =   18
            Top             =   1005
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   330
            Left            =   420
            TabIndex        =   15
            Top             =   390
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Observació dret."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   195
            TabIndex        =   19
            Top             =   795
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Marge dret:"
            Height          =   300
            Left            =   375
            TabIndex        =   17
            Top             =   150
            Width           =   990
         End
         Begin VB.Label Label3 
            Caption         =   "Cm."
            Height          =   300
            Left            =   1095
            TabIndex        =   16
            Top             =   450
            Width           =   255
         End
      End
      Begin VB.Frame Framesortida1 
         BackColor       =   &H00F1B75F&
         BorderStyle     =   0  'None
         Caption         =   "2"
         Height          =   7005
         Left            =   180
         TabIndex        =   11
         Top             =   285
         Width           =   9195
         Begin VB.Shape Shape9 
            BackColor       =   &H00F1B75F&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   1440
            Left            =   7455
            Top             =   1035
            Width           =   1275
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00F1B75F&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   1440
            Left            =   6600
            Top             =   405
            Width           =   1275
         End
         Begin VB.Shape Shape8 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            Height          =   1485
            Left            =   105
            Shape           =   2  'Oval
            Top             =   375
            Width           =   1650
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H0000C0C0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            Height          =   510
            Left            =   660
            Shape           =   2  'Oval
            Top             =   855
            Width           =   525
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            Height          =   1470
            Left            =   7290
            Shape           =   2  'Oval
            Top             =   375
            Width           =   1455
         End
         Begin VB.Line Line9 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   1050
            X2              =   7845
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Line Line8 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   1725
            X2              =   1740
            Y1              =   1110
            Y2              =   6795
         End
         Begin VB.Line Line7 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   8715
            X2              =   8730
            Y1              =   1260
            Y2              =   6780
         End
         Begin VB.Line Line6 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   1725
            X2              =   8730
            Y1              =   6795
            Y2              =   6810
         End
      End
      Begin VB.Frame Framesortida2 
         BackColor       =   &H00F1B75F&
         BorderStyle     =   0  'None
         Caption         =   "2"
         Height          =   7005
         Left            =   1770
         TabIndex        =   10
         Top             =   270
         Width           =   8595
         Begin VB.Line Line5 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   135
            X2              =   7155
            Y1              =   6795
            Y2              =   6810
         End
         Begin VB.Line Line4 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   7155
            X2              =   7140
            Y1              =   1860
            Y2              =   6810
         End
         Begin VB.Line Line3 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   105
            X2              =   120
            Y1              =   1125
            Y2              =   6810
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   1050
            X2              =   7845
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00F1B75F&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   1440
            Left            =   6570
            Top             =   405
            Width           =   1275
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            Height          =   1470
            Left            =   7050
            Shape           =   2  'Oval
            Top             =   375
            Width           =   1455
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   1050
            X2              =   7815
            Y1              =   1845
            Y2              =   1845
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H0000C0C0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            Height          =   510
            Left            =   660
            Shape           =   2  'Oval
            Top             =   855
            Width           =   525
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            Height          =   1485
            Left            =   105
            Shape           =   2  'Oval
            Top             =   375
            Width           =   1650
         End
      End
      Begin VB.Data dataclientsvinculats_linies 
         Caption         =   "dataclientsvinculats_linies"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   5445
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   9195
         Visible         =   0   'False
         Width           =   3255
      End
      Begin MSDBGrid.DBGrid reixa 
         Bindings        =   "formImpClient.frx":10CA
         Height          =   2805
         Left            =   315
         OleObjectBlob   =   "formImpClient.frx":10F0
         TabIndex        =   13
         Top             =   7380
         Width           =   11100
      End
      Begin VB.Frame frameimgpdf 
         BackColor       =   &H00F1B75F&
         BorderStyle     =   0  'None
         Height          =   4905
         Left            =   1935
         TabIndex        =   12
         Top             =   2145
         Width           =   6960
         Begin VB.Image imgpdf 
            Height          =   4860
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   6930
         End
      End
      Begin VB.CommandButton brotarpdf 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Girar"
         Height          =   690
         Left            =   12555
         Picture         =   "formImpClient.frx":193D
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gira 90 graus el PDF"
         Top             =   1155
         Width           =   735
      End
      Begin VB.CommandButton bcanvisortida 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Canvi Sortida"
         Height          =   1155
         Left            =   12540
         Picture         =   "formImpClient.frx":1EC7
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Canvia el tipus de sortida de bobina."
         Top             =   1875
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00EAD9CE&
         Height          =   600
         Left            =   12450
         Picture         =   "formImpClient.frx":2F91
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Guarda els canvis fets a la sortida."
         Top             =   195
         Width           =   915
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF80FF&
         Caption         =   "Borrar"
         Height          =   690
         Left            =   12540
         Picture         =   "formImpClient.frx":351B
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Possa tots els parametres a defecte i recarrega la imatge."
         Top             =   3060
         Width           =   735
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00ED823A&
         Caption         =   "Ajustar el PDF "
         Height          =   1545
         Left            =   11655
         TabIndex        =   1
         Top             =   3975
         Width           =   1620
         Begin VB.CommandButton Command6 
            Height          =   420
            Left            =   135
            Picture         =   "formImpClient.frx":39A5
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Retalla el PDF en direcció a la fletxa."
            Top             =   630
            Width           =   465
         End
         Begin VB.CommandButton Command5 
            Height          =   420
            Left            =   1050
            Picture         =   "formImpClient.frx":3F2F
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Retalla el PDF en direcció a la fletxa."
            Top             =   630
            Width           =   465
         End
         Begin VB.CommandButton Command4 
            Height          =   420
            Left            =   600
            Picture         =   "formImpClient.frx":44B9
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Retalla el PDF en direcció a la fletxa."
            Top             =   1065
            Width           =   465
         End
         Begin VB.CommandButton Command3 
            Height          =   420
            Left            =   585
            Picture         =   "formImpClient.frx":4A43
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Retalla el PDF en direcció a la fletxa."
            Top             =   210
            Width           =   465
         End
      End
   End
End
Attribute VB_Name = "formImpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bcanvisortida_Click()
  If Framesortida1.visible Then
      ' imgpdf.Left = 60 + Framesortida.Left
       Framesortida2.visible = True
       Framesortida1.visible = False
       
       Exit Sub
   End If
   If Framesortida2.visible Then
      ' imgpdf.Left = 2700 + Framesortida.Left
       Framesortida1.visible = True
       Framesortida2.visible = False
       
       Exit Sub
   End If
End Sub
Sub generarminiaturapdf()
  Dim vrutapdfmini As String
  Dim vrutapdf As String
  Dim vdatapdf As String
  Dim vnomfitxerini As String
  Set imgpdf = LoadPicture("")
  imgpdf.tag = ""
  Kill "c:\temp\pdfmini_tmp.gif"
  vrutapdfmini = formclixes.rutapdftreball
  If Not existeix(vrutapdfmini) Then MsgBox "Encara no hi ha el PDF linkat al treball, no es pot fer la sortida.", vbCritical, "Error": Exit Sub
  vrutapdf = vrutapdfmini
  vrutapdfmini = substituir(vrutapdfmini, ".pdf", "_mini.gif")
  vnomfitxerini = substituir(atrim(vrutapdfmini), ".gif", ".ini")
  If vdatapdf <> FileDateTime(vrutapdf) Then If existeix(vrutapdfmini) Then Kill vrutapdfmini
  If Not existeix(vrutapdfmini) Then
     ConvertirFormats vrutapdf, vrutapdfmini, 50
  End If
  Set imgpdf = LoadPicture(vrutapdfmini)
  imgpdf.tag = vrutapdfmini
  FileCopy imgpdf.tag, "c:\temp\pdfmini_tmp.gif"
  escriure_ini "General", "datapdf", FileDateTime(vrutapdf), vnomfitxerini
  escriure_ini "General", "nompdf", vrutapdf, vnomfitxerini
End Sub
'Function substituir(cadena As String, buscar As String, canviar As String) As String
'   comença = InStr(1, cadena, buscar) - 1
'   If comença < 1 Then substituir = cadena: Exit Function
'   acaba = comença + Len(buscar) + 1
'   cadena = Mid(cadena, 1, comença) + canviar + Mid(cadena, acaba)
'   substituir = cadena
 
'End Function

Private Sub brotarpdf_Click()
  RotarImatge "c:\temp\pdfmini_tmp.gif", "c:\temp\pdfmini_tmp.gif", 90
  Set imgpdf = LoadPicture("c:\temp\pdfmini_tmp.gif")
  brotarpdf.tag = cadbl(brotarpdf.tag) + 90
  If cadbl(brotarpdf.tag) >= 360 Then brotarpdf.tag = "0"
End Sub


Private Sub Command7_Click()
  If MsgBox("Vols guardar aquests canvis?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
     If formclientsvinculats.datavinculats.Recordset.EditMode = 0 Then formclientsvinculats.datavinculats.Recordset.Edit
     formclientsvinculats.datavinculats.Recordset!grausrotaciopdfsortida = cadbl(brotarpdf.tag)
     formclientsvinculats.datavinculats.Recordset!tipusdesortidadebobina = IIf(Framesortida1.visible, 1, 2)
     FileCopy "c:\temp\pdfmini_tmp.gif", imgpdf.tag
     formclientsvinculats.datavinculats.Recordset.Update
  End If
End Sub

Private Sub Command8_Click()
 If MsgBox("Aixó borrarà la imatge del PDF i borra la configuració de la sortida de bobina." + Chr(10) + "SEGUR QUE VOLS FER-HO?", vbExclamation + vbDefaultButton2 + vbYesNo, "A T E N C I Ó") = vbYes Then
       If existeix(imgpdf.tag) Then Kill imgpdf.tag
   '    generarminiaturapdf
       brotarpdf.tag = "0"
       Framesortida1.visible = True
       Framesortida2.visible = False
       If formclientsvinculats.datavinculats.Recordset.EditMode = 0 Then formclientsvinculats.datavinculats.Recordset.Edit
       formclientsvinculats.datavinculats.Recordset!grausrotaciopdfsortida = cadbl(brotarpdf.tag)
       formclientsvinculats.datavinculats.Recordset!tipusdesortidadebobina = 1
       formclientsvinculats.datavinculats.Recordset.Update
       formclientsvinculats.datavinculats.Recordset.Move 0
       Me.Hide
  End If
End Sub

Private Sub Form_Load()
        generarminiaturapdf
        frameimgpdf.ZOrder 0
        dataclientsvinculats_linies.DatabaseName = formclientsvinculats.datavinculats.DatabaseName
        dataclientsvinculats_linies.RecordSource = "select * from clientsvinculats_linies where idclientvinculat=" + atrim(formclientsvinculats.datavinculats.Recordset!ID)
        dataclientsvinculats_linies.Refresh
       If imgpdf.tag = "" Then Exit Sub
       brotarpdf.tag = cadbl(formclientsvinculats.datavinculats.Recordset!grausrotaciopdfsortida)
       Framesortida1.visible = False
       Framesortida2.visible = False
       If cadbl(formclientsvinculats.datavinculats.Recordset!tipusdesortidadebobina) = 2 Then
           Framesortida2.visible = True
          ' imgpdf.Left = 60 + Framesortida.Left
             Else: Framesortida1.visible = True ': imgpdf.Left = 2700 + Framesortida.Left
       End If
End Sub


Private Sub Command3_Click()
  TallarImatge "c:\temp\pdfmini_tmp.gif", "c:\temp\pdfmini_tmp.gif", 0, 20
  Set imgpdf = LoadPicture("c:\temp\pdfmini_tmp.gif")
End Sub

Private Sub Command4_Click()
TallarImatge "c:\temp\pdfmini_tmp.gif", "c:\temp\pdfmini_tmp.gif", 0, -20
  Set imgpdf = LoadPicture("c:\temp\pdfmini_tmp.gif")
End Sub

Private Sub Command5_Click()
  TallarImatge "c:\temp\pdfmini_tmp.gif", "c:\temp\pdfmini_tmp.gif", -20, 0
  Set imgpdf = LoadPicture("c:\temp\pdfmini_tmp.gif")
End Sub

Private Sub Command6_Click()
TallarImatge "c:\temp\pdfmini_tmp.gif", "c:\temp\pdfmini_tmp.gif", 20, 0
  Set imgpdf = LoadPicture("c:\temp\pdfmini_tmp.gif")
End Sub

Private Sub reixa_OnAddNew()
    dataclientsvinculats_linies.Recordset!idclientvinculat = formclientsvinculats.datavinculats.Recordset!ID
End Sub

