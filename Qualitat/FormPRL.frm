VERSION 5.00
Begin VB.Form FormPRL 
   Caption         =   "Control PRL"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14220
   Icon            =   "FormPRL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bbuscar 
      Height          =   480
      Left            =   8115
      Picture         =   "FormPRL.frx":048A
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Buscar proveidor"
      Top             =   270
      Width           =   765
   End
   Begin VB.TextBox cnomproveidor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1095
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Escullir proveidor"
      Top             =   300
      Width           =   6945
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operaris"
      Height          =   3375
      Left            =   135
      TabIndex        =   2
      Top             =   4800
      Width           =   10575
      Begin VB.Frame Frame3 
         Caption         =   "Documentació"
         Height          =   2985
         Left            =   2820
         TabIndex        =   6
         Top             =   150
         Width           =   7575
         Begin VB.CommandButton Command1 
            Caption         =   "Imprimir Targeta"
            Height          =   555
            Left            =   5865
            Picture         =   "FormPRL.frx":0A14
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2340
            Width           =   1650
         End
         Begin VB.CommandButton bDocumentacioTreballadors 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Formació en alçada"
            Height          =   555
            Index           =   7
            Left            =   3210
            OLEDropMode     =   1  'Manual
            Style           =   1  'Graphical
            TabIndex        =   29
            Tag             =   "FORMACIO_ALÇADA"
            Top             =   2070
            Width           =   2565
         End
         Begin VB.CommandButton bDocumentacioTreballadors 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Cessió carretó elevador"
            Height          =   555
            Index           =   6
            Left            =   3225
            OLEDropMode     =   1  'Manual
            Style           =   1  'Graphical
            TabIndex        =   28
            Tag             =   "CESSIO_TORO"
            Top             =   1470
            Width           =   2565
         End
         Begin VB.CommandButton bDocumentacioTreballadors 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Cessió Plataforma"
            Height          =   555
            Index           =   5
            Left            =   3240
            OLEDropMode     =   1  'Manual
            Style           =   1  'Graphical
            TabIndex        =   27
            Tag             =   "CESSIO_PLAT"
            Top             =   870
            Width           =   2565
         End
         Begin VB.CommandButton bDocumentacioTreballadors 
            BackColor       =   &H0025EFAD&
            Caption         =   "BP. manteniment"
            Height          =   555
            Index           =   4
            Left            =   3225
            OLEDropMode     =   1  'Manual
            Style           =   1  'Graphical
            TabIndex        =   26
            Tag             =   "BP_TREBALLADOR"
            Top             =   285
            Width           =   2565
         End
         Begin VB.CommandButton bDocumentacioTreballadors 
            BackColor       =   &H0025EFAD&
            Caption         =   "Certificat Aptitud Mèdica"
            Height          =   555
            Index           =   3
            Left            =   240
            OLEDropMode     =   1  'Manual
            Style           =   1  'Graphical
            TabIndex        =   25
            Tag             =   "CERT-METGE"
            Top             =   2070
            Width           =   2565
         End
         Begin VB.CommandButton bDocumentacioTreballadors 
            BackColor       =   &H00C0C0FF&
            Caption         =   "EPIS"
            Height          =   555
            Index           =   2
            Left            =   255
            OLEDropMode     =   1  'Manual
            Style           =   1  'Graphical
            TabIndex        =   24
            Tag             =   "EPIS"
            Top             =   1470
            Width           =   2565
         End
         Begin VB.CommandButton bDocumentacioTreballadors 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Formacions"
            Height          =   555
            Index           =   1
            Left            =   255
            OLEDropMode     =   1  'Manual
            Style           =   1  'Graphical
            TabIndex        =   23
            Tag             =   "FORMACIONS"
            Top             =   870
            Width           =   2565
         End
         Begin VB.CommandButton bDocumentacioTreballadors 
            BackColor       =   &H00C0C0FF&
            Caption         =   "PRL"
            Height          =   555
            Index           =   0
            Left            =   255
            OLEDropMode     =   1  'Manual
            Style           =   1  'Graphical
            TabIndex        =   22
            Tag             =   "PRL"
            Top             =   285
            Width           =   2565
         End
         Begin VB.Image ImageOK 
            Height          =   1290
            Left            =   6015
            Picture         =   "FormPRL.frx":0F9E
            Stretch         =   -1  'True
            Top             =   465
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Image imatgeprohibit 
            Height          =   1500
            Left            =   5925
            Picture         =   "FormPRL.frx":17D7
            Top             =   330
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label etcaducitatrevisiometge 
            BackStyle       =   0  'Transparent
            Caption         =   "Caducitat Certificat: 12/12/26"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   300
            TabIndex        =   30
            Top             =   2625
            Width           =   2475
         End
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   105
         Picture         =   "FormPRL.frx":24D7
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   300
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   540
         Picture         =   "FormPRL.frx":2A61
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Edicio del  Registres"
         Top             =   300
         Width           =   420
      End
      Begin VB.ListBox cllistatreballadors 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         ItemData        =   "FormPRL.frx":2FEB
         Left            =   105
         List            =   "FormPRL.frx":2FFE
         TabIndex        =   3
         Top             =   690
         Width           =   2580
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Documentació Empresa"
      Height          =   3870
      Left            =   135
      TabIndex        =   0
      Top             =   870
      Width           =   10560
      Begin VB.CommandButton ccanvidata 
         Height          =   315
         Left            =   10140
         Picture         =   "FormPRL.frx":303A
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Edicio del  Registres"
         Top             =   150
         Width           =   360
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H00C0C0FF&
         Height          =   495
         Index           =   11
         Left            =   4515
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H00C0C0FF&
         Height          =   495
         Index           =   10
         Left            =   4515
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2535
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Registre d'entrega de documentació als treballadors"
         Height          =   495
         Index           =   9
         Left            =   4530
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "ENTREGA_DOCS"
         Top             =   1980
         Width           =   3420
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Riscos generals del centre de treball"
         Height          =   495
         Index           =   8
         Left            =   4515
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "RISCOS_EMPRESA"
         Top             =   1440
         Width           =   3420
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Normes generals empresa"
         Height          =   495
         Index           =   7
         Left            =   4500
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "NORMES"
         Top             =   915
         Width           =   3420
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H0025EFAD&
         Caption         =   "BP. Manipulació"
         Height          =   495
         Index           =   6
         Left            =   4500
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "BP_EMPRESA"
         Top             =   390
         Width           =   3420
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H00C0C0FF&
         Caption         =   "RNT i RLC"
         Height          =   495
         Index           =   5
         Left            =   660
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "RNT_RLC"
         Top             =   3135
         Width           =   3420
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Fitxes de Seguretat"
         Height          =   495
         Index           =   4
         Left            =   660
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "FITXES"
         Top             =   2550
         Width           =   3420
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Designació RP"
         Height          =   495
         Index           =   3
         Left            =   675
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "RP"
         Top             =   1995
         Width           =   3420
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Declaració durada PRL i CAE"
         Height          =   495
         Index           =   2
         Left            =   660
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "PRL"
         Top             =   1455
         Width           =   3420
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H0025EFAD&
         Caption         =   "Avaluació de Riscos"
         Height          =   495
         Index           =   1
         Left            =   645
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "RISCOS"
         Top             =   930
         Width           =   3420
      End
      Begin VB.CommandButton bDocumentacio 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Documentacio CAE"
         Height          =   495
         Index           =   0
         Left            =   645
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "CAE"
         Top             =   405
         Width           =   3420
      End
      Begin VB.CommandButton beliminarproveidor 
         Height          =   375
         Left            =   45
         Picture         =   "FormPRL.frx":35C4
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Eliminar el proveïdor"
         Top             =   195
         Width           =   465
      End
      Begin VB.Label etcaducitat 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Caducitat: 12/12/26"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   7935
         TabIndex        =   32
         Top             =   225
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Proveïdor:"
      Height          =   285
      Left            =   210
      TabIndex        =   1
      Top             =   420
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   10890
      Picture         =   "FormPRL.frx":3B4E
      Stretch         =   -1  'True
      Top             =   1125
      Width           =   3030
   End
End
Attribute VB_Name = "FormPRL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vrutaFitxersPRL As String
Dim dbqualitat As Database
Dim vTotaLaDocumentacioEmpresa As Boolean

Private Sub ccanvidata_Click()
   Dim v As String
   v = InputBox("Entra la data de caducitat.", "Data caducitat", Format(DateAdd("yyyy", 1, Now), "dd/mm/yy"))
   If IsDate(v) Then dbqualitat.Execute "update proveidors_prl set datacaducitat=#" + Format(v, "mm/dd/yy") + "# where id=" + atrim(cadbl(cnomproveidor.Tag))
   carregar_informacio
   carregainfotreballador
End Sub

Private Sub Command1_Click()
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim camp As TextObject
  Dim f  As OLEObject
  Dim vformula As String
  Dim i As Byte
  Dim vcopies As Byte
  Dim vnomfitxerRPT As String
  Dim vnumtreballador As Double
  
  If cllistatreballadors.ListIndex < 0 Then Exit Sub
  If ImageOK.Visible = False Then MsgBox "No hi han totes les dades necesaries per poder deixar entrar aquest treballador.", vbCritical, "Error": Exit Sub
  vnumtreballador = cllistatreballadors.ItemData(cllistatreballadors.ListIndex)
  Set oapp = New CRAXDDRT.Application
  vnomfitxerRPT = llegir_ini("General", "rutallistats", "comandes.ini") + "TargetaTreballadorPRL.rpt"
  Set oreport = oapp.OpenReport(vnomfitxerRPT, 1)
  If existeix(vrutaFitxersPRL + atrim(cnomproveidor.Tag) + "\" + "\" + atrim(vnumtreballador) + "\CESSIO_PLAT.PDF") Then
        oreport.FormulaFields.GetItemByName("toru").Text = 1
  End If
  If existeix(vrutaFitxersPRL + atrim(cnomproveidor.Tag) + "\" + "\" + atrim(vnumtreballador) + "\CESSIO_TORO.PDF") Then
        oreport.FormulaFields.GetItemByName("elavador").Text = 1
  End If
  If existeix(vrutaFitxersPRL + atrim(cnomproveidor.Tag) + "\" + "\" + atrim(vnumtreballador) + "\FORMACIO_ALÇADA.PDF") Then
        oreport.FormulaFields.GetItemByName("altura").Text = 1
  End If
  oreport.FormulaFields.GetItemByName("nomempresa").Text = "'" + cnomproveidor + "'"
  oreport.FormulaFields.GetItemByName("nomoperari").Text = "'" + cllistatreballadors + "'"
  
  oreport.FormulaFields.GetItemByName("datacaducitat").Text = "'Caduca: " + etcaducitat.Tag + "'"
  oreport.FormulaFields.GetItemByName("datacaducitatrevisio").Text = "'Cad.Rv.Metge: " + etcaducitatrevisiometge.Tag + "'"
  
  oreport.PrintOut False
  MsgBox "Targeta impresa.", vbInformation, "Impresió"
End Sub

Private Sub alta_Click()
  Dim v As String
  Dim rst As Recordset
  v = UCase(InputBox("Escriu el nom del treballador.", "Nom treballador"))
  If atrim(v) = "" Then Exit Sub
  
  Set rst = dbqualitat.OpenRecordset("select * from treballadors_prl where idproveidor=" + atrim(cadbl(cnomproveidor.Tag)))
  rst.FindFirst "nomtreballador='" + v + "'"
  If Not rst.NoMatch Then MsgBox "Aquest treballador ja està donat d'alta.", vbCritical, "Error": Exit Sub
  rst.AddNew
  rst!idproveidor = cadbl(cnomproveidor.Tag)
  rst!nomtreballador = atrim(v)
  rst.Update
  rst.MoveLast
  cllistatreballadors.AddItem atrim(v)
  cllistatreballadors.ItemData(cllistatreballadors.NewIndex) = rst!ID
  cllistatreballadors.ListIndex = cllistatreballadors.NewIndex
  carregainfotreballador
  Set rst = Nothing
End Sub

Private Sub bbuscar_Click()
  Dim v As String
  Dim rst As Recordset
  ratoli "normal"
  Unload formseleccio
  Load formseleccio
  'formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "qualitat.mdb"
  formseleccio.Data1.RecordSource = "SELECT id,nom from proveidors_prl"
  formseleccio.refrescar
  formseleccio.alta.Visible = True
  formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 10000
  formseleccio.DBGrid2.Font.Name = "Arial"
  formseleccio.DBGrid2.Font.Size = 20
  formseleccio.DBGrid2.Width = 10000
  formseleccio.Width = 12000
  formseleccio.Height = 6000
  formseleccio.Caption = "Escull Opció"
  formseleccio.Show 1
  If seleccioret = 2 Then
       '  crearnova:
    v = atrim(InputBox("Escriu el nom del NOU proveïdor", "Nou Proveïdor"))
    If v = "" Then Exit Sub
    v = UCase(treure_apostruf(v))
    Set rst = dbqualitat.OpenRecordset("select * from proveidors_prl")
    v = UCase(treure_apostruf(v))
    rst.FindFirst "nom='" + atrim(v) + "'"
    If Not rst.NoMatch Then MsgBox "Aquest proveïdor ja existeix.", vbCritical, "Error": GoTo fi
    rst.AddNew: rst!nom = v: rst.Update
    cnomproveidor = v
    rst.FindFirst "nom='" + atrim(v) + "'"
    cnomproveidor.Tag = atrim(rst!ID)
    
  End If
  If seleccioret = 1 Then
      cnomproveidor = UCase(formseleccio.Data1.Recordset!nom)
      cnomproveidor.Tag = atrim(formseleccio.Data1.Recordset!ID)
  End If
fi:
  Set rst = Nothing
  Unload formseleccio
  carregar_informacio
  carregar_treballadors
End Sub
Sub carregar_informacio()
   Dim i As Byte
   Dim vnomfitxer As String
   Dim rst As Recordset
   etcaducitat = "Caducitat: "
   etcaducitat.Tag = ""
   vTotaLaDocumentacioEmpresa = True
   Set rst = dbqualitat.OpenRecordset("select * from proveidors_prl where id=" + atrim(cadbl(cnomproveidor.Tag)))
   If Not rst.EOF Then
           etcaducitat = "Caducitat: " + Format(atrim(rst!datacaducitat), "dd/mm/yy")
           etcaducitat.Tag = Format(atrim(rst!datacaducitat), "dd/mm/yy")
   End If
   For i = 0 To bDocumentacio.Count - 1
       vnomfitxer = vrutaFitxersPRL + atrim(cadbl(cnomproveidor.Tag)) + "\" + bDocumentacio(i).Tag + ".PDF"
       If existeix(vnomfitxer) Then
           bDocumentacio(i).BackColor = &H25EFAD
             Else:
               If bDocumentacio(i).Visible Then vTotaLaDocumentacioEmpresa = False
               bDocumentacio(i).BackColor = &HC0C0FF
       End If
   Next i
   
End Sub

Private Sub bDocumentacio_Click(Index As Integer)
  Dim vnomfitxer As String
  vnomfitxer = vrutaFitxersPRL + atrim(cnomproveidor.Tag) + "\" + bDocumentacio(Index).Tag + ".PDF"
  If existeix(vnomfitxer) Then obrir_document vnomfitxer
End Sub

Private Sub bDocumentacio_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim vnumtreballador As Double
  Dim vnomfitxers As String
  If Shift = 2 Then
     vnomfitxers = vrutaFitxersPRL + atrim(cnomproveidor.Tag) + "\" + bDocumentacio(Index).Tag + ".PDF"
     If existeix(vnomfitxers) Then
        If MsgBox("Segur que vols eliminar el document " + bDocumentacio(Index).Caption + "?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
            If existeix(vnomfitxers) Then Kill vnomfitxers
            carregar_informacio
        End If
     End If
  End If
End Sub

Private Sub bDocumentacio_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim vnomfitxer As String
   Dim vnomfitxer_desti As String
   If cadbl(atrim(cnomproveidor.Tag)) = 0 Then MsgBox "Primer has d'escullir un proveidor.", vbCritical, "Error": Exit Sub
   vnomfitxer_desti = vrutaFitxersPRL + atrim(cnomproveidor.Tag) + "\" + bDocumentacio(Index).Tag + ".PDF"
   vnomfitxer = UCase(Data.Files(1))
   If existeix(vnomfitxer_desti) Then If MsgBox("Aquest fitxer ja existeix, vols sustituir-lo?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then Kill vnomfitxer_desti Else Exit Sub
   Copiar_Fitxer vnomfitxer, vnomfitxer_desti
   carregar_informacio
End Sub

Private Sub bDocumentacioTreballadors_Click(Index As Integer)
  Dim vnomfitxer As String
  Dim vnumtreballador As Integer
  vnumtreballador = cllistatreballadors.ItemData(cllistatreballadors.ListIndex)
  vnomfitxer = vrutaFitxersPRL + atrim(cnomproveidor.Tag) + "\" + "\" + atrim(vnumtreballador) + "\" + bDocumentacioTreballadors(Index).Tag + ".PDF"
  If existeix(vnomfitxer) Then obrir_document vnomfitxer
End Sub

Private Sub bDocumentacioTreballadors_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim vnumtreballador As Double
  Dim vnomfitxers As String
  If Shift = 2 Then
     vnumtreballador = cllistatreballadors.ItemData(cllistatreballadors.ListIndex)
     vnomfitxers = vrutaFitxersPRL + atrim(cadbl(cnomproveidor.Tag)) + "\" + atrim(vnumtreballador) + "\" + bDocumentacioTreballadors(Index).Tag + ".PDF"
     If existeix(vnomfitxers) Then
        If MsgBox("Segur que vols eliminar el document " + bDocumentacioTreballadors(Index).Caption + "?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
            If existeix(vnomfitxers) Then Kill vnomfitxers
            carregar_treballadors
        End If
     End If
  End If
End Sub

Private Sub bDocumentacioTreballadors_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim vnomfitxer As String
   Dim vnomfitxer_desti As String
   Dim vnumtreballador As Double
   If cadbl(atrim(cnomproveidor.Tag)) = 0 Then MsgBox "Primer has d'escullir un proveidor.", vbCritical, "Error": Exit Sub
   If cllistatreballadors.ListIndex < 0 Then MsgBox "Primer has d'escullir un treballador.", vbCritical, "Error": Exit Sub
   vnumtreballador = cllistatreballadors.ItemData(cllistatreballadors.ListIndex)
   If bDocumentacioTreballadors(Index).Tag = "CERT-METGE" Then
          vdata = InputBox("Entra la data de caducitat del certificat metge.", "Atenció")
          If DateDiff("d", vdata, Now) > 0 Then MsgBox "Ha de ser una data futura, aquesta data ja està passada.", vbCritical, "Erro": Exit Sub
          dbqualitat.Execute "update  treballadors_prl set caducitatrevisiometge=#" + Format(vdata, "mm/dd/yy") + "# where idproveidor=" + atrim(cadbl(cnomproveidor.Tag)) + " and id=" + atrim(vnumtreballador)
   End If
   vnomfitxer_desti = vrutaFitxersPRL + atrim(cadbl(cnomproveidor.Tag)) + "\" + atrim(vnumtreballador) + "\" + bDocumentacioTreballadors(Index).Tag + ".PDF"
   vnomfitxer = UCase(Data.Files(1))
   If existeix(vnomfitxer_desti) Then If MsgBox("Aquest fitxer ja existeix, vols sustituir-lo?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then Kill vnomfitxer_desti Else Exit Sub
   Copiar_Fitxer vnomfitxer, vnomfitxer_desti
   carregainfotreballador
End Sub

Private Sub beliminarproveidor_Click()
  Dim vnomfitxers As String
  If MsgBox("Segur que vols eliminar aquest proveidor?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
       'mirar si te documentació relacionada
       vnomfitxers = vrutaFitxersPRL + atrim(cnomproveidor.Tag)
       If existeix(vnomfitxers) Then CreateObject("Scripting.FileSystemObject").DeleteFolder vnomfitxers, True
       dbqualitat.Execute "delete * from treballadors_prl where idproveidor=" + atrim(cadbl(cnomproveidor.Tag))
       dbqualitat.Execute "delete * from proveidors_prl where nom='" + atrim(cnomproveidor) + "'"
       MsgBox "Proveïdor eliminat.", vbCritical, "Atenció"
       cnomproveidor.Text = "Escullir proveidor"
       cnomproveidor.Tag = "0"
       carregar_informacio
       carregar_treballadors
       carregainfotreballador
  End If
End Sub
Sub carregainfotreballador()
   Dim rst As Recordset
   Dim vnumtreballador As Byte
   Dim vnomfitxer As String
   Dim vtotaladocumentacioOperari As Boolean
   vnumtreballador = 0
   vtotaladocumentacioOperari = True
   etcaducitatrevisiometge = ""
   etcaducitatrevisiometge.Tag = ""
   If cllistatreballadors.ListIndex >= 0 Then vnumtreballador = cllistatreballadors.ItemData(cllistatreballadors.ListIndex)
   For i = 0 To bDocumentacioTreballadors.Count - 1
       vnomfitxer = vrutaFitxersPRL + atrim(cadbl(cnomproveidor.Tag)) + "\" + atrim(vnumtreballador) + "\" + bDocumentacioTreballadors(i).Tag + ".PDF"
       If existeix(vnomfitxer) Then
                bDocumentacioTreballadors(i).BackColor = &H25EFAD
                   Else:
                     bDocumentacioTreballadors(i).BackColor = &HC0C0FF
                     If bDocumentacioTreballadors(i).Tag <> "CESSIO_PLAT" And bDocumentacioTreballadors(i).Tag <> "CESSIO_TORO" And bDocumentacioTreballadors(i).Tag <> "FORMACIO_ALÇADA" Then
                          vtotaladocumentacioOperari = False
                     End If
       End If
   Next i

   
   
   Set rst = dbqualitat.OpenRecordset("select * from treballadors_prl where idproveidor=" + atrim(cadbl(cnomproveidor.Tag)) + " and id=" + atrim(vnumtreballador))
   If Not rst.EOF Then
        If Not IsNull(rst!caducitatrevisiometge) Then
            etcaducitatrevisiometge = "Caducitat Certificat: " + Format(rst!caducitatrevisiometge, "dd/mm/yy")
            etcaducitatrevisiometge.Tag = Format(rst!caducitatrevisiometge, "dd/mm/yy")
        End If
   End If
   
   ImageOK.Visible = False: imatgeprohibit.Visible = True
   If vTotaLaDocumentacioEmpresa And vtotaladocumentacioOperari Then
       If IsDate(etcaducitatrevisiometge.Tag) And IsDate(etcaducitat.Tag) Then
         If DateDiff("d", Now, CVDate(etcaducitat.Tag)) > 0 And DateDiff("d", Now, CVDate(etcaducitat.Tag)) > 0 Then
             ImageOK.Visible = True: imatgeprohibit.Visible = False
         End If
       End If
   End If
   Set rst = Nothing
End Sub
Private Sub cllistatreballadors_Click()
   carregainfotreballador
End Sub

Private Sub eliminar_Click()

End Sub

Private Sub Form_Load()
   cami = llegir_ini("General", "cami", "comandes.ini")
   vrutaFitxersPRL = "\\ord_copies\DadesProduccio\Arxius Produccio\DadesGenerals\PRL_Proveidors\"
   Set dbqualitat = OpenDatabase(rutadelfitxer(cami) + "qualitat.mdb")
   carregar_informacio
   carregar_treballadors
   carregainfotreballador
End Sub
Sub carregar_treballadors()
   Dim rst As Recordset
   cllistatreballadors.Clear
   If cadbl(cnomproveidor.Tag) = 0 Then Exit Sub
   Set rst = dbqualitat.OpenRecordset("select * from treballadors_prl where idproveidor=" + atrim(cadbl(cnomproveidor.Tag)) + " order by nomtreballador")
   While Not rst.EOF
     cllistatreballadors.AddItem rst!nomtreballador
     cllistatreballadors.ItemData(cllistatreballadors.NewIndex) = rst!ID
     rst.MoveNext
   Wend
   Set rst = Nothing
   If cllistatreballadors.ListCount > 0 Then cllistatreballadors.ListIndex = 0
   carregainfotreballador
End Sub

Private Sub modificar_Click()
   Dim v As String
   Dim vrutafitxers As String
   Dim vnumtreballador As Double
   If cllistatreballadors.ListIndex < 0 Then MsgBox "Primer escull un operari.", vbCritical, "Atenció": Exit Sub
   vnumtreballador = cadbl(cllistatreballadors.ItemData(cllistatreballadors.ListIndex))
   v = UCase(treure_apostruf(InputBox("Escriu el nom de l'operari." + vbNewLine + "Escriu [ELIMINAR] per eliminar-lo i els documents també.", "Nom operari", cllistatreballadors.List(cllistatreballadors.ListIndex))))
   If v = "ELIMINAR" Then
         'eliminar registre i fitxers vinculats
        vrutafitxers = vrutaFitxersPRL + atrim(cnomproveidor.Tag) + "\" + atrim(vnumtreballador)
        If existeix(vrutafitxers) Then
            If MsgBox("Segur que vols eliminar tots els arxius vinculats?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
               'MsgBox vrutafitxers
               CreateObject("Scripting.FileSystemObject").DeleteFolder vrutafitxers, True
                 Else: GoTo fi
            End If
        End If
        dbqualitat.Execute "delete * from treballadors_prl where idproveidor=" + atrim(cadbl(cnomproveidor.Tag)) + " and id=" + atrim(vnumtreballador)
        carregar_treballadors
      Else
         'modificar
        dbqualitat.Execute "update treballadors_prl set nomtreballador='" + v + "' where idproveidor=" + atrim(cadbl(cnomproveidor.Tag)) + " and id=" + atrim(vnumtreballador)
        cllistatreballadors.List(cllistatreballadors.ListIndex) = v
   End If
fi:
End Sub
