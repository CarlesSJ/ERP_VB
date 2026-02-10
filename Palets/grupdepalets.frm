VERSION 5.00
Begin VB.Form grupdepalets 
   Caption         =   "Grups de palets"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton sortir 
      Height          =   390
      Left            =   5250
      Picture         =   "grupdepalets.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Sortir"
      Top             =   150
      Width           =   390
   End
   Begin VB.Data grupdpalets 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   2895
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   165
      Width           =   2205
   End
   Begin VB.Frame Grups 
      Caption         =   "Grups"
      Height          =   6165
      Left            =   135
      TabIndex        =   0
      Top             =   630
      Width           =   5415
      Begin VB.Frame Framecompatibles 
         BackColor       =   &H00F1B75F&
         Height          =   1350
         Left            =   45
         TabIndex        =   41
         Top             =   4455
         Width           =   5310
         Begin VB.ComboBox Combocompatibles 
            DataField       =   "nomgrupmaterialscompatibles"
            DataSource      =   "grupdpalets"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00ED823A&
            Height          =   360
            Left            =   255
            TabIndex        =   42
            Top             =   465
            Width           =   4995
         End
         Begin VB.Label etcodicompatible 
            BackStyle       =   0  'Transparent
            Caption         =   "Label4"
            DataField       =   "codigrupmaterialscompatibles"
            DataSource      =   "grupdpalets"
            Height          =   225
            Left            =   465
            TabIndex        =   44
            Top             =   1005
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Grup de materials compatibles"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   43
            Top             =   180
            Width           =   3270
         End
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "estocminim"
         DataSource      =   "grupdpalets"
         Height          =   285
         Index           =   7
         Left            =   4410
         TabIndex        =   39
         Top             =   1200
         Width           =   915
      End
      Begin VB.ComboBox seccio 
         DataField       =   "seccio"
         DataSource      =   "grupdpalets"
         Height          =   315
         ItemData        =   "grupdepalets.frx":058A
         Left            =   4740
         List            =   "grupdepalets.frx":0594
         TabIndex        =   37
         Top             =   375
         Width           =   540
      End
      Begin VB.TextBox idreserva 
         BackColor       =   &H00C0C0C0&
         DataField       =   "idreserva"
         DataSource      =   "grupdpalets"
         Height          =   285
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   35
         ToolTipText     =   "Aquest ID identifica la reserva a la taula (NO TE CAP REFERENCIA AMB RES MES)"
         Top             =   825
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         DataField       =   "nomdelgrup"
         DataSource      =   "grupdpalets"
         Height          =   285
         Index           =   6
         Left            =   1425
         TabIndex        =   28
         Top             =   390
         Width           =   3195
      End
      Begin VB.TextBox txtFields 
         DataField       =   "numerogrup"
         DataSource      =   "grupdpalets"
         Height          =   285
         Index           =   5
         Left            =   540
         TabIndex        =   26
         Top             =   390
         Width           =   795
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "paletexemple"
         DataSource      =   "grupdpalets"
         Height          =   285
         Index           =   0
         Left            =   1710
         TabIndex        =   23
         Top             =   780
         Width           =   915
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00D29F7D&
         Enabled         =   0   'False
         Height          =   2730
         Left            =   255
         TabIndex        =   1
         Top             =   1590
         Width           =   4935
         Begin VB.TextBox txtFields 
            DataField       =   "codimatprognou"
            DataSource      =   "grupdpalets"
            Height          =   285
            Index           =   1
            Left            =   1380
            TabIndex        =   12
            Top             =   360
            Width           =   390
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Ample"
            DataSource      =   "grupdpalets"
            Height          =   285
            Index           =   2
            Left            =   1365
            TabIndex        =   11
            Top             =   690
            Width           =   795
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Plegat"
            DataSource      =   "grupdpalets"
            Height          =   285
            Index           =   3
            Left            =   2790
            TabIndex        =   10
            Top             =   705
            Width           =   780
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Solapa"
            DataSource      =   "grupdpalets"
            Height          =   285
            Index           =   4
            Left            =   1380
            TabIndex        =   9
            Top             =   1320
            Width           =   795
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "semielaborat"
            DataSource      =   "grupdpalets"
            Height          =   315
            ItemData        =   "grupdepalets.frx":059E
            Left            =   450
            List            =   "grupdepalets.frx":05A8
            TabIndex        =   8
            Top             =   2265
            Width           =   615
         End
         Begin VB.TextBox nommaterial 
            DataSource      =   "grupdpalets"
            Height          =   285
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   360
            Width           =   2505
         End
         Begin VB.ComboBox Combo2 
            DataField       =   "carestractat"
            DataSource      =   "grupdpalets"
            Height          =   315
            ItemData        =   "grupdepalets.frx":05B2
            Left            =   1080
            List            =   "grupdepalets.frx":05BF
            TabIndex        =   6
            Top             =   2265
            Width           =   615
         End
         Begin VB.ComboBox Combo3 
            DataField       =   "obert"
            DataSource      =   "grupdpalets"
            Height          =   315
            ItemData        =   "grupdepalets.frx":05CC
            Left            =   1890
            List            =   "grupdepalets.frx":05D9
            TabIndex        =   5
            Top             =   2265
            Width           =   615
         End
         Begin VB.CheckBox microp 
            BackColor       =   &H00D29F7D&
            Caption         =   "Microperforat"
            DataField       =   "microperforat"
            DataSource      =   "grupdpalets"
            Height          =   300
            Left            =   2655
            TabIndex        =   4
            Top             =   2265
            Width           =   1470
         End
         Begin VB.TextBox txtFields 
            DataField       =   "micres"
            DataSource      =   "grupdpalets"
            Height          =   285
            Index           =   11
            Left            =   1365
            TabIndex        =   3
            Top             =   1005
            Width           =   780
         End
         Begin VB.TextBox txtFields 
            DataField       =   "grmsm2"
            DataSource      =   "grupdpalets"
            Height          =   285
            Index           =   13
            Left            =   3645
            TabIndex        =   2
            Top             =   1020
            Width           =   660
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Producte:"
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   22
            Top             =   405
            Width           =   1005
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Ample:"
            Height          =   255
            Index           =   2
            Left            =   420
            TabIndex        =   21
            Top             =   735
            Width           =   1005
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00D29F7D&
            Caption         =   "Plegat:"
            Height          =   255
            Index           =   3
            Left            =   2205
            TabIndex        =   20
            Top             =   735
            Width           =   1005
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Solapa:"
            Height          =   255
            Index           =   4
            Left            =   420
            TabIndex        =   19
            Top             =   1365
            Width           =   1005
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "T/L"
            Height          =   300
            Index           =   3
            Left            =   555
            TabIndex        =   18
            Top             =   2010
            Width           =   765
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cares Tractat"
            Height          =   300
            Index           =   0
            Left            =   930
            TabIndex        =   17
            Top             =   2010
            Width           =   1020
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Obert"
            Height          =   300
            Index           =   1
            Left            =   1980
            TabIndex        =   16
            Top             =   2025
            Width           =   540
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Espesor:"
            Height          =   255
            Index           =   15
            Left            =   435
            TabIndex        =   15
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Label Label2 
            BackColor       =   &H00D29F7D&
            Caption         =   "Micres"
            Height          =   195
            Left            =   2235
            TabIndex        =   14
            Top             =   1035
            Width           =   675
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Grms/m2:"
            Height          =   255
            Index           =   16
            Left            =   2895
            TabIndex        =   13
            Top             =   1050
            Width           =   1005
         End
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Estoc Minim (Mtrs):"
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   40
         Top             =   1230
         Width           =   1590
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Seccio"
         Height          =   255
         Index           =   7
         Left            =   4770
         TabIndex        =   38
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID Reserva:"
         Height          =   270
         Left            =   3435
         TabIndex        =   36
         Top             =   825
         Width           =   1110
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom del Grup"
         Height          =   255
         Index           =   6
         Left            =   1935
         TabIndex        =   29
         Top             =   165
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº del Grup"
         Height          =   255
         Index           =   5
         Left            =   540
         TabIndex        =   27
         Top             =   165
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Palet d'Exemple:"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   24
         Top             =   810
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Left            =   60
      TabIndex        =   30
      Top             =   15
      Width           =   5625
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   30
         Picture         =   "grupdepalets.frx":05E6
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Alta  Registres"
         Top             =   135
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   900
         Picture         =   "grupdepalets.frx":0B70
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Eliminacio Registres"
         Top             =   135
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   465
         Picture         =   "grupdepalets.frx":10FA
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Edicio del  Registres"
         Top             =   135
         Width           =   420
      End
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   1380
         Picture         =   "grupdepalets.frx":1684
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Acceptar canvis"
         Top             =   150
         Width           =   420
      End
   End
End
Attribute VB_Name = "grupdepalets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
 Dim rstpalets As Recordset
  Dim elgran As Double
  'If palets.Recordset.EOF Then Exit Sub
  If grupdpalets.Recordset.EditMode > 0 Then MsgBox "Estas editant. Primer finalitza l'edicio.": Exit Sub
  activarframes True
  grupdpalets.Recordset.AddNew
  activarframes True
  txtFields(5).SetFocus
End Sub

Private Sub Combocompatibles_Click()
  If Combocompatibles.ListIndex > -1 Then
    etcodicompatible = Combocompatibles.ItemData(Combocompatibles.ListIndex)
      Else: etcodicompatible = "0"
  End If
End Sub

Private Sub Command1_Click()
  If grupdpalets.Recordset.EditMode > 0 Then
     If atrim(seccio) = "" Then MsgBox "Falta possar la seccio": Exit Sub
     If cadbl(idreserva) = 0 And cadbl(cadbl(txtFields(0))) > 0 Then idreserva = atrim(crear_reserva_delgrup)
     If cadbl(txtFields(0)) = 0 Then txtFields(0) = "0"
     grupdpalets.Recordset.Update
     grupdpalets.Recordset.Bookmark = grupdpalets.Recordset.LastModified
     
     activarframes False
  End If
End Sub
Function crear_reserva_delgrup() As Double
  Dim rstres As Recordset
  Dim rstpalet As Recordset
  Dim rstmat As Recordset
  Set rstpalet = dbtmp.OpenRecordset("select * from palets where idpalet=" + atrim(cadbl(txtFields(0))))
  If rstpalet.EOF Then GoTo errorreserva
  Set rstmat = dbtmpb.OpenRecordset("select * from materials where codi=" + atrim(cadbl(rstpalet!codimatprognou)))
  If rstmat.EOF Then GoTo errorreserva
  Set rstres = dbtmp.OpenRecordset("reserves")
  rstres.AddNew
  rstres!ample = rstpalet!ample
  rstres!plegat = rstpalet!plegat
  rstres!carestractat = rstpalet!carestractat
  rstres!obert = rstpalet!obert
  rstres!microperforat = cabool(rstpalet!microperforat)
  rstres!semielaborat = rstpalet!semielaborat
  rstres!espesor = rstpalet!micres
  rstres!familia = rstmat!familia
  rstres!subfamilia = rstmat!subfamilia
  rstres!familiacol = rstmat!familiacol
  rstres!subfamiliacol = rstmat!subfamiliacol
  rstres!familiaad = rstmat!familiaad
  rstres!subfamiliaad = rstmat!subfamiliaad
  
  rstres.Update
  rstres.Bookmark = rstres.LastModified
  crear_reserva_delgrup = rstres!idreserva
  Set rstres = Nothing
  Set rstpalet = Nothing
  Exit Function
errorreserva:
  MsgBox "Palet no trobat no s'ha creat la reserva.": Exit Function
End Function
Private Sub eliminar_Click()
  If grupdpalets.Recordset.EOF Then MsgBox "No hi ha registres": Exit Sub
  If MsgBox("Segur que vols borrar aquest grup de palets?", vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
     'If Not grupdpalets.Recordset.EOF Then
      grupdpalets.Recordset.Delete
      grupdpalets.Refresh
     'End If
  End If
  activarframes False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Command1_Click
  If KeyCode = 27 Then
   If grupdpalets.Recordset.EditMode > 0 Then
     grupdpalets.Recordset.CancelUpdate
     activarframes False
   End If
  End If
End Sub
Sub carregar_combo_compatibles()
  Dim rst As Recordset
  Combocompatibles = Clear
  Set rst = dbtmp.OpenRecordset("select * from grupsmaterialscompatibles order by nomdelgrup")
  While Not rst.EOF
      Combocompatibles.AddItem rst!nomdelgrup
      Combocompatibles.ItemData(Combocompatibles.NewIndex) = cadbl(rst!numerodegrup)
      rst.MoveNext
  Wend
  Combocompatibles.AddItem "-  CAP  -"
  Set rst = Nothing
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
   grupdpalets.RecordSource = "grupsdepalets"
   grupdpalets.DatabaseName = Form1.palets.DatabaseName
   grupdpalets.Refresh
   If Not grupdpalets.Recordset.EOF Then grupdpalets.Recordset.MoveLast: grupdpalets.Recordset.MoveFirst
   carregar_combo_compatibles
End Sub
Sub activarframes(estat As Boolean)
  Grups.Enabled = estat
 
End Sub
Private Sub grupdpalets_Reposition()
grupdpalets.Caption = "Grups " + atrim(grupdpalets.Recordset.AbsolutePosition + 1) + " / " + atrim(grupdpalets.Recordset.RecordCount)
 If grupdpalets.Recordset.EditMode <> 3 Then activarframes False
 actualitzar_vinculats
End Sub
Sub actualitzar_vinculats()
 Set rst = dbtmpb.OpenRecordset("select descripcio from materials where codi=" + atrim(cadbl(grupdpalets.Recordset!codimatprognou)))
   If Not rst.EOF Then
      nommaterial = rst!descripcio
     Else: nommaterial = ""
   End If

End Sub

Private Sub modificar_Click()
  If grupdpalets.Recordset.EOF Then Exit Sub
    activarframes True
    grupdpalets.Recordset.Edit
    txtFields(5).SetFocus
End Sub

Private Sub seccio_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub seccio_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub sortir_Click()
  If grupdpalets.Recordset.EditMode > 0 Then grupdpalets.Recordset.Update
  Unload grupdepalets
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
   If Index = 0 And grupdpalets.Recordset.EditMode > 0 Then
   
      copiar_campsdelpalet
   End If
   If Index = 7 Then txtFields(7) = cadbl(txtFields(7))
End Sub
Sub copiar_campsdelpalet()
  Dim rstpalet As Recordset
  Dim nopassar As String
  
  nopassar = "nomdelgrupnumerogruppaletexempleidreservaseccioestocminimcodigrupmaterialscompatiblesnomgrupmaterialscompatibles"
  Set rstpalet = dbtmp.OpenRecordset("select * from palets where idpalet=" + atrim(cadbl(txtFields(0))))
  If Not rstpalet.EOF Then
      For i = 0 To grupdpalets.Recordset.Fields.Count - 1
        If InStr(1, nopassar, grupdpalets.Recordset.Fields(i).Name) = 0 Then
           grupdpalets.Recordset.Fields(i) = rstpalet.Fields(grupdpalets.Recordset.Fields(i).Name)
        End If
      Next i
      Command1_Click
      
     Else:
       MsgBox "Aquest palet no existeix"
       
  End If
End Sub

