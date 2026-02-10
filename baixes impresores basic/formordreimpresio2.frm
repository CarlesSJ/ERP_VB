VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formordreimpresio 
   BorderStyle     =   0  'None
   Caption         =   "Ordre Impressió"
   ClientHeight    =   10860
   ClientLeft      =   240
   ClientTop       =   900
   ClientWidth     =   9570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   10860
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox cimatgemissatgeCHAT_Verd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   7  'Invert
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   7320
      Picture         =   "formordreimpresio.frx":0000
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   42
      Top             =   135
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Frame framellegenda 
      Height          =   10890
      Left            =   3270
      TabIndex        =   20
      Top             =   10605
      Visible         =   0   'False
      Width           =   9570
      Begin VB.Frame Frame1 
         Caption         =   "Comandes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2385
         Left            =   165
         TabIndex        =   22
         Top             =   180
         Width           =   9030
         Begin VB.Frame Frame4 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Si el fondo ès:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1395
            Left            =   165
            TabIndex        =   25
            Top             =   495
            Width           =   8775
            Begin VB.Label Label10 
               BackColor       =   &H00F1B75F&
               Caption         =   "Blau: Falten metres assignats al packinglist."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3780
               TabIndex        =   38
               Top             =   1005
               Width           =   4740
            End
            Begin VB.Label Label5 
               BackColor       =   &H0017D062&
               Caption         =   "Verd : Si les bobines estan a impresores."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3780
               TabIndex        =   29
               Top             =   705
               Width           =   4740
            End
            Begin VB.Label Label4 
               BackColor       =   &H0000FFFF&
               Caption         =   "Groc : Bobines al magatzem d'impresores."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3795
               TabIndex        =   28
               Top             =   390
               Width           =   4710
            End
            Begin VB.Label Label3 
               BackColor       =   &H005C31DD&
               Caption         =   "Vermell: Bobines al magatzem."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   255
               TabIndex        =   27
               Top             =   780
               Width           =   3390
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Blanc: Bobines d'estoc."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   270
               TabIndex        =   26
               Top             =   465
               Width           =   3390
            End
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lletra Fucsia: PackingList no fet o No firmat."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   285
            Left            =   435
            TabIndex        =   45
            Top             =   1965
            Width           =   4935
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   """B""al davant: La comanda porta blanc."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   225
            TabIndex        =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   4020
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Texte d'impresió"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2130
         Left            =   210
         TabIndex        =   24
         Top             =   5085
         Width           =   8580
         Begin VB.Frame Frame7 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Si la lletra ès de color:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   225
            TabIndex        =   43
            Top             =   1290
            Width           =   8235
            Begin VB.Label Label12 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Lletra Fucsia: Falten firmes d'oficines."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   285
               Left            =   120
               TabIndex        =   44
               Top             =   285
               Width           =   3870
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00EAD9CE&
            Caption         =   "Si el fondo ès:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   225
            TabIndex        =   33
            Top             =   300
            Width           =   8250
            Begin VB.Label Label14 
               BackColor       =   &H000000FF&
               Caption         =   "Vermell: Falta marcar els clixes."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   3750
               TabIndex        =   46
               Top             =   645
               Width           =   3405
            End
            Begin VB.Label Label11 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Blanc: Comanda no muntada."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   135
               TabIndex        =   36
               Top             =   270
               Width           =   3390
            End
            Begin VB.Label Label9 
               BackColor       =   &H0000FFFF&
               Caption         =   "Groc : Comanda programada."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3765
               TabIndex        =   35
               Top             =   375
               Width           =   3390
            End
            Begin VB.Label Label8 
               BackColor       =   &H0017D062&
               Caption         =   "Verd : Comanda muntada."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   135
               TabIndex        =   34
               Top             =   585
               Width           =   3390
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tintes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   195
         TabIndex        =   23
         Top             =   2640
         Width           =   8970
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1515
            Left            =   405
            MultiLine       =   -1  'True
            TabIndex        =   32
            Text            =   "formordreimpresio.frx":01C5
            Top             =   690
            Width           =   7305
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   255
            Left            =   120
            Top             =   1290
            Width           =   270
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H0017D062&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   255
            Left            =   120
            Top             =   1560
            Width           =   270
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   255
            Left            =   120
            Top             =   1830
            Width           =   270
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   255
            Left            =   120
            Top             =   735
            Width           =   270
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   255
            Left            =   120
            Top             =   1020
            Width           =   270
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Situació de les tintes:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   255
            TabIndex        =   31
            Top             =   345
            Width           =   2445
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "D'acord"
         Height          =   525
         Left            =   7590
         TabIndex        =   21
         Top             =   10140
         Width           =   1605
      End
   End
   Begin VB.CommandButton bcomentariscomanda 
      BackColor       =   &H00FF80FF&
      Height          =   690
      Left            =   90
      Picture         =   "formordreimpresio.frx":02B5
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Comentaris de comanda pels operaris."
      Top             =   1575
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   135
      Top             =   6990
   End
   Begin VB.CommandButton botoPDFmodificacions 
      BackColor       =   &H0080FF80&
      Height          =   690
      Left            =   30
      Picture         =   "formordreimpresio.frx":05AF
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Veure full de modificacions MK."
      Top             =   705
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton bcanvimaquinareprint 
      BackColor       =   &H005C31DD&
      Height          =   690
      Left            =   0
      Picture         =   "formordreimpresio.frx":08A9
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Canvi de màquina REPRINT"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton bbobinesamaquina 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   8565
      Picture         =   "formordreimpresio.frx":1773
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2670
      Width           =   855
   End
   Begin VB.TextBox cpostit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1185
      Left            =   195
      Locked          =   -1  'True
      MouseIcon       =   "formordreimpresio.frx":1E2B
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   3930
      Visible         =   0   'False
      Width           =   8040
   End
   Begin VB.CommandButton Command17 
      Height          =   270
      Left            =   9165
      Picture         =   "formordreimpresio.frx":23B5
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Opcions d'Encarregat."
      Top             =   10515
      Width           =   360
   End
   Begin VB.CommandButton bcanvimaquina 
      Height          =   480
      Left            =   8580
      Picture         =   "formordreimpresio.frx":293F
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Canvia la màquina de la llista."
      Top             =   9810
      Width           =   915
   End
   Begin VB.Frame Framemodificacions 
      Height          =   3705
      Left            =   8550
      TabIndex        =   4
      Top             =   3405
      Visible         =   0   'False
      Width           =   975
      Begin VB.CommandButton Command7 
         Height          =   690
         Left            =   45
         Picture         =   "formordreimpresio.frx":2EC9
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Canvi de màquina de la comanda"
         Top             =   2955
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFC0&
         Height          =   705
         Left            =   60
         Picture         =   "formordreimpresio.frx":3D93
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2250
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFC0&
         Height          =   690
         Left            =   60
         Picture         =   "formordreimpresio.frx":4C5D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1545
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Height          =   705
         Left            =   60
         Picture         =   "formordreimpresio.frx":5B27
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   825
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Height          =   705
         Left            =   60
         Picture         =   "formordreimpresio.frx":69F1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Height          =   810
      Left            =   8580
      Picture         =   "formordreimpresio.frx":78BB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   39.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   8580
      Picture         =   "formordreimpresio.frx":80F4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton bimprimir 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   8565
      Picture         =   "formordreimpresio.frx":8999
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1815
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   10260
      Left            =   15
      TabIndex        =   0
      Top             =   105
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   18098
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame6 
      Caption         =   "Zoom"
      Height          =   585
      Left            =   8565
      TabIndex        =   39
      Top             =   9195
      Width           =   945
      Begin VB.CommandButton Command10 
         Height          =   300
         Left            =   510
         Picture         =   "formordreimpresio.frx":8CD4
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   225
         Width           =   345
      End
      Begin VB.CommandButton Command9 
         Height          =   300
         Left            =   75
         Picture         =   "formordreimpresio.frx":925E
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   225
         Width           =   345
      End
   End
   Begin VB.Label ethoraimpresio 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora d'impresió aprox.: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005C31DD&
      Height          =   270
      Left            =   195
      TabIndex        =   16
      Top             =   10590
      Width           =   8895
   End
   Begin VB.Label ethoresmuntades 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005C31DD&
      Height          =   270
      Left            =   165
      TabIndex        =   14
      Top             =   10335
      Width           =   3285
   End
   Begin VB.Label etmaquina 
      BackStyle       =   0  'Transparent
      Caption         =   "FW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005C31DD&
      Height          =   315
      Left            =   8745
      TabIndex        =   11
      Top             =   10260
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "Màq:"
      Height          =   195
      Left            =   8235
      TabIndex        =   10
      Top             =   10335
      Width           =   435
   End
   Begin VB.Menu mopcions 
      Caption         =   "Opcions"
      Begin VB.Menu mfingerprint 
         Caption         =   "Fer la comanda seleccionada com a FingerPrint."
      End
      Begin VB.Menu mprogramar 
         Caption         =   "Programar la comanda seleccionada (dia i hora)"
      End
      Begin VB.Menu mpossarcomentari 
         Caption         =   "Possar comentari a una comanda."
      End
      Begin VB.Menu mfeinesengparar 
         Caption         =   "Manteniment Feines Engegar- Parar Impresores"
      End
      Begin VB.Menu mllistatproduccio 
         Caption         =   "Llistat de producció."
      End
      Begin VB.Menu mclixesmuntats 
         Caption         =   "Llistat clixes muntats"
      End
   End
   Begin VB.Menu mchatencarregat 
      Caption         =   "CHAT ENCARREGAT"
   End
   Begin VB.Menu mllegenda 
      Caption         =   "Llegenda"
   End
End
Attribute VB_Name = "formordreimpresio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vtempsultimcanvi As Date
Dim dbplanificacio As Database
Dim vpasswordtintes As Boolean
Dim vX As Double
Dim vY As Double

Private Sub bbobinesamaquina_Click()
   Dim vnumc As Double
   
   vnumc = cadbl(numerodecomandaseleccionada)
  ' MsgBox atrim(vnumc)
   If notepackinglist(vnumc) Then
       If MsgBox("Tens seleccionada la comanda " + atrim(vnumc) + " es aquesta comanda la que vols Des/Marcar?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then GoTo fi
       dbtmpb.Execute "update impresores_ordreimpresio set estaamaquina=not [estaamaquina] where comanda=" + atrim(vnumc)
       carregar_reixaordre vnumc
   End If
fi:
End Sub
Function notepackinglist(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbstocks.OpenRecordset("SELECT Palets.Idpalet, bobines.sit,Bobines.Idbobina, materials.descripcio, Bobines.Numcomrev, Parcials.comanda FROM materials RIGHT JOIN (Palets LEFT JOIN (Bobines LEFT JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)) ON Palets.Idpalet = Bobines.Idpalet) ON materials.codi = Palets.codimatprognou WHERE (((Parcials.comanda)='" + atrim(vnumc) + "'));")
   If rst.EOF Then notepackinglist = True
   Set rst = Nothing
End Function

Private Sub bcanvimaquina_Click()
  Dim vmaq As String
  Dim vnommaq As String
  Unload formseleccio
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "SELECT codi,descripcio FROM maquines where maquina='I' and donadadebaixa=null and mid(descripcio,1,1)<>'#'"
  formseleccio.caption = "Escullir impresora"
  formseleccio.refrescar
  While Not formseleccio.Data1.Recordset.EOF
     If formseleccio.Data1.Recordset!codi = cadbl(etmaquina.tag) Then GoTo surtir
     formseleccio.Data1.Recordset.MoveNext
  Wend
surtir:
  formseleccio.DBGrid2.Columns(0).width = 1000
  formseleccio.DBGrid2.Columns(1).width = 5000
  'formseleccio.Left = 1000
  'formseleccio.Top = 1000
  formseleccio.Show 1
  If seleccioret = 1 Then
   vmaq = cadbl(formseleccio.Data1.Recordset!codi)
   vnommaq = atrim(formseleccio.Data1.Recordset!descripcio) + "    "
   vnommaq = Mid(vnommaq, 1, InStr(1, vnommaq, " "))
   etmaquina = vnommaq
   etmaquina.tag = vmaq
   Unload formseleccio
   DoEvents
   carregar_reixaordre
  End If
  Unload formseleccio
  
End Sub

Private Sub bcanvimaquinareprint_Click()
   Command7_Click
End Sub

Private Sub bcomentariscomanda_Click()
  mpossarcomentari_Click
End Sub

Private Sub bimprimir_Click()
  seleccioret = 5
  If Command5.visible = False Then
    dbtmpb.Execute "update muntadoratot set packinglistimpresaimpresores=true where comanda=" + atrim(cadbl(numerodecomandaseleccionada))
    dbtmpb.Execute "update impresores_ordreimpresio set imprespackinglist=true where comanda=" + atrim(cadbl(numerodecomandaseleccionada))
  End If
  'Me.Hide
End Sub

Private Sub bPDFmodificacions_Click()
  obrir_fitxer_modificacions numerodecomandaseleccionada
End Sub
Sub obrir_fitxer_modificacions(vnumc As Double)
   Dim vpdfmodifi As String
  ' carregaravisosmanteniment False
  ' avisosxrseccio.Show 1
  vpdfmodifi = rutamodifispdftreball(vnumc)
  If existeix(vpdfmodifi) Then obrir_document vpdfmodifi
End Sub
Function rutamodifispdftreball(vnumc As Double) As String
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
   If rst.EOF Then Exit Function
   On Error Resume Next
   MkDir ruta_documentacio_clixes + "\" + Format(rst!numtreball, "00000")
   rutamodifispdftreball = ruta_documentacio_clixes + "\" + Format(rst!numtreball, "00000") + "\MODIFI" + Format(rst!numtreball, "00000") + "-" + Format(rst!numordremodificacio, "000") + ".pdf"
   
End Function



Private Sub cimatgemissatgeCHAT_Verd_DblClick()
mchatencarregat_Click
End Sub

Sub canvilloc_iconaCHAT(X As Single, Y As Single)

  cimatgemissatgeCHAT_Verd.Left = X
  cimatgemissatgeCHAT_Verd.Top = Y
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  seleccioret = 0
  If Shift = 2 And Button = 2 Then seleccioret = 99
  gravar_posicio
  Unload veurereport
  Me.Hide
End Sub

Private Sub Command10_Click()
If reixa.Font.Size < 40 Then reixa.Font.Size = reixa.Font.Size + 1
End Sub

Private Sub Command17_Click()
  Dim vpassword As String
  vpassword = UCase(InputBoxEx("Entra la contrasenya de modificació.", "Entra contrasenya", , , , , , SPassword))
  If vpassword = "INPLACSA" Then  'contrasenya de L'OPERARI
      Framemodificacions.visible = True
      Framemodificacions.tag = "operari"
      Command17.visible = False
      Command5.visible = False
      Command6.visible = False
      Exit Sub
  End If
  If vpassword = "PICASO" Then  'contrasenya ENCARREGAT
      Framemodificacions.visible = True
'      mchatencarregat.visible = True
      Framemodificacions.tag = ""
      Command17.visible = False
      Command5.visible = True
      Command6.visible = True
      Exit Sub
  End If
  
  MsgBox "Contrasenya erronea.", vbCritical, "Error"
  
End Sub

Private Sub Command2_Click()
   Dim i As Integer
   Dim vnumc As Double
   Dim rst As Recordset
   Dim rstmodi As Recordset
   Dim vtexteimpresio As String
   Dim vtipuscomanda As String
   
   If cadbl(etmaquina.tag) = 0 Then MsgBox "No hi ha màquina seleccionada.", vbCritical, "Error": Exit Sub
   vnumc = cadbl(InputBox("Escriu el numero de comanda que vols afegir a la llista.", "Nova comanda"))
   If vnumc = 0 Then Exit Sub
   If Not esvalidaaquestacomanda(vnumc) Then Exit Sub
   
   Set rst = dbtmpb.OpenRecordset("select numtreball,impressio,numordremodificacio from comandes where comanda=" + atrim(vnumc))
   If rst.EOF Then Exit Sub
      'si la comadna es nova o modificada pregunto si ha revisat clixes i surto si no
   If UCase(rst!impressio) <> "R" Then
       If MsgBox("Aquesta comanda es nova o modificada." + Chr(10) + "HAS REVISAT ELS CLIXES JA?", vbExclamation + vbDefaultButton2 + vbYesNo, "NOVA/MODIFICADA") = vbNo Then Exit Sub
   End If
   vtipuscomanda = UCase(rst!impressio)
   Set rstmodi = dbtmp.OpenRecordset("select reimpres from modificacions where id_treball=" + atrim(cadbl(rst!numtreball)) + " and ordre=" + atrim(rst!numordremodificacio))
   Set rst = dbtmp.OpenRecordset("select id_treball,marca,linia from clixes where id_treball=" + atrim(rst!numtreball))
   If rst.EOF Or rstmodi.EOF Then MsgBox "Error al localitzar el Clixé", vbCritical, "Error": Exit Sub
   vtexteimpresio = atrim(rst!marca) + "-" + atrim(rst!linia)
   avisar_canvisoperari "Nova comanda entrada per un operari a impresores --> " + atrim(vnumc) + " Op:" + atrim(numop)
   vmetresminutultimcop = buscarmetresminutcomanda(cadbl(rst!id_treball))
   vmetresminutultimcop = Redondejar(cadbl(vmetresminutultimcop), 0)
   dbtmpb.Execute "insert into impresores_ordreimpresio (ordre,maquina,nommaquina,comanda,texteimpresio,metresminutultimcop,esreprint,modificada) values (999," + etmaquina.tag + ",'" + etmaquina + "'," + atrim(vnumc) + ",'" + vtexteimpresio + "'," + atrim(vmetresminutultimcop) + "," + IIf(rstmodi!reimpres, "True", "False") + ",'" + vtipuscomanda + "')"
   If Hour(Now) > 15 Then enviaremailgeneric "expedicions@inplacsa.com", "S'ha afegit la comanda " + atrim(vnumc) + " a l'ordre d'impressió de comandes a les " + atrim(Format(Now, "dd/mm/yy hh:nn")) + ".", ""
   'aixó s'ha d'activar quan sigui definitiu
   afegircomandaaplanificacioimpresoresoperaris vnumc, cadbl(etmaquina.tag), etmaquina
   dbtmpb.Execute "insert into muntadora_ordremuntatge (ordre,nummaquina,comanda,comandavisual) values (999,'" + etmaquina + "'," + atrim(vnumc) + ",'" + atrim(vnumc) + "-" + atrim(etmaquina) + "')"
   actualitzar_estatdelesllistes_demuntadoraiimpresora
   reordenallista
   carregar_reixaordre vnumc
   possar_comentaricomandasicorrespon vnumc
   Set rstmodi = Nothing
   Set rst = Nothing
End Sub
Sub possar_comentaricomandasicorrespon(vnumc As Double)
   Dim rstc As Recordset
   Dim rstcant As Recordset
   bcomentariscomanda.tag = ""
   Set rstc = dbtmp.OpenRecordset("select comandaduplicadade from comandes_extres where comanda=" + atrim(vnumc))
   If rstc.EOF Then Exit Sub
   'Set rstcant = dbtmp.OpenRecordset("SELECT comandes.comanda,comandes.materialex, trim(familiesmaterials.descripcio)&'-'& trim(subfamiliesmaterials.descripcio)&'-'& trim(familiescolorants.descripcio) as nommaterial FROM (((materials RIGHT JOIN comandes ON materials.codi = comandes.materialex) LEFT JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi) LEFT JOIN familiescolorants ON materials.familiacol = familiescolorants.codi WHERE (((comandes.comanda)=" + atrim(cadbl(rstc!comandaduplicadade)) + "));")
   'Set rstc = dbtmp.OpenRecordset("SELECT comandes.comanda,comandes.materialex, trim(familiesmaterials.descripcio)&'-'& trim(subfamiliesmaterials.descripcio)&'-'& trim(familiescolorants.descripcio) as nommaterial FROM (((materials RIGHT JOIN comandes ON materials.codi = comandes.materialex) LEFT JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi) LEFT JOIN familiescolorants ON materials.familiacol = familiescolorants.codi WHERE (((comandes.comanda)=" + atrim(vnumc) + "));")
   Set rstcant = dbtmp.OpenRecordset("SELECT comandes.comanda,comandes.materialex, trim(familiesmaterials.descripcio) as nommaterial FROM (((materials RIGHT JOIN comandes ON materials.codi = comandes.materialex) LEFT JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi) LEFT JOIN familiescolorants ON materials.familiacol = familiescolorants.codi WHERE (((comandes.comanda)=" + atrim(cadbl(rstc!comandaduplicadade)) + "));")
   Set rstc = dbtmp.OpenRecordset("SELECT comandes.comanda,comandes.materialex, trim(familiesmaterials.descripcio) as nommaterial FROM (((materials RIGHT JOIN comandes ON materials.codi = comandes.materialex) LEFT JOIN familiesmaterials ON materials.familia = familiesmaterials.codi) LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi) LEFT JOIN familiescolorants ON materials.familiacol = familiescolorants.codi WHERE (((comandes.comanda)=" + atrim(vnumc) + "));")
   If rstcant.EOF Or rstc.EOF Then Exit Sub
   If rstcant!materialex <> rstc!materialex Then
     bcomentariscomanda.tag = "CANVI DE MATERIAL:" + vbNewLine + "ANT: " + atrim(rstcant!nommaterial) + vbNewLine + "NOU: " + atrim(rstc!nommaterial) + vbNewLine
     mpossarcomentari_Click
   End If
   Set rstc = Nothing
   Set rstcant = Nothing
End Sub
Function buscarmetresminutcomanda(vnumtreball As Double) As Double
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select comanda from comandes where (proximaseccio<>'E' and proximaseccio<>'I') and numtreball=" + atrim(vnumtreball) + " order by comanda desc")
   If Not rst.EOF Then
       Set rst = dbtmpb.OpenRecordset("select metresmin from impressorestot where comanda=" + atrim(rst!comanda))
       If Not rst.EOF Then buscarmetresminutcomanda = cadbl(rst!metresmin)
   End If
   Set rst = Nothing
End Function
Sub afegircomandaaplanificacioimpresoresoperaris(numc As Double, vnummaq As Double, vnommaquina As String)
  Dim rstplanificacio As Recordset

  'vnummaq = IIf(vnommaquina = "FW", 7, 9)
  Set rstplanificacio = dbplanificacio.OpenRecordset("select * from planificacioimp where comanda=" + atrim(numc)) ' + " and maquina=" + atrim(vnummaq))
  If rstplanificacio.EOF Then
      dbplanificacio.Execute "insert into planificacioimp (comanda,ordre,maquina) values (" + atrim(numc) + ",998," + atrim(vnummaq) + ")"
     Else
         rstplanificacio.Edit
         If rstplanificacio!ordre = 999 Then rstplanificacio!ordre = 998
         rstplanificacio!maquina = vnummaq
         rstplanificacio.Update
  End If
  Set rstplanificacio = Nothing
End Sub
Function jaestaalallista(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio where comanda=" + atrim(vnumc), , ReadOnly)
   If Not rst.EOF Then jaestaalallista = True
End Function
Function esvalidaaquestacomanda(comanda As Double) As Boolean
    Dim msg As String
    esvalidaaquestacomanda = True
    If jaestaalallista(comanda) Then MsgBox "Aquesta comanda ja està entrada a la llista.", vbCritical, "Error": esvalidaaquestacomanda = False: GoTo fi
    If Not form1.comandavalida(cadbl(comanda), msg, True) Then
          '"Aquesta comanda ESTÀ PARADA O HI HA ALGUN MOTIU PER PARAR-LA."
        If InStr(1, UCase(msg), "FALTA AUTORITZAR") > 0 Then MsgBox msg, vbCritical, "Atenció": comanda = "0": esvalidaaquestacomanda = False: GoTo fi
        If MsgBox(msg + Chr(10) + "VOLS CONTINUAR IGUALMENT?", vbCritical + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbNo Then esvalidaaquestacomanda = False: GoTo fi
    End If
fi:
End Function
Sub avisar_canvisoperari(vmsg As String, Optional vaviscanviordre As Boolean)
  
   If Framemodificacions.tag = "operari" Then
       enviaremailgeneric "impresoresi@inplacsa.com;tintes@inplacsa.com;expedicions@inplacsa.com", vmsg, vmsg + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Data: " + atrim(Now)
   End If
   If UCase(arguments(3)) = "TINTES" Then
       enviaremailgeneric "impresores@inplacsa.com;tintes@inplacsa.com", "TINTES - " + vmsg, vmsg + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Data: " + atrim(Now)
   End If

   
End Sub

Sub reordenallista()
   Dim rst As Recordset
   Dim vgran As Double
   Dim vcont As Integer
   
   dbtmpb.Execute "UPDATE impresores_ordreimpresio RIGHT JOIN muntadora_ordremuntatge ON impresores_ordreimpresio.comanda = muntadora_ordremuntatge.comanda SET muntadora_ordremuntatge.ordre = [impresores_ordreimpresio].[ordre];"
   dbtmpb.Execute "UPDATE muntadora_ordremuntatge LEFT JOIN impresores_ordreimpresio ON muntadora_ordremuntatge.comanda = impresores_ordreimpresio.comanda SET muntadora_ordremuntatge.comandavisual = Trim([impresores_ordreimpresio].[comanda]) & ' [' & Format([impresores_ordreimpresio].[dataprogramada],'dd/mm hh:nn')&']-'+trim([muntadora_ordremuntatge].[nummaquina]) WHERE (((impresores_ordreimpresio.dataprogramada) Is Not Null));"
   
  ' Set rst = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio where ordre<999 order by ordre")
  ' If rst.EOF Then
  '    vgran = 1
  '     Else: rst.MoveLast: vgran = rst!ordre
  ' End If
   Set rst = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio where dataprogramada=null order by ordre")
   If rst.EOF Then Exit Sub
  ' rst.MoveLast
  ' If rst!ordre = 999 Then rst.Edit: rst!ordre = vgran + 1: rst.Update
  ' If vgran > 500 Then
      rst.MoveFirst
      vcont = 1
      While Not rst.EOF
        rst.Edit
        rst!ordre = vcont
        rst.Update
        vcont = vcont + 1
        rst.MoveNext
      Wend
   'End If
   Set rst = Nothing
   
   'aixó s 'ha d'activar quan ja començi a funcionar
   dbtmpb.Execute "UPDATE muntadorA_ordremuntatge INNER JOIN impresores_ordreimpresio ON muntadorA_ordremuntatge.comanda = impresores_ordreimpresio.comanda SET muntadora_ordremuntatge.ordre = [impresores_ordreimpresio].[ordre];"

End Sub
Private Sub Command3_Click()
  acceptar
End Sub
Sub acceptar()
Unload veurereport
  seleccioret = 1
  Me.Hide
End Sub

Private Sub Command4_Click()
   Dim vnumc As Double
   vnumc = cadbl(numerodecomandaseleccionada)
    If MsgBox("Segur que vols eliminar la comanda " + atrim(vnumc) + "?", vbExclamation + vbDefaultButton2 + vbYesNo, "Eliminar comanda de la llista") = vbYes Then
         dbtmpb.Execute "delete * from impresores_ordreimpresio where comanda=" + atrim(vnumc)
         avisar_canvisoperari "Comanda ELIMINADA per un operari a impresores --> " + atrim(vnumc) + " Op:" + atrim(numop)
       's'ha d'activar en el moment que ja sigui funcional
         dbtmpb.Execute "delete * from muntadora_ordremuntatge where comanda=" + atrim(vnumc)
        reordenallista
        carregar_reixaordre vnumc
        calcular_horesmuntades
    End If
End Sub
Function verificarpasswordtintes() As Boolean
  If UCase(arguments(3)) = "TINTES" Then
     If vpasswordtintes = False Then
          If UCase(InputBoxEx("Escriu el password per modificar l'ordre.", "Atenció", , , , , , SPassword)) = "INPLACSA" Then
             vpasswordtintes = True
              Else: Exit Function
          End If
     End If
       Else: vpasswordtintes = True
  End If
End Function
Private Sub Command5_Click()
    Dim v As Double
    Dim va As String
    v = 1
    If Framemodificacions.tag <> "operari" Then
        va = InputBox("De quan vols fer el salt? " + vbNewLine + "o escriu a la comanda on vols saltar.", "Salt de comandes", 1)
        v = cadbl(va)
        If v = 0 Then Exit Sub
    End If
  mourecomanda numerodecomandaseleccionada, "-", cadbl(v)
  
End Sub
Sub mourecomanda(vnumc As Double, vdireccio As String, Optional vsalts As Double)
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim vordre As Double
   verificarpasswordtintes
   If Not vpasswordtintes Then MsgBox "Error de password": Exit Sub
   
   If reixa.CellFontItalic Then Exit Sub
   Set rst = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio where maquina=" + atrim(etmaquina.tag) + " order by ordre")
   If vsalts > 999 Then
        Set rst2 = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio where maquina=" + atrim(etmaquina.tag) + " and comanda=" + atrim(vsalts) + " order by ordre")
        If rst2.EOF Then MsgBox "No he trobat la comanda on s'ha de saltar.", vbCritical, "Error": GoTo fi
        rst.FindFirst "comanda=" + atrim(vnumc)
        If Not rst.NoMatch Then
            rst.Edit
            rst!ordre = rst2!ordre - 0.5
            rst.Update
        End If
        GoTo fi
          Else: If vsalts > 0 Then vsalts = vsalts - 1
   End If
  
   If Not rst.NoMatch Then
       vordre = rst!ordre
       If vdireccio = "+" Then
            rst.MoveNext
          Else: rst.MovePrevious
       End If
       If Not rst.BOF And Not rst.EOF Then
          vordre = rst!ordre
       End If
       rst.FindFirst "comanda=" + atrim(vnumc)
       rst.Edit
       vordre = vordre + IIf(vdireccio = "+", vsalts + 0.5, (vsalts * -1) + -0.5)
       If vordre < 0 Then vordre = 0.5
       rst!ordre = vordre
       rst.Update
   End If
fi:
   If Not rst.NoMatch Then avisar_canvisoperari "Comanda CANVIADA D´ORDRE per un operari a impresores --> " + atrim(vnumc) + " Op:" + atrim(numop)
   
   reordenallista
   carregar_reixaordre vnumc
   Set rst = Nothing
   Set rst2 = Nothing
End Sub

Private Sub Command6_Click()
    Dim v As String
    v = 1
    If Framemodificacions.tag <> "operari" Then
        v = InputBox("De quan vols fer el salt? " + vbNewLine + "o escriu a la comanda on vols saltar.", "Salt de comandes", 1)
        If cadbl(v) = 0 Then Exit Sub
    End If
    mourecomanda numerodecomandaseleccionada, "+", cadbl(v)

End Sub

Private Sub Command7_Click()
   Dim vmaq As String
  Dim vnommaq As String
  Dim rst As Recordset
  Dim vnumc As Double
  vnumc = cadbl(numerodecomandaseleccionada)
  If vnumc = 0 Then MsgBox "Primer escull una comanda.", vbExclamation, "Atenció": Exit Sub
  Unload formseleccio
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "SELECT codi,descripcio FROM maquines where maquina='I' and donadadebaixa=null and mid(descripcio,1,1)<>'#'"
  formseleccio.caption = "Canvi d´impresora. Comanda: " + atrim(vnumc)
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 1000
  formseleccio.DBGrid2.Columns(1).width = 5000
  'formseleccio.Left = 1000
  'formseleccio.Top = 1000
  formseleccio.Show 1
  If seleccioret = 1 Then
   vmaq = cadbl(formseleccio.Data1.Recordset!codi)
   vnommaq = atrim(formseleccio.Data1.Recordset!descripcio) + "    "
   vnommaq = Mid(vnommaq, 1, InStr(1, vnommaq, " "))
   Set rst = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
       rst.Edit
       rst!ordre = 999
       rst!maquina = vmaq
       rst!nommaquina = vnommaq
       rst.Update
       avisar_canvisoperari "Comanda CANVIADA D´ORDRE per un operari a impresores --> " + atrim(vnumc) + " Op:" + atrim(numop)
   End If
    ' aixó s'ha d'activar quan sigui definitiu
   Set rst = dbtmpb.OpenRecordset("select * from muntadora_ordremuntatge where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
       rst.Edit
       rst!ordre = 999
       rst!nummaquina = vnommaq
       rst!comandavisual = atrim(vnumc) + "-" + vnommaq
       rst.Update
   End If
   carregar_reixaordre
  End If
  Unload formseleccio
End Sub
Sub calcular_horesmuntades()
   Dim rst As Recordset
   Dim vsql As String
   Dim vtotalhores As Double
   Dim vmetresminut As Double
   Dim vnommaquina As String
   
   ethoresmuntades = ""
   vsql = "SELECT comandes.cantitatex AS metres, impresores_ordreimpresio.nommaquina,impresores_ordreimpresio.metresminutultimcop "
   vsql = vsql + " FROM comandes RIGHT JOIN impresores_ordreimpresio ON comandes.comanda = impresores_ordreimpresio.comanda"
   vsql = vsql + " where (comandes.comanda In (select comanda from muntadoratot where acabada)) and maquina=9 "
   vsql = vsql + " ORDER BY impresores_ordreimpresio.nommaquina;"
calcular:
   Set rst = dbtmpb.OpenRecordset(vsql)
   vtotalhores = 0
   vnommaquina = ""
   If Not rst.EOF Then vnommaquina = IIf(atrim(rst!nommaquina) = "", "F2óFW", atrim(rst!nommaquina))
   While Not rst.EOF
     vmetresminut = 200
     If cadbl(rst!metresminutultimcop) > 0 Then vmetresminut = cadbl(rst!metresminutultimcop)
     'converteixo les comandes a minuts comptat 200metres per minut i faix 1,5h de canvi per comanda
     vtotalhores = vtotalhores + (cadbl(rst!metres) / vmetresminut) + 90
     
     rst.MoveNext
   Wend
   vtotalhores = Redondejar(vtotalhores / 60, 0)
   ethoresmuntades = ethoresmuntades + vnommaquina + ": " + atrim(vtotalhores) + "H. "
   If InStr(1, vsql, "maquina=7") = 0 Then
        vsql = "SELECT comandes.cantitatex AS metres, impresores_ordreimpresio.nommaquina,impresores_ordreimpresio.metresminutultimcop "
        vsql = vsql + " FROM comandes RIGHT JOIN impresores_ordreimpresio ON comandes.comanda = impresores_ordreimpresio.comanda"
        vsql = vsql + " where (comandes.comanda In (select comanda from muntadoratot where acabada)) and maquina=7 "
        vsql = vsql + " ORDER BY impresores_ordreimpresio.nommaquina;"
        GoTo calcular
   End If
   ethoresmuntades = ethoresmuntades
End Sub
Sub calcular_horesmuntades_NOVALID()
   Dim rst As Recordset
   Dim vsql As String
   Dim vtotalhores As Double
   Dim vmetresminut As Double
   
   ethoresmuntades = ""
   vsql = "SELECT Sum(comandes.cantitatex) AS Tmetres,  impresores_ordreimpresio.nommaquina, Count(comandes.comanda) AS Tcomandes "
   vsql = vsql + " FROM comandes RIGHT JOIN impresores_ordreimpresio ON comandes.comanda = impresores_ordreimpresio.comanda"
   vsql = vsql + " WHERE (((comandes.comanda) In (select comanda from muntadoratot where acabada)) AND ((comandes.proximaseccio)='I'))"
   vsql = vsql + " GROUP BY impresores_ordreimpresio.nommaquina;"

   Set rst = dbtmpb.OpenRecordset(vsql)
   While Not rst.EOF
     vmetresminut = 200
     If cadbl(rst!metresminutultimcop) > 0 Then vmetresminut = cadbl(rst!metresminutultimcop)
     'converteixo les comandes a minuts comptat 200metres per minut i faix 1,5h de canvi per comanda
     vtotalhores = (cadbl(rst!tmetres) / vmetresminut) + (90 * cadbl(rst!Tcomandes))
     vtotalhores = Redondejar(vtotalhores / 60, 0)
     vnommaquina = IIf(atrim(rst!nommaquina) = "", "F2óFW", atrim(rst!nommaquina))
     ethoresmuntades = ethoresmuntades + vnommaquina + ": " + atrim(vtotalhores) + "H. "
     rst.MoveNext
   Wend
   ethoresmuntades = ethoresmuntades
End Sub



Private Sub Command8_Click()
   framellegenda.visible = False
End Sub

Private Sub Command9_Click()
  If reixa.Font.Size > 8 Then reixa.Font.Size = reixa.Font.Size - 1
  
End Sub

Private Sub Form_Activate()
comprovarsihihamissatgesCHAT
  If arguments(1) = "ORDREIMPRESSIO" Then
    bimprimir.visible = False
    Framemodificacions.visible = False
    Command17.visible = True
    If UCase(arguments(3)) = "TINTES" Then
        Framemodificacions.visible = True
        Command17.visible = False
    End If
  End If
  If existeix("c:\ordprog.ini") Then
    Framemodificacions.visible = True
    'mchatencarregat.visible = True
    Command17.visible = False
  End If
  If Not esmuntadora Then
    If UCase(arguments(1)) = "ORDREIMPRESSIO" Then
        Timer1.Enabled = True
        If cadbl(llegir_ini("PosicioOrdreimpresio-" + atrim(nummaq), "left", "comandes.ini")) <> 0 Then
           'If cadbl(llegir_ini("PosicioOrdreimpresio-" + atrim(nummaq), "left", "comandes.ini")) < Screen.width Then
            formordreimpresio.Left = cadbl(llegir_ini("PosicioOrdreimpresio-" + atrim(nummaq), "left", "comandes.ini"))
            formordreimpresio.Top = cadbl(llegir_ini("PosicioOrdreimpresio-" + atrim(nummaq), "top", "comandes.ini"))
           'End If
        End If
          Else:
           If UCase(arguments(1)) <> "DESBOBINADORS" Then
            Timer1.Enabled = False
            formordreimpresio.Top = 225
            formordreimpresio.Left = 0
           End If
    End If
      Else
        formordreimpresio.Top = 500
        formordreimpresio.Left = 0
        If etmaquina.tag = "" Then bcanvimaquina_Click
  End If
  comprovarsihihamissatgesCHAT
End Sub
Function numerodecomandaseleccionada() As String
  Dim vnumc As String
  vnumc = reixa.TextMatrix(reixa.row, 0)
  If Not IsNumeric(Mid(vnumc, 1, 1)) Then vnumc = Mid(vnumc, 2)
numerodecomandaseleccionada = cadbl(vnumc)
End Function

Private Sub Form_Click()
'  MsgBox unaBsiteblanc(206709)
End Sub

Private Sub Form_Load()
  Dim rst As Recordset
  Dim vnommaq As String
  'Framemodificacions.visible = True
  
  framellegenda.visible = False
  Set rst = dbtmp.OpenRecordset("select descripcio from maquines where maquina='I' and codi=" + atrim(nummaq), , ReadOnly)
  If Not rst.EOF And Not esmuntadora Then
      vnommaq = atrim(rst!descripcio) + "    "
      vnommaq = Mid(vnommaq, 1, InStr(1, vnommaq, " "))
      etmaquina = vnommaq
      etmaquina.tag = nummaq
  End If
  Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificaciooperaris.mdb")
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
  actualitzar_estatdelesllistes_demuntadoraiimpresora
  
  Set rst = Nothing
  carregar_reixaordre
  
  formordreimpresio.visible = False
End Sub
Sub actualitzar_estatdelesllistes_demuntadoraiimpresora()
  dbtmpb.Execute "UPDATE muntadoratot INNER JOIN impresores_ordreimpresio ON muntadoratot.comanda = impresores_ordreimpresio.comanda SET impresores_ordreimpresio.muntada = [muntadoratot].[acabada];"
  dbtmpb.Execute "UPDATE muntadoratot INNER JOIN muntadora_ordremuntatge ON muntadoratot.comanda = muntadora_ordremuntatge.comanda SET muntadora_ordremuntatge.muntada = [muntadoratot].[acabada];"
  'actualitzo les comandes de les llistes si n'hi ha alguna de ja feta l'elimino de la llista
  dbtmpb.Execute "DELETE impresores_ordreimpresio.* FROM comandes RIGHT JOIN impresores_ordreimpresio ON comandes.comanda = impresores_ordreimpresio.comanda WHERE (((comandes.proximaseccio)<>'E' And (comandes.proximaseccio)<>'I'));"
  dbtmpb.Execute "DELETE muntadora_ordremuntatge.* FROM comandes RIGHT JOIN muntadora_ordremuntatge ON comandes.comanda = muntadora_ordremuntatge.comanda WHERE (((comandes.proximaseccio)<>'E' And (comandes.proximaseccio)<>'I'));"
  calcular_horesmuntades
End Sub
Function tempsxrcomanda(vnumc As Double, vmetresxrminut As Double) As Double
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select cantitatex from comandes where comanda=" + atrim(vnumc), , ReadOnly)
  If rst.EOF Then Exit Function
  If vmetresxrminut = 0 Then vmetresxrminut = 200
  tempsxrcomanda = (cadbl(rst!cantitatex) / vmetresxrminut) + 90
  Set rst = Nothing
End Function
Function unaBsiteblanc(vnumc As String) As String
   Dim rst As Recordset
   Dim vmodificacio As String
   Dim vtreball As String
   Dim vsql As String
'   If vdesplaçant Then
'      Exit Function
'   End If
   
   Set rst = dbtmp.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
   If rst.EOF Then Exit Function
   vtreball = rst!numtreball: vmodificacio = rst!numordremodificacio
   
   Set rst = dbtintes.OpenRecordset("SELECT Tintesclixesnous.id_treball, Tintesclixesnous.ordremodificacio, familiescolors.descripcio FROM (Tintesclixesnous LEFT JOIN tintes ON Tintesclixesnous.coditinta = tintes.codi) LEFT JOIN familiescolors ON tintes.idfamcolor = familiescolors.codi WHERE (((Tintesclixesnous.id_treball)=" + vtreball + ") AND ((Tintesclixesnous.ordremodificacio)=" + vmodificacio + ") AND ((familiescolors.descripcio)='BLANCO'));")
   If Not rst.EOF Then unaBsiteblanc = "B "
   If unaBsiteblanc = "" Then
      vsql = "SELECT tinterlinkambid_treball FROM (Tintesclixesnous LEFT JOIN tintes ON Tintesclixesnous.coditinta = tintes.codi) LEFT JOIN familiescolors ON tintes.idfamcolor = familiescolors.codi WHERE (Tintesclixesnous.id_treball=" + vtreball + " AND Tintesclixesnous.ordremodificacio=" + vmodificacio + ") "
      Set rst = dbtintes.OpenRecordset("SELECT Tintesclixesnous.id_treball, Tintesclixesnous.ordremodificacio, familiescolors.descripcio FROM (Tintesclixesnous LEFT JOIN tintes ON Tintesclixesnous.coditinta = tintes.codi) LEFT JOIN familiescolors ON tintes.idfamcolor = familiescolors.codi WHERE familiescolors.descripcio='BLANCO' AND id_tinter IN (" + vsql + ");")
'      Clipboard.Clear
'      Clipboard.SetText "SELECT Tintesclixesnous.id_treball, Tintesclixesnous.ordremodificacio, familiescolors.descripcio FROM (Tintesclixesnous LEFT JOIN tintes ON Tintesclixesnous.coditinta = tintes.codi) LEFT JOIN familiescolors ON tintes.idfamcolor = familiescolors.codi WHERE familiescolors.descripcio='BLANCO' AND id_tinter IN (" + vsql + ");"
      If Not rst.EOF Then unaBsiteblanc = "B "
   End If
   Set rst = Nothing
End Function
Sub carregar_reixaordre(Optional vnumc As Double)
  Dim rst As Recordset
  Dim vcolor As Double
  Dim vrow As Integer
  Dim vfiladelacomanda As Integer
  Dim vmaq As Integer
  Dim vdataimpresio As Date
  Dim vtoteslesbobines As Byte
  Dim rsttemps As Recordset
  
  vtempsultimcanvi = Now
  vmaq = cadbl(etmaquina.tag)
  If vmaq < 1 Then Exit Sub
  reixa.visible = False
  
  configurar_reixa
  Set rst = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio where maquina=" + atrim(vmaq) + " order by ordre,dataprogramada")
  Set rsttemps = dbtmpb.OpenRecordset("select * from horarisimpresores where maquina=7 and year(dataihora)=year(now) order by dataihora")
  vrow = 1
  vdataimpresio = Now
  While Not rst.EOF
    reixa.Rows = vrow + 1
    reixa.row = vrow
    If rst!comanda = vnumc Then vfiladelacomanda = vrow
    If Not esmuntadora Then
       'reixa.TextMatrix(vrow, 0) = unaBsiteblanc(rst!comanda) + atrim(rst!comanda)
        Else: reixa.TextMatrix(vrow, 0) = atrim(rst!comanda)
    End If
    reixa.TextMatrix(vrow, 0) = atrim(rst!comanda)
    reixa.TextMatrix(vrow, 2) = atrim(rst!texteimpresio)
    If rst!ordre > 0 And Not esmuntadora Then
        vdataimpresio = DateAdd("n", tempsxrcomanda(rst!comanda, cadbl(rst!metresminutultimcop)), vdataimpresio)
        'vdataimpresio = passardataaadiadetreball(vdataimpresio, rsttemps)
        
        If Format(vdataimpresio, "w", vbSunday) = 1 Then  'si es diumenge afegeix un dia
          vdataimpresio = DateAdd("d", 1, vdataimpresio)
        End If
        reixa.TextMatrix(vrow, 3) = atrim(vdataimpresio)
        If Not esmuntadora Then dbtmpb.Execute "update impresores_ordreimpresio set dataprevistaimpresiocalculada=#" + Format(vdataimpresio, "mm/dd/yy hh:nn") + "# where comanda=" + atrim(rst!comanda)
    End If
    If rst!muntada Then
       reixa.col = 2: reixa.CellBackColor = &H6BEBB1        'verd xulu si està muntada
       reixa.col = 0: reixa.CellBackColor = &H6BEBB1        'verd xulu si està muntada
    End If
    'posso l'estat de gestió de la tinta
    reixa.col = 1
    vcolor = QBColor(15)
    vgestiotinta = estatdelagestiodelatinta(rst!comanda)
 '15/02/22 no volen color a la columna d'estat de tinta
    If vgestiotinta = "S" Then vcolor = &H5C31DD
    If vgestiotinta = "F" Then vcolor = &H5C31DD
    If vgestiotinta = "M" Then vcolor = &H6BEBB1
    If vgestiotinta = "P" Then vcolor = &H80C0FF
    If vgestiotinta = "N" Then vcolor = &H5C31DD
    reixa.TextMatrix(vrow, 1) = vgestiotinta
    reixa.CellBackColor = vcolor
    
    reixa.col = 0
    vtoteslesbobines = hihatoteslesbobines(rst!comanda)
    If vtoteslesbobines = 4 Then reixa.CellBackColor = &HF3B378 ' blau si no hi ha prou material
    If vtoteslesbobines = 3 Then
       reixa.CellBackColor = QBColor(15) 'si no hi ha res assignat es quea en BLANC (ESTOC per exemple)
       If rst!estaamaquina Then vtoteslesbobines = 2  'si han marcat que tenen les bobines a maquina ho passo a verd
    End If
    If vtoteslesbobines = 2 Then reixa.CellBackColor = &H80FF80 'si hi ha totes les bobines A MAQUINA  poso la comanda amb verd txillon
    If vtoteslesbobines = 1 Then reixa.CellBackColor = QBColor(14) 'si les bobines estan baixades PERO NO AESTAN A MAQUINA ho posso en groc
    If vtoteslesbobines = 0 Then reixa.CellBackColor = &H5C31DD 'si NO ESTAN BAIXADES totes les bobines poso la comanda amb vermell
   ' If rst!imprespackinglist Then reixa.col = 0: reixa.CellBackColor = &H5C31DD 'vermell xulu si ja han impres packinglist
    If IsDate(rst!dataprogramada) Then
       reixa.col = 2: reixa.CellFontBold = True: reixa.CellFontItalic = True: reixa.CellForeColor = QBColor(14)
       If reixa.CellBackColor = 0 Then reixa.CellBackColor = QBColor(8)
       reixa.col = 0: reixa.CellFontBold = True: reixa.CellFontItalic = True: reixa.CellForeColor = QBColor(0)
       If reixa.CellBackColor = 0 Then reixa.CellBackColor = QBColor(8)
    End If
    If clixesnomarcats(rst!comanda) Then
         reixa.col = 2: reixa.CellBackColor = &H5C31DD: reixa.col = 0
    End If
    If nohihapackinglistonorevisat(rst!comanda) Then
         reixa.col = 0: reixa.CellForeColor = QBColor(13): reixa.col = 0
    End If
    If faltenfirmesaoficina(rst!comanda) Then
         reixa.col = 2: reixa.CellForeColor = QBColor(13): reixa.col = 0
    End If
    vrow = vrow + 1
    rst.MoveNext
  Wend
  reixa.visible = True
'  MsgBox reixa.TextMatrix(2, 0)
  If vfiladelacomanda > 0 Then
     reixa.SetFocus
     reixa.row = vfiladelacomanda
     reixa.RowSel = vfiladelacomanda
     reixa.col = 0: reixa.ColSel = reixa.Cols - 1
     If Not reixa.RowIsVisible(IIf(reixa.row + 1 < reixa.Rows, reixa.row + 1, reixa.row)) And (reixa.row - 5) > 0 Then reixa.TopRow = reixa.row - 5
  End If
  Set rst = Nothing
  If reixa.Rows = 1 Then reixa.Rows = 2
  If vnumc = 0 Then reixa.row = 1
  Set rsttemps = Nothing
End Sub
Function nohihapackinglistonorevisat(vnumc As Double) As Boolean
  Dim rst As Recordset
  Set rst = dbtmp.OpenRecordset("select * from comandes_firmes where comanda=" + atrim(vnumc) + " and tipus='PK2' and not anulada")
  If rst.EOF Then nohihapackinglistonorevisat = True
  Set rst = Nothing
End Function

Function faltenfirmesaoficina(vnumc As Double) As Boolean
  Dim rst As Recordset
  Dim vtipusimpressio As String
  Dim vwhere As String
  Set rst = dbtmp.OpenRecordset("select impressio from comandes where comanda=" + atrim(vnumc))
  If rst.EOF Then GoTo fi
  vtipusimpressio = atrim(rst!impressio)
  vwhere = "tipus='INI' or tipus='TEC' or tipus='COM'"
  If vtipusimpressio = "N" Or vtipusimpressio = "M" Then vwhere = vwhere + " or tipus='IM2' or tipus='PK2'"
  Set rst = dbtmp.OpenRecordset("select * from comandes_firmes where comanda=" + atrim(vnumc) + " and not anulada and (" + vwhere + ")")
  If rst.EOF Then faltenfirmesaoficina = True
fi:
  Set rst = Nothing
End Function

Function clixesnomarcats(vnumc As Double) As Boolean
  Dim rst As Recordset
  clixesnomarcats = True
  Set rst = dbtmp.OpenRecordset("select numtreball,numordremodificacio,impressio from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
      If rst!impressio = "R" Then clixesnomarcats = False: GoTo fi
      Set rst = dbtmpb.OpenRecordset("select datarepas from clixesentrats_control where numtreball=" + atrim(rst!numtreball) + " and versio=" + atrim(rst!numordremodificacio))
      If Not rst.EOF Then
          If Not IsNull(rst!datarepas) Then clixesnomarcats = False
      End If
  End If
fi:
  Set rst = Nothing
End Function
Function passardataaadiadetreball(vdata As Date, rsttemps As Recordset) As Date
   Dim i As Byte
   i = 0
   rsttemps.FindFirst "month(dataihora)=" + atrim(Month(vdata))
   While rsttemps.NoMatch And (Month(vdata) + i) < 13
     i = i + 1
     rsttemps.FindFirst "month(dataihora)=" + atrim(Month(DateAdd("m", i, vdata)))
   Wend
   If rsttemps.NoMatch Then Exit Function

   While rsttemps!dataihora < vdata And Not rsttemps.EOF
     If Not rsttemps.EOF Then rsttemps.MoveNext
   Wend
   If Not rsttemps.EOF Then
     If DateDiff("n", vdata, rsttemps!dataihora) > 60 Then
        passardataaadiadetreball = Format(rsttemps!dataihora, "dd/mm/yy") + " " + Format(vdata, "hh:nn")
          Else: passardataaadiadetreball = vdata
     End If
       Else: passardataaadiadetreball = vdata
   End If
End Function
Function estatdelagestiodelatinta(vnumc As Double)
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select estatgestio from comandesrevisadesatintes where comanda=" + atrim(vnumc))
   If Not rst.EOF Then estatdelagestiodelatinta = rst!estatgestio
   Set rst = Nothing
End Function
Function hihatoteslesbobines(vnumc As Double) As Byte
   Dim rst As Recordset
   Dim rstb As Recordset
   Dim vnotrobat As Boolean
   Dim valgunaNOIMP  As Boolean
   Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "Palets.mdb")
   Set rstb = dbtmpb.OpenRecordset("select top 300 * from impresores_bobinesamaquina order by data desc")
   Set rst = dbstocks.OpenRecordset("SELECT Palets.Idpalet, bobines.sit,Bobines.Idbobina, materials.descripcio, Bobines.Numcomrev, Parcials.comanda FROM materials RIGHT JOIN (Palets LEFT JOIN (Bobines LEFT JOIN Parcials ON (Bobines.Idbobina = Parcials.idbobina) AND (Bobines.Idpalet = Parcials.idpalet)) ON Palets.Idpalet = Bobines.Idpalet) ON materials.codi = Palets.codimatprognou WHERE (((Parcials.comanda)='" + atrim(vnumc) + "'));")
   
   If rst.EOF Then
     hihatoteslesbobines = 3: GoTo fi
   End If
   While Not rst.EOF And Not vnotrobat
     rstb.FindFirst "numbobina='" + atrim(rst!idpalet) + "/" + atrim(rst!idbobina) + "'"
     If rstb.NoMatch Then vnotrobat = True
     If InStr(1, UCase(atrim(rst!sit)), "IMP") = 0 Then valgunaNOIMP = True  'SI NO ESTÀ AMB IMP ACTIVO LA VARIALBE PER POSAR AMB VERMELL A PANTALLA
     rst.MoveNext
   Wend
   hihatoteslesbobines = IIf(Not vnotrobat, 2, 1)
   If valgunaNOIMP Then hihatoteslesbobines = 0
fi:
   If nohihaproumetresassignats(vnumc) And Not comandacomençada(vnumc) Then hihatoteslesbobines = 4
   Set rst = Nothing
   Set rstb = Nothing
End Function
Function comandacomençada(vnumc As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select * from impressores where tipus='F' and comanda=" + atrim(vnumc))
   If Not rst.EOF Then comandacomençada = True
   Set rst = Nothing
End Function
Function nohihaproumetresassignats(vnumc As Double) As Boolean
   Dim rstp As Recordset
   Dim vmetresactualts As Double
   Dim vmetresassignats As Double
   
   Set rstp = dbtmp.OpenRecordset("select * from comandes_extres where comanda=" + atrim(vnumc))
   If Not rstp.EOF Then vmetresassignats = cadbl(rstp!metresassignatspackinglist)
   If vmetresassignats = 0 Then GoTo fi
   Set rstp = dbstocks.OpenRecordset("select sum(metres) as Tmetres from parcials where utilitzada=false and comanda='" + atrim(vnumc) + "'")
   If Not rstp.EOF Then vmetresactuals = cadbl(rstp!tmetres)
   
   If vmetresactuals < vmetresassignats Then
       nohihaproumetresassignats = True
   End If
   If Not nohihaproumetresassignats Then
     Set rstp = dbstocks.OpenRecordset("select * from parcials where not utilitzada and comanda='" + atrim(vnumc) + "'")
     While Not rstp.EOF
        If bobinesdentrada.calcular_mtrsdispreals(rstp!idpalet, rstp!idbobina) < rstp!metres Then nohihaproumetresassignats = True
        rstp.MoveNext
     Wend
   End If
fi:
   Set rstp = Nothing
End Function
Sub configurar_reixa()
  reixa.Clear
  reixa.Rows = 1
  reixa.Cols = 4
  reixa.TextMatrix(0, 0) = "Comanda"
  reixa.TextMatrix(0, 1) = "Tintes"
  reixa.TextMatrix(0, 2) = "Texte d 'impresió"
  reixa.TextMatrix(0, 3) = "Hora Prevista"
  reixa.ColWidth(0) = 1700
  reixa.ColWidth(1) = 500
  reixa.ColWidth(2) = 8200
  reixa.ColWidth(3) = 0
  reixa.ColAlignment(0) = 7
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub MSFlexGrid1_DblClick()
  acceptar
End Sub

Sub gravar_posicio()
  If UCase(arguments(1)) = "ORDREIMPRESSIO" Then
      escriure_ini "PosicioOrdreimpresio-" + atrim(nummaq), "left", formordreimpresio.Left, "comandes.ini"
      escriure_ini "PosicioOrdreimpresio-" + atrim(nummaq), "top", formordreimpresio.Top, "comandes.ini"
  End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload veurereport
End Sub

Private Sub mchatencarregat_Click()
    If atrim(UCase(InputBoxEx("Entra la contrasenya per accedir al XAT", "CONTRASENYA", , , , , , SPassword))) <> "LEONARDO" Then Exit Sub
    formCHAT.carregar_missatges_operari "I", 0
    formCHAT.Show 1
    form1.comprovarsihihamissatgesCHAT
End Sub

Private Sub mclixesmuntats_Click()
  formclixesmuntats.Show 1
End Sub

Private Sub mfeinesengparar_Click()
form1.feines_parar_engegar_maquina "ENGEGAR", cadbl(etmaquina.tag)
End Sub

Private Sub mfingerprint_Click()
  seleccioret = 2
  Me.Hide
End Sub

Private Sub mllegenda_Click()
  framellegenda.Top = 0
  framellegenda.Left = 15
  
  framellegenda.visible = True
End Sub

Private Sub mllistatproduccio_Click()
  If UCase(InputBox("Escriu la contrasenya per accedir.", "Llistat de producció.")) <> "PICASO" Then Exit Sub
  Load Llistatproduccio
  Llistatproduccio.seccio.Enabled = False
  Llistatproduccio.combooperaris.Enabled = False
  Llistatproduccio.CheckExcel.Enabled = False
  Llistatproduccio.Checkexcelnou.Enabled = False
  Llistatproduccio.acceptar.Enabled = False
  Llistatproduccio.Show 1
End Sub

Private Sub mpossarcomentari_Click()
  Dim vnumc As Double
  Dim rst As Recordset
  vnumc = cadbl(numerodecomandaseleccionada)
  Load obsidtreball
  Set rst = dbtmpb.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
            Set rst = dbtmpb.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(rst!numtreball) + " and ordre=" + atrim(rst!numordremodificacio) + " order by id")
            While Not rst.EOF
              obsidtreball.obsid = obsidtreball.obsid + vbNewLine + atrim(rst!observacio)
              rst.MoveNext
            Wend
  End If
  Set rst = dbtmpb.OpenRecordset("select * from Impresores_ObsXoperaris where numcomanda=" + atrim(vnumc))
  obsidtreball.caption = "Comentari de la comanda " + atrim(vnumc) + IIf(Not Framemodificacions.visible, " (NOMÉS LECTURA)", "")
  If Not rst.EOF Then
      obsidtreball.obsid = obsidtreball.obsid + vbNewLine + vbNewLine + "[Observacions editables a partir d'aquí.] " + vbNewLine + "[No borrar] Comentari de la comanda: " + vbNewLine + rst!obscomanda
       Else: obsidtreball.obsid = obsidtreball.obsid + vbNewLine + "[No borrar] Comentari de la comanda: " + vbNewLine
  End If
  If bcomentariscomanda.tag <> "" Then obsidtreball.obsid = bcomentariscomanda.tag
  obsidtreball.obsid = obsidtreball.obsid + vbNewLine
  obsidtreball.obsid.SelStart = Len(obsidtreball.obsid)
  obsidtreball.obsid.SelLength = Len(obsidtreball.obsid)
  obsidtreball.Show 1
  If Not Framemodificacions.visible Then GoTo fi
  If InStr(1, r, "Comentari de la comanda: ") = 0 Then GoTo fi
  r = Mid(r, InStr(1, r, "Comentari de la comanda: ") + 26)
  If atrim(r) = "" Then
        If Not rst.EOF Then: rst.Delete
        GoTo fi
  End If
  r = treure_entersiniciifi(r)
  If r = "" Then
    If Not rst.EOF Then rst.Delete
    GoTo fi
  End If
  If Not rst.EOF Then
       rst.Edit
           Else: rst.AddNew
  End If
  rst!numcomanda = vnumc
  rst!obscomanda = r
  rst.Update
fi:
  Set rst = Nothing
  bcomentariscomanda.tag = ""
End Sub
Function treure_entersiniciifi(r) As String
   Dim i As Byte
   Dim v As String
   v = r
   While Len(v) > 1
      If Asc(Mid(v, 1, 1)) < 32 Then
              v = Mid(v, 2)
             Else: GoTo treureultims
      End If
   Wend
treureultims:
   While Len(v) >= 1
      If Asc(Mid(v, Len(v), 1)) < 32 Then
              v = Mid(v, 1, Len(v) - 1)
          Else: GoTo fi
      End If
   Wend
fi:
   treure_entersiniciifi = v
End Function
Private Sub mprogramar_Click()
  Dim vdata As String
  Dim vhora As String
  Dim vmotiu As String
  Dim vnumc As Double
  Dim rst As Recordset
  vnumc = cadbl(numerodecomandaseleccionada)
  'vnumc = formordreimpresio.reixa.TextMatrix(formordreimpresio.reixa.row, 0)
  If vnumc = 0 Then MsgBox "No hi ha cap comanda selecionada.", vbExclamation, "Error": Exit Sub
  vdata = InputBox("Entra la data que vols fer la impresió. ex:01/10/17" + Chr(10) + "ESCRIU [BORRAR] PER BORRAR LA DATA DE PROGRAMACIÓ.", "Entra data programació", Format(DateAdd("d", 1, Now), "dd/mm/yy"))
  If UCase(vdata) = "BORRAR" Then vdata = "": GoTo gravar
  If Not IsDate(vdata) Then Exit Sub
  vhora = InputBox("Entra la hora que vols fer la impresió. ex:09:00", "Entra la hora de programació", "09:00")
  If Not IsDate(vdata + " " + vhora) Then MsgBox "Data o hora mal entrada o erronea.", vbCritical, "Error"
  vmotiu = InputBox("Explica el motiu perquè es fa aquesta programació.", "Programació impresió")
gravar:
  Set rst = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio where comanda=" + atrim(vnumc))
  If rst.EOF Then MsgBox "Comanda no trobada.": GoTo fi
  If vdata <> "" Then
    rst.Edit
    rst!dataprogramada = CVDate(vdata + " " + vhora)
    rst!ordre = 0
    rst!motiuprogramacio = treure_apostruf(Mid(vmotiu, 1, 255))
    rst.Update
      Else:
        dbtmpb.Execute "update impresores_ordreimpresio set ordre=999,motiuprogramacio='',dataprogramada=null where comanda=" + atrim(vnumc)
        MsgBox "Programació eliminada.", vbInformation, "Borrar programació."
        GoTo fi
  End If
  
fi:
  Set rst = Nothing
  carregar_reixaordre vnumc
End Sub
Sub canvimaquinaREPRINT(vActivar As Boolean)
   bcanvimaquinareprint.Left = reixa.width - bcanvimaquinareprint.width '- IIf(bPDFmodificacions.visible, bPDFmodificacions.width + 10, 0)
   bcanvimaquinareprint.Top = reixa.CellTop
   bcanvimaquinareprint.visible = vActivar
End Sub
Sub botoPD+Fmodificacions(vActivar As Boolean)
   bPDFmodificacions.Left = reixa.width - bPDFmodificacions.width - IIf(bcanvimaquinareprint.visible, bcanvimaquinareprint.width + 10, 0)
   bPDFmodificacions.Top = reixa.CellTop
   bPDFmodificacions.visible = vActivar
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub reixa_DblClick()
 If UCase(arguments(1)) = "ORDREIMPRESSIO" Then Exit Sub
If vX > 0 And vX < reixa.ColWidth(0) Then
   Unload veurereport
   form1.imprimir_packinglistTICKET cadbl(reixa.TextMatrix(reixa.row, 0)), True
   veurereport.Top = formordreimpresio.Top
   If UCase(arguments(1)) = "DESBOBINADORS" Then
       veurereport.Left = formordreimpresio.Left
       
        Else: veurereport.Left = formordreimpresio.Left + formordreimpresio.width + 100
   End If
End If
End Sub

Private Sub reixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vX = X
    vY = Y
End Sub
Sub botoCOMENTARISCOMADA(vnumc As Double)
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select * from Impresores_ObsXoperaris where numcomanda=" + atrim(vnumc))
   bcomentariscomanda.visible = False
   If Not rst.EOF Then
      bcomentariscomanda.Left = reixa.width - bcomentariscomanda.width - IIf(bcanvimaquinareprint.visible, bcanvimaquinareprint.width + 10, 0) - IIf(bPDFmodificacions.visible, bPDFmodificacions.width + 10, 0)
      bcomentariscomanda.Top = reixa.CellTop
      bcomentariscomanda.visible = True
       Else
          Set rst = dbtmpb.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
          If Not rst.EOF Then
            Set rst = dbtmpb.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(rst!numtreball) + " and ordre=" + atrim(rst!numordremodificacio) + " order by ordre")
            If Not rst.EOF Then
                bcomentariscomanda.Left = reixa.width - bcomentariscomanda.width - IIf(bcanvimaquinareprint.visible, bcanvimaquinareprint.width + 10, 0) - IIf(bPDFmodificacions.visible, bPDFmodificacions.width + 10, 0)
                bcomentariscomanda.Top = reixa.CellTop
                bcomentariscomanda.visible = True
            End If
          End If
   End If
   Set rst = Nothing
End Sub
Sub reixa_refrescarfila()
   Dim rst As Recordset
   Dim vdataf As String
   If formordreimpresio.visible = False Then Exit Sub
   canvimaquinaREPRINT False
   botoPDFmodificacions False
   ethoraimpresio = ""
   Set rst = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio where comanda=" + atrim(cadbl(numerodecomandaseleccionada)))
   If reixa.CellFontItalic Then
    'Set rst = dbtmpb.OpenRecordset("select * from impresores_ordreimpresio where comanda=" + atrim(reixa.TextMatrix(reixa.row, 0)))
    If Not rst.EOF Then
        cpostit.text = " Dia programat: " + UCase(Format(rst!dataprogramada, "dd ""de"" mmmm hh:nn") + "H.") + Chr(13) + Chr(10) + " Motiu: " + atrim(rst!motiuprogramacio)
        cpostit.Top = reixa.RowPos(reixa.row) + reixa.Top + reixa.RowHeight(reixa.row)
        cpostit.Left = reixa.Left + 700
        cpostit.visible = True
    End If
     Else: cpostit.visible = False
     If Not rst.EOF Then If rst!esreprint Then canvimaquinaREPRINT True
   End If
   If Not rst.EOF Then
    If UCase(rst!modificada) = "M" Then
       botoPDFmodificacions True
        Else: botoPDFmodificacions False
    End If
    botoCOMENTARISCOMADA cadbl(numerodecomandaseleccionada)
   End If
   If reixa.Cols > 2 Then
    If Format(reixa.TextMatrix(reixa.row, 3), "dd") = Format(Now, "dd") Then
      vdataf = "Avui a les " + Format(reixa.TextMatrix(reixa.row, 3), " hh:nn")
        Else: vdataf = Format(reixa.TextMatrix(reixa.row, 3), "dddd hh:nn")
    End If
    If reixa.TextMatrix(reixa.row, 3) <> "" Then ethoraimpresio = "Finalització impresió aprox: " + vdataf
   End If
   Set rst = Nothing
End Sub
Private Sub reixa_RowColChange()
   reixa_refrescarfila
   Unload veurereport
'   If arguments(1) = "" Then If bcomentariscomanda.visible Then mpossarcomentari_Click
End Sub
Function esmuntadora() As Boolean
   If InStr(1, LCase(App.EXEName), "muntadora") > 0 Then esmuntadora = True
End Function

Private Sub Timer1_Timer()
   If DateDiff("n", vtempsultimcanvi, Now) > 0 Then
      carregar_reixaordre
   End If
   comprovarsihihamissatgesCHAT
End Sub
Sub comprovarsihihamissatgesCHAT()
   Dim rst As Recordset
   cimatgemissatgeCHAT_Verd.visible = False
   
   If Framemodificacions.visible Then
      Set rst = dbmissatges.OpenRecordset("select * from converses_assumpte where datalectura=null and operariultimcanvi='T'")
      If Not rst.EOF Then
          cimatgemissatgeCHAT_Verd.visible = True
           
      End If
   End If
   Set rst = Nothing
End Sub



