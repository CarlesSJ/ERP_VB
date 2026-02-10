VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formtintes 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manteniment de Tintes"
   ClientHeight    =   11175
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15360
   ClipControls    =   0   'False
   FillColor       =   &H80000001&
   Icon            =   "formtintes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport llistat 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab pestanyes 
      Height          =   11070
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   19526
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tintes"
      TabPicture(0)   =   "formtintes.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label22(38)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "etajudabusqueda"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "etfiltrarper"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "framedadestintes"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "reixatintes"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "reixallaunes"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "reixarecarregues"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command55"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "botoestocminim"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command44"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command40"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame11"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "bcontrolestocminim"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command31"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command30"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command29"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Command28"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame6"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "bimportarllaunes"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Frame2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "tintes"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "buscador"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Command2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Command10"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "combosionookcarta"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Command16"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "filtretinta(0)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Command21"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Command36"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Llaunes"
      TabPicture(1)   =   "formtintes.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command22"
      Tab(1).Control(1)=   "Command23"
      Tab(1).Control(2)=   "fnumllauna"
      Tab(1).Control(3)=   "fdesctintallauna"
      Tab(1).Control(4)=   "datadellaunes"
      Tab(1).Control(5)=   "fcoditintallauna"
      Tab(1).Control(6)=   "Command24"
      Tab(1).Control(7)=   "fsituaciollauna"
      Tab(1).Control(8)=   "Command25"
      Tab(1).Control(9)=   "Command26"
      Tab(1).Control(10)=   "Command32"
      Tab(1).Control(11)=   "Command33"
      Tab(1).Control(12)=   "brecalcularpesllaunes"
      Tab(1).Control(13)=   "fmostracolor"
      Tab(1).Control(14)=   "checkactives"
      Tab(1).Control(15)=   "checkimpresores"
      Tab(1).Control(16)=   "MSComm1"
      Tab(1).Control(17)=   "DBGrid1"
      Tab(1).Control(18)=   "Label22(0)"
      Tab(1).Control(19)=   "ettotalllaunes"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "Formules"
      TabPicture(2)   =   "formtintes.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pestanyesforumes"
      Tab(2).Control(1)=   "Frame5(6)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Albarans"
      TabPicture(3)   =   "formtintes.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "checktots"
      Tab(3).Control(1)=   "Checkultims30"
      Tab(3).Control(2)=   "Command41"
      Tab(3).Control(3)=   "llistaalbarans"
      Tab(3).Control(4)=   "Command38"
      Tab(3).Control(5)=   "Command39"
      Tab(3).Control(6)=   "Label5"
      Tab(3).Control(7)=   "Label23"
      Tab(3).Control(8)=   "etcreantllaunes"
      Tab(3).Control(9)=   "campabuscar"
      Tab(3).Control(10)=   "etquant"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Comandes Actives"
      TabPicture(4)   =   "formtintes.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5(1)"
      Tab(4).Control(1)=   "framecuatricomia"
      Tab(4).Control(2)=   "Command67(13)"
      Tab(4).Control(3)=   "Command67(12)"
      Tab(4).Control(4)=   "Command67(11)"
      Tab(4).Control(5)=   "buscador_estoc(1)"
      Tab(4).Control(6)=   "Command67(9)"
      Tab(4).Control(7)=   "Command19"
      Tab(4).Control(8)=   "cobsoperari"
      Tab(4).Control(9)=   "Command17"
      Tab(4).Control(10)=   "Command13"
      Tab(4).Control(11)=   "Command9"
      Tab(4).Control(12)=   "Command8"
      Tab(4).Control(13)=   "framebotons"
      Tab(4).Control(14)=   "llistacomandes"
      Tab(4).Control(15)=   "Command42"
      Tab(4).Control(16)=   "llistatintes"
      Tab(4).Control(17)=   "Command47"
      Tab(4).Control(18)=   "Command52"
      Tab(4).Control(19)=   "filtre(0)"
      Tab(4).Control(20)=   "Command56"
      Tab(4).Control(21)=   "Frame12"
      Tab(4).Control(22)=   "Command62"
      Tab(4).Control(23)=   "Command64"
      Tab(4).Control(24)=   "Command65"
      Tab(4).Control(25)=   "reixacomandes"
      Tab(4).Control(26)=   "Command54"
      Tab(4).Control(27)=   "Command14"
      Tab(4).Control(28)=   "Check1(2)"
      Tab(4).Control(29)=   "Label22(29)"
      Tab(4).Control(30)=   "Label22(28)"
      Tab(4).Control(31)=   "Label22(27)"
      Tab(4).Control(32)=   "etextensio"
      Tab(4).Control(33)=   "Label25"
      Tab(4).Control(34)=   "Label26"
      Tab(4).Control(35)=   "Label27"
      Tab(4).Control(36)=   "etactualitzant"
      Tab(4).Control(37)=   "etajudadosclics"
      Tab(4).Control(38)=   "ettotalcomandes"
      Tab(4).Control(39)=   "Label30"
      Tab(4).Control(40)=   "etquancolor"
      Tab(4).ControlCount=   41
      TabCaption(5)   =   "Estoc Tintes"
      TabPicture(5)   =   "formtintes.frx":0616
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "ettotalestocs"
      Tab(5).Control(1)=   "etfiltrarperestoc"
      Tab(5).Control(2)=   "reixaestocs"
      Tab(5).Control(3)=   "Command7"
      Tab(5).Control(4)=   "buscador_estoc(0)"
      Tab(5).Control(5)=   "checkinclou"
      Tab(5).Control(6)=   "checknomesfora"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "Compres"
      TabPicture(6)   =   "formtintes.frx":0632
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label1"
      Tab(6).Control(1)=   "Label2"
      Tab(6).Control(2)=   "Label3"
      Tab(6).Control(3)=   "etenviantemail"
      Tab(6).Control(4)=   "llistacompres"
      Tab(6).Control(5)=   "Command67(3)"
      Tab(6).Control(6)=   "Command67(4)"
      Tab(6).Control(7)=   "Command12"
      Tab(6).ControlCount=   8
      Begin VB.Frame Frame5 
         Height          =   10110
         Index           =   6
         Left            =   -74850
         TabIndex        =   297
         Top             =   855
         Visible         =   0   'False
         Width           =   14730
         Begin VB.CommandButton Command67 
            Caption         =   "Tancar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Index           =   29
            Left            =   12240
            Style           =   1  'Graphical
            TabIndex        =   299
            Top             =   9015
            Width           =   2010
         End
         Begin MSFlexGridLib.MSFlexGrid reixagrmskilo 
            Height          =   7755
            Left            =   1485
            TabIndex        =   298
            Top             =   975
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   13679
            _Version        =   393216
            Rows            =   12
            Cols            =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Nivell cuatricomia"
         Height          =   2565
         Index           =   1
         Left            =   -66855
         TabIndex        =   249
         Top             =   7440
         Visible         =   0   'False
         Width           =   5025
         Begin VB.CommandButton Command67 
            BackColor       =   &H80000005&
            Caption         =   "F2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   31
            Left            =   3450
            Style           =   1  'Graphical
            TabIndex        =   301
            ToolTipText     =   "Nº Maq. Apreta mes de 3 segons per copiar els valors de l'altra màquina."
            Top             =   30
            Width           =   870
         End
         Begin VB.CommandButton Command67 
            BackColor       =   &H80000005&
            Caption         =   "FW"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   30
            Left            =   2565
            Style           =   1  'Graphical
            TabIndex        =   300
            ToolTipText     =   "Nº Maq. Apreta mes de 3 segons per copiar els valors de l'altra màquina."
            Top             =   30
            Width           =   870
         End
         Begin VB.CommandButton Command67 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   27
            Left            =   4395
            Picture         =   "formtintes.frx":064E
            Style           =   1  'Graphical
            TabIndex        =   295
            ToolTipText     =   "Copiar dades versió anterior"
            Top             =   600
            Width           =   480
         End
         Begin VB.CommandButton Command67 
            BackColor       =   &H0025EFAD&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   26
            Left            =   4410
            Picture         =   "formtintes.frx":0BD8
            Style           =   1  'Graphical
            TabIndex        =   294
            ToolTipText     =   "Exportar dades a excel."
            Top             =   2115
            Width           =   480
         End
         Begin VB.Frame Frame5 
            Caption         =   "Detall de cargues i lots "
            Height          =   2235
            Index           =   2
            Left            =   225
            TabIndex        =   251
            Top             =   2010
            Visible         =   0   'False
            Width           =   4320
            Begin VB.CommandButton Command67 
               BackColor       =   &H00EAD9CE&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   720
               Index           =   20
               Left            =   3330
               Picture         =   "formtintes.frx":1162
               Style           =   1  'Graphical
               TabIndex        =   293
               ToolTipText     =   "Guarda canvis cuatricomia"
               Top             =   1335
               Width           =   810
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H006BEBB1&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   7
               Left            =   930
               TabIndex        =   263
               Tag             =   "Negre_toleranciaV"
               Top             =   255
               Width           =   615
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   8
               Left            =   1665
               TabIndex        =   262
               Tag             =   "Negre_toleranciaT"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H005C31DD&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   9
               Left            =   2370
               TabIndex        =   261
               Tag             =   "Negre_toleranciaM"
               Top             =   255
               Width           =   540
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H006BEBB1&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   10
               Left            =   915
               TabIndex        =   260
               Tag             =   "Groc_toleranciaV"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   11
               Left            =   1665
               TabIndex        =   259
               Tag             =   "Groc_toleranciaT"
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H005C31DD&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   12
               Left            =   2370
               TabIndex        =   258
               Tag             =   "Groc_toleranciaM"
               Top             =   720
               Width           =   540
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H006BEBB1&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   13
               Left            =   915
               TabIndex        =   257
               Tag             =   "Magenta_toleranciaV"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   14
               Left            =   1665
               TabIndex        =   256
               Tag             =   "Magenta_toleranciaT"
               Top             =   1200
               Width           =   615
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H005C31DD&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   15
               Left            =   2370
               TabIndex        =   255
               Tag             =   "Magenta_toleranciaM"
               Top             =   1200
               Width           =   540
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H006BEBB1&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   16
               Left            =   915
               TabIndex        =   254
               Tag             =   "Cyan_toleranciaV"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   17
               Left            =   1665
               TabIndex        =   253
               Tag             =   "Cyan_toleranciaT"
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H005C31DD&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   18
               Left            =   2370
               TabIndex        =   252
               Tag             =   "Cyan_toleranciaM"
               Top             =   1680
               Width           =   540
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   31
               Left            =   720
               TabIndex        =   267
               Top             =   270
               Width           =   180
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tol:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   30
               Left            =   225
               TabIndex        =   271
               Top             =   330
               Width           =   420
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00000000&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   7
               Left            =   120
               Top             =   225
               Width           =   2985
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tol:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   32
               Left            =   225
               TabIndex        =   270
               Top             =   795
               Width           =   420
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tol:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   34
               Left            =   225
               TabIndex        =   269
               Top             =   1275
               Width           =   420
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tol:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   240
               Index           =   36
               Left            =   225
               TabIndex        =   268
               Top             =   1755
               Width           =   420
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               Index           =   33
               Left            =   720
               TabIndex        =   266
               Top             =   735
               Width           =   180
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   360
               Index           =   35
               Left            =   720
               TabIndex        =   265
               Top             =   1215
               Width           =   180
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   360
               Index           =   37
               Left            =   720
               TabIndex        =   264
               Top             =   1695
               Width           =   180
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0000FFFF&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   6
               Left            =   120
               Top             =   705
               Width           =   2985
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FF00FF&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   5
               Left            =   120
               Top             =   1185
               Width           =   2985
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FFFF00&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   4
               Left            =   120
               Top             =   1665
               Width           =   2985
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Valors Delta"
            Height          =   2190
            Index           =   5
            Left            =   1695
            TabIndex        =   279
            Top             =   345
            Width           =   1320
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFF80&
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
               Index           =   22
               Left            =   285
               TabIndex        =   283
               Top             =   1680
               Width           =   705
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H00FF80FF&
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
               Index           =   21
               Left            =   285
               TabIndex        =   282
               Top             =   1200
               Width           =   705
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H0080FFFF&
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
               Index           =   20
               Left            =   285
               TabIndex        =   281
               Top             =   735
               Width           =   705
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H00808080&
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
               Index           =   19
               Left            =   285
               TabIndex        =   280
               Top             =   255
               Width           =   705
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FF00FF&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   11
               Left            =   105
               Top             =   1170
               Width           =   1095
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0000FFFF&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   10
               Left            =   105
               Top             =   690
               Width           =   1095
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00000000&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   9
               Left            =   105
               Top             =   210
               Width           =   1095
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FFFF00&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   8
               Left            =   105
               Top             =   1650
               Width           =   1095
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00989FF8&
            Caption         =   "Valors Delta E"
            Height          =   2190
            Index           =   4
            Left            =   3045
            TabIndex        =   274
            Top             =   345
            Width           =   1320
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H00808080&
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
               Index           =   3
               Left            =   285
               TabIndex        =   278
               Top             =   255
               Width           =   705
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H0080FFFF&
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
               Index           =   4
               Left            =   285
               TabIndex        =   277
               Top             =   735
               Width           =   705
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H00FF80FF&
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
               Index           =   5
               Left            =   285
               TabIndex        =   276
               Top             =   1200
               Width           =   705
            End
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFF80&
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
               Index           =   6
               Left            =   270
               TabIndex        =   275
               Top             =   1695
               Width           =   705
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FFFF00&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   3
               Left            =   105
               Top             =   1650
               Width           =   1095
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00000000&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   0
               Left            =   105
               Top             =   210
               Width           =   1095
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0000FFFF&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   1
               Left            =   105
               Top             =   690
               Width           =   1095
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FF00FF&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   2
               Left            =   105
               Top             =   1170
               Width           =   1095
            End
         End
         Begin VB.CommandButton Command67 
            BackColor       =   &H00FF8080&
            Caption         =   "Tol."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   21
            Left            =   4425
            Style           =   1  'Graphical
            TabIndex        =   273
            ToolTipText     =   "Tolerancies GENERALS dels Delta"
            Top             =   1560
            Width           =   480
         End
         Begin VB.CommandButton Command67 
            BackColor       =   &H00FFC0FF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   19
            Left            =   4395
            Style           =   1  'Graphical
            TabIndex        =   272
            ToolTipText     =   "Sortir sense guardar canvis"
            Top             =   180
            Width           =   480
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H80000005&
            Caption         =   "Anilox i volum"
            Height          =   2190
            Index           =   3
            Left            =   45
            TabIndex        =   250
            Top             =   345
            Width           =   1635
            Begin VB.CommandButton Command67 
               BackColor       =   &H80000005&
               Caption         =   "+"
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
               Index           =   25
               Left            =   1245
               TabIndex        =   292
               ToolTipText     =   "Sortir sense guardar canvis"
               Top             =   1695
               Width           =   330
            End
            Begin VB.CommandButton Command67 
               BackColor       =   &H80000005&
               Caption         =   "+"
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
               Index           =   24
               Left            =   1245
               TabIndex        =   291
               ToolTipText     =   "Sortir sense guardar canvis"
               Top             =   1200
               Width           =   330
            End
            Begin VB.CommandButton Command67 
               BackColor       =   &H80000005&
               Caption         =   "+"
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
               Index           =   23
               Left            =   1245
               TabIndex        =   290
               ToolTipText     =   "Sortir sense guardar canvis"
               Top             =   720
               Width           =   330
            End
            Begin VB.CommandButton Command67 
               BackColor       =   &H80000005&
               Caption         =   "+"
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
               Index           =   22
               Left            =   1260
               TabIndex        =   289
               ToolTipText     =   "Sortir sense guardar canvis"
               Top             =   240
               Width           =   330
            End
            Begin VB.ComboBox Combo 
               BackColor       =   &H00FFFF80&
               Height          =   315
               Index           =   3
               Left            =   165
               TabIndex        =   288
               Top             =   1725
               Width           =   1005
            End
            Begin VB.ComboBox Combo 
               BackColor       =   &H00FF80FF&
               Height          =   315
               Index           =   2
               Left            =   165
               TabIndex        =   287
               Top             =   1245
               Width           =   1005
            End
            Begin VB.ComboBox Combo 
               BackColor       =   &H0080FFFF&
               Height          =   315
               Index           =   1
               Left            =   165
               TabIndex        =   286
               Top             =   750
               Width           =   1005
            End
            Begin VB.ComboBox Combo 
               BackColor       =   &H00808080&
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Index           =   0
               Left            =   165
               TabIndex        =   285
               Top             =   300
               Width           =   1005
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FF00FF&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   15
               Left            =   120
               Top             =   1170
               Width           =   1095
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0000FFFF&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   14
               Left            =   120
               Top             =   690
               Width           =   1095
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00000000&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   13
               Left            =   120
               Top             =   210
               Width           =   1095
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00FFFF00&
               BackStyle       =   1  'Opaque
               Height          =   465
               Index           =   12
               Left            =   120
               Top             =   1650
               Width           =   1095
            End
         End
      End
      Begin VB.Frame framecuatricomia 
         BackColor       =   &H80000005&
         Caption         =   "Nivell cuatricomia"
         Height          =   2220
         Left            =   -69315
         TabIndex        =   109
         Top             =   8085
         Visible         =   0   'False
         Width           =   7530
      End
      Begin VB.CheckBox checktots 
         Caption         =   "Veure tots."
         Height          =   195
         Left            =   -63570
         TabIndex        =   98
         Top             =   1695
         Width           =   1425
      End
      Begin VB.CheckBox Checkultims30 
         Caption         =   "Ultims 30"
         Height          =   195
         Left            =   -64620
         TabIndex        =   241
         Top             =   1695
         Width           =   1425
      End
      Begin VB.CommandButton Command67 
         Caption         =   "Ordre F2"
         Height          =   420
         Index           =   13
         Left            =   -66765
         Style           =   1  'Graphical
         TabIndex        =   240
         ToolTipText     =   "Assignar linies d'impresió 001#1"
         Top             =   540
         Width           =   660
      End
      Begin VB.CommandButton Command67 
         Caption         =   "Ordre FW"
         Height          =   420
         Index           =   12
         Left            =   -67440
         Style           =   1  'Graphical
         TabIndex        =   239
         ToolTipText     =   "Assignar linies d'impresió 001#1"
         Top             =   540
         Width           =   660
      End
      Begin VB.CommandButton Command67 
         Height          =   420
         Index           =   11
         Left            =   -68115
         Picture         =   "formtintes.frx":16EC
         Style           =   1  'Graphical
         TabIndex        =   237
         ToolTipText     =   "Assignar linies d'impresió 001#1"
         Top             =   525
         Width           =   660
      End
      Begin VB.TextBox buscador_estoc 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Index           =   1
         Left            =   -73935
         Locked          =   -1  'True
         TabIndex        =   234
         Top             =   10290
         Width           =   12090
      End
      Begin VB.CommandButton Command67 
         Height          =   420
         Index           =   9
         Left            =   -68775
         Picture         =   "formtintes.frx":1FF1
         Style           =   1  'Graphical
         TabIndex        =   231
         ToolTipText     =   "Verificació de tintes de treballs nous o modificats."
         Top             =   525
         Width           =   660
      End
      Begin VB.CommandButton Command36 
         Height          =   390
         Left            =   2580
         Picture         =   "formtintes.frx":257B
         Style           =   1  'Graphical
         TabIndex        =   228
         ToolTipText     =   "Comandes on s'ha utilitzat aquesta llauna."
         Top             =   8730
         Width           =   390
      End
      Begin VB.CommandButton Command19 
         Height          =   420
         Left            =   -69435
         Picture         =   "formtintes.frx":2B05
         Style           =   1  'Graphical
         TabIndex        =   225
         ToolTipText     =   "Ordre impresió de les comandes"
         Top             =   525
         Width           =   660
      End
      Begin VB.CommandButton Command12 
         Height          =   420
         Left            =   -60735
         Picture         =   "formtintes.frx":2BD7
         Style           =   1  'Graphical
         TabIndex        =   204
         ToolTipText     =   "Re-Enviar correu de compres pendents al Departament de Compres."
         Top             =   10515
         Width           =   675
      End
      Begin VB.CommandButton Command67 
         Caption         =   "Afegir Compres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   4
         Left            =   -62445
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   2475
         Width           =   1950
      End
      Begin VB.CommandButton Command67 
         Caption         =   "Actualitzar Compres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   3
         Left            =   -62460
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   1815
         Width           =   1950
      End
      Begin VB.ListBox llistacompres 
         BackColor       =   &H00C7CBF5&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   8940
         Left            =   -74250
         TabIndex        =   198
         Top             =   1695
         Width           =   11535
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "+Lots"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   -60990
         Picture         =   "formtintes.frx":3161
         Style           =   1  'Graphical
         TabIndex        =   165
         ToolTipText     =   "Afegir lots a una llauna."
         Top             =   5535
         Width           =   975
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H0080FFFF&
         Caption         =   "Tintes Semblants"
         Height          =   360
         Left            =   11595
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   6420
         Width           =   3105
      End
      Begin VB.TextBox cobsoperari 
         BackColor       =   &H00FDDECE&
         Height          =   285
         Left            =   -73920
         MaxLength       =   255
         TabIndex        =   163
         Top             =   9960
         Width           =   12045
      End
      Begin VB.CommandButton Command17 
         Height          =   420
         Left            =   -70095
         Picture         =   "formtintes.frx":343B
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "Modificar tinters del treball de la comanda seleccionada"
         Top             =   525
         Width           =   660
      End
      Begin VB.TextBox filtretinta 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Index           =   0
         Left            =   915
         TabIndex        =   160
         Top             =   990
         Width           =   930
      End
      Begin VB.CommandButton Command16 
         Height          =   270
         Left            =   195
         Picture         =   "formtintes.frx":39C5
         Style           =   1  'Graphical
         TabIndex        =   159
         ToolTipText     =   "Eliminar totes les linies"
         Top             =   990
         Width           =   540
      End
      Begin VB.ComboBox combosionookcarta 
         BackColor       =   &H00FFC0FF&
         Height          =   315
         ItemData        =   "formtintes.frx":3F4F
         Left            =   930
         List            =   "formtintes.frx":3F59
         TabIndex        =   151
         Top             =   660
         Width           =   660
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H008080FF&
         Enabled         =   0   'False
         Height          =   405
         Left            =   -61755
         Picture         =   "formtintes.frx":3F65
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   6930
         Width           =   1545
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Command10"
         Height          =   300
         Left            =   11775
         TabIndex        =   149
         Top             =   510
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.CommandButton Command9 
         Height          =   420
         Left            =   -70755
         Picture         =   "formtintes.frx":41D0
         Style           =   1  'Graphical
         TabIndex        =   147
         ToolTipText     =   "Llista de Marques i Linies de treballs."
         Top             =   525
         Width           =   660
      End
      Begin VB.CommandButton Command8 
         Height          =   420
         Left            =   -71415
         Picture         =   "formtintes.frx":475A
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Ensenyar nomes reprint."
         Top             =   525
         Width           =   660
      End
      Begin VB.CheckBox checknomesfora 
         Caption         =   "Només tintes fora"
         Height          =   210
         Left            =   -62205
         TabIndex        =   145
         Top             =   585
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CheckBox checkinclou 
         Caption         =   "Inclou tintes comandes"
         Height          =   210
         Left            =   -62205
         TabIndex        =   144
         Top             =   795
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.TextBox buscador_estoc 
         Height          =   315
         Index           =   0
         Left            =   -73680
         TabIndex        =   142
         Top             =   1125
         Width           =   3525
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Actualitzar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -62295
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   1050
         Width           =   1710
      End
      Begin MSFlexGridLib.MSFlexGrid reixaestocs 
         Height          =   9135
         Left            =   -74745
         TabIndex        =   139
         Top             =   1485
         Width           =   14310
         _ExtentX        =   25241
         _ExtentY        =   16113
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin VB.Frame framebotons 
         Height          =   3285
         Left            =   -61785
         TabIndex        =   122
         Top             =   7260
         Width           =   1710
         Begin VB.CommandButton Command67 
            BackColor       =   &H00C78DFA&
            Caption         =   "Imprimir Combinació Llaunes"
            Height          =   450
            Index           =   8
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   227
            ToolTipText     =   "Combinació de llaunes fetes per aconseguir aquest color."
            Top             =   2340
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton balternatives 
            Caption         =   "Tintes Alternatives"
            Height          =   480
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   130
            ToolTipText     =   "Assigna una llauna a la tinta sel.leccionada."
            Top             =   1410
            Width           =   1590
         End
         Begin VB.CommandButton Command59 
            Caption         =   "Valors Cuatricomia "
            Height          =   480
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   900
            Width           =   1590
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Excel"
            Height          =   720
            Left            =   30
            TabIndex        =   126
            Top             =   120
            Width           =   1635
            Begin VB.CommandButton Command49 
               BackColor       =   &H0080FF80&
               Caption         =   "Nou"
               Height          =   420
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   128
               ToolTipText     =   "Fer nou Excel"
               Top             =   195
               Width           =   450
            End
            Begin VB.CommandButton Command46 
               Caption         =   "Passar-ho a Excel"
               Height          =   420
               Left            =   510
               TabIndex        =   127
               ToolTipText     =   "Passar informació sel.leccionada a Excel"
               Top             =   195
               Width           =   1020
            End
         End
         Begin VB.CommandButton Command45 
            Caption         =   "Tintes Semblants"
            Height          =   420
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   125
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton Command48 
            Caption         =   "Marcar com a Extensió Feta"
            Height          =   420
            Left            =   1380
            TabIndex        =   124
            Top             =   2370
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.CommandButton Command51 
            Caption         =   " Extensió Feta manualment"
            Height          =   420
            Left            =   60
            TabIndex        =   123
            Top             =   2820
            Width           =   1545
         End
      End
      Begin VB.ListBox llistacomandes 
         BackColor       =   &H00EAD9CE&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   -74850
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   121
         Top             =   6405
         Visible         =   0   'False
         Width           =   14670
      End
      Begin VB.CommandButton Command42 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Actualitzar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -61800
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   450
         Width           =   1710
      End
      Begin VB.ListBox llistatintes 
         BackColor       =   &H00EAD9CE&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2490
         Left            =   -74850
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   119
         Top             =   7425
         Width           =   13050
      End
      Begin VB.CommandButton Command47 
         Height          =   315
         Left            =   -74955
         Picture         =   "formtintes.frx":4854
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Buscar comandes"
         Top             =   375
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command52 
         Height          =   420
         Left            =   -74715
         Picture         =   "formtintes.frx":4DDE
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Veure la comanda"
         Top             =   525
         Width           =   660
      End
      Begin VB.TextBox filtre 
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00808080&
         Height          =   285
         Index           =   0
         Left            =   -74745
         TabIndex        =   115
         ToolTipText     =   "Pots buscar valors separats per comes i a client pots posar el codi de client."
         Top             =   960
         Width           =   630
      End
      Begin VB.CommandButton Command56 
         Height          =   270
         Left            =   -74940
         Picture         =   "formtintes.frx":5368
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Eliminar totes les linies"
         Top             =   945
         Width           =   240
      End
      Begin VB.Frame Frame12 
         Height          =   675
         Left            =   -65895
         TabIndex        =   110
         Top             =   315
         Width           =   4065
         Begin VB.CommandButton Command67 
            BackColor       =   &H00FFC0FF&
            Caption         =   "P4"
            Height          =   225
            Index           =   17
            Left            =   3030
            Style           =   1  'Graphical
            TabIndex        =   247
            ToolTipText     =   "Botons filtres programables"
            Top             =   405
            Width           =   960
         End
         Begin VB.CommandButton Command67 
            BackColor       =   &H00FFC0FF&
            Caption         =   "P3"
            Height          =   225
            Index           =   16
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   246
            ToolTipText     =   "Botons filtres programables"
            Top             =   405
            Width           =   975
         End
         Begin VB.CommandButton Command67 
            BackColor       =   &H00FFC0FF&
            Caption         =   "P2"
            Height          =   225
            Index           =   15
            Left            =   1095
            Style           =   1  'Graphical
            TabIndex        =   245
            ToolTipText     =   "Botons filtres programables"
            Top             =   405
            Width           =   900
         End
         Begin VB.CommandButton Command67 
            BackColor       =   &H00FFC0FF&
            Caption         =   "P1"
            Height          =   225
            Index           =   14
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   244
            ToolTipText     =   "Botons filtres programables"
            Top             =   405
            Width           =   990
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H0000FFFF&
            Caption         =   "Treb. Blaus"
            Height          =   270
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   162
            ToolTipText     =   "Treballs de color blau (Clixes entrats)"
            Top             =   120
            Width           =   1005
         End
         Begin VB.CommandButton Command63 
            BackColor       =   &H0000FFFF&
            Caption         =   "Te CallOff"
            Height          =   270
            Left            =   3030
            Style           =   1  'Graphical
            TabIndex        =   113
            ToolTipText     =   "Tintes comprades a fora que no estan a muntadora."
            Top             =   120
            Width           =   960
         End
         Begin VB.CommandButton Command57 
            BackColor       =   &H0000FFFF&
            Caption         =   "Munt. No"
            Height          =   270
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   112
            ToolTipText     =   "Comandes a muntadora No revisades."
            Top             =   120
            Width           =   915
         End
         Begin VB.CommandButton Command58 
            BackColor       =   &H0000FFFF&
            Caption         =   "Tintes fora"
            Height          =   270
            Left            =   2025
            Style           =   1  'Graphical
            TabIndex        =   111
            ToolTipText     =   "Tintes comprades a fora que no estan a muntadora."
            Top             =   120
            Width           =   960
         End
      End
      Begin VB.CommandButton Command62 
         Height          =   420
         Left            =   -73395
         Picture         =   "formtintes.frx":58F2
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Veure la comanda"
         Top             =   525
         Width           =   660
      End
      Begin VB.CommandButton Command64 
         Height          =   420
         Left            =   -72735
         Picture         =   "formtintes.frx":5E7C
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Ordenar per data d'Expedició."
         Top             =   525
         Width           =   660
      End
      Begin VB.CommandButton Command65 
         Height          =   420
         Left            =   -72075
         Picture         =   "formtintes.frx":6406
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Informació del treball."
         Top             =   525
         Width           =   660
      End
      Begin VB.CommandButton Command41 
         Caption         =   "No crear llaunes d'aquesta tinta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -62805
         TabIndex        =   105
         Top             =   3300
         Width           =   1710
      End
      Begin VB.ListBox llistaalbarans 
         BackColor       =   &H00EAD9CE&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   8940
         Left            =   -74475
         TabIndex        =   101
         Top             =   1920
         Width           =   11535
      End
      Begin VB.CommandButton Command38 
         Caption         =   "Actualitzar llista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -62805
         TabIndex        =   100
         Top             =   2610
         Width           =   1710
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Crear les llaunes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -62790
         TabIndex        =   99
         Top             =   1965
         Width           =   1710
      End
      Begin VB.CommandButton Command2 
         Height          =   345
         Left            =   4170
         Picture         =   "formtintes.frx":6990
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   360
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox buscador 
         Height          =   345
         Left            =   1125
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.Data tintes 
         Caption         =   "Tintes"
         Connect         =   "Access"
         DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\tintes.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   9510
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tintes_tot"
         Top             =   285
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   210
         TabIndex        =   33
         Top             =   5310
         Width           =   14550
         Begin VB.CommandButton bdescatalogar 
            Caption         =   " Descatalogar"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13050
            TabIndex        =   39
            ToolTipText     =   "Descatalogar o Activar la Tinta"
            Top             =   165
            Width           =   1365
         End
         Begin VB.PictureBox fotodescatalogar 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   13065
            Picture         =   "formtintes.frx":6F1A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   38
            ToolTipText     =   "Descatalogar o Activar la Tinta"
            Top             =   240
            Width           =   240
         End
         Begin VB.Timer rellotge1 
            Interval        =   800
            Left            =   4440
            Top             =   135
         End
         Begin VB.CommandButton Command1 
            Height          =   390
            Left            =   1245
            Picture         =   "formtintes.frx":74A4
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Acceptar canvis"
            Top             =   135
            Width           =   390
         End
         Begin VB.CommandButton modificar 
            Height          =   390
            Left            =   465
            Picture         =   "formtintes.frx":7A2E
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Modificació de registre"
            Top             =   135
            Width           =   390
         End
         Begin VB.CommandButton eliminar 
            Height          =   390
            Left            =   855
            Picture         =   "formtintes.frx":7FB8
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Eliminacio Registres"
            Top             =   135
            Width           =   390
         End
         Begin VB.CommandButton alta 
            Height          =   390
            Left            =   75
            Picture         =   "formtintes.frx":8542
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Alta  Registres"
            Top             =   135
            Width           =   390
         End
         Begin VB.Label estattaula 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2325
            TabIndex        =   40
            Top             =   150
            Width           =   1995
         End
      End
      Begin VB.CommandButton Command23 
         Height          =   315
         Left            =   -74640
         Picture         =   "formtintes.frx":8ACC
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Neteja filtre"
         Top             =   1455
         Width           =   360
      End
      Begin VB.TextBox fnumllauna 
         Height          =   345
         Left            =   -74205
         TabIndex        =   31
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox fdesctintallauna 
         Height          =   345
         Left            =   -71055
         TabIndex        =   30
         Top             =   1440
         Width           =   5925
      End
      Begin VB.Data datadellaunes 
         Caption         =   "datadellaunes"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67815
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   945
         Visible         =   0   'False
         Width           =   3150
      End
      Begin VB.TextBox fcoditintallauna 
         Height          =   345
         Left            =   -72990
         TabIndex        =   29
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Barrejar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   -60990
         Picture         =   "formtintes.frx":9056
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Barrejar dos llaunes."
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox fsituaciollauna 
         Height          =   345
         Left            =   -65115
         TabIndex        =   27
         Top             =   1440
         Width           =   1065
      End
      Begin VB.CommandButton bimportarllaunes 
         Caption         =   "Importar Llaunes"
         Height          =   210
         Left            =   13395
         TabIndex        =   26
         Top             =   540
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Frame Frame6 
         Caption         =   "Historia de la llauna"
         Height          =   2235
         Left            =   7365
         TabIndex        =   22
         Top             =   8640
         Width           =   7230
         Begin VB.CommandButton Command27 
            Height          =   300
            Left            =   30
            Picture         =   "formtintes.frx":A8F0
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Eliminacio Registres"
            Top             =   225
            Width           =   300
         End
         Begin MSDBGrid.DBGrid reixahistoria 
            Bindings        =   "formtintes.frx":AE7A
            Height          =   1950
            Left            =   345
            OleObjectBlob   =   "formtintes.frx":AE91
            TabIndex        =   24
            Top             =   195
            Width           =   6765
         End
      End
      Begin VB.CommandButton Command25 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Veure Llauna"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   -60990
         Picture         =   "formtintes.frx":BD7A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Barrejar dos llaunes."
         Top             =   1035
         Width           =   975
      End
      Begin VB.CommandButton Command26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Retorn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   -60990
         Picture         =   "formtintes.frx":C304
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Retorn de les llaunes."
         Top             =   2805
         Width           =   975
      End
      Begin VB.CommandButton Command28 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Lots Base"
         Height          =   390
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Numeros de lots afectats."
         Top             =   8715
         Width           =   840
      End
      Begin VB.CommandButton Command32 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Convertir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   -60990
         Picture         =   "formtintes.frx":D426
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Reconvertir la llauna a una altra tinta."
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton Command33 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Situació"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   -60990
         Picture         =   "formtintes.frx":D9B0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Reconvertir la tinta."
         Top             =   4635
         Width           =   975
      End
      Begin VB.CommandButton Command29 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Referencies Proveidors"
         Height          =   360
         Left            =   11595
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6045
         Width           =   3105
      End
      Begin VB.CommandButton Command30 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Sit"
         Height          =   390
         Left            =   5790
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Alta  Registres"
         Top             =   8715
         Width           =   420
      End
      Begin VB.CommandButton Command31 
         Height          =   300
         Left            =   7395
         Picture         =   "formtintes.frx":EA6A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Actualització"
         Top             =   9195
         Width           =   300
      End
      Begin VB.CommandButton Command6 
         Height          =   390
         Left            =   600
         Picture         =   "formtintes.frx":EFF4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Alta  Registres"
         Top             =   8730
         Width           =   390
      End
      Begin VB.CommandButton bcontrolestocminim 
         Caption         =   "Consultar Estoc mínim"
         Height          =   420
         Left            =   12900
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   375
         Visible         =   0   'False
         Width           =   2070
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   3030
         TabIndex        =   9
         Top             =   8670
         Width           =   1380
         Begin VB.Label etkgtotals 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   30
            MouseIcon       =   "formtintes.frx":F57E
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   30
            Width           =   1320
         End
      End
      Begin VB.CommandButton brecalcularpesllaunes 
         Caption         =   "Recalcular pes llaunes"
         Height          =   690
         Left            =   -60885
         TabIndex        =   8
         Top             =   8520
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Frame fmostracolor 
         Caption         =   "Mostra del Color"
         Height          =   720
         Left            =   -62475
         TabIndex        =   7
         Top             =   495
         Width           =   1485
      End
      Begin VB.CommandButton Command40 
         Height          =   390
         Left            =   2175
         Picture         =   "formtintes.frx":FB08
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8730
         Width           =   390
      End
      Begin VB.CommandButton Command44 
         Height          =   270
         Left            =   9270
         Picture         =   "formtintes.frx":10092
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Pots filtrar totes les tintes amb aquestes families de tinta i color."
         Top             =   6840
         Width           =   360
      End
      Begin VB.CommandButton botoestocminim 
         BackColor       =   &H0080FF80&
         Caption         =   "Estoc mínim"
         Height          =   285
         Left            =   9630
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   6825
         Width           =   1440
      End
      Begin VB.CheckBox checkactives 
         Caption         =   "Només actives."
         Height          =   195
         Left            =   -63915
         TabIndex        =   3
         Top             =   1530
         Value           =   1  'Checked
         Width           =   1530
      End
      Begin VB.CheckBox checkimpresores 
         Caption         =   "Només a Impresores"
         Height          =   195
         Left            =   -63915
         TabIndex        =   2
         Top             =   1305
         Width           =   2400
      End
      Begin VB.CommandButton Command55 
         BackColor       =   &H00EAD9CE&
         Caption         =   "+Lots"
         Height          =   390
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Afegir lots a la llauna seleccionada."
         Top             =   8715
         Width           =   510
      End
      Begin MSDBGrid.DBGrid reixarecarregues 
         Bindings        =   "formtintes.frx":1061C
         Height          =   1860
         Left            =   6375
         OleObjectBlob   =   "formtintes.frx":10636
         TabIndex        =   12
         Top             =   8925
         Width           =   870
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   -74865
         Top             =   990
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   327680
         DTREnable       =   -1  'True
      End
      Begin MSDBGrid.DBGrid reixallaunes 
         Bindings        =   "formtintes.frx":10E82
         Height          =   1665
         Left            =   510
         OleObjectBlob   =   "formtintes.frx":10E98
         TabIndex        =   25
         Top             =   9165
         Width           =   5685
      End
      Begin MSDBGrid.DBGrid reixatintes 
         Bindings        =   "formtintes.frx":120E0
         Height          =   4005
         Left            =   180
         OleObjectBlob   =   "formtintes.frx":120F1
         TabIndex        =   43
         Top             =   1275
         Width           =   14535
      End
      Begin TabDlg.SSTab pestanyesforumes 
         Height          =   9615
         Left            =   -74745
         TabIndex        =   44
         Top             =   1365
         Width           =   14145
         _ExtentX        =   24950
         _ExtentY        =   16960
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "Components"
         TabPicture(0)   =   "formtintes.frx":13A30
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Command67(28)"
         Tab(0).Control(1)=   "Command67(18)"
         Tab(0).Control(2)=   "Command67(6)"
         Tab(0).Control(3)=   "Frame5(0)"
         Tab(0).Control(4)=   "datacomponents"
         Tab(0).Control(5)=   "datalotsbase"
         Tab(0).Control(6)=   "bactualitzacargues"
         Tab(0).Control(7)=   "Command50"
         Tab(0).Control(8)=   "reixacomponents"
         Tab(0).Control(9)=   "etnumlot"
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Fòrmula"
         TabPicture(1)   =   "formtintes.frx":13A4C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Check1(4)"
         Tab(1).Control(1)=   "Check1(3)"
         Tab(1).Control(2)=   "checknomesseleccionats"
         Tab(1).Control(3)=   "Frame7"
         Tab(1).Control(4)=   "datadetallformules"
         Tab(1).Control(5)=   "dataformules"
         Tab(1).Control(6)=   "Command11"
         Tab(1).Control(7)=   "frameactualitzacio"
         Tab(1).Control(8)=   "filtreformuladesc"
         Tab(1).Control(9)=   "filtreformulacodi"
         Tab(1).Control(10)=   "filtreformulaserie"
         Tab(1).Control(11)=   "treurefiltre"
         Tab(1).Control(12)=   "Command43"
         Tab(1).Control(13)=   "dllistadecomponents"
         Tab(1).Control(14)=   "llistallaunesformula"
         Tab(1).Control(15)=   "reixaformules"
         Tab(1).Control(16)=   "Label22(7)"
         Tab(1).Control(17)=   "Label22(6)"
         Tab(1).Control(18)=   "Label22(5)"
         Tab(1).Control(19)=   "Label31"
         Tab(1).ControlCount=   20
         TabCaption(2)   =   " Busqueda de fòrmules semblants"
         TabPicture(2)   =   "formtintes.frx":13A68
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "etiquetatotalkg"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label22(2)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label22(3)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Label22(4)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Line1(0)"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Line1(1)"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "Line1(2)"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "Line1(3)"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "Line1(4)"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "Line1(5)"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "Label22(8)"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).Control(11)=   "Label22(9)"
         Tab(2).Control(11).Enabled=   0   'False
         Tab(2).Control(12)=   "Label22(10)"
         Tab(2).Control(12).Enabled=   0   'False
         Tab(2).Control(13)=   "Label22(11)"
         Tab(2).Control(13).Enabled=   0   'False
         Tab(2).Control(14)=   "Label22(12)"
         Tab(2).Control(14).Enabled=   0   'False
         Tab(2).Control(15)=   "Line1(6)"
         Tab(2).Control(15).Enabled=   0   'False
         Tab(2).Control(16)=   "Line1(7)"
         Tab(2).Control(16).Enabled=   0   'False
         Tab(2).Control(17)=   "Label22(22)"
         Tab(2).Control(17).Enabled=   0   'False
         Tab(2).Control(18)=   "Label22(23)"
         Tab(2).Control(18).Enabled=   0   'False
         Tab(2).Control(19)=   "Line1(8)"
         Tab(2).Control(19).Enabled=   0   'False
         Tab(2).Control(20)=   "Line1(9)"
         Tab(2).Control(20).Enabled=   0   'False
         Tab(2).Control(21)=   "Line1(10)"
         Tab(2).Control(21).Enabled=   0   'False
         Tab(2).Control(22)=   "Label22(24)"
         Tab(2).Control(22).Enabled=   0   'False
         Tab(2).Control(23)=   "Label22(1)"
         Tab(2).Control(23).Enabled=   0   'False
         Tab(2).Control(24)=   "Label22(26)"
         Tab(2).Control(24).Enabled=   0   'False
         Tab(2).Control(25)=   "formulasemblanta"
         Tab(2).Control(25).Enabled=   0   'False
         Tab(2).Control(26)=   "reixaformulacio"
         Tab(2).Control(26).Enabled=   0   'False
         Tab(2).Control(27)=   "Command67(0)"
         Tab(2).Control(27).Enabled=   0   'False
         Tab(2).Control(28)=   "Command68"
         Tab(2).Control(28).Enabled=   0   'False
         Tab(2).Control(29)=   "kgxrecuperar(0)"
         Tab(2).Control(29).Enabled=   0   'False
         Tab(2).Control(30)=   "kgformula"
         Tab(2).Control(30).Enabled=   0   'False
         Tab(2).Control(31)=   "Command69"
         Tab(2).Control(31).Enabled=   0   'False
         Tab(2).Control(32)=   "Command70"
         Tab(2).Control(32).Enabled=   0   'False
         Tab(2).Control(33)=   "botorelacioguardada(0)"
         Tab(2).Control(33).Enabled=   0   'False
         Tab(2).Control(34)=   "botorelacioguardada(1)"
         Tab(2).Control(34).Enabled=   0   'False
         Tab(2).Control(35)=   "llista(0)"
         Tab(2).Control(35).Enabled=   0   'False
         Tab(2).Control(36)=   "llista(1)"
         Tab(2).Control(36).Enabled=   0   'False
         Tab(2).Control(37)=   "formulaacomparar"
         Tab(2).Control(37).Enabled=   0   'False
         Tab(2).Control(38)=   "Command67(1)"
         Tab(2).Control(38).Enabled=   0   'False
         Tab(2).Control(39)=   "Command67(2)"
         Tab(2).Control(39).Enabled=   0   'False
         Tab(2).Control(40)=   "botorelacioguardada(2)"
         Tab(2).Control(40).Enabled=   0   'False
         Tab(2).Control(41)=   "llista(2)"
         Tab(2).Control(41).Enabled=   0   'False
         Tab(2).Control(42)=   "formulaacomparar2"
         Tab(2).Control(42).Enabled=   0   'False
         Tab(2).Control(43)=   "kgxrecuperar(1)"
         Tab(2).Control(43).Enabled=   0   'False
         Tab(2).Control(44)=   "timercontrolfocus"
         Tab(2).Control(44).Enabled=   0   'False
         Tab(2).Control(45)=   "Command67(5)"
         Tab(2).Control(45).Enabled=   0   'False
         Tab(2).Control(46)=   "botorelacioguardada(3)"
         Tab(2).Control(46).Enabled=   0   'False
         Tab(2).Control(47)=   "botorelacioguardada(4)"
         Tab(2).Control(47).Enabled=   0   'False
         Tab(2).Control(48)=   "Command67(7)"
         Tab(2).Control(48).Enabled=   0   'False
         Tab(2).Control(49)=   "Command67(10)"
         Tab(2).Control(49).Enabled=   0   'False
         Tab(2).ControlCount=   50
         Begin VB.CommandButton Command67 
            Caption         =   "Càlcul Grams per Kilo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Index           =   28
            Left            =   -67530
            Style           =   1  'Graphical
            TabIndex        =   296
            Top             =   7740
            Width           =   2010
         End
         Begin VB.CommandButton Command67 
            BackColor       =   &H006BEBB1&
            Caption         =   "Comprovar que els LOTS dels dosificadors estiguin correctament entrats"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1275
            Index           =   18
            Left            =   -65055
            Style           =   1  'Graphical
            TabIndex        =   248
            Top             =   6375
            Width           =   3780
         End
         Begin VB.CheckBox Check1 
            Caption         =   "<-Amb llaunes."
            Height          =   195
            Index           =   4
            Left            =   -67230
            TabIndex        =   243
            Top             =   735
            Value           =   1  'Checked
            Width           =   1620
         End
         Begin VB.CheckBox Check1 
            Caption         =   "<-Nomes Bases."
            Height          =   195
            Index           =   3
            Left            =   -64050
            TabIndex        =   242
            Top             =   675
            Width           =   1470
         End
         Begin VB.CommandButton Command67 
            Height          =   435
            Index           =   10
            Left            =   4965
            Picture         =   "formtintes.frx":13A84
            Style           =   1  'Graphical
            TabIndex        =   232
            Top             =   1770
            Width           =   510
         End
         Begin VB.CommandButton Command67 
            BackColor       =   &H00C78DFA&
            Height          =   450
            Index           =   7
            Left            =   13230
            Picture         =   "formtintes.frx":1400E
            Style           =   1  'Graphical
            TabIndex        =   226
            ToolTipText     =   "Guardar relació amb la comanda activa."
            Top             =   1080
            Width           =   645
         End
         Begin VB.CommandButton Command67 
            Caption         =   "Verificar Lots dosificadors Inkmaker"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Index           =   6
            Left            =   -67530
            Style           =   1  'Graphical
            TabIndex        =   224
            Top             =   7125
            Width           =   2010
         End
         Begin VB.CommandButton botorelacioguardada 
            BackColor       =   &H005C31DD&
            Caption         =   "Kg X"
            Height          =   435
            Index           =   4
            Left            =   12315
            Style           =   1  'Graphical
            TabIndex        =   223
            ToolTipText     =   "Buscar el valor mes proxim a X Kg"
            Top             =   2355
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.CommandButton botorelacioguardada 
            Caption         =   "Kg X"
            Height          =   435
            Index           =   3
            Left            =   13050
            Style           =   1  'Graphical
            TabIndex        =   222
            ToolTipText     =   "Buscar el valor mes proxim a X Kg"
            Top             =   2190
            Width           =   510
         End
         Begin VB.CommandButton Command67 
            Height          =   330
            Index           =   5
            Left            =   5640
            Picture         =   "formtintes.frx":14598
            Style           =   1  'Graphical
            TabIndex        =   221
            ToolTipText     =   "Filtra valors de la sel.lecció de la pestanya de fòrmula."
            Top             =   2040
            Width           =   300
         End
         Begin VB.Timer timercontrolfocus 
            Interval        =   1000
            Left            =   11490
            Top             =   2175
         End
         Begin VB.TextBox kgxrecuperar 
            Alignment       =   2  'Center
            BackColor       =   &H00F1B75F&
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
            Index           =   1
            Left            =   10665
            TabIndex        =   216
            Top             =   2205
            Width           =   705
         End
         Begin VB.ComboBox formulaacomparar2 
            BackColor       =   &H00F1B75F&
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
            Left            =   5955
            TabIndex        =   213
            Top             =   2220
            Width           =   4680
         End
         Begin VB.ListBox llista 
            BackColor       =   &H00F1B75F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2220
            Index           =   2
            Left            =   9795
            TabIndex        =   211
            Top             =   7290
            Width           =   4245
         End
         Begin VB.CommandButton botorelacioguardada 
            Height          =   435
            Index           =   2
            Left            =   13065
            Picture         =   "formtintes.frx":14B22
            Style           =   1  'Graphical
            TabIndex        =   210
            ToolTipText     =   "Relacions guardades -> Combo"
            Top             =   1710
            Width           =   510
         End
         Begin VB.CommandButton Command67 
            Height          =   435
            Index           =   2
            Left            =   12180
            Picture         =   "formtintes.frx":150AC
            Style           =   1  'Graphical
            TabIndex        =   192
            Top             =   1095
            Width           =   1035
         End
         Begin VB.CommandButton Command67 
            Height          =   435
            Index           =   1
            Left            =   12000
            Picture         =   "formtintes.frx":15636
            Style           =   1  'Graphical
            TabIndex        =   190
            ToolTipText     =   "Filtra valors de la sel.lecció de la pestanya de fòrmula."
            Top             =   1710
            Width           =   510
         End
         Begin VB.ComboBox formulaacomparar 
            BackColor       =   &H00F8FDB5&
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
            ItemData        =   "formtintes.frx":15BC0
            Left            =   5970
            List            =   "formtintes.frx":15BC2
            TabIndex        =   189
            Top             =   1755
            Width           =   4650
         End
         Begin VB.ListBox llista 
            BackColor       =   &H00F8FDB5&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2220
            Index           =   1
            Left            =   5205
            TabIndex        =   184
            Top             =   7290
            Width           =   4245
         End
         Begin VB.ListBox llista 
            BackColor       =   &H00EAD9CE&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2220
            Index           =   0
            Left            =   630
            TabIndex        =   183
            Top             =   7290
            Width           =   4245
         End
         Begin VB.CommandButton botorelacioguardada 
            Height          =   345
            Index           =   1
            Left            =   4785
            Picture         =   "formtintes.frx":15BC4
            Style           =   1  'Graphical
            TabIndex        =   182
            ToolTipText     =   "Relacions guardades -> llista"
            Top             =   1380
            Width           =   510
         End
         Begin VB.CommandButton botorelacioguardada 
            Height          =   345
            Index           =   0
            Left            =   4245
            Picture         =   "formtintes.frx":1614E
            Style           =   1  'Graphical
            TabIndex        =   181
            Top             =   1380
            Width           =   510
         End
         Begin VB.CommandButton Command70 
            BackColor       =   &H00C0FFC0&
            Height          =   435
            Left            =   11340
            Picture         =   "formtintes.frx":166D8
            Style           =   1  'Graphical
            TabIndex        =   174
            ToolTipText     =   "Busca la combinació correcte."
            Top             =   1095
            Width           =   825
         End
         Begin VB.CommandButton Command69 
            Height          =   435
            Left            =   10830
            Picture         =   "formtintes.frx":16C62
            Style           =   1  'Graphical
            TabIndex        =   172
            ToolTipText     =   "Recalcula la formulació"
            Top             =   1095
            Width           =   495
         End
         Begin VB.TextBox kgformula 
            Alignment       =   2  'Center
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
            Left            =   1890
            TabIndex        =   171
            Top             =   945
            Width           =   1140
         End
         Begin VB.TextBox kgxrecuperar 
            Alignment       =   2  'Center
            BackColor       =   &H00F8FDB5&
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
            Index           =   0
            Left            =   10680
            TabIndex        =   170
            Top             =   1740
            Width           =   675
         End
         Begin VB.CommandButton Command68 
            Height          =   420
            Left            =   11475
            Picture         =   "formtintes.frx":171EC
            Style           =   1  'Graphical
            TabIndex        =   169
            ToolTipText     =   "Buscar semblants a la Fòrmula 1"
            Top             =   1725
            Width           =   510
         End
         Begin VB.CommandButton Command67 
            Height          =   435
            Index           =   0
            Left            =   12540
            Picture         =   "formtintes.frx":17776
            Style           =   1  'Graphical
            TabIndex        =   168
            Top             =   1710
            Width           =   510
         End
         Begin MSFlexGridLib.MSFlexGrid reixaformulacio 
            Height          =   3615
            Left            =   150
            TabIndex        =   167
            Top             =   2850
            Width           =   13620
            _ExtentX        =   24024
            _ExtentY        =   6376
            _Version        =   393216
            Cols            =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox formulasemblanta 
            BackColor       =   &H00EAD9CE&
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   166
            Top             =   1770
            Width           =   4845
         End
         Begin VB.CheckBox checknomesseleccionats 
            Caption         =   "Només seleccionats"
            Height          =   195
            Left            =   -65640
            TabIndex        =   155
            Top             =   345
            Visible         =   0   'False
            Width           =   2310
         End
         Begin VB.Frame Frame5 
            Caption         =   "Detall de cargues i lots "
            Height          =   3030
            Index           =   0
            Left            =   -74880
            TabIndex        =   61
            Top             =   6300
            Width           =   7020
            Begin MSDBGrid.DBGrid reixalotsbase 
               Bindings        =   "formtintes.frx":17D00
               Height          =   2220
               Left            =   240
               OleObjectBlob   =   "formtintes.frx":17D17
               TabIndex        =   62
               Top             =   255
               Width           =   6540
            End
         End
         Begin VB.Data datacomponents 
            Caption         =   "datacomponents"
            Connect         =   "Access"
            DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\tintes.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -69750
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Componentsbase"
            Top             =   1620
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Data datalotsbase 
            Caption         =   "datalotsbase"
            Connect         =   "Access"
            DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\tintes.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -69360
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "SELECT* FROM detallnumeroslotsbase"
            Top             =   6240
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Frame Frame7 
            Caption         =   "Detall de la fòrmula"
            Height          =   3075
            Left            =   -74775
            TabIndex        =   57
            Top             =   6480
            Width           =   7890
            Begin VB.TextBox kgxrecuperar 
               Alignment       =   2  'Center
               BackColor       =   &H00F1B75F&
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
               Index           =   2
               Left            =   6990
               TabIndex        =   218
               Top             =   750
               Width           =   810
            End
            Begin MSDBGrid.DBGrid reixadetallformules 
               Bindings        =   "formtintes.frx":186FC
               Height          =   2505
               Left            =   60
               OleObjectBlob   =   "formtintes.frx":18719
               TabIndex        =   58
               Top             =   255
               Width           =   6885
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Anilox Formulat"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   25
               Left            =   6720
               TabIndex        =   219
               Top             =   240
               Width           =   1350
            End
            Begin VB.Label ltotalpercent 
               BackStyle       =   0  'Transparent
               Caption         =   "Total :     100%"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4215
               TabIndex        =   59
               Top             =   2745
               Width           =   2820
            End
         End
         Begin VB.Data datadetallformules 
            Caption         =   "datadetallformules"
            Connect         =   "Access"
            DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\tintes.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -68040
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   6210
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Data dataformules 
            Caption         =   "dataformules"
            Connect         =   "Access"
            DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\tintes.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -66135
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Formules"
            Top             =   315
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Actualitzar Formules amb l'InkMaker"
            Height          =   480
            Left            =   -62505
            TabIndex        =   56
            Top             =   495
            Width           =   1530
         End
         Begin VB.Frame frameactualitzacio 
            Height          =   2400
            Left            =   -70860
            TabIndex        =   54
            Top             =   2175
            Visible         =   0   'False
            Width           =   5145
            Begin VB.Label msgactualitzacio 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   12
                  Charset         =   0
                  Weight          =   900
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   255
               TabIndex        =   55
               Top             =   885
               Width           =   4695
            End
         End
         Begin VB.TextBox filtreformuladesc 
            Height          =   345
            Left            =   -72285
            TabIndex        =   53
            Top             =   675
            Width           =   3405
         End
         Begin VB.TextBox filtreformulacodi 
            Height          =   345
            Left            =   -74175
            TabIndex        =   52
            Top             =   675
            Width           =   1905
         End
         Begin VB.TextBox filtreformulaserie 
            Height          =   345
            Left            =   -68865
            TabIndex        =   51
            Top             =   675
            Width           =   1560
         End
         Begin VB.CommandButton treurefiltre 
            Height          =   315
            Left            =   -74580
            Picture         =   "formtintes.frx":19494
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Neteja filtre"
            Top             =   675
            Width           =   360
         End
         Begin VB.CommandButton bactualitzacargues 
            Caption         =   "Actualtizar Cargues Components"
            Height          =   585
            Left            =   -67530
            TabIndex        =   49
            Top             =   8370
            Width           =   2055
         End
         Begin VB.CommandButton Command43 
            Caption         =   "Filtrar formula amb component concret"
            Height          =   465
            Left            =   -65640
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   510
            Width           =   1575
         End
         Begin VB.ListBox dllistadecomponents 
            BackColor       =   &H00EAD9CE&
            Height          =   4785
            ItemData        =   "formtintes.frx":19A1E
            Left            =   -65625
            List            =   "formtintes.frx":19A25
            Style           =   1  'Checkbox
            TabIndex        =   47
            Top             =   1035
            Visible         =   0   'False
            Width           =   4320
         End
         Begin VB.CommandButton Command50 
            Caption         =   "Imp etiqueta dosificador"
            Height          =   585
            Left            =   -67545
            Picture         =   "formtintes.frx":19A3E
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Imprimeix la etiqueta del dosificador per utilitzar-la a impresores."
            Top             =   6375
            Width           =   2055
         End
         Begin VB.ListBox llistallaunesformula 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2220
            Left            =   -66465
            TabIndex        =   45
            Top             =   6735
            Width           =   4245
         End
         Begin MSDBGrid.DBGrid reixacomponents 
            Bindings        =   "formtintes.frx":19FC8
            Height          =   5535
            Left            =   -74820
            OleObjectBlob   =   "formtintes.frx":19FE1
            TabIndex        =   60
            Top             =   465
            Width           =   13650
         End
         Begin MSDBGrid.DBGrid reixaformules 
            Bindings        =   "formtintes.frx":1B250
            Height          =   5295
            Left            =   -74700
            OleObjectBlob   =   "formtintes.frx":1B267
            TabIndex        =   63
            Top             =   1035
            Width           =   13650
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "F3:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   26
            Left            =   5655
            TabIndex        =   220
            Top             =   2355
            Width           =   405
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   330
            Index           =   1
            Left            =   9780
            TabIndex        =   217
            Top             =   7005
            Width           =   4320
         End
         Begin VB.Label Label22 
            Caption         =   " Kg"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   24
            Left            =   10965
            TabIndex        =   215
            Top             =   1440
            Width           =   405
         End
         Begin VB.Line Line1 
            Index           =   10
            X1              =   11430
            X2              =   11430
            Y1              =   1605
            Y2              =   2685
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   5580
            X2              =   11445
            Y1              =   2685
            Y2              =   2685
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   5580
            X2              =   5580
            Y1              =   1575
            Y2              =   2700
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "F2:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   23
            Left            =   5655
            TabIndex        =   214
            Top             =   1800
            Width           =   405
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "F3:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   22
            Left            =   9465
            TabIndex        =   212
            Top             =   7380
            Width           =   375
         End
         Begin VB.Line Line1 
            Index           =   7
            X1              =   13245
            X2              =   13245
            Y1              =   1590
            Y2              =   1755
         End
         Begin VB.Line Line1 
            Index           =   6
            X1              =   8745
            X2              =   13245
            Y1              =   1590
            Y2              =   1590
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   12
            Left            =   4830
            TabIndex        =   191
            Top             =   1110
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "F2:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   11
            Left            =   4875
            TabIndex        =   188
            Top             =   7380
            Width           =   375
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "F1:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   10
            Left            =   300
            TabIndex        =   187
            Top             =   7335
            Width           =   375
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   330
            Index           =   9
            Left            =   5145
            TabIndex        =   186
            Top             =   7005
            Width           =   4320
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   330
            Index           =   8
            Left            =   615
            TabIndex        =   185
            Top             =   7005
            Width           =   4320
         End
         Begin VB.Line Line1 
            Index           =   5
            X1              =   6165
            X2              =   6315
            Y1              =   1650
            Y2              =   1575
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   6165
            X2              =   6330
            Y1              =   1500
            Y2              =   1590
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   3180
            X2              =   3315
            Y1              =   1575
            Y2              =   1620
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   3150
            X2              =   3330
            Y1              =   1560
            Y2              =   1500
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   4950
            X2              =   6315
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   3195
            X2              =   4350
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Sèrie"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   -68640
            TabIndex        =   180
            Top             =   405
            Width           =   1050
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcio fòrmula"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   -71970
            TabIndex        =   179
            Top             =   420
            Width           =   3210
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Codi fòrmula"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   -74055
            TabIndex        =   178
            Top             =   390
            Width           =   1620
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Fòrmules (a comparar)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   6405
            TabIndex        =   177
            Top             =   1455
            Width           =   3345
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Fòrmula 1 (Comparada)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   780
            TabIndex        =   176
            Top             =   1425
            Width           =   2310
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Kg formula:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   600
            TabIndex        =   175
            Top             =   1005
            Width           =   1875
         End
         Begin VB.Label etiquetatotalkg 
            Caption         =   "Total Grms dosificador: "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   8730
            TabIndex        =   173
            Top             =   6420
            Width           =   5115
         End
         Begin VB.Label etnumlot 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74700
            TabIndex        =   65
            Top             =   6075
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Label Label31 
            Caption         =   "Llaunes disponibles"
            Height          =   240
            Left            =   -66360
            TabIndex        =   64
            Top             =   6510
            Width           =   2730
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "formtintes.frx":1C164
         Height          =   8865
         Left            =   -74805
         OleObjectBlob   =   "formtintes.frx":1C17C
         TabIndex        =   66
         Top             =   2040
         Width           =   13785
      End
      Begin MSFlexGridLib.MSFlexGrid reixacomandes 
         Height          =   5700
         Left            =   -74745
         TabIndex        =   116
         Top             =   1230
         Width           =   14625
         _ExtentX        =   25797
         _ExtentY        =   10054
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.CommandButton Command54 
         Caption         =   "--> Llista activades"
         Height          =   420
         Left            =   -61710
         TabIndex        =   131
         ToolTipText     =   "Passa la/es comanda/es a activades."
         Top             =   7380
         Width           =   1500
      End
      Begin VB.Frame framedadestintes 
         Caption         =   "Manteniment de la Tinta"
         Enabled         =   0   'False
         Height          =   5070
         Left            =   195
         TabIndex        =   67
         Top             =   5910
         Width           =   14565
         Begin VB.CommandButton Command15 
            Height          =   285
            Left            =   6375
            Picture         =   "formtintes.frx":1D073
            Style           =   1  'Graphical
            TabIndex        =   158
            ToolTipText     =   "Buscar noms iguals per agrupar."
            Top             =   795
            Width           =   345
         End
         Begin VB.TextBox cnominplacsa 
            BackColor       =   &H00C0FFC0&
            DataField       =   "nominplacsa"
            DataSource      =   "tintes"
            Height          =   300
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   157
            ToolTipText     =   "Nom que utilitzem a inplacsa per agrupar tintes iguals però amb nom diferent."
            Top             =   780
            Width           =   4920
         End
         Begin VB.TextBox cdataokcarta 
            DataField       =   "dataokcarta"
            DataSource      =   "tintes"
            Height          =   285
            Index           =   0
            Left            =   13725
            TabIndex        =   153
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ok Carta. (Digital)"
            DataField       =   "okcarta"
            DataSource      =   "tintes"
            Height          =   375
            Index           =   0
            Left            =   12315
            TabIndex        =   148
            Top             =   1065
            Width           =   1080
         End
         Begin VB.Frame Frame4 
            Caption         =   "Colors de la tinta"
            Height          =   1035
            Left            =   5490
            TabIndex        =   82
            Top             =   1080
            Width           =   5430
            Begin VB.ComboBox cfamiliacolor 
               DataField       =   "descripciofamcol"
               Height          =   315
               Left            =   1020
               TabIndex        =   84
               Top             =   225
               Width           =   4380
            End
            Begin VB.ComboBox csubfamiliacolor 
               DataField       =   "descripciosubfamcol"
               Height          =   315
               Left            =   1005
               TabIndex        =   83
               Top             =   570
               Width           =   4380
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "SubFamilia:"
               Height          =   330
               Index           =   21
               Left            =   120
               TabIndex        =   209
               Top             =   600
               Width           =   810
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Familia:"
               Height          =   330
               Index           =   20
               Left            =   150
               TabIndex        =   208
               Top             =   270
               Width           =   810
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Families de la tinta"
            Height          =   1035
            Left            =   270
            TabIndex        =   79
            Top             =   1080
            Width           =   5205
            Begin VB.ComboBox csubfamilia 
               DataField       =   "descripciosubfam"
               DragMode        =   1  'Automatic
               Height          =   315
               Left            =   915
               TabIndex        =   81
               Top             =   615
               Width           =   4230
            End
            Begin VB.ComboBox nomfamilia 
               DataField       =   "descripciofam"
               Height          =   315
               Left            =   930
               TabIndex        =   80
               Top             =   240
               Width           =   4230
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "SubFamilia:"
               Height          =   330
               Index           =   19
               Left            =   105
               TabIndex        =   207
               Top             =   645
               Width           =   810
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Familia:"
               Height          =   330
               Index           =   18
               Left            =   105
               TabIndex        =   206
               Top             =   270
               Width           =   810
            End
         End
         Begin VB.TextBox crefcolor 
            DataField       =   "referenciacolor"
            DataSource      =   "tintes"
            Height          =   315
            Left            =   6360
            MaxLength       =   25
            TabIndex        =   78
            Top             =   450
            Width           =   1605
         End
         Begin VB.TextBox descripciotinta 
            BackColor       =   &H00E0E0E0&
            DataField       =   "descripcio"
            DataSource      =   "tintes"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   77
            Top             =   450
            Width           =   4905
         End
         Begin VB.TextBox ccoditinta 
            BackColor       =   &H00E0E0E0&
            DataField       =   "codi"
            DataSource      =   "tintes"
            Enabled         =   0   'False
            Height          =   315
            Left            =   270
            MaxLength       =   30
            TabIndex        =   76
            Top             =   450
            Width           =   1020
         End
         Begin VB.Data datatintesformules 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\tintes.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   13125
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "tintesformules"
            Top             =   1620
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Frame Frame9 
            Caption         =   "Formules relacionades"
            Height          =   1725
            Left            =   11100
            TabIndex        =   71
            Top             =   885
            Width           =   3390
            Begin VB.TextBox cdataokcarta 
               DataField       =   "dataokcartapaper"
               DataSource      =   "tintes"
               Height          =   285
               Index           =   1
               Left            =   2640
               TabIndex        =   230
               Top             =   285
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Ok Carta. (Paper)"
               DataField       =   "okcartapaper"
               DataSource      =   "tintes"
               Height          =   375
               Index           =   1
               Left            =   2295
               TabIndex        =   229
               Top             =   120
               Width           =   1080
            End
            Begin VB.CommandButton Command34 
               Height          =   330
               Left            =   405
               Picture         =   "formtintes.frx":1D5FD
               Style           =   1  'Graphical
               TabIndex        =   74
               ToolTipText     =   "Eliminar formula"
               Top             =   225
               Width           =   345
            End
            Begin VB.CommandButton Command35 
               Height          =   330
               Left            =   765
               Picture         =   "formtintes.frx":1DB87
               Style           =   1  'Graphical
               TabIndex        =   73
               ToolTipText     =   "Marcar formula principal."
               Top             =   225
               Width           =   345
            End
            Begin VB.CommandButton Command37 
               Height          =   360
               Left            =   30
               Picture         =   "formtintes.frx":1E111
               Style           =   1  'Graphical
               TabIndex        =   72
               ToolTipText     =   "Vincular tinta amb formula"
               Top             =   195
               Width           =   360
            End
            Begin MSDBGrid.DBGrid reixatintesformules 
               Bindings        =   "formtintes.frx":1E69B
               Height          =   915
               Left            =   75
               OleObjectBlob   =   "formtintes.frx":1E6B8
               TabIndex        =   75
               Top             =   570
               Width           =   3255
            End
         End
         Begin VB.ComboBox nomserie 
            DataField       =   "DescripcioSerie"
            Height          =   315
            Left            =   8100
            TabIndex        =   70
            Top             =   420
            Width           =   1305
         End
         Begin VB.Frame fmostracolortinta 
            Caption         =   "Mostra del Color"
            Height          =   540
            Left            =   9660
            TabIndex        =   69
            Top             =   105
            Width           =   1455
         End
         Begin VB.TextBox cobservacions 
            Height          =   300
            Left            =   285
            MaxLength       =   200
            TabIndex        =   68
            Top             =   2325
            Width           =   10635
         End
         Begin VB.Frame Frame1 
            Caption         =   "Llaunes per tinta"
            Height          =   2400
            Left            =   285
            TabIndex        =   85
            Top             =   2610
            Width           =   14175
            Begin VB.CommandButton Command3 
               Height          =   390
               Left            =   1305
               Picture         =   "formtintes.frx":1F0C3
               Style           =   1  'Graphical
               TabIndex        =   89
               Top             =   210
               Width           =   390
            End
            Begin VB.CommandButton Command4 
               Height          =   390
               Left            =   525
               Picture         =   "formtintes.frx":1F64D
               Style           =   1  'Graphical
               TabIndex        =   88
               ToolTipText     =   "Consulta Registres"
               Top             =   210
               Width           =   390
            End
            Begin VB.CommandButton Command5 
               Height          =   390
               Left            =   915
               Picture         =   "formtintes.frx":1FBD7
               Style           =   1  'Graphical
               TabIndex        =   87
               ToolTipText     =   "Eliminacio Registres"
               Top             =   210
               Width           =   390
            End
            Begin VB.Data datallaunes 
               Caption         =   "datallaunes"
               Connect         =   "Access"
               DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\tintes.mdb"
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   345
               Left            =   765
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   $"formtintes.frx":20161
               Top             =   -135
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.Data datahistoria 
               Caption         =   "datahistoria"
               Connect         =   "Access"
               DatabaseName    =   "\\serverprodu\dades\progcomandes\dades\tintes.mdb"
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   345
               Left            =   9735
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   "historiallauna"
               Top             =   690
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.Frame Frame10 
               Caption         =   "Nº Rec."
               Height          =   2235
               Left            =   5820
               TabIndex        =   86
               Top             =   120
               Width           =   1005
               Begin VB.Data datarecarregues 
                  Caption         =   "datarecarregues"
                  Connect         =   "Access"
                  DatabaseName    =   ""
                  DefaultCursorType=   0  'DefaultCursor
                  DefaultType     =   2  'UseODBC
                  Exclusive       =   0   'False
                  Height          =   345
                  Left            =   150
                  Options         =   0
                  ReadOnly        =   0   'False
                  RecordsetType   =   1  'Dynaset
                  RecordSource    =   ""
                  Top             =   1890
                  Visible         =   0   'False
                  Width           =   1140
               End
            End
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Nom serie"
            Height          =   330
            Index           =   17
            Left            =   8310
            TabIndex        =   197
            Top             =   210
            Width           =   1035
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Referència del color"
            Height          =   330
            Index           =   16
            Left            =   6435
            TabIndex        =   196
            Top             =   225
            Width           =   1875
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripció Tinta"
            Height          =   330
            Index           =   15
            Left            =   1800
            TabIndex        =   195
            Top             =   225
            Width           =   1875
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Codi Tinta"
            Height          =   330
            Index           =   14
            Left            =   390
            TabIndex        =   194
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Nom Inplacsa:"
            Height          =   330
            Index           =   13
            Left            =   240
            TabIndex        =   193
            Top             =   795
            Width           =   1875
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Observacions:"
            Height          =   240
            Left            =   360
            TabIndex        =   90
            Top             =   2100
            Width           =   1275
         End
      End
      Begin VB.CommandButton Command14 
         Height          =   420
         Left            =   -74055
         Picture         =   "formtintes.frx":2028F
         Style           =   1  'Graphical
         TabIndex        =   154
         ToolTipText     =   "Historia d'impresions d'una comanda."
         Top             =   525
         Width           =   660
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calcular Consum T"
         Height          =   195
         Index           =   2
         Left            =   -65205
         TabIndex        =   233
         Top             =   6930
         Width           =   1830
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº màquina"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   29
         Left            =   -73290
         TabIndex        =   238
         Top             =   6915
         Width           =   825
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Tintes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   28
         Left            =   -74820
         TabIndex        =   236
         Top             =   10275
         Width           =   900
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Obs.Op:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   27
         Left            =   -74835
         TabIndex        =   235
         Top             =   9975
         Width           =   900
      End
      Begin VB.Label etenviantemail 
         BackStyle       =   0  'Transparent
         Caption         =   "Enviant E-mail a compres..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   -66390
         TabIndex        =   205
         Top             =   1140
         Visible         =   0   'False
         Width           =   4260
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Doble click per modificar "
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   -68190
         TabIndex        =   201
         Top             =   1500
         Width           =   2745
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "       Ref:         Descripció de la compra                         Observació"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74220
         TabIndex        =   200
         Top             =   1485
         Width           =   11580
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Llista de compres pendents"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -72840
         TabIndex        =   199
         Top             =   585
         Width           =   10395
      End
      Begin VB.Label etextensio 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -70125
         TabIndex        =   156
         Top             =   7035
         Width           =   60
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Ok Carta?"
         Height          =   270
         Left            =   870
         TabIndex        =   152
         Top             =   420
         Width           =   915
      End
      Begin VB.Label etfiltrarperestoc 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar per descripció"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -69960
         TabIndex        =   143
         Tag             =   "descripcio"
         Top             =   1080
         Width           =   4080
      End
      Begin VB.Label ettotalestocs 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   -74745
         TabIndex        =   141
         Top             =   10605
         Width           =   5250
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   " Màq/Ord Lot/Nova  Munt Treball     Metres          Nom del client                           Texte Imp."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74580
         TabIndex        =   138
         Top             =   6075
         Width           =   11580
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Munt"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -72585
         MouseIcon       =   "formtintes.frx":20819
         MousePointer    =   99  'Custom
         OLEDropMode     =   1  'Manual
         TabIndex        =   137
         Top             =   6075
         Width           =   630
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "#T Codi  Ani  Nom de la tinta ( * té extensió feta)   NºLlauna/Sit    Consum T/Aquesta  Estoc Kg"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74835
         TabIndex        =   136
         Top             =   7230
         Width           =   13125
      End
      Begin VB.Label etactualitzant 
         BackStyle       =   0  'Transparent
         Caption         =   "Actualitzant..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   345
         Left            =   -69240
         TabIndex        =   135
         Top             =   300
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.Label etajudadosclics 
         BackColor       =   &H0000FFFF&
         Caption         =   "Dos clics per anar a la tinta sel.leccionada. Multiple sel.lecció per exportar a XLS els que vulguis."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74805
         TabIndex        =   134
         Top             =   10650
         Visible         =   0   'False
         Width           =   8685
      End
      Begin VB.Label ettotalcomandes 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   -74640
         TabIndex        =   133
         Top             =   6915
         Width           =   1785
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "+10Kg"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -63975
         TabIndex        =   132
         Top             =   7080
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Llista de comandes pendents de fer etiquetes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -73065
         TabIndex        =   104
         Top             =   810
         Width           =   10395
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Comanda Prov.      Descripció de la compra                 Bidó   Kilos"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74460
         TabIndex        =   103
         Top             =   1740
         Width           =   11580
      End
      Begin VB.Label etcreantllaunes 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   -73545
         TabIndex        =   102
         Top             =   1320
         Width           =   8880
      End
      Begin VB.Label etfiltrarper 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar per descripció"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4740
         TabIndex        =   97
         Tag             =   "descripcio"
         Top             =   330
         Visible         =   0   'False
         Width           =   4080
      End
      Begin VB.Label campabuscar 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   405
         Left            =   -67935
         TabIndex        =   96
         Top             =   1515
         Width           =   2820
      End
      Begin VB.Label etquant 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -71055
         TabIndex        =   95
         Top             =   10485
         Width           =   3885
      End
      Begin VB.Label etquancolor 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -71430
         TabIndex        =   94
         Top             =   10365
         Width           =   3885
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   $"formtintes.frx":20DA3
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   -74175
         TabIndex        =   93
         Top             =   1200
         Width           =   10125
      End
      Begin VB.Label etajudabusqueda 
         BackColor       =   &H0080FFFF&
         Caption         =   "Podeu buscar per paraules separades... ex: AMARILLO RF"
         Height          =   210
         Left            =   1605
         TabIndex        =   92
         Top             =   750
         Width           =   5070
      End
      Begin VB.Label ettotalllaunes 
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F3B378&
         Height          =   195
         Left            =   -74250
         TabIndex        =   91
         Top             =   1800
         Width           =   11565
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Sense connexió a Inkmaker"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   38
         Left            =   4920
         TabIndex        =   284
         Top             =   375
         Visible         =   0   'False
         Width           =   5280
      End
   End
   Begin VB.Menu m_manteniments 
      Caption         =   "Manteniments"
      Begin VB.Menu m_families 
         Caption         =   "Families Tintes"
      End
      Begin VB.Menu m_familiescolors 
         Caption         =   "Families colors"
         Begin VB.Menu mfamiliacolors 
            Caption         =   "Familia Colors"
         End
         Begin VB.Menu msubfamcolors 
            Caption         =   "Subfamilia Colors"
         End
      End
      Begin VB.Menu mseries 
         Caption         =   "Series"
      End
      Begin VB.Menu mlotserronis 
         Caption         =   "Lots marcats com a Erronis"
      End
      Begin VB.Menu mdetalldelstinters 
         Caption         =   "Detalls dels tinters"
      End
      Begin VB.Menu menu_mantdebido 
         Caption         =   "Manteniment de Bidons/Contenidors"
         Begin VB.Menu mtipusbidons 
            Caption         =   "Tipus de Bidons/Contenidors"
         End
         Begin VB.Menu m_tipusdecontenidorsmaterials 
            Caption         =   "Tipus de contenidors (Material amb que estan fets)"
         End
      End
      Begin VB.Menu msituacions 
         Caption         =   "Situacions de les llaunes"
      End
      Begin VB.Menu mbidonsperllençar 
         Caption         =   "Bidons per llençar/reciclar"
         Begin VB.Menu mpaletsxrllençar 
            Caption         =   "Palets per Llençar"
         End
         Begin VB.Menu mpaletsperreciclar 
            Caption         =   "Palets per Reciclar"
         End
      End
   End
   Begin VB.Menu mllistats 
      Caption         =   "Llistats"
      Begin VB.Menu mllistatdellaunes 
         Caption         =   "Llistat de llaunes"
         Begin VB.Menu mlltoteslesllaunes 
            Caption         =   "Totes les llaunes"
         End
         Begin VB.Menu mllllaunes20i25 
            Caption         =   "Llaunes de 20kg i 25 kg (Filtre kg)"
         End
         Begin VB.Menu mestadisticallaunesinplacsa 
            Caption         =   "Estadistica llaunes inplacsa"
         End
      End
      Begin VB.Menu mllistatllauneaajuntar 
         Caption         =   "Llistat de llaunes per ajuntar (de 2 fer-ne 1)"
      End
      Begin VB.Menu mllistatdellaunesxrpalet 
         Caption         =   "Llistat de llaunes per palet"
      End
      Begin VB.Menu llistatokcartaentredates 
         Caption         =   "Llistat ok carta entre dates"
      End
      Begin VB.Menu mllistatcontenidors 
         Caption         =   "Llistat de contenidors"
      End
      Begin VB.Menu mllistatnoactivesambkg 
         Caption         =   "Llistat de llaunes No actives amb Kg"
      End
      Begin VB.Menu mllistatllaunesamb1_7kg 
         Caption         =   "Llistat de llaunes amb  -1,7 Kg"
      End
      Begin VB.Menu mllistatdellaunesambasterisc 
         Caption         =   "Llistat de llaunes a impresores (*)"
      End
      Begin VB.Menu mrelaciodedeltes 
         Caption         =   "Relació de deltes d'una comanda concreta."
      End
      Begin VB.Menu mtotalKgrec 
         Caption         =   "Total de Kg recuperats entre dates"
      End
   End
   Begin VB.Menu mutils 
      Caption         =   "Utilitats"
      Begin VB.Menu mregularitzacioinventarillaunes 
         Caption         =   "Regularització d'Inventari de Llaunes"
      End
      Begin VB.Menu mcalcularpreukg 
         Caption         =   "Calcular preu/kg llaunes a zero"
      End
      Begin VB.Menu mcanvirecuperador 
         Caption         =   "Canvi de recuperador del contenidor"
      End
      Begin VB.Menu mguardaramplescomandes 
         Caption         =   "Guardar amples reixa comandes"
      End
   End
   Begin VB.Menu mveuredeltes 
      Caption         =   "Veure Deltes"
      Visible         =   0   'False
   End
   Begin VB.Menu mllegenda 
      Caption         =   "Llegenda (Ajuda)"
      Visible         =   0   'False
   End
   Begin VB.Menu mtintesrevisades 
      Caption         =   "Tintes Revisades"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu msubmenutintesrevisades 
         Caption         =   "Tintes Revisades"
         Begin VB.Menu msi 
            Caption         =   "Si"
         End
         Begin VB.Menu mno 
            Caption         =   "No"
         End
      End
   End
   Begin VB.Menu m_menucopiarpegar 
      Caption         =   "menucopiarpegar"
      Visible         =   0   'False
      Begin VB.Menu mcopiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu mpegar 
         Caption         =   "Pegar"
      End
   End
End
Attribute VB_Name = "formtintes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Dim vtempsbotoapretat As Date
Dim vmsgcontroldosificadors As String
Dim dbplanificacioalicia As Database
Dim rstCdL As Recordset
Dim rstCdLestats As Recordset
Dim vcolreixacomandes As Long
Dim werescomandes As String
Dim vcodiformules2() As String * 50
Dim vcodiformules() As String * 50
Dim campsestoc(50) As String
Dim ordreestoc As String
Dim ultimwerescomandes As String
Dim vfocusultimcontrol As String
Dim vkgllaunaf2 As Double
Dim vkgllaunaf3 As Double
Dim vcontrasenyavalida As Boolean
Dim vllistatllaunesautomatic As Boolean
Dim vrstCloneComandes As Recordset

Private Function EstaCorriendo(ByVal NombreDelProceso As String) As Boolean
    Const MAX_PATH As Long = 260
    Dim con As Byte
    Dim lProcesses() As Long, lModules() As Long, N As Long, lRet As Long, hProcess As Long
    Dim sName As String
    NombreDelProceso = UCase$(NombreDelProceso)
    ReDim lProcesses(1023) As Long
 con = 0
    If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
        For N = 0 To (lRet \ 4) - 1
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(N))
            If hProcess Then
                ReDim lModules(1023)
                If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
                    sName = String$(MAX_PATH, vbNullChar)
                    GetModuleBaseName hProcess, lModules(0), sName, MAX_PATH
                    sName = Left$(sName, InStr(sName, vbNullChar) - 1)
 
                    If Len(sName) = Len(NombreDelProceso) Then
                        If NombreDelProceso = UCase$(sName) Then con = con + 1
                        If con = 1 Then EstaCorriendo = True: Exit Function
                    End If
                End If
            End If
            CloseHandle hProcess
        Next N
    End If
End Function

Private Sub bactualitzarcompres_Click()
    
End Sub

Private Sub balternatives_Click()
  Dim rst As Recordset
  Dim vnumtreball As String
  Dim vmsg As String
  vnumtreball = reixacomandes.TextMatrix(reixacomandes.Row, numcol("NºTreball"))
  Set rst = dbclixes.OpenRecordset("SELECT Tintes.id_treball, Tintes_alternatives.* FROM Tintes RIGHT JOIN Tintes_alternatives ON Tintes.id_tinter = Tintes_alternatives.id_tinter where id_treball=" + atrim(vnumtreball))
  While Not rst.EOF
    vmsg = vmsg + atrim(rst!coditinta) + "-" + atrim(rst!color) + vbNewLine
    rst.MoveNext
  Wend
  If vmsg <> "" Then MsgBox vmsg, vbInformation, "Tintes Alternatives del treball " + vnumtreball
  Set rst = Nothing
End Sub

Private Sub botorelacioguardada_Click(Index As Integer)
   Dim vkg As Double
   If Index = 0 Then guardarrelacióformules
   If Index = 1 Then obrirrelacióformules
   If Index = 2 Then carregarrelacióformules
   If Index = 3 Then
      vkg = cadbl(InputBox("Entra els Kg de tinta totals que vols tenir.", "Kg Tinta"))
      buscarkgoptimsperlabarreja vkg
   End If
   If Index = 4 Then buscarkgoptimspertoteslesformules
End Sub
Sub buscarkgoptimspertoteslesformules()
    Dim vformula2 As String
    Dim vformula3 As String
    Dim vindexf2 As Long
    Dim vindexf3 As Long
    Dim vkg As Double
    vkg = cadbl(InputBox("Entra els Kg de tinta totals que vols tenir.", "Kg Tinta"))
    If vkg = 0 Then Exit Sub
    
    For f2 = 0 To formulaacomparar.ListCount - 1
      formulaacomparar.ListIndex = f2
      For f3 = 0 To formulaacomparar2.ListCount - 1
        formulaacomparar2.ListIndex = f3
        buscarkgoptimsperlabarreja vkg
        If vultimvalor = 0 Then vultimvalor = etiquetatotalkg.WhatsThisHelpID
        If vultimvalor > etiquetatotalkg.WhatsThisHelpID Then
          vultimf2 = kgxrecuperar(0)
          vultimf3 = kgxrecuperar(1)
          vformula2 = formulaacomparar.tag
          vformula3 = formulaacomparar2.tag
          vindexf2 = f2
          vindexf3 = f3
        End If
      Next f3
    Next f2
    formulaacomparar2.ListIndex = vindexf3
    formulaacomparar.ListIndex = vindexf2
    kgxrecuperar(0) = vultimf2
    kgxrecuperar(1) = vultimf3
    buscarkgoptimsperlabarreja vkg
End Sub
Sub buscarkgoptimsperlabarreja(vkg As Double)
  Dim vkgmaxf2 As Double
  Dim vkgmaxf3 As Double
  Dim vultimvalor As Double
  Dim vultimf2 As Double
  Dim vultimf3 As Double
  Dim vcont As Double
  
  If formulaacomparar <> "" And formulaacomparar2 <> "" And vkgllaunaf2 > 0 And vkgllaunaf3 > 0 Then
      reixaformulacio.Enabled = False
      
      If vkg = 0 Then Exit Sub
      If (vkgllaunaf3 + vkgllaunaf2) < vkg Then vkg = (vkgllaunaf3 + vkgllaunaf2)
      vkgmaxf2 = IIf(vkgllaunaf2 < vkg, vkgllaunaf2, vkg)
      vkgmaxf3 = IIf(vkgllaunaf3 < vkg, vkgllaunaf3, vkg)
       kgxrecuperar(0) = Redondejar(vkgmaxf2, 1)
       kgxrecuperar(1) = Redondejar(vkg - vkgmaxf2, 1)
       Command70_Click
       vultimvalor = etiquetatotalkg.WhatsThisHelpID
       vultimf2 = kgxrecuperar(0)
       vultimf3 = kgxrecuperar(1)
       While kgxrecuperar(0) > 0 And kgxrecuperar(1) <= vkgmaxf3
          kgxrecuperar(0) = cadbl(kgxrecuperar(0)) - 0.5
          kgxrecuperar(1) = cadbl(kgxrecuperar(1)) + 0.5
          Command70_Click
          If etiquetatotalkg.WhatsThisHelpID < vultimvalor Then
             vultimvalor = etiquetatotalkg.WhatsThisHelpID
             vultimf2 = kgxrecuperar(0)
             vultimf3 = kgxrecuperar(1)
          End If
       Wend
       reixaformulacio.Enabled = True
       kgxrecuperar(0) = vultimf2
       kgxrecuperar(1) = vultimf3
       Command70_Click
        Else: MsgBox "Per utilitzar aquesta funció hi ha d'haver formula 2 i formula 3 i han de tenir llaunes amb KG.", vbCritical, "Error"
  End If
  
End Sub
Sub carregarrelacióformules()
  Dim rst As Recordset
  Dim rst2 As Recordset
 ' Unload formseleccio
 ' Load formseleccio
 ' formseleccio.caption = "Escullir formula semblant"
  'formseleccio.Data1.DatabaseName = camitintes
  Set rst = dbtintes.OpenRecordset("SELECT * from tintes_semblants where coditintarelacio='" + atrim(llista(0).tag) + "' order by nomdelatinta")
  formulaacomparar.Clear
  formulaacomparar2.Clear
  If rst.EOF Then Exit Sub
  rst.MoveLast
  rst.MoveFirst
  If Not rst.EOF Then
    ReDim vcodiformules2(0)
    ReDim vcodiformules(0)
    ReDim vcodiformules2(rst.RecordCount)
    ReDim vcodiformules(rst.RecordCount)
  End If
  While Not rst.EOF
     Set rst2 = dbtintes.OpenRecordset("SELECT tintes.codi, tintesformules.numformula, tintesformules.predeterminada FROM tintesformules LEFT JOIN tintes ON tintesformules.idtinta = tintes.idtinta where codi='" + atrim(rst!coditinta) + "'")
     formulaacomparar.AddItem atrim(rst!nomdelatinta)
     formulaacomparar.ItemData(formulaacomparar.NewIndex) = rst!coditinta
     vcodiformules(formulaacomparar.NewIndex) = atrim(rst2!numformula)
     'formulaacomparar2.AddItem atrim(rst!nomdelatinta)
     'formulaacomparar2.ItemData(formulaacomparar.NewIndex) = rst!coditinta
     Set rst2 = dbtintes.OpenRecordset("SELECT tintes.codi, tintesformules.numformula, tintesformules.predeterminada FROM tintesformules LEFT JOIN tintes ON tintesformules.idtinta = tintes.idtinta where codi='" + atrim(rst!coditinta_F2) + "'")
     formulaacomparar2.AddItem atrim(rst!nomdelatinta_F2)
     formulaacomparar2.ItemData(formulaacomparar.NewIndex) = cadbl(rst!coditinta_F2)
     vcodiformules2(formulaacomparar2.NewIndex) = atrim(rst2!numformula)
     rst.MoveNext
  Wend
  If formulaacomparar.ListCount > 0 Then
    botorelacioguardada(2).tag = "semblants"
    formulaacomparar.ListIndex = 0: formulaacomparar.Text = formulaacomparar.List(formulaacomparar.ListIndex)
    formulaacomparar.SetFocus
  End If
  Set rst2 = Nothing
  Set rst = Nothing
End Sub
Sub obrirrelacióformules()
  Dim rst As Recordset
  Unload formseleccio
  Load formseleccio
  formseleccio.caption = "Escullir formula semblant"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "SELECT * from tintes_semblants where coditintarelacio='" + atrim(llista(0).tag) + "' order by nomdelatinta"
  formseleccio.sortirs.tag = "filtre"
  formseleccio.bborrar.visible = True
  formseleccio.refrescar
  formseleccio.caption = "Escullir formula semblant"
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(6).visible = False
  formseleccio.DBGrid2.Columns(1).width = 800
  formseleccio.DBGrid2.Columns(2).width = 4000
  formseleccio.DBGrid2.Columns(3).width = 800
  formseleccio.DBGrid2.Columns(4).width = 4000
  formseleccio.width = 15000
  formseleccio.Show 1
  If seleccioret = 1 Then
   If formseleccio.Data1.Recordset.EOF Then GoTo fi
   carregar_reixaformulacio
   Set rst = dbtintes.OpenRecordset("SELECT tintes.codi, tintesformules.numformula, tintesformules.predeterminada FROM tintesformules LEFT JOIN tintes ON tintesformules.idtinta = tintes.idtinta where codi='" + atrim(formseleccio.Data1.Recordset!coditinta) + "'")
   If rst.EOF Then GoTo fi
   formulaacomparar.tag = atrim(rst!numformula)
   Set rst = dbtintes.OpenRecordset("SELECT tintes.codi, tintesformules.numformula, tintesformules.predeterminada FROM tintesformules LEFT JOIN tintes ON tintesformules.idtinta = tintes.idtinta where codi='" + atrim(formseleccio.Data1.Recordset!coditinta_F2) + "'")
   If Not rst.EOF Then formulaacomparar2.tag = atrim(rst!numformula)
   If formulaacomparar.tag <> "" Then
        Label22(12) = atrim(formseleccio.Data1.Recordset!observacions)
        Label22(12).visible = True
        formulaacomparar.Text = formseleccio.DBGrid2.Columns(2)
        formulaacomparar2.Text = formseleccio.DBGrid2.Columns(4)
        carregar_componentssemblants_formula2 2
        carregar_componentssemblants_formula2 3
        pestanyesforumes.Tab = 1
        pestanyesforumes.Tab = 2
        recalcularformulacio
   End If
  End If
  If seleccioret = 8 Then
     If formseleccio.Data1.Recordset.EOF Then GoTo fi
     If MsgBox("Segur que vols borrar la tinta semblant " + atrim(formseleccio.DBGrid2.Columns(2)) + "?", vbExclamation + vbDefaultButton2 + vbYesNo, "Eliminar relació tinta semblant.") = vbYes Then
        formseleccio.Data1.Recordset.Delete
        MsgBox "Relació de tinta semblant esborrada.", vbInformation, "Borrar relació"
     End If
  End If
fi:

End Sub
Sub guardarrelacióformules()
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim rst3 As Recordset
   Dim vvalues As String
   Dim vdescF2 As String
   Dim vobservacio As String
   Set rst = dbtintes.OpenRecordset("SELECT * from tintes_semblants where coditintarelacio='" + atrim(llista(0).tag) + "' and coditinta='" + atrim(llista(1).tag) + "' and coditinta_F2='" + atrim(llista(2).tag) + "' order by nomdelatinta")
   Set rst2 = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(llista(1).tag) + "'")
   Set rst3 = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(llista(2).tag) + "'")
   If Not rst2.EOF Then
      vobservacio = InputBox("Escriu una observació per aquesta relació.", "Observació", Label22(12))
      If vobservacio = "" And Label22(12) <> "" Then If MsgBox("Vols borrar la observació anterior d'aquesta relació?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then vobservacio = Label22(12)
      vobservacio = Mid(treure_apostruf(vobservacio), 1, 80)
      If rst3.EOF Then
           vdescF2 = ""
        Else: vdescF2 = treure_apostruf(rst3!descripcio)
      End If
      vvalues = "(" + atrim(cadbl(llista(1).tag)) + ",'" + treure_apostruf(rst2!descripcio) + "'," + atrim(cadbl(llista(2).tag)) + ",'" + vdescF2 + "'," + atrim(cadbl(llista(0).tag)) + ",'" + vobservacio + "')"
      If rst.EOF Then
         dbtintes.Execute "insert into tintes_semblants (coditinta,nomdelatinta,coditinta_F2,nomdelatinta_F2,coditintarelacio,observacions) values " + vvalues
           Else:
             dbtintes.Execute "update tintes_semblants set observacions='" + vobservacio + "' where id=" + atrim(rst!id)
             Label22(12) = vobservacio
      End If
   End If
   mirar_semblants_formulessemblants cadbl(llista(0).tag)
End Sub

Private Sub Check2_Click()

End Sub

Private Sub buscador_estoc_Change(Index As Integer)
  filtrar_estocs
End Sub

Private Sub buscador_estoc_DblClick(Index As Integer)
   If Index = 1 Then
       MsgBox buscador_estoc(1), vbInformation, "Observacions tintes"
   End If
End Sub

Private Sub Check1_Click(Index As Integer)
  If Index = 2 Then
      If Check1(2).Value = 1 Then
        carregar_liniadelareixaseleccionada
      End If
  End If
  If Index = 0 Then
    If Check1(0).tag <> "1" Then Exit Sub
    If Screen.ActiveControl.Name = "Check1" Then
        If Check1(0).Value = 1 Then
            cdataokcarta(0) = Now
                Else: cdataokcarta(0) = "00:00:00"
        End If
    End If
  End If
  If Index = 1 Then
    If Check1(1).tag <> "1" Then Exit Sub
    If Screen.ActiveControl.Name = "Check1" Then
        If Check1(1).Value = 1 Then
            cdataokcarta(1) = Now
                Else: cdataokcarta(1) = "00:00:00"
        End If
    End If
  End If
  If Index = 4 Then
      filtrarformules
  End If
End Sub

Private Sub cobsoperari_LostFocus()
  Dim vnumc As Double
  Dim vnumtreball As Double
  'vnumc = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("Comanda")))
  'vnumtreball = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("NºTreball")))
  vnumc = cobsoperari.WhatsThisHelpID
  vnumtreball = cadbl(cobsoperari.tag)
  If vnumc > 0 Then
       dbtintes.Execute "insert into comandesrevisadesatintes (comanda,numtreball,estatgestio) values (" + atrim(vnumc) + "," + atrim(cadbl(vnumtreball)) + ",'N')"
       dbtintes.Execute "update comandesrevisadesatintes set observacions='" + treure_apostruf(cobsoperari) + "' where comanda=" + atrim(vnumc)
  End If
End Sub

Sub triar_tintes_compres()
  Dim rstmat As Recordset
  Dim vdesc As String
  Dim vq As Double
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
  formseleccio.Data1.RecordSource = "SELECT tintes.codi, tintes.descripcio, tintes.referenciacolor, tintesreferencies.referencia, tipusbidons.nombido, tintesreferencies.nomproveidor,tipusbidons.litrescompres,tintesreferencies.id FROM (tintesreferencies INNER JOIN tintes ON tintesreferencies.idtinta = tintes.idtinta) INNER JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id where tintesreferencies.predeterminada=true AND  tintesreferencies.nomproveidor<>'INPLACSA'"
  formseleccio.refrescar
  formseleccio.width = 13200
  formseleccio.DBGrid2.Columns(0).width = 800
  formseleccio.DBGrid2.Columns(1).width = 4000
  formseleccio.DBGrid2.Columns(2).width = 2000
  formseleccio.DBGrid2.Columns(3).width = 1800
  formseleccio.DBGrid2.Columns(4).width = 1200
  formseleccio.DBGrid2.Columns(5).width = 1200
  formseleccio.DBGrid2.Columns(6).width = 0
  formseleccio.DBGrid2.Columns(7).width = 0
  'formseleccio.DBGrid2.Columns(8).width = 0
  formseleccio.Show 1
  If seleccioret = 1 Then
      vq = cadbl(InputBox("Quantes vols comprar?", "Quantitat", 1))
      If vq = 0 Then Exit Sub
      vdataexecucio = InputBox("SI VOLS ENVIAR LA COMPRA UN ALTRA DIA QUE NO SIGUI INMEDIATAMENT ESCRIU LA DATA." + vbNewLine + "Ex: " + Format(Now, "25/mm/yy"), "ENVIAMENT DIFERIT DE LA COMPRA")
      vdesc = treure_apostruf(Mid(atrim(formseleccio.Data1.Recordset!descripcio), 1, 40))
      If Not IsDate(vdataexecucio) Then
              dbtintes.Execute "insert into comprespendents (descripcio,referencia,coditinta,quantitat) values ('" + vdesc + "','" + atrim(formseleccio.Data1.Recordset!referencia) + "','" + atrim(cadbl(formseleccio.Data1.Recordset!codi)) + "'," + atrim(vq) + ")"
              enviarlescompresaldepartamentdecompres
           Else: dbtintes.Execute "insert into compresdiferides (dataexecuciocompra,descripcio,referencia,coditinta,quantitat) values (#" + Format(vdataexecucio, "mm/dd/yy") + "#,'" + treure_apostrof(vdesc) + "','" + atrim(formseleccio.Data1.Recordset!referencia) + "','" + atrim(cadbl(formseleccio.Data1.Recordset!codi)) + "'," + atrim(vq) + ")"
      End If
  End If
  Unload formseleccio
End Sub
Sub enviarlescompresaldepartamentdecompres(Optional vmodificat As String)
  Dim vmsg As String
  Dim rst As Recordset
  Dim vassumpte As String
  etenviantemail.visible = True: DoEvents
  Set rst = dbtintes.OpenRecordset("select * from comprespendents where demanat=false")
  If Not rst.EOF Then
     vmsg = "Quant. CodiT-Ref:         Descripció de la compra                         Observació" + Chr(13) + Chr(10)
     vmsg = vmsg + "======================================================================================" + Chr(13) + Chr(10)
  End If
  While Not rst.EOF
     vmsg = vmsg + Chr(13) + Chr(10) + justificar("Q: " + atrim(cadbl(rst!quantitat)), 7, "E") + "(" + justificar(atrim(rst!coditinta), 5, "E") + ") " + justificar(atrim(rst!referencia), 15, "E") + " " + justificar(atrim(rst!descripcio), 40, "E") + " " + atrim(rst!observacio)
     rst.MoveNext
  Wend
  vassumpte = "Compres de tintes desde la Sala de Tintes."
  If vmodificat <> "" Then
     vassumpte = "[ACTUALITZACIÓ] Compres de tintes desde la Sala de Tintes."
     vmsg = vmodificat + Chr(13) + Chr(10) + Chr(13) + Chr(10) + vmsg
  End If
  If vmsg <> "" Then
      enviaremailgeneric "comprestintesAdptmcompres", vassumpte, vmsg
      'enviaremailgeneric "miquel.inplacsa@gmail.com", vassumpte, vmsg
  End If
  wait 2
  etenviantemail.visible = False
End Sub
Sub actualitzar_llista_compres()
  'Dim rstcompres As Recordset
  Dim rstlinies As Recordset
  Dim rstbido As Recordset
  Dim vreferencia As String
  Dim vcont As Byte
  llistacompres.Clear
  ' Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
  Set rst = dbtintes.OpenRecordset("select * from comprespendents where not demanat")
  llistacompres.AddItem "  ---  Pendent de demanar  -----"
  While Not rst.EOF
        llistacompres.AddItem Format(rst!Datademanat, "dd/mm") + " " + justificar("Q: " + atrim(cadbl(rst!quantitat)), 7, "E") + "(" + justificar(atrim(rst!coditinta), 5, "E") + ")" + justificar(atrim(rst!referencia), 20, "E") + justificar(atrim(rst!descripcio), 40, "E") + " " + atrim(rst!observacio)
        llistacompres.ItemData(llistacompres.NewIndex) = rst!id
        rst.MoveNext
  Wend
  llistacompres_programades
  llistacompres.AddItem "  "
  Set rst = dbtintes.OpenRecordset("select * from comprespendents where demanat and data>#" + Format(DateAdd("m", -1, Now), "mm/dd/yy") + "# order by data desc")
  llistacompres.AddItem "  ---  Demanat  -----"
  
  While Not rst.EOF
        llistacompres.AddItem Format(rst!Data, "dd/mm") + " " + justificar("Q: " + atrim(cadbl(rst!quantitat)), 7, "E") + "(" + justificar(atrim(rst!coditinta), 5, "E") + ")" + justificar(atrim(rst!referencia), 15, "E") + justificar(atrim(rst!descripcio), 50, "E")
        llistacompres.ItemData(llistacompres.NewIndex) = rst!id
        rst.MoveNext
  Wend
  afegir_compres_alallista
  
  Set rst = Nothing
End Sub


Sub llistacompres_programades()
   Dim rst As Recordset
   Dim vlinia As String
   
   Set rst = dbtintes.OpenRecordset("select * from compresdiferides")
   If rst.EOF Then Exit Sub
   llistacompres.AddItem " "
   llistacompres.AddItem "   --- Comandes Pendents Programades ---"
   With rst
   While Not .EOF
      vlinia = "Data Exec: " + justificar(Format(!Dataexecuciocompra, "dd/mm/yy"), 10, "D") + "  " + justificar("Q: " + atrim(cadbl(!quantitat)), 7, "E") + "   " + justificar(!coditinta, 6, "E") + " " + justificar(atrim(!descripcio), 20, "E")
      llistacompres.AddItem vlinia
      llistacompres.ItemData(llistacompres.NewIndex) = rst!id
      .MoveNext
   Wend
   End With
   Set rst = Nothing
   
End Sub

Sub afegir_compres_alallista()
   Dim rstcompres As Recordset
   Dim i As Byte
   Dim vlinia As String
   Dim vsql As String
   Dim rst As Recordset
   Dim ventrataalbaransbip As Boolean
   
   Set rst = dbcompres.OpenRecordset("select * from albaransbip")
   'Set rstcompres = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, capcalera.data, capcalera.dataentrega, capcalera.nomprovcomercial, liniescompra.codimaterial, liniescompra.nommaterial, liniescompra.quantitatkg, liniescompra.kgentregats FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((capcalera.numcomanda)>0) AND ((liniescompra.totentregat)=False) AND ((liniescompra.tipusmaterialcomprat)='T')) order by dataentrega;")
     'hi havia limit de 15 dies a la consulta i he pujat a 20 dies no sé perque estava a 15
   Set rstcompres = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, capcalera.data, capcalera.dataentrega, capcalera.nomprovcomercial, liniescompra.codimaterial, liniescompra.nommaterial, liniescompra.quantitatkg, liniescompra.kgentregats,liniescompra.idliniacompra,liniescompra.totentregat FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((capcalera.numcomanda)>0) AND ((liniescompra.tipusmaterialcomprat)='T')) and (data>dateadd('d',-100,now))  order by dataentrega;")  'and not totentregat
   llistacompres.AddItem " "
   llistacompres.AddItem "   --- Comandes fetes al proveïdor ---"
   
   llistacompres.AddItem "Cmnda  DataCom/Prev Proveidor CodiT  Descripció                                  Kg     Pndts"
   llistacompres.AddItem "================================================================================================"
   With rstcompres
   While Not .EOF
      rst.FindFirst "idliniacompra=" + atrim(rstcompres!idliniacompra)
      If Not rst.NoMatch Then
          ventrataalbaransbip = True
          If rst!albaraescanejat = True And rstcompres!totentregat Then GoTo proxima
      End If
      If InStr(1, atrim(.Fields(3)), "MORCHEM") > 0 Then GoTo proxima
      If InStr(1, atrim(.Fields(3)), "HENKEL AG") > 0 Then GoTo proxima
      vlinia = justificar(.Fields(0), 6, "E") + justificar(Format(.Fields(1), "dd/mm"), 6, "D") + justificar(Format(.Fields(2), "dd/mm"), 6, "D") + "   " + justificar(.Fields(3), 6, "E") + " " + IIf(ventrataalbaransbip And cadbl(!kgentregats) = 0, "[", "") + IIf(Not rst!albaraescanejat And rstcompres!totentregat, "(NoEsc)", justificar(.Fields(4), 7, "E")) + justificar(.Fields(5), 40, "E") + IIf(ventrataalbaransbip And cadbl(!kgentregats) = 0, "]", "") + justificar(.Fields(6), 8, "D") + justificar(.quantitatkg - cadbl(!kgentregats), 8, "D")
      llistacompres.AddItem vlinia
      
proxima:
      .MoveNext
      ventrataalbaransbip = False
   Wend
   End With
   Set rstcompres = Nothing
   
End Sub

Private Sub Combo_Click(Index As Integer)
    If Index = 0 Then carregar_cuatricomia_combo "N"
    If Index = 1 Then carregar_cuatricomia_combo "G"
    If Index = 2 Then carregar_cuatricomia_combo "M"
    If Index = 3 Then carregar_cuatricomia_combo "C"
End Sub
Sub carregar_cuatricomia_combo(vcolor As String)
  Dim rst As Recordset
  Dim vanilox As String
  Dim vnumtreballiversio As String
  Dim vnummaquina As Double
  vnumtreballiversio = Frame5(1).tag
  If vcolor = "N" Then vanilox = Combo(0)
  If vcolor = "G" Then vanilox = Combo(1)
  If vcolor = "M" Then vanilox = Combo(2)
  If vcolor = "C" Then vanilox = Combo(3)
  vnummaquina = maquinaescullidaPerCuatricomia
  Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where color='" + atrim(vcolor) + "' and aniloxivolum='" + atrim(vanilox) + "' and numtreballiversio='" + vnumtreballiversio + "' and nummaq=" + atrim(vnummaquina))
  If Not rst.EOF Then
    possar_color_cuatricomia rst
  End If
  Set rst = Nothing
End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index < 4 Then KeyCode = 0
End Sub

Private Sub Combo_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index < 4 Then KeyAscii = 0
End Sub

Private Sub Command12_Click()
  If MsgBox("Segur que vols RE-ENVIAR el correu de compres pendents al DEPARTAMENT DE COMPRES?" + Chr(10) + "Aquesta operació també es fa automàticament al afegir quelcom a la llista.", vbYesNo + vbDefaultButton2 + vbInformation, "Atenció") = vbNo Then Exit Sub
  enviarlescompresaldepartamentdecompres
End Sub

Private Sub Command66_Click()
  
End Sub
Sub escullir_formula_semblant_filtre(Optional vformula As String, Optional vnomformula As String)
  Label22(8) = "": vsitprimerallauna = ""
  Label22(9) = "": vsitprimerallaunaf2 = "": vkgllaunaf2 = 0
  Label22(1) = "": vsitprimerallaunaf3 = "": vkgllaunaf3 = 0
  llista(1).Clear: llista(2).Clear

  kgformula = ""
  Label22(12) = ""
  Label22(12).visible = False
  If vformula = "" Then
     vformula = escullir_formula
     If vformula <> "" Then vnomformula = formseleccio.DBGrid2.Columns(2)
  End If
  If vformula <> "" Then
    formulasemblanta.tag = vformula
    formulasemblanta.Text = vnomformula
    formulaacomparar.Clear
    formulaacomparar2.Clear
    kgxrecuperar(0) = "0"
    kgxrecuperar(1) = "0"
  End If
  carregar_components_semblants
  mirar_semblants_formulessemblants cadbl(llista(0).tag)
  Unload formseleccio
End Sub
Sub mirar_semblants_formulessemblants(vcoditinta As Double)
  Dim vcodi As String
  Dim rstc As Recordset
  Dim vidtinter As String
  botorelacioguardada(1).BackColor = Command14.BackColor
  botorelacioguardada(2).BackColor = Command14.BackColor
  Set rstc = dbtintes.OpenRecordset("select * from tintes_semblants where coditintarelacio='" + atrim(vcoditinta) + "'")
  If Not rstc.EOF Then
        Command45.tag = cadbl(vcoditinta)
        botorelacioguardada(1).BackColor = QBColor(12)
        botorelacioguardada(2).BackColor = QBColor(12)
  End If
  Set rstc = Nothing
End Sub
Sub carregar_components_semblants()
    Dim vsql As String
    Dim vformula As String
    Dim rst As Recordset
    Dim r As Byte
    reixaformulacio.ColAlignment(0) = 1
    vformula = formulasemblanta.tag
    vsql = "SELECT Formules.codiformula, Componentsbase.nomcomponent, DetallFormules.[%decomponent] as tanx100 FROM (Formules INNER JOIN DetallFormules ON Formules.idformula = DetallFormules.IDFormula) INNER JOIN Componentsbase ON DetallFormules.IdComponente = Componentsbase.idcomponent"
    vsql = vsql + " WHERE (((Formules.codiformula)='" + vformula + "')) order by esbase DESC  ;"
   ' Clipboard.Clear
   ' Clipboard.SetText vsql
    Set rst = dbtintes.OpenRecordset(vsql)
    reixaformulacio.Rows = 1
    r = 1
    While Not rst.EOF
       reixaformulacio.Rows = r + 1
       reixaformulacio.TextMatrix(r, 0) = atrim(rst!nomcomponent)
       reixaformulacio.TextMatrix(r, 1) = rst!tanx100
       r = r + 1
       rst.MoveNext
    Wend
    
    Set rst = Nothing
    reixaformulacio.tag = reixaformulacio.Rows
    
    'buscar el codi de tinta i carregar les llaunes
     llista(0).tag = ""
    Set rst = dbtintes.OpenRecordset("SELECT tintesformules.numformula, tintes.codi, tintes.descripcio FROM tintesformules LEFT JOIN tintes ON tintesformules.idtinta = tintes.idtinta WHERE (((tintesformules.numformula)='" + vformula + "'));")
    If Not rst.EOF Then
        llista(0).tag = atrim(rst!codi)
    End If
    carregar_llaunes_semblants llista(0), vformula
    
End Sub
Sub carregar_llaunes_semblants(llista As ListBox, vformula As String)
    Dim vcodi As String
    Dim rst As Recordset
    Dim rstll As Recordset
    vcodi = llista.tag
    llista.Clear
    If llista.Index = 0 Then Label22(8) = "": vsitprimerallauna = ""
    If llista.Index = 1 Then Label22(9) = "": vsitprimerallaunaf2 = "": vkgllaunaf2 = 0
    If llista.Index = 2 Then Label22(1) = "": vsitprimerallaunaf3 = "": vkgllaunaf3 = 0
    Set rst = dbtintes.OpenRecordset("select * from tintes_tot where tintes_tot.idtinta in (select idtinta from tintesformules where numformula='" + atrim(vformula) + "')")
    'poso l'etiqueta de la tinta sobre de la llista de llaunes disponibles
    If Not rst.EOF Then
       If llista.Index = 0 Then Label22(8) = atrim(rst!codi) + " - " + atrim(rst!descripcio)
       If llista.Index = 1 Then Label22(9) = atrim(rst!codi) + " - " + atrim(rst!descripcio)
       If llista.Index = 2 Then Label22(1) = atrim(rst!codi) + " - " + atrim(rst!descripcio)
    End If
    'empleno les caixes amb les llaunes disponibles
    While Not rst.EOF
    ' Set rstll = dbtintes.OpenRecordset("select * from llaunes where activa=true and idtinta=" + atrim(cadbl(rst!idtinta)))
     
     vsql = "SELECT Llaunes.id, Llaunes.numllauna, Llaunes.preuxrkilo, Llaunes.capacitatactual AS kgactuals, comandesrevisadesatintes.estatgestio, [Llaunes].[situacio]+IIf([llaunes].[aimpresores],'*'+Trim(' ' & [comandesrevisadesatintes].[estatgestio]),'') AS situacioiimp, tintesreferencies.referencia, tipusbidons.capacitat, Contenidors_material.descripcio AS nomcontenidor, Llaunes.activa FROM ((assignaciollaunesacomandes RIGHT JOIN ((Llaunes LEFT JOIN tintesreferencies ON Llaunes.id_refproveidor = tintesreferencies.id) LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id) ON assignaciollaunesacomandes.numllauna = Llaunes.numllauna) LEFT JOIN comandesrevisadesatintes ON assignaciollaunesacomandes.comanda = comandesrevisadesatintes.comanda) LEFT JOIN Contenidors_material ON Llaunes.idmaterialcontenidor = Contenidors_material.codi "
     Set rstll = dbtintes.OpenRecordset(vsql + " Where (((Llaunes.idtinta) = " + atrim(cadbl(rst!idtinta)) + ") and activa=true )ORDER BY Llaunes.activa;")
     
     While Not rstll.EOF
       llista.AddItem rstll!numllauna + " --> " + justificar(atrim(rstll!situacioiimp), 6, "D") + " Sit  " + atrim(rstll!kgactuals) + "Kg"
       If llista.Index = 1 Then If vkgllaunaf2 = 0 Then vkgllaunaf2 = cadbl(rstll!kgactuals)
       If llista.Index = 2 Then If vkgllaunaf3 = 0 Then vkgllaunaf3 = cadbl(rstll!kgactuals)
       If llista.Index = 0 And vsitprimerallauna = "" Then vsitprimerallauna = atrim(rstll!situacioiimp)
       If llista.Index = 1 And vsitprimerallaunaf2 = "" Then vsitprimerallaunaf2 = atrim(rstll!situacioiimp)
       If llista.Index = 2 And vsitprimerallaunaf3 = "" Then vsitprimerallaunaf3 = atrim(rstll!situacioiimp)
       rstll.MoveNext
     Wend
     rst.MoveNext
    Wend
    Set rstll = Nothing
    Set rst = Nothing
End Sub
Function escullir_formula_semblant() As String
  Unload formseleccio
  Load formseleccio
  formseleccio.caption = "Escullir formula semblant"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = dataformules.RecordSource
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.caption = "Escullir formula"
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 2000
  formseleccio.DBGrid2.Columns(2).width = 5000
  formseleccio.DBGrid2.Columns(3).width = 2000
  formseleccio.width = 9500
  formseleccio.Show 1
  If seleccioret = 1 Then
   If formseleccio.Data1.Recordset.EOF Then GoTo fi
   escullir_formula_semblant = atrim(formseleccio.Data1.Recordset!codiformula)
  End If
fi:

End Function

Function escullir_formula() As String
  Unload formseleccio
  Load formseleccio
  formseleccio.caption = "Escullir formula"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "SELECT idformula,codiformula,descripcioformula from formules order by descripcioformula"
  formseleccio.sortirs.tag = "filtre"
  formseleccio.cmissatge.tag = "2"
  formseleccio.refrescar
  formseleccio.caption = "Escullir formula"
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 3000
  formseleccio.DBGrid2.Columns(2).width = 5000
  formseleccio.width = 9500
  formseleccio.Show 1
  If seleccioret = 1 Then
   If formseleccio.Data1.Recordset.EOF Then GoTo fi
   escullir_formula = atrim(formseleccio.Data1.Recordset!codiformula)
  End If
fi:

End Function

Sub intercanviarvalorsentreformules()
    Dim vcombotemp(1000) As String
    Dim vcombotempDATA(1000) As Double
    Dim vvalortag As String
    Dim vkg As Double
    Dim vitems As Double
    Dim vdesc As String
    Dim vIndex As Long
    
    vkg = cadbl(kgxrecuperar(0))
    vvalortag = formulaacomparar.tag
    vitems = formulaacomparar.ListCount
    vdesc = formulaacomparar
    vIndex = formulaacomparar.ListIndex
    For i = 0 To formulaacomparar.ListCount - 1
       vcombotemp(i) = formulaacomparar.List(i)
       vcombotempDATA(i) = formulaacomparar.ItemData(i)
    Next i
    
    formulaacomparar.Clear
    formulaacomparar.tag = formulaacomparar2.tag
    formulaacomparar = formulaacomparar2
    For i = 0 To formulaacomparar2.ListCount - 1
       formulaacomparar.AddItem formulaacomparar2.List(i)
       formulaacomparar.ItemData(i) = formulaacomparar2.ItemData(i)
    Next i
    formulaacomparar.ListIndex = formulaacomparar2.ListIndex
    
    formulaacomparar2.Clear
    formulaacomparar2.tag = vvalortag
    formulaacomparar2 = vdesc
    For i = 0 To vitems - 1
       formulaacomparar2.AddItem vcombotemp(i)
       formulaacomparar2.ItemData(i) = vcombotempDATA(i)
    Next i
    formulaacomparar2.ListIndex = vIndex
    
    
    kgxrecuperar(0) = kgxrecuperar(1)
    kgxrecuperar(1) = atrim(vkg)
    carregar_componentssemblants_formula2 2
    carregar_componentssemblants_formula2 3
    If vIndex < formulaacomparar.ListCount Then
         formulaacomparar.ListIndex = vIndex
         formulaacomparar.SetFocus
    End If
    recalcularformulacio
End Sub

Sub verificarlotsInmaker()
   formverificarlotsinkmaker.Show 1
End Sub
Sub verificaciotintestreballsnousomodificats()
  Static vdins As Boolean
  If vdins Then MsgBox "Hi ha una altra finestre de verificació de tintes oberta.", vbCritical, "Error": Exit Sub
inici:
  vdins = True
  Load formseleccio
  formseleccio.caption = "Treballs per revisar"
  formseleccio.Data1.DatabaseName = rutadelfitxer(camitintes) + "CLIXESNOUS.MDB"
  formseleccio.Data1.RecordSource = "SELECT Modificacions.id_treball, Modificacions.ordre, [marca] & ' - ' & [linia] AS Marcailinia, Modificacions.estatrevisiotintes FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball WHERE (((InStr(1,[estatrevisiotintes],'DISSENY'))>0) AND ((InStr(1,[estatrevisiotintes],'+TIN'))=0));"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 1000
  formseleccio.DBGrid2.Columns(1).width = 500
  formseleccio.DBGrid2.Columns(2).width = 5500
  formseleccio.DBGrid2.Columns(3).width = 3500
  formseleccio.width = 12000
  formseleccio.Show 1
  If seleccioret = 1 Then
   ShellAndWait "\\serverprodu\dades\progcomandes\aplicacio\clixesnous.exe " + "comandes.ini ''  modificartintes " + atrim(formseleccio.Data1.Recordset!id_treball) + " " + atrim(formseleccio.Data1.Recordset!ordre) + " +TIN", vbNormalFocus
'   wait 2
'   AppActivate "Manteniment de les Tintes    " + atrim(formseleccio.Data1.Recordset!id_treball) + "/" + atrim(formseleccio.Data1.Recordset!ordre)
   GoTo inici
  End If
  
  Unload formseleccio
  vdins = False
End Sub
Sub posar_ordre_impresio(vmaq As Double)
  Dim rst As Recordset
  Dim v As String
  Set rst = dbbaixes.OpenRecordset("select * from impresores_ordreimpresio where maquina=" + atrim(vmaq) + " order by ordre,dataprogramada")
  While Not rst.EOF
    v = v + IIf(v <> "", ",", "") + atrim(rst!comanda)
    rst.MoveNext
  Wend
  If v <> "" Then Command56_Click: wait 1: filtre(0) = v: filtre_LostFocus 0
  Set rst = Nothing
End Sub
Sub botofiltreprogramable(vIndex As Integer, vbutton As Integer)
  If vbutton = 1 Then  'botó esquerra
      carregar_filtre vIndex - 14
  End If
  If vbutton = 2 Then  'botó dret
      If MsgBox("Vols programar aquest botó amb el filtre que hi ha ara personalitzat?", vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
          guardar_filtre vIndex - 14
      End If
  End If
End Sub
Sub carregar_filtre(vIndex As Integer)
  Dim i As Integer
  Dim v As String
  Command56_Click  'borro el filtre que hi ha ara
  For i = 0 To filtre.Count - 1
    v = atrim(llegir_ini("Filtres_" + atrim(vIndex), "Filtre" + atrim(i), "comandes.ini"))
    If v = "{[}]" Then v = ""
    If v <> "" Then filtre(i).Text = v
  Next i
  filtre(0).SetFocus
  filtre_LostFocus (0)
End Sub
Sub guardar_filtre(vIndex As Integer)
  Dim i As Integer
  For i = 0 To filtre.Count - 1
   If filtre(i).Text <> reixacomandes.TextMatrix(0, i) And atrim(filtre(i).Text) <> "" Then
      escriure_ini "Filtres_" + atrim(vIndex), "Filtre" + atrim(i), filtre(i), "comandes.ini"
       Else: escriure_ini "Filtres_" + atrim(vIndex), "Filtre" + atrim(i), "", "comandes.ini"
   End If
  Next i
  MsgBox "Guardat"
End Sub
Sub guardar_cuatricomia()
  Dim vnumtreballiversio As String
  Dim v As String
'  vnumtreballiversio = framecuatricomia.tag
'  If InStr(2, vnumtreballiversio, "/") = 0 Then Exit Sub
'  dbtintes.Execute "delete * from valorscuatricomia_treball where numtreballiversio='" + vnumtreballiversio + "'"
'  v = "'" + atrim(vnumtreballiversio) + "'," + passaradecimalpunt(atrim(cadbl(kgxrecuperar(3)))) + "," + passaradecimalpunt(cadbl(kgxrecuperar(4))) + "," + passaradecimalpunt(cadbl(kgxrecuperar(5))) + "," + passaradecimalpunt(cadbl(kgxrecuperar(6)))
'  dbtintes.Execute "insert into valorscuatricomia_treball (numtreballiversio,negre,groc,magenta,cyan) values (" + v + ")"
  guardar_tolerancies_cuatricomia
  MsgBox "Valors guardats del treball " + vnumtreballiversio
  framecuatricomia.visible = False
End Sub
Sub guardar_tolerancies_cuatricomia()
  Dim i As Byte
  Dim vini As String
  vini = rutadelfitxer(cami) + "valorsprograma.ini"
  For i = 7 To 18
    escriure_ini "ToleranciesCuatricomia", kgxrecuperar(i).tag, kgxrecuperar(i), vini
  Next i
End Sub
Sub llegir_tolerancies_cuatricomia()
  Dim i As Byte
  Dim v As String
  Dim vini As String
  vini = rutadelfitxer(cami) + "valorsprograma.ini"
  For i = 7 To 18
    v = llegir_ini("ToleranciesCuatricomia", kgxrecuperar(i).tag, vini)
    If v = "{[}]" Then v = 0
    kgxrecuperar(i) = v
  Next i
End Sub
Function triaranilox() As Double
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select lineatura as liniatura,sum(quantitat) as Quants_tenim from aniloxos group by lineatura order by lineatura"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   'formseleccio.Width = formseleccio.Width + 1000
   'formseleccio.DBGrid2.Width = formseleccio.DBGrid2.Width + 1000
   formseleccio.DBGrid2.Columns(0).width = 1000
   formseleccio.DBGrid2.Columns(1).width = 1500
   'formseleccio.DBGrid2.Columns(2).Width = 1000
   'formseleccio.DBGrid2.Columns(3).Width = 1000
   'formseleccio.DBGrid2.Columns("id_estat").Width = 0
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           triaranilox = formseleccio.DBGrid2.Columns("liniatura")
        End If
   End If
    If seleccioret = 9 Then
        triaranilox = 0
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Function
Function triarvolum(vanilox As Double) As Double
   Load formseleccio
   formseleccio.Data1.DatabaseName = cami
   formseleccio.Data1.RecordSource = "select distinct nommaquina,volum  from aniloxos where lineatura=" + atrim(vanilox) + " order by volum"
   formseleccio.DBGrid2.AllowDelete = False
   formseleccio.refrescar
   formseleccio.DBGrid2.Columns(0).width = 2000
   formseleccio.DBGrid2.Columns(1).width = 500
   formseleccio.Show 1
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           triarvolum = formseleccio.DBGrid2.Columns("volum")
        End If
   End If
    If seleccioret = 9 Then
        triarvolum = 0
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Function
Function maquinaescullidaPerCuatricomia() As Double
  If Command67(30).BackColor <> &H80000005 Then maquinaescullidaPerCuatricomia = 7
  If Command67(31).BackColor <> &H80000005 Then maquinaescullidaPerCuatricomia = 9
  If Command67(30).BackColor <> &H80000005 And Command67(31).BackColor <> &H80000005 Then maquinaescullidaPerCuatricomia = 1
End Function
Sub escullir_anilox_cuatricomia(vcolor As String)
  Dim vanilox As Double
  Dim vaniloxivolum As String
  Dim rst As Recordset
  Dim vnumtreballiversio As String
  If maquinaescullidaPerCuatricomia = 0 Then MsgBox "No has escullit l'impressora que has agafat la mostra.", vbCritical, "Error": Exit Sub
  vnumtreballiversio = Frame5(1).tag
  vanilox = triaranilox
  If vanilox > 0 Then
       vaniloxivolum = atrim(triarvolum(vanilox))
       If cadbl(vaniloxivolum) > 0 Then
           vaniloxivolum = atrim(vanilox) + " v" + atrim(vaniloxivolum)
             Else: vaniloxivolum = ""
       End If
  End If
  If vaniloxivolum <> "" Then
      Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where nummaq=" + atrim(maquinaescullidaPerCuatricomia) + " and color='" + vcolor + "' and aniloxivolum='" + atrim(vaniloxivolum) + "' and numtreballiversio='" + vnumtreballiversio + "'")
      If Not rst.EOF Then MsgBox "Aquest anilox ja està entrat.", vbCritical, "Error": GoTo fi
      rst.AddNew
      rst!aniloxivolum = vaniloxivolum
      rst!color = vcolor
      rst!numtreballiversio = vnumtreballiversio
      rst!delta = 0
      rst!deltaE = 0
      rst!nummaq = maquinaescullidaPerCuatricomia
      rst.Update
      rst.MoveLast
      possar_color_cuatricomia rst
      dbtintes.Execute "delete * from valorscuatricomia_treball where nummaq=1 and aniloxivolum='" + atrim(vaniloxivolum) + "' and numtreballiversio='" + vnumtreballiversio + "'"
  End If
fi:
  Set rst = Nothing
End Sub
Sub afegir_combo_cuatricomia(vcombo As ComboBox, vanilox As String)
    Dim i As Byte
    Dim vtrobat As Boolean
    If vcombo.ListCount > 0 Then
        For i = 0 To vcombo.ListCount - 1
            If vcombo.Text = vanilox Then vtrobat = True
        Next i
    End If
    If Not vtrobat Then vcombo.AddItem vanilox
End Sub
Sub possar_color_cuatricomia(rst As Recordset)
    Dim vcolor As String
    escullir_maq_cuatricomia rst!nummaq, True
    vcolor = rst!color
     If vcolor = "N" Then afegir_combo_cuatricomia Combo(0), rst!aniloxivolum: Combo(0) = rst!aniloxivolum: kgxrecuperar(19) = rst!delta: kgxrecuperar(3) = rst!deltaE
     If vcolor = "G" Then afegir_combo_cuatricomia Combo(1), rst!aniloxivolum: Combo(1) = rst!aniloxivolum: kgxrecuperar(20) = rst!delta: kgxrecuperar(4) = rst!deltaE
     If vcolor = "M" Then afegir_combo_cuatricomia Combo(2), rst!aniloxivolum: Combo(2) = rst!aniloxivolum: kgxrecuperar(21) = rst!delta: kgxrecuperar(5) = rst!deltaE
     If vcolor = "C" Then afegir_combo_cuatricomia Combo(3), rst!aniloxivolum: Combo(3) = rst!aniloxivolum: kgxrecuperar(22) = rst!delta: kgxrecuperar(6) = rst!deltaE
End Sub
Sub exportar_dades_cuatricomia()
  Dim rst As Recordset
  Dim vnumtreballiversio As String
  Dim vlinia As String
  vnumtreballiversio = Frame5(1).tag
  Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball ORDER BY numtreballiversio,color desc,aniloxivolum")
  Open "c:\temp\~LlistatDeltaE.csv" For Output As #1
  Print #1, "   Valors cuatricomia DELTAE "
  vlinia = "Treball/Versió;Color;Anilox/Volum;Delta;DeltaE"
  Print #1, vlinia
  While Not rst.EOF
    vlinia = atrim(rst!numtreballiversio) + ";" + atrim(rst!color) + ";" + atrim(rst!aniloxivolum) + ";" + atrim(rst!delta) + ";" + atrim(rst!deltaE)
    Print #1, vlinia
    rst.MoveNext
  Wend
  Close #1
  If existeix("c:\temp\~LlistatDeltaE.csv") Then obrir_document "c:\temp\~LlistatDeltaE.csv"
  Set rst = Nothing
End Sub
Sub copiar_dades_versio_anterior()
  Dim vnumversio As Double
  Dim vnumtreball As Double
  Dim vnumtreballacopiar As Double
  Dim vnumversioacopiar As Double
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim i As Byte
  
  vnumversio = cadbl(Mid(Frame5(1).tag, InStr(1, Frame5(1).tag, "/") + 1))
  vnumtreball = cadbl(Mid(Frame5(1).tag, 1, InStr(1, Frame5(1).tag, "/") - 1))
  'If vnumversio = 1 Then MsgBox "No hi ha cap versió anterior.", vbCritical, "Error": Exit Sub
  'If MsgBox("Vols copiar dades DELTA de la versió anterior d'aquest treball?", vbYesNo + vbDefaultButton2 + vbExclamation, "Atenció") = vbYes Then
  vnumtreballacopiar = cadbl(InputBox("Entra el numero de treball ON VOLS COPIAR AQUESTES DADES.", "Copiar dades Delta"))
  If cadbl(vnumtreballacopiar) > 0 Then
     vnumversioacopiar = InputBox("Entra el numero de versió ON VOLS copiar.", "Copiar dades Delta")
     If cadbl(vnumversioacopiar) = 0 Then GoTo fi
     Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where numtreballiversio='" + atrim(vnumtreball) + "/" + atrim(vnumversio) + "'")
     If Not rst.EOF Then
         Set rst2 = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where numtreballiversio='" + atrim(vnumtreballacopiar) + "/" + atrim(vnumversioacopiar) + "'")
         If Not rst2.EOF Then
            If MsgBox("Aquesta versió de treball ja té valors Delta entrats." + vbNewLine + "  VOLS SOBRE-ESCRIURE AQUESTS VALORS PELS DE LA VERSIÓ ANTERIOR?", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then GoTo fi
            dbtintes.Execute "delete * from valorscuatricomia_treball where numtreballiversio='" + atrim(vnumtreballacopiar) + "/" + atrim(vnumversioacopiar) + "'"
         End If
           Else: MsgBox "No he trobat dades d'aquest treball i versió.", vbCritical, "Errro"
     End If
     While Not rst.EOF
        rst2.AddNew
        For i = 0 To rst.Fields.Count - 1
             rst2.Fields(i) = rst.Fields(i)
        Next i
        rst2!numtreballiversio = atrim(vnumtreballacopiar) + "/" + atrim(vnumversioacopiar)
        rst2.Update
        rst.MoveNext
     Wend
  End If
fi:
  Set rst = Nothing
  Set rst2 = Nothing
End Sub

Sub tancar_framegrmkilo()
   Frame5(6).visible = False
End Sub

Sub calcul_grms_kilo()
  Dim i As Byte
  Frame5(6).Top = 885
  Frame5(6).Left = 75
  Frame5(6).visible = True
  reixagrmskilo.Clear
  reixagrmskilo.ColWidth(0) = 3000
  reixagrmskilo.ColWidth(1) = 3000
  reixagrmskilo.ColWidth(2) = 3000
  reixagrmskilo.TextMatrix(1, 0) = "Llauna:(Kg)"
  reixagrmskilo.TextMatrix(2, 0) = "Groc:"
  reixagrmskilo.TextMatrix(3, 0) = "Blau:"
  reixagrmskilo.TextMatrix(4, 0) = "Negre:"
  reixagrmskilo.TextMatrix(5, 0) = "Magenta:"
  reixagrmskilo.TextMatrix(6, 0) = "Violeta:"
  reixagrmskilo.TextMatrix(7, 0) = "V.Mig:"
  reixagrmskilo.TextMatrix(8, 0) = "Taronja:"
  reixagrmskilo.TextMatrix(9, 0) = "Rosa:"
  reixagrmskilo.TextMatrix(10, 0) = "Verd:"
  reixagrmskilo.TextMatrix(11, 0) = "Blanc:"
  reixagrmskilo.TextMatrix(0, 1) = "Grams"
  reixagrmskilo.TextMatrix(0, 2) = "-Total-"
  reixagrmskilo.col = 2
  For i = 1 To 11
    reixagrmskilo.Row = i
    reixagrmskilo.CellBackColor = &H8000000F
  Next i
  reixagrmskilo.ColAlignment(1) = 3
  
  
End Sub
Sub escullir_maq_cuatricomia(vIndex As Integer, Optional nocarregar As Boolean)
  If vIndex = 1 Then
        Command67(30).BackColor = QBColor(15)
        Command67(31).BackColor = QBColor(15)
        GoTo cont
  End If
  If Command67(30).BackColor = QBColor(15) And Command67(31).BackColor = QBColor(15) And vIndex >= 30 Then
      preguntar_per_gravarvalors_a_la_maquina vIndex
      vtempsbotoapretat = Now
  End If
  If vIndex = 7 Then vIndex = 30
  If vIndex = 9 Then vIndex = 31
  Command67(30).BackColor = &H80000005
  Command67(31).BackColor = &H80000005
  If vIndex = 30 Then Command67(30).BackColor = &H5C31DD
  If vIndex = 31 Then Command67(31).BackColor = &H5C31DD
cont:
  If Not nocarregar Then carregar_valors_cuatricomia Frame5(1).tag 'frame5(1).tag es el valor de treball i versió que he escullit
End Sub
Sub preguntar_per_gravarvalors_a_la_maquina(vmaquina As Integer)
  Dim rst As Recordset
  If vmaquina = 30 Then vmaquina = 7
  If vmaquina = 31 Then vmaquina = 9
  If MsgBox("Vols passar aquestes dades sense màquina assignada a la Impresora " + atrim(vmaquina) + "?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  dbtintes.Execute "update  valorscuatricomia_treball set nummaq=" + atrim(vmaquina) + " where nummaq=1 and numtreballiversio='" + Frame5(1).tag + "'"
  Set rst = Nothing
End Sub
Private Sub Command67_Click(Index As Integer)
  Dim vformula As String
  Dim vformulajaposada As Boolean
  If Index = 31 Or Index = 30 Then escullir_maq_cuatricomia Index
  If Index = 29 Then tancar_framegrmkilo
  If Index = 28 Then calcul_grms_kilo
  If Index = 27 Then copiar_dades_versio_anterior
  If Index = 26 Then exportar_dades_cuatricomia
  If Index = 22 Then escullir_anilox_cuatricomia "N"
  If Index = 23 Then escullir_anilox_cuatricomia "G"
  If Index = 24 Then escullir_anilox_cuatricomia "M"
  If Index = 25 Then escullir_anilox_cuatricomia "C"
  If Index = 21 Then Frame5(2).visible = Not Frame5(2).visible: Frame5(2).ZOrder 0: Frame5(2).Left = 60: Frame5(2).Top = 210
  If Index = 20 Then guardar_cuatricomia
  If Index = 19 Then Frame5(1).visible = False
  If Index = 12 Then posar_ordre_impresio 7
  If Index = 13 Then posar_ordre_impresio 9
  If Index = 10 Then escullir_formula_semblant_filtre
  If Index = 11 Then Formagrupartreballs.Show
  If Index = 9 Then verificaciotintestreballsnousomodificats
  If Index = 6 Then verificarlotsInmaker
  If Index = 5 Then intercanviarvalorsentreformules
  If Index = 18 Then comprovar_lesllaunesdelsdosificadors
  If Index = 0 Then
   If vfocusultimcontrol = "formulaacomparar" Then formulaacomparar = ""
   If vfocusultimcontrol = "formulaacomparar2" Then formulaacomparar2 = ""
   kgformula = ""
   vformula = escullir_formula
   If vformula <> "" Then
        If formulaacomparar.Text = "" Then
             formulaacomparar.tag = vformula
             formulaacomparar.Text = formseleccio.DBGrid2.Columns(2)
             carregar_componentssemblants_formula2 2
             vformulajaposada = True
        End If
        If formulaacomparar2.Text = "" Then
            If Not vformulajaposada Then
             formulaacomparar2.tag = vformula
             formulaacomparar2.Text = formseleccio.DBGrid2.Columns(2)
            End If
            carregar_componentssemblants_formula2 3
        End If
        pestanyesforumes.Tab = 1
        pestanyesforumes.Tab = 2
        recalcularformulacio
        botorelacioguardada(2).tag = ""
   End If
  End If
  If Index = 1 Then
     If vfocusultimcontrol = "formulaacomparar" Then formulaacomparar = ""
     If vfocusultimcontrol = "formulaacomparar2" Then formulaacomparar2 = ""
     kgformula = ""
     carregarvalorsdepestanyaformulaalcombo
     If formulaacomparar.List(0) = "" Then
        carregarvalorformulaacomparar vcodiformules(formulaacomparar.ListIndex), 2
        'carregarvalorformulaacomparar formulaacomparar.List(0), 2
     End If
     
  End If
  If Index = 2 Then imprimir_ticket_semblants
  If Index = 3 Then actualitzar_llista_compres
  If Index = 4 Then afegir_llista_compres
  If Index = 7 Then guardar_ticket_relaciocomandaactiva
  If Index = 8 Then imprimir_totes_lescombinaciollaunes
End Sub
Function escullir_tinter(vnumc As Double) As Double
Load formseleccio
  formseleccio.caption = "Selecciona Color"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "select idtinter,formula,valorformula as Color from tintes_semblants_relacioambcomandaactiva where comanda=" + atrim(vnumc) + " and formula='l4' order by idtinter"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).visible = False
  formseleccio.DBGrid2.Columns(2).width = 4500
  formseleccio.Show 1
  If seleccioret = 1 Then
   escullir_tinter = atrim(formseleccio.Data1.Recordset!idtinter)
  End If
  Unload formseleccio
     
End Function
Sub imprimir_totes_lescombinaciollaunes()
   Dim rst As Recordset
   Dim vnumc As Double
   Dim vidtinter As Double
   vnumc = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, col))
   If vnumc = 0 Then Exit Sub
   Set rst = dbtintes.OpenRecordset("select distinct idtinter from tintes_semblants_relacioambcomandaactiva where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
      rst.MoveLast
      rst.MoveFirst
      If rst.RecordCount > 1 Then vidtinter = escullir_tinter(vnumc)
   End If
      
   While Not rst.EOF
     If cadbl(rst!idtinter) = vidtinter Then
       imprimir_combinaciollaunes cadbl(rst!idtinter)
     End If
     rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub imprimir_combinaciollaunes(vidtinter As Double)
  Dim rst As Recordset
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vlinia As String
  Dim vtotallinia As Double
  Dim vnumc As Double
  vnumc = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, col))
  Set rst = dbtintes.OpenRecordset("select * from tintes_semblants_relacioambcomandaactiva where comanda=" + atrim(vnumc) + IIf(vidtinter > 0, " and idtinter=" + atrim(vidtinter), ""))
  If rst.EOF Then MsgBox "No s'ha trobat cap relació de llaunes guardada.", vbCritical, "Error": Exit Sub
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "ticket_impresio.rpt", 1)
  While Not rst.EOF
    oreport.FormulaFields.GetItemByName(rst!formula).Text = "'" + rst!valorformula + "'"
    rst.MoveNext
  Wend
  oreport.FormulaFields.GetItemByName("comanda").Text = "'Comanda: " + atrim(vnumc) + "'"
  oreport.DiscardSavedData
  'If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
   ' Else
   '   oreport.DisplayProgressDialog = False
   '   oreport.PrintOut False, 1
  'End If
  Set rstt = Nothing
  Set rstf = Nothing
End Sub
Sub guardar_ticket_relaciocomandaactiva()
  Dim vlinia As String
  Dim vtotallinia As Double
  Dim vnumc As Double
  Dim vidtinter As Double
  Dim rst As Recordset
  
  vnumc = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, col))
  If llistatintes.ListIndex = -1 Then MsgBox "No hi ha tinta escullida a la llista de Comandes Actives.", vbCritical, "Error": Exit Sub
  vidtinter = cadbl(llistatintes.ItemData(llistatintes.ListIndex))
  If vnumc = 0 Then MsgBox "No hi ha comanda seleccionada a la reixa de comandes actives.", vbCritical, "Error": Exit Sub
  Set rst = dbtintes.OpenRecordset("select * from tintes_semblants_relacioambcomandaactiva where comanda=" + atrim(vnumc) + " and idtinter=" + atrim(vidtinter))
  If Not rst.EOF Then
     If MsgBox("La comanda " + atrim(vnumc) + "amb la tinta:" + Chr(13) + Mid(llistatintes, 13, 30) + Chr(13) + "ja te una relació de tintes guardada..." + Chr(13) + "VOLS BORRAR PRIMER AQUESTA RELACIÓ?", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then GoTo fi
     dbtintes.Execute "delete * from tintes_semblants_relacioambcomandaactiva where comanda=" + atrim(vnumc) + " and idtinter=" + atrim(vidtinter)
     Set rst = dbtintes.OpenRecordset("select * from tintes_semblants_relacioambcomandaactiva where comanda=" + atrim(vnumc))
     If rst.EOF Then
        dbtintes.Execute "update comandesrevisadesatintes set combinaciollaunesfeta=false where comanda=" + atrim(vnumc)
        dbtintes.Execute "update comandesactives set combinaciollaunesfeta=false where comanda=" + atrim(vnumc)
     End If
     MsgBox "Ja s'ha borrat la relació, si vols guardar la nova combinació torna a apretar el botó de guardar.", vbInformation, "Relació borrada"
     GoTo fi
  End If
  If MsgBox("Vols relacionar aquest combinació de tintes amb la comanda " + atrim(vnumc) + Chr(13) + Mid(llistatintes, 13, 30) + "?", vbExclamation + vbDefaultButton2 + vbYesNo, "RELACIÓ") = vbNo Then GoTo fi
  
  dbtintes.Execute "insert into tintes_semblants_relacioambcomandaactiva (idtinter,comanda,formula,valorformula) values (" + atrim(vidtinter) + "," + atrim(vnumc) + ",'l1'," + "'Sit Llauna F1: " + atrim(vsitprimerallauna) + "'" + ")"
  dbtintes.Execute "insert into tintes_semblants_relacioambcomandaactiva (idtinter,comanda,formula,valorformula) values (" + atrim(vidtinter) + "," + atrim(vnumc) + ",'l1.2'," + "'Sit Llauna F2: " + atrim(vsitprimerallaunaf2) + " (" + atrim(kgxrecuperar(0)) + "K)'" + ")"
  dbtintes.Execute "insert into tintes_semblants_relacioambcomandaactiva (idtinter,comanda,formula,valorformula) values (" + atrim(vidtinter) + "," + atrim(vnumc) + ",'l1.3'," + "'Sit Llauna F3: " + atrim(vsitprimerallaunaf3) + " (" + atrim(kgxrecuperar(1)) + "K)'" + ")"
  dbtintes.Execute "insert into tintes_semblants_relacioambcomandaactiva (idtinter,comanda,formula,valorformula) values (" + atrim(vidtinter) + "," + atrim(vnumc) + ",'l4'," + "'F1: " + atrim(Label22(8)) + "')"
  dbtintes.Execute "insert into tintes_semblants_relacioambcomandaactiva (idtinter,comanda,formula,valorformula) values (" + atrim(vidtinter) + "," + atrim(vnumc) + ",'l5'," + "'F2: " + atrim(Label22(9)) + "')"
  dbtintes.Execute "insert into tintes_semblants_relacioambcomandaactiva (idtinter,comanda,formula,valorformula) values (" + atrim(vidtinter) + "," + atrim(vnumc) + ",'l6'," + "'F3: " + atrim(Label22(1)) + "')"
  
  vlinia = 8
  For i = 1 To reixaformulacio.Rows - 1
      If InStr(1, reixaformulacio.TextMatrix(i, 0), "BASE ") > 0 And cadbl(reixaformulacio.TextMatrix(i, 1)) > 0 Then
            vtotallinia = cadbl(reixaformulacio.TextMatrix(i, 6))
            dbtintes.Execute "insert into tintes_semblants_relacioambcomandaactiva (idtinter,comanda,formula,valorformula) values (" + atrim(vidtinter) + "," + atrim(vnumc) + ",'" + "l" + atrim(vlinia) + "'," + "'" + reixaformulacio.TextMatrix(i, 0) + "->" + atrim(vtotallinia) + "')"
            vlinia = vlinia + 1
      End If
  Next i
  dbtintes.Execute "update comandesrevisadesatintes set combinaciollaunesfeta=true where comanda=" + atrim(vnumc)
  dbtintes.Execute "update comandesactives set combinaciollaunesfeta=true where comanda=" + atrim(vnumc)
  carregar_liniadelareixaseleccionada
fi:
  Set rst = Nothing
End Sub
Sub imprimir_ticket_semblants()
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim vlinia As String
  Dim vtotallinia As Double
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "ticket_impresio.rpt", 1)
  oreport.FormulaFields.GetItemByName("l1").Text = "'Sit Llauna F1: " + atrim(vsitprimerallauna) + "'"
  oreport.FormulaFields.GetItemByName("l1.2").Text = "'Sit Llauna F2: " + atrim(vsitprimerallaunaf2) + " (" + atrim(kgxrecuperar(0)) + "K)'"
  oreport.FormulaFields.GetItemByName("l1.3").Text = "'Sit Llauna F3: " + atrim(vsitprimerallaunaf3) + " (" + atrim(kgxrecuperar(1)) + "K)'"
  'oreport.FormulaFields.GetItemByName("l2").Text = "'Kg a recuperar: " + atrim(kgxrecuperar(0)) + "'"
  oreport.FormulaFields.GetItemByName("l4").Text = "'F1: " + atrim(Label22(8)) + "'"
  oreport.FormulaFields.GetItemByName("l5").Text = "'F2: " + atrim(Label22(9)) + "'"
  oreport.FormulaFields.GetItemByName("l6").Text = "'F3: " + atrim(Label22(1)) + "'"
  oreport.FormulaFields.GetItemByName("t1").Text = "'Lot: " + atrim(cadbl(reixacomandes.TextMatrix(reixacomandes.Row, 0))) + "'"
  oreport.FormulaFields.GetItemByName("t2").Text = "'" + etextensio + "'"
  
  vlinia = 8
  For i = 1 To reixaformulacio.Rows - 1
      If InStr(1, reixaformulacio.TextMatrix(i, 0), "BASE ") > 0 And cadbl(reixaformulacio.TextMatrix(i, 1)) > 0 Then
            vtotallinia = cadbl(reixaformulacio.TextMatrix(i, 6))
            oreport.FormulaFields.GetItemByName("l" + atrim(vlinia)).Text = "'" + reixaformulacio.TextMatrix(i, 0) + "->" + atrim(vtotallinia) + "'"
            vlinia = vlinia + 1
      End If
      
  Next i
  
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

Private Sub Command67_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Command67(Index).ToolTipText = "Botons filtres programables" Then
       botofiltreprogramable Index, Button
  End If
  If Index = 30 Then vtempsbotoapretat = Now
End Sub

Private Sub Command67_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If DateDiff("s", Now, vtempsbotoapretat) <= -2 And DateDiff("s", Now, vtempsbotoapretat) > -5 Then
     If Index = 30 Or Index = 31 Then
      If MsgBox("Vols copiar els Delta de l'altra màquina?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
           copiar_deltes_maquina IIf(Index = 30, 7, IIf(Index = 31, 9, 0)), Frame5(1).tag
            carregar_valors_cuatricomia Frame5(1).tag
      End If
     End If
  End If
  vtempsbotoapretat = 31 / 12 / 1  'posso la data mes antiga per deixar sense valor
End Sub

Sub copiar_deltes_maquina(vnummaqoncopiar As Double, vnumtreballiversio As String)
   Dim vnumtreball As Double
   Dim vnumversio As Double
   Dim vnummaqorigen As Double
   separartreballiversio vnumtreballiversio, vnumtreball, vnumversio
   If vnummaqoncopiar = 7 Then vnummaqorigen = 9
   If vnummaqoncopiar = 9 Then vnummaqorigen = 7
    Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where nummaq=" + atrim(vnummaqorigen) + " and numtreballiversio='" + atrim(vnumtreball) + "/" + atrim(vnumversio) + "'")
     If Not rst.EOF Then
         Set rst2 = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where nummaq=" + atrim(vnummaqoncopiar) + " and numtreballiversio='" + atrim(vnumtreball) + "/" + atrim(vnumversio) + "'")
         If Not rst2.EOF Then
            If MsgBox("Aquesta versió de treball ja té valors Delta entrats per aquesta màquina..." + vbNewLine + "  VOLS SOBRE-ESCRIURE AQUESTS VALORS?", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbNo Then GoTo fi
            dbtintes.Execute "delete * from valorscuatricomia_treball where nummaq=" + atrim(vnummaqoncopiar) + " and numtreballiversio='" + atrim(vnumtreball) + "/" + atrim(vnumversio) + "'"
         End If
         Set rst2 = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where nummaq=" + atrim(vnummaqoncopiar) + " and numtreballiversio='" + atrim(vnumtreball) + "/" + atrim(vnumversio) + "'")
         Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where nummaq=" + atrim(vnummaqorigen) + " and numtreballiversio='" + atrim(vnumtreball) + "/" + atrim(vnumversio) + "'")
           Else: MsgBox "No he trobat dades d'aquest treball i versió.", vbCritical, "Errro"
     End If
     While Not rst.EOF
        rst2.AddNew
        For i = 0 To rst.Fields.Count - 1
             rst2.Fields(i) = rst.Fields(i)
        Next i
        rst2!numtreballiversio = atrim(vnumtreball) + "/" + atrim(vnumversio)
        rst2!nummaq = vnummaqoncopiar
        rst2.Update
        rst.MoveNext
     Wend
fi:
End Sub
Private Sub Command68_Click()
   Dim vsql As String
   Dim rst As Recordset
   Dim rstcomp As Recordset
   Dim rstcomp2 As Recordset
   Dim i As Integer
   Dim j As Integer
   kgformula = ""
   vsitprimerallaunaf2 = ""
   If vfocusultimcontrol = "formulaacomparar" Then formulaacomparar = ""
   If vfocusultimcontrol = "formulaacomparar2" Then formulaacomparar2 = ""
   Command43_Click
   vsql = "SELECT esbase,Formules.codiformula, Componentsbase.nomcomponent, DetallFormules.[%decomponent] as tanx100 FROM (Formules INNER JOIN DetallFormules ON Formules.idformula = DetallFormules.IDFormula) INNER JOIN Componentsbase ON DetallFormules.IdComponente = Componentsbase.idcomponent"
   vsql = vsql + " WHERE  esbase<>' ' and (((Formules.codiformula)='" + formulasemblanta.tag + "')) ;"
  ' Clipboard.Clear
  ' Clipboard.SetText vsql
   Set rstcomp = dbtintes.OpenRecordset("select * from componentsbase")
   Set rst = dbtintes.OpenRecordset(vsql)
   While Not rst.EOF
      rstcomp.FindFirst "nomcomponent='" + atrim(rst!nomcomponent) + "'"
      If Not rstcomp.NoMatch Then
             Set rstcomp2 = dbtintes.OpenRecordset("select * from componentsbase where esbase='" + atrim(rst!esbase) + "'")
              Else: Set rstcomp2 = dbtintes.OpenRecordset("select * from componentsbase where 1=2")
      End If
      For i = 0 To dllistadecomponents.ListCount - 1
          If dllistadecomponents.List(i) = atrim(rst!nomcomponent) Then
               dllistadecomponents.Selected(i) = True
               While Not rstcomp2.EOF
                 For j = 0 To dllistadecomponents.ListCount - 1
                      If dllistadecomponents.List(j) = atrim(rstcomp2!nomcomponent) Then dllistadecomponents.Selected(j) = True: Exit For
                 Next j
                 rstcomp2.MoveNext
               Wend
               Exit For
          End If
      Next i
      rst.MoveNext
   Wend
   checknomesseleccionats.Value = 1
   Set rst = Nothing
   Set rstcomp = Nothing
   Set rstcomp2 = Nothing
   Command43_Click
   
   carregarvalorsdepestanyaformulaalcombo
  ' formulaacomparar.tag = escullir_formula_semblant
  ' If formulaacomparar.tag <> "" Then
  '      formulaacomparar.Text = formseleccio.DBGrid2.Columns(2)
  '      carregar_componentssemblants_formula2
  '      pestanyesforumes.Tab = 1
  '      pestanyesforumes.Tab = 2
  '      recalcularformulacio
  ' End If
End Sub
Sub carregarvalorsdepestanyaformulaalcombo()
    Dim rst As Recordset
    Dim v As String
    v = formulaacomparar
    formulaacomparar.Clear
    formulaacomparar = v
    v = formulaacomparar2
    formulaacomparar2.Clear
    formulaacomparar2 = v
    botorelacioguardada(2).tag = ""
    
    If dataformules.Recordset.EOF Then Exit Sub
    Set rst = dataformules.Recordset.Clone
    rst.MoveLast
    rst.MoveFirst
    'redimensiono per borrar els valors anteriors i creo espai per la nova consulta
    ReDim vcodiformules2(0)
    ReDim vcodiformules(0)
    ReDim vcodiformules2(rst.RecordCount)
    ReDim vcodiformules(rst.RecordCount)
    
    While Not rst.EOF
        formulaacomparar.AddItem rst!descripcioformula
        vcodiformules(formulaacomparar.NewIndex) = atrim(rst!codiformula)
        formulaacomparar2.AddItem rst!descripcioformula
        vcodiformules2(formulaacomparar.NewIndex) = atrim(rst!codiformula)
        rst.MoveNext
    Wend
    Set rst = Nothing
    If formulaacomparar.ListCount > 0 And formulaacomparar = "" Then
       vsitprimerallauna2 = "": formulaacomparar.SetFocus
       formulaacomparar.ListIndex = 0
       formulaacomparar.tag = ""
    End If
    
    If formulaacomparar2.ListCount > 0 And formulaacomparar2 = "" Then
       vsitprimerallauna2 = "": formulaacomparar2.SetFocus
       formulaacomparar2.ListIndex = 0
       formulaacomparar2.tag = ""
    End If
    
End Sub
Sub carregar_componentssemblants_formula2(vnumformula As Byte)
    Dim vsql As String
    Dim vformula As String
    Dim rst As Recordset
    Dim r As Byte
    Dim vcolformula As Byte
    Dim vlistboxllista As ListBox
    If vnumformula = 2 Then vformula = formulaacomparar.tag
    If vnumformula = 3 Then vformula = formulaacomparar2.tag
    vcolformula = IIf(vnumformula = 2, 2, 3)
    vsql = "SELECT Formules.codiformula, Componentsbase.nomcomponent, DetallFormules.[%decomponent] as tanx100 FROM (Formules INNER JOIN DetallFormules ON Formules.idformula = DetallFormules.IDFormula) INNER JOIN Componentsbase ON DetallFormules.IdComponente = Componentsbase.idcomponent"
    vsql = vsql + " WHERE (((Formules.codiformula)='" + vformula + "')) order by instr(1,nomcomponent,'BASE') DESC  ;"
    Set rst = dbtintes.OpenRecordset(vsql)
'    reixaformulacio.Rows = cadbl(reixaformulacio.tag)
    If cadbl(reixaformulacio.tag) = 0 Then
       reixaformulacio.Rows = 1
        'Else: reixaformulacio.Rows = cadbl(reixaformulacio.tag)
    End If
    'If reixaformulacio.Rows = 0 Then Exit Sub
    'netejo els % de la formula abans de possar els altres per evitar que algun valor quedi escrit
    For r = 1 To reixaformulacio.Rows - 1
       reixaformulacio.TextMatrix(r, vnumformula) = ""
       reixaformulacio.col = vnumformula
    Next r
    r = 1
    While Not rst.EOF
       r = buscarcomponentalareixaformulacio(atrim(rst!nomcomponent))
       If r = 0 Then
         r = reixaformulacio.Rows
         reixaformulacio.Rows = r + 1
         reixaformulacio.Row = r
         reixaformulacio.col = vnumformula
         reixaformulacio.CellBackColor = &HC0C0FF 'vermell clar
       End If
       reixaformulacio.TextMatrix(r, 0) = atrim(rst!nomcomponent)
       reixaformulacio.TextMatrix(r, vnumformula) = rst!tanx100
       rst.MoveNext
    Wend
    
    
      'buscar el codi de tinta i carregar les llaunes
    If vnumformula = 2 Then Set vlistboxllista = llista(1)
    If vnumformula = 3 Then Set vlistboxllista = llista(2)
    vlistboxllista.tag = ""
    Set rst = dbtintes.OpenRecordset("SELECT tintesformules.numformula, tintes.codi, tintes.descripcio FROM tintesformules LEFT JOIN tintes ON tintesformules.idtinta = tintes.idtinta WHERE (((tintesformules.numformula)='" + vformula + "'));")
    If Not rst.EOF Then
          vlistboxllista.tag = atrim(rst!codi)
    End If
    carregar_llaunes_semblants vlistboxllista, vformula
    
End Sub
Function buscarcomponentalareixaformulacio(vnomcomponent As String) As Integer
    Dim i As Byte
    i = 1
    While i < reixaformulacio.Rows
      If reixaformulacio.TextMatrix(i, 0) = vnomcomponent Then GoTo fi
      i = i + 1
    Wend
    i = 0
fi:
   buscarcomponentalareixaformulacio = i
End Function

Private Sub Command69_Click()
   recalcularformulacio
End Sub
Sub recalcularformulacio()
   Dim kgF2 As Double
   Dim kgF3 As Double
   Dim kg2 As Double
   Dim vtotaldosificador As Double
   Dim vtotverd As Boolean
   Dim vdiftanxcent As Double
   
   vtotverd = True
   etiquetatotalkg.tag = ""
   etiquetatotalkg.WhatsThisHelpID = 0
   kgF2 = cadbl(kgxrecuperar(0))
   kgF3 = cadbl(kgxrecuperar(1))
   kg2 = cadbl(kgformula)
   etiquetatotalkg = "Total dosificador: "
   'If kgF2 = 0 Then Exit Sub
   With reixaformulacio
   
   .col = 6
   For i = 1 To .Rows - 1
      'If InStr(1, .TextMatrix(i, 0), "BASE ") > 0 And cadbl(.TextMatrix(i, 1)) > 0 Then
     If i < .Rows Then
      If cadbl(.TextMatrix(i, 1)) > 0 Then
          .TextMatrix(i, 4) = "0"
          If kgF2 > 0 Then
            vdiftanxcent = cadbl(.TextMatrix(i, 1) - cadbl(.TextMatrix(i, 2))) 'dif % F2
            .TextMatrix(i, 4) = (cadbl(vdiftanxcent * 100) / 10) * kgF2  'dif grms  F2
          End If
          If kgF3 > 0 Then
            vdiftanxcent = cadbl(.TextMatrix(i, 1) - cadbl(.TextMatrix(i, 3))) 'dif % F3
            .TextMatrix(i, 4) = cadbl(.TextMatrix(i, 4)) + ((cadbl(vdiftanxcent * 100) / 10) * kgF3) 'dif grms F3
          End If
          .TextMatrix(i, 4) = Redondejar(cadbl(.TextMatrix(i, 4)), 5) 'redondejo a 5 digits el grams
          .TextMatrix(i, 5) = (cadbl(.TextMatrix(i, 1)) / 0.1) * kg2 'dif kg formula
          .TextMatrix(i, 6) = Redondejar(cadbl(.TextMatrix(i, 4)) + cadbl(.TextMatrix(i, 5)), 4) 'total
          If InStr(1, .TextMatrix(i, 0), "BASE ") > 0 Or InStr(1, .TextMatrix(i, 0), " PRIMAR ") > 0 Then
             .Row = i
             If cadbl(.TextMatrix(i, 6)) < 0 Then
               .CellBackColor = &HC0C0FF 'vermell clar
               vtotverd = False
                 Else:
                   .CellBackColor = &H6BEBB1   'verd clar
                   etiquetatotalkg.WhatsThisHelpID = etiquetatotalkg.WhatsThisHelpID + cadbl(.TextMatrix(i, 6))
             End If
          End If
          vtotaldosificador = vtotaldosificador + cadbl(.TextMatrix(i, 6))
      End If
      If cadbl(.TextMatrix(i, 1)) = 0 And cadbl(.TextMatrix(i, 2)) = 0 And cadbl(.TextMatrix(i, 3)) = 0 Then
           If i > 1 Then .RemoveItem i
      End If
     End If
   Next i
   End With
   vtotaldosificador = Redondejar(vtotaldosificador, 2)
   etiquetatotalkg = "Total dosificador: " + atrim(vtotaldosificador + kgF2 + kgF3) + " Grms."
   If vtotverd Then etiquetatotalkg.tag = "totverd"
End Sub

Private Sub Command70_Click()
    etiquetatotalkg.tag = ""
    If cadbl(kgformula) = 50 Then kgformula = 0
    reixaformulacio.Redraw = False
    While etiquetatotalkg.tag <> "totverd" And cadbl(kgformula) < 50
       kgformula = cadbl(kgformula) + 0.1
       recalcularformulacio
    Wend
    While etiquetatotalkg.tag = "totverd" And cadbl(kgformula) > 0
       kgformula = cadbl(kgformula) - 0.001
       recalcularformulacio
       If etiquetatotalkg.tag <> "totverd" Then
          kgformula = cadbl(kgformula) + 0.001
       End If
       DoEvents
       
    Wend
    While etiquetatotalkg.tag = "totverd" And cadbl(kgformula) > 0
       kgformula = cadbl(kgformula) - 0.0001
       recalcularformulacio
       If etiquetatotalkg.tag <> "totverd" Then
          kgformula = cadbl(kgformula) + 0.0001
       End If
       DoEvents
    Wend
    recalcularformulacio
    reixaformulacio.Redraw = True
End Sub

Sub carregarvalorformulaacomparar(vvalor As String, vnumformula As Byte)

   If formulaacomparar = "" Then
     formulaacomparar.tag = ""
     kgformula = ""
     vsitprimerallauna2 = "":
   End If
   dataformules.Recordset.FindFirst "codiformula='" + vvalor + "'"
   If dataformules.Recordset.NoMatch Then dataformules.Recordset.FindFirst "descripcioformula='" + vvalor + "'"
   If Not dataformules.Recordset.NoMatch Then
      If vnumformula = 2 Then formulaacomparar.tag = atrim(dataformules.Recordset!codiformula)
      If vnumformula = 3 Then formulaacomparar2.tag = atrim(dataformules.Recordset!codiformula)
   End If
   carregar_componentssemblants_formula2 vnumformula
   recalcularformulacio
End Sub

Private Sub dllistadecomponents_Click()
  Dim rst As Recordset
  Dim i As Byte
  Dim vSioNo As Boolean
  Dim vIndex As Long

  vIndex = dllistadecomponents.ListIndex
  vSioNo = dllistadecomponents.Selected(dllistadecomponents.ListIndex)
  Set rst = dbtintes.OpenRecordset("select * from componentsbase")
  If rst.EOF Then Exit Sub
  rst.FindFirst "idcomponent=" + atrim(dllistadecomponents.ItemData(dllistadecomponents.ListIndex))
  If Not rst.NoMatch Then
      Set rst = dbtintes.OpenRecordset("select * from componentsbase where esbase='" + atrim(rst!esbase) + "'")
      While Not rst.EOF
         For i = 0 To dllistadecomponents.ListCount - 1
           If dllistadecomponents.ItemData(i) = rst!idcomponent Then dllistadecomponents.Selected(i) = vSioNo: Exit For
         Next i
         rst.MoveNext
      Wend
  End If
  Set rst = Nothing
  dllistadecomponents.ListIndex = vIndex
  
End Sub

Private Sub ettinta_Click(Index As Integer)
End Sub

Private Sub Form_Activate()
   If Not vcontrasenyavalida Then
    If UCase(Environ("computername")) = "ORD_TINTES" Or UCase(Environ("computername")) = "14-00372" Then
            If InputBoxEx("Escriu la contrasenya per treballar amb el programa.", "Contrasenya", , , , , , SPassword) <> "0429" Then End
            enviar_compres_programades
    End If
   End If
   vcontrasenyavalida = True
   
End Sub

Private Sub formulaacomparar_Click()
   Dim rst As Recordset
    kgformula = "":  vsitprimerallauna2 = "":
    If formulaacomparar.ItemData(formulaacomparar.ListIndex) > 0 Then
      Set rst = dbtintes.OpenRecordset("SELECT tintes.codi, tintesformules.numformula, tintesformules.predeterminada FROM tintesformules LEFT JOIN tintes ON tintesformules.idtinta = tintes.idtinta where codi='" + atrim(formulaacomparar.ItemData(formulaacomparar.ListIndex)) + "'")
      If Not rst.EOF Then
         formulaacomparar.tag = rst!numformula
         carregar_componentssemblants_formula2 2
         
         
      End If
        Else: 'carregarvalorformulaacomparar formulaacomparar.List(formulaacomparar.ListIndex), 2
           carregarvalorformulaacomparar vcodiformules(formulaacomparar.ListIndex), 2
    End If
    If botorelacioguardada(2).tag = "semblants" Then
          If formulaacomparar2.ListCount >= formulaacomparar.ListCount Then
               formulaacomparar2.ListIndex = formulaacomparar.ListIndex
               formulaacomparar2_Click
          End If
    End If
    recalcularformulacio
End Sub

Private Sub formulaacomparar_GotFocus()
    vfocusultimcontrol = "formulaacomparar"
End Sub

Private Sub formulaacomparar_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 40 And formulaacomparar.ListIndex < formulaacomparar.ListCount - 1 Then
      'carregarvalorformulaacomparar formulaacomparar.List(formulaacomparar.ListIndex + 1), 2
      carregarvalorformulaacomparar vcodiformules(formulaacomparar.ListIndex + 1), 2
   End If
   If KeyCode = 38 And formulaacomparar.ListIndex > 0 Then
      carregarvalorformulaacomparar vcodiformules(formulaacomparar.ListIndex - 1), 2
      'carregarvalorformulaacomparar formulaacomparar.List(formulaacomparar.ListIndex - 1), 2
   End If
   If KeyCode = 46 Or KeyCode = 8 Then formulaacomparar = ""
   If KeyCode <> 40 And KeyCode <> 38 Then KeyCode = 0

End Sub

Private Sub formulaacomparar_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub formulaacomparar_LostFocus()
   If formulaacomparar = "" Then
      formulaacomparar.tag = ""
      carregar_componentssemblants_formula2 2
      timercontrolfocus.Enabled = True
      recalcularformulacio
   End If
End Sub

Private Sub formulaacomparar2_Click()
   Dim rst As Recordset
    kgformula = "":  vsitprimerallauna3 = "":
    If formulaacomparar2.ItemData(formulaacomparar2.ListIndex) > 0 Then
      Set rst = dbtintes.OpenRecordset("SELECT tintes.codi, tintesformules.numformula, tintesformules.predeterminada FROM tintesformules LEFT JOIN tintes ON tintesformules.idtinta = tintes.idtinta where codi='" + atrim(formulaacomparar2.ItemData(formulaacomparar2.ListIndex)) + "'")
      If Not rst.EOF Then
         formulaacomparar2.tag = rst!numformula
         carregar_componentssemblants_formula2 3
      End If
        Else:
         formulaacomparar2.tag = ""
         carregarvalorformulaacomparar vcodiformules2(formulaacomparar2.ListIndex), 3
         'formulaacomparar2.List(formulaacomparar2.ListIndex), 3
    End If
    recalcularformulacio
    
End Sub

Private Sub formulaacomparar2_GotFocus()
   vfocusultimcontrol = "formulaacomparar2"
End Sub

Private Sub formulaacomparar2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Or KeyCode = 8 Then formulaacomparar2 = ""
   If KeyCode <> 40 And KeyCode <> 38 Then KeyCode = 0
End Sub

Private Sub formulaacomparar2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub formulaacomparar2_LostFocus()
   If formulaacomparar2 = "" Then
      formulaacomparar2.tag = ""
      carregar_componentssemblants_formula2 3
      timercontrolfocus.Enabled = True
      recalcularformulacio
   End If
End Sub

Private Sub framedensitats_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub kgformula_LostFocus()
  If cadbl(kgformula) > 999 Then MsgBox "Aquest número no pot ser tan gran", vbCritical, "Error": kgformula = "0"
End Sub

Private Sub kgxrecuperar_Change(Index As Integer)
   If Index = 0 Or Index = 1 Then
     kgformula = "0"
     recalcularformulacio
   End If
End Sub

Private Sub kgxrecuperar_DblClick(Index As Integer)
   Dim v As String
   If Index = 2 Then
       v = InputBox("Escriu l'anilox que s'ha utilitzat per formular aquesta formula.", "Anilox utilitzat")
       If StrPtr(v) = 0 Then Exit Sub
         dataformules.Recordset.Edit
         dataformules.Recordset!aniloxformulada = cadbl(v)
         dataformules.Recordset.Update
   End If
End Sub

Private Sub kgxrecuperar_GotFocus(Index As Integer)
  If Index > 2 And Index < 7 Then
     kgxrecuperar(Index).SelStart = 0
     kgxrecuperar(Index).SelLength = Len(kgxrecuperar(Index)) + 1
  End If
End Sub

Private Sub kgxrecuperar_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 2 Then KeyAscii = 0
End Sub

Private Sub kgxrecuperar_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 2 Then Exit Sub
   If KeyCode = 38 And Index < 3 Then kgxrecuperar(Index) = cadbl(kgxrecuperar(Index)) + 0.5
   If KeyCode = 40 And Index < 3 Then kgxrecuperar(Index) = cadbl(kgxrecuperar(Index)) - 0.5
   If cadbl(kgxrecuperar(Index)) < 0 And Index < 3 Then kgxrecuperar(Index) = 0
   'If KeyCode = 38 Or KeyCode = 40 Then: recalcularformulacio
End Sub

Private Sub kgxrecuperar_LostFocus(Index As Integer)
   If (Index >= 19 And Index <= 22) Or (Index >= 3 And Index <= 6) Then
        If Index = 19 Or Index = 3 Then guardar_cuatricomia_color "N", cadbl(kgxrecuperar(19)), cadbl(kgxrecuperar(3)), Combo(0)
        If Index = 20 Or Index = 4 Then guardar_cuatricomia_color "G", cadbl(kgxrecuperar(20)), cadbl(kgxrecuperar(4)), Combo(1)
        If Index = 21 Or Index = 5 Then guardar_cuatricomia_color "M", cadbl(kgxrecuperar(21)), cadbl(kgxrecuperar(5)), Combo(2)
        If Index = 22 Or Index = 6 Then guardar_cuatricomia_color "C", cadbl(kgxrecuperar(22)), cadbl(kgxrecuperar(6)), Combo(3)
        dbtintes.Execute "delete * from valorscuatricomia_treball where numtreballiversio='" + Frame5(1).tag + "' and delta=0 and deltaE=0"
   End If
End Sub
Sub separartreballiversio(vnumtreballiversio As String, vnumtreball As Double, vnumversio As Double)
    If vnumtreballiversio = "" Then Exit Sub
    If InStr(1, vnumtreballiversio, "/") = 0 Then Exit Sub
    vnumtreball = cadbl(Mid(vnumtreballiversio, 1, InStr(1, vnumtreballiversio, "/") - 1))
    vnumversio = cadbl(substituir(atrim(vnumtreballiversio), atrim(vnumtreball) + "/", ""))
End Sub

Sub guardar_cuatricomia_color(vcolor As String, vd As Double, vdE As Double, vanilox As String)
   Dim rst As Recordset
   Dim vnumtreballiversio As String
   vnumtreballiversio = Frame5(1).tag
   Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where nummaq=" + atrim(maquinaescullidaPerCuatricomia) + " and color='" + atrim(vcolor) + "' and  aniloxivolum='" + vanilox + "' and numtreballiversio='" + vnumtreballiversio + "'")
   If Not rst.EOF Then
       rst.Edit
       rst!delta = vd
       rst!deltaE = vdE
       rst.Update
   End If
   Set rst = Nothing
End Sub

Private Sub llista_DblClick(Index As Integer)
 Dim vnumllauna As String
 vnumllauna = atrim(Mid(" " + llista(Index), 1, InStr(1, llista(Index), " --")))
 colocarsealallauna vnumllauna
End Sub
Sub modificar_dades_compraprogramada()
   Dim vq As Double
   Dim vdata As String
   Dim rst As Recordset
   Dim v As String
   
   Set rst = dbtintes.OpenRecordset("select * from compresdiferides where id=" + atrim(llistacompres.ItemData(llistacompres.ListIndex)))
   If Not rst.EOF Then
       vdata = atrim(rst!Dataexecuciocompra)
       vq = atrim(rst!quantitat)
       v = InputBox("Escriu la data que vols fer la comanda:", "Data", vdata)
       If IsDate(v) Then vdata = v Else GoTo fi
       v = InputBox("Escriu la quantitat que vols demanar:" + vbNewLine + vbNewLine + "SI POSES 0 S'ELIMINARÀ LA COMANDA", "Quantitat", vq)
       If StrPtr(v) = 0 Then GoTo fi
       vq = cadbl(v)
       If vq = 0 Then dbtintes.Execute "delete * from compresdiferides where id=" + atrim(llistacompres.ItemData(llistacompres.ListIndex))
       dbtintes.Execute "update  compresdiferides set dataexecuciocompra=#" + Format(vdata, "mm/dd/yy") + "#,quantitat=" + atrim(passaradecimalpunt(atrim(vq))) + " where id=" + atrim(llistacompres.ItemData(llistacompres.ListIndex))
       actualitzar_llista_compres
   End If
   
fi:
  Set rst = Nothing
End Sub
Private Sub llistacompres_DblClick()
   Dim vq As Double
   Dim r As String
   Dim rst As Recordset
   Dim vmodificat As String
   Dim vant As String
   
   If llistacompres.ItemData(llistacompres.ListIndex) = 0 Then MsgBox "Aquesta linia no es pot modificar.", vbCritical, "Error": Exit Sub
   If Mid(llistacompres.Text, 1, 10) = "Data Exec:" Then
          modificar_dades_compraprogramada
          GoTo fi
      Else
        r = InputBox("Entra la quantitat que vols demanar." + Chr(10) + " escriu 0 per eliminar la compra." + Chr(10) + "Escriu D per passar a [D]emanat o P per [P]endent de demanar." + Chr(10) + "Escriu [OBS] per modificar l'observació", "Quantitat")
   End If
   If StrPtr(r) = 0 Then Exit Sub
   r = UCase(r)
   If r = "P" Or r = "D" Then
      If r = "P" Then dbtintes.Execute "update comprespendents set demanat=false,data=null where id=" + atrim(llistacompres.ItemData(llistacompres.ListIndex))
      If r = "D" Then dbtintes.Execute "update comprespendents set demanat=true,data=now where id=" + atrim(llistacompres.ItemData(llistacompres.ListIndex))
      GoTo fi
   End If
   Set rst = dbtintes.OpenRecordset("select * from comprespendents where id=" + atrim(llistacompres.ItemData(llistacompres.ListIndex)))
   If r = "OBS" Then
      'Set rst = dbtintes.OpenRecordset("select * from comprespendents where id=" + atrim(llistacompres.ItemData(llistacompres.ListIndex)))
      If Not rst.EOF Then
         vant = atrim(rst!observacio)
         r = InputBox("Escriu la modificació per aquesta linia", "Observació", atrim(rst!observacio))
         
         If StrPtr(r) = 0 Then Exit Sub
         r = treure_apostruf(r)
         If atrim(vant) <> atrim(r) Then
            dbtintes.Execute "update comprespendents set observacio='" + atrim(r) + "' where id=" + atrim(llistacompres.ItemData(llistacompres.ListIndex))
            vmodificat = "Compra: " + atrim(rst!descripcio) + Chr(13) + Chr(10)
            vmodificat = vmodificat + "Modificacio observació: [" + vant + "] --> [" + atrim(r) + "]"
         End If
         Set rst = Nothing
         GoTo fi
      End If
   End If
   vq = cadbl(r)
   If vq = 0 Then
      If MsgBox("Segur que vols eliminar aquesta compra?", vbCritical + vbDefaultButton2 + vbYesNo, "Eliminar") = vbNo Then Exit Sub
      vmodificat = "Compra: " + atrim(rst!descripcio) + Chr(13) + Chr(10)
      vmodificat = vmodificat + "****  COMPRA ELIMINADA  ****"
      dbtintes.Execute "delete * from comprespendents where id=" + atrim(llistacompres.ItemData(llistacompres.ListIndex))
        Else:
          If cadbl(vq) <> cadbl(rst!quantitat) Then
           vmodificat = "Compra: " + atrim(rst!descripcio) + Chr(13) + Chr(10)
           vmodificat = vmodificat + "Modificacio QUANTITAT: [" + atrim(rst!quantitat) + "] --> [" + atrim(vq) + "]"
           dbtintes.Execute "update comprespendents set quantitat=" + atrim(vq) + " where id=" + atrim(llistacompres.ItemData(llistacompres.ListIndex))
          End If
   End If
fi:
   If vmodificat <> "" Then enviarlescompresaldepartamentdecompres vmodificat
   actualitzar_llista_compres
End Sub
Sub afegir_llista_compres()
    triar_tintes_compres
    actualitzar_llista_compres
End Sub

Private Sub llistallaunesformula_DblClick()
 Dim vnumllauna As String
 vnumllauna = atrim(Mid(" " + llistallaunesformula, 1, InStr(1, llistallaunesformula, " --")))
  colocarsealallauna vnumllauna
End Sub

Private Sub mcanvirecuperador_Click()
  Dim vnumllauna As String
  Dim vidproveidorrecuperador As Long
  Dim capacitatllauna As Double
  Dim rst As Recordset
  vnumllauna = InputBox("Entra el numero de llauna/Contenidor que vols fer el canvi de RECUPERADOR.", "Canvi de recuperador")
  If atrim(vnumllauna) = "" Then Exit Sub
  If noespotcanviarelrecuperador(vnumllauna) Then Exit Sub
  colocarsealallauna vnumllauna
  If tintes.Recordset.EOF Then Exit Sub
  If datallaunes.Recordset!numllauna <> UCase(vnumllauna) Then MsgBox "No he trobat la llauna " + vnumllauna, vbCritical, "Error": Exit Sub
  escullir_proveidorrecuperador vidproveidorrecuperador
  If cadbl(vidproveidorrecuperador) = 0 Then Exit Sub
  dbtintes.Execute "update llaunes set idproveidorrecuperador=" + atrim(cadbl(vidproveidorrecuperador)) + " where numllauna='" + vnumllauna + "'"
  MsgBox "Canvi fet.", vbInformation, "Canvi de Recuperador"
End Sub
Function noespotcanviarelrecuperador(vnumllauna As String) As Boolean
  Dim rst As Recordset
  noespotcanviarelrecuperador = False
  Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, recuperadorsdecontenidors.noespermetcanvis FROM Llaunes LEFT JOIN recuperadorsdecontenidors ON Llaunes.idproveidorrecuperador = recuperadorsdecontenidors.Id where numllauna='" + atrim(vnumllauna) + "'")
  If Not rst.EOF Then
      If atrim(rst!noespermetcanvis) = "S" Then
            If UCase(InputBox("Aquest recuperador no permet canvis perquè ells recullen els seus contenidors." + vbNewLine + "SI REALMENT VOLS FER AQUEST CANVI ESCRIU [CANVI DE RECUPERADOR].", "RECUPERADOR - PROVEIDOR")) <> "CANVI DE RECUPERADOR" Then noespotcanviarelrecuperador = True
      End If
  End If
  Set rst = Nothing
End Function
Private Sub mestadisticallaunesinplacsa_Click()
  Dim vregistres As Integer
  Dim rsttintes As Recordset
  Dim vmsg As String
  Dim vqtotal As Double
  Dim vkgtotal As Double
  
  vregistres = cadbl(InputBox("Quantes linies d'estadistica vols veure?", "Estadistica", 10))
  If vregistres = 0 Then Exit Sub
  
  
  Set rsttintes = dbtintes.OpenRecordset("select * from estadistica_llaunesinplacsa order by dataestadistica desc")
  Open "c:\temp\~llistatllaunesestadistica.csv" For Output As #1
  Print #1, "ESTADISTICA LLAUNES INPLACSA"
  Print #1, ""
  Print #1, "Data;Q LlaunesTotal;Kgs Totals;;Q 0-10kg;Kgs 0-10Kg;;Q 10-15kg;Kgs 10-15Kg;;Q 15-18kg;Kgs 15-18Kg;;Q 18-25kg;Kgs 18-25Kg;;"
  Print #1, ""
  For i = 1 To vregistres
     vqtotal = cadbl(rsttintes![unitats-llaunes0-10]) + cadbl(rsttintes![unitats-llaunes10-15]) + cadbl(rsttintes![unitats-llaunes15-18]) + cadbl(rsttintes![unitats-llaunes18-25])
     vkgtotal = cadbl(rsttintes![kg-llaunes0-10]) + cadbl(rsttintes![kg-llaunes10-15]) + cadbl(rsttintes![kg-llaunes15-18]) + cadbl(rsttintes![kg-llaunes18-25])
     vmsg = Format(rsttintes!dataestadistica, "dd/mm/yy") + ";" + atrim(Redondejar(vqtotal, 0)) + ";" + atrim(Redondejar(vkgtotal, 0)) + ";;" + atrim(rsttintes![unitats-llaunes0-10]) + ";" + atrim(rsttintes![kg-llaunes0-10]) + ";"
     vmsg = vmsg + ";" + atrim(rsttintes![unitats-llaunes10-15]) + ";" + atrim(rsttintes![kg-llaunes10-15]) + ";"
     vmsg = vmsg + ";" + atrim(rsttintes![unitats-llaunes15-18]) + ";" + atrim(rsttintes![kg-llaunes15-18]) + ";"
     vmsg = vmsg + ";" + atrim(rsttintes![unitats-llaunes18-25]) + ";" + atrim(rsttintes![kg-llaunes18-25])
     Print #1, vmsg
     rsttintes.MoveNext
     If rsttintes.EOF Then GoTo fi
     
  Next i
fi:
  Close #1
  obrir_document "c:\temp\~llistatllaunesestadistica.csv"
  Set rsttintes = Nothing
  
  
End Sub

Private Sub mguardaramplescomandes_Click()
   guardar_amples_reixa
   MsgBox "Amples de la reixa de comandes guardades.", vbInformation, "Ample reixa"
End Sub

Private Sub mllistatcontenidors_Click()
  Dim vemail As String
  vemail = atrim(InputBox("Escriu a quin email vols el llistat.", "Llistat de contenidors", "tintes@inplacsa.com"))
  If vemail <> "" Then
     escriure_ini "accionsglobals", "tirarllistatcontenidors", treure_apostruf(vemail), rutadelfitxer(cami) + "valorsprograma.ini"
     MsgBox "Llistat demanat, en pocs moments hauria d'arribar al correu.", vbInformation, "Llistat de contenidors"
  End If
End Sub

Private Sub mllistatdellaunesambasterisc_Click()
  '
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatdellaunesnoactivesambkgdetinta.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "tintes.mdb"
  oreport.FormulaFields.GetItemByName("titol").Text = "'Llistat de llaunes amb [*] estan a Impresores.'"

  oreport.RecordSelectionFormula = "{Llaunes.aimpresores}=true"
  'oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  oreport.DiscardSavedData
 ' If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
  '  Else
  '    oreport.DisplayProgressDialog = False
 '     oreport.PrintOut False, 1
 ' End If
End Sub

Private Sub mllistatllaunesamb1_7kg_Click()
    '
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatdellaunesnoactivesambkgdetinta.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "tintes.mdb"
  oreport.FormulaFields.GetItemByName("titol").Text = "'Llistat de llaunes amb 1,7Kg Kilos de tinta.'"
  oreport.RecordSelectionFormula = "{Llaunes.capacitatactual}=-1.7"
  'oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  oreport.DiscardSavedData
 ' If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
  '  Else
  '    oreport.DisplayProgressDialog = False
 '     oreport.PrintOut False, 1
 ' End If
End Sub

Private Sub mllistatnoactivesambkg_Click()
   '
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatdellaunesnoactivesambkgdetinta.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "tintes.mdb"
  oreport.FormulaFields.GetItemByName("titol").Text = "'Llistat de llaunes No actives amb Kilos de tinta.'"

  'oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + atrim(numllauna) + "'"
  'oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  oreport.DiscardSavedData
 ' If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
  '  Else
  '    oreport.DisplayProgressDialog = False
 '     oreport.PrintOut False, 1
 ' End If
End Sub

Private Sub mlotserronis_Click()
 Load formaltarep
  formaltarep.caption = "Lots erronis "
  formaltarep.width = formaltarep.width * 2
  formaltarep.Data1.DatabaseName = camitintes
  formaltarep.Data1.RecordSource = "select  numerolot as [Nºde lot erroni] from lotserronis order by numerolot"
  formaltarep.refrescar
  formaltarep.Data1.tag = "select  numerolot as [Nºde lot erroni] from lotserronis order by numerolot"
  'formaltarep.DBGrid1.Columns(0).visible = False
  formaltarep.DBGrid1.Columns(0).width = 6000
  'formaltarep.DBGrid1.Columns(2).width = 3000
  'formaltarep.DBGrid1.Columns(3).width = 1200
  'formaltarep.DBGrid1.Columns(4).width = 1500
  formaltarep.DBGrid1.width = 8000
  formaltarep.width = 8200
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub mpaletsperreciclar_Click()
 Dim vbidonsxrllençar As String
   vbidonsxrllençar = cadbl(llegir_ini("BidonsPerLlençar", "numerodepaletsperreciclar", rutadelfitxer(cami) + "valorsprograma.ini"))
   vbidonsxrllençar = InputBox("Entre quants PALETS hi ha per RECICLAR ara mateix.", "Bidons per llençar", vbidonsxrllençar)
   If vbidonsxrllençar = "" Then Exit Sub
   If cadbl(vbidonsxrllençar) = 0 Then If MsgBox("Has posat que no hi ha cap PALET, ès correcte?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   escriure_ini "BidonsPerLlençar", "numerodepaletsperreciclar", vbidonsxrllençar, rutadelfitxer(cami) + "valorsprograma.ini"
End Sub

Private Sub mpaletsxrllençar_Click()
 Dim vbidonsxrllençar As String
   vbidonsxrllençar = cadbl(llegir_ini("BidonsPerLlençar", "numerodepaletsperllençar", rutadelfitxer(cami) + "valorsprograma.ini"))
   vbidonsxrllençar = InputBox("Entre quants PALETS hi ha per LLENÇAR ara mateix.", "Bidons per llençar", vbidonsxrllençar)
   If vbidonsxrllençar = "" Then Exit Sub
   If cadbl(vbidonsxrllençar) = 0 Then If MsgBox("Has posat que no hi ha cap PALET, ès correcte?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   escriure_ini "BidonsPerLlençar", "numerodepaletsperllençar", vbidonsxrllençar, rutadelfitxer(cami) + "valorsprograma.ini"
End Sub

Private Sub mrelaciodedeltes_Click()
   Dim vnumc As Double
   vnumc = cadbl(InputBox("Escriu la comanda que vols consultar els deltes utilitzats.", "Deltes"))
   If vnumc = 0 Then Exit Sub
   Load formseleccio
  formseleccio.caption = "Deltes de la comanda " + atrim(vnumc)
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "baixes.mdb"
  formseleccio.Data1.RecordSource = "select comanda,hora,coditinta,nomdelatinta,valordelta,numbobina from impresores_valorsdelta where comanda=" + atrim(vnumc) + " order by numbobina,hora"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 700
  formseleccio.DBGrid2.Columns(1).width = 1000
  formseleccio.DBGrid2.Columns(2).width = 700
  formseleccio.DBGrid2.Columns(3).width = 3000
  formseleccio.DBGrid2.Columns(4).width = 400
  formseleccio.DBGrid2.Columns(5).width = 400
  formseleccio.Show 1
  Unload formseleccio
End Sub

Private Sub msubmenutintesrevisades_click()
  Dim v As String
  v = atrim(reixacomandes.TextMatrix(reixacomandes.Row, numcol("Nova/Repetida")))
  If Len(v) > 1 Then v = Mid(v, 1, 1)
  If v <> "N" And v <> "M" Then MsgBox "La comanda ha de ser Nova o Modificada per canviar l'estat", vbCritical, "Atenció": mtintesrevisades.visible = False: Exit Sub
End Sub
Private Sub msi_click()
   Dim vnumc As Double
   Dim vnumtreball As Double
   vnumc = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("Comanda")))
   vnumtreball = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("NºTreball")))
   canviarestatnovamodificades_reixacomandes vnumc, vnumtreball, "S"
End Sub
Private Sub mno_click()
   Dim vnumc As Double
   Dim vnumtreball As Double
   vnumc = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("Comanda")))
   vnumtreball = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("NºTreball")))
   canviarestatnovamodificades_reixacomandes vnumc, vnumtreball, "N"
End Sub
Private Sub alta_Click()
  If tintes.Recordset.EditMode = 0 Then
    tintes.RecordSource = "tintes_tot"

    tintes.Refresh
    tintes.Recordset.AddNew
    framedadestintes.Enabled = True
    ccoditinta = coditintamesun
    nomserie = ""
    nomfamilia = ""
    csubfamilia = ""
    cfamiliacolor = ""
    csubfamiliacolor = ""
    crefcolor.SetFocus
      Else: MsgBox "Ja estàs editant...", vbCritical, "Error": Exit Sub
  End If
End Sub
Function coditintamesun() As String
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select max(codi) as gran from tintes ")
   coditintamesun = cadbl(rst!gran) + 1
End Function

Private Sub Combo2_Change()

End Sub
Sub escullir_familiacolor()
  Static ultimcodi As String
  Load formseleccio
  formseleccio.caption = "Selecciona Familia Color"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "select codi,descripcio from familiescolors order by descripcio"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 4500
  If cadbl(ultimcodi) > 0 Then formseleccio.Data1.Recordset.FindFirst "codi=" + atrim(ultimcodi)
  
  formseleccio.Show 1
  If seleccioret = 1 Then
   cfamiliacolor = atrim(formseleccio.Data1.Recordset!descripcio)
   tintes.Recordset!idfamcolor = formseleccio.Data1.Recordset!codi
   ultimcodi = atrim(tintes.Recordset!idfamcolor)
  End If
  Unload formseleccio
  
End Sub

Private Sub Combo2_DropDown()

End Sub


Private Sub bactualitzacargues_Click()
  actualitzarcarguescomponents
  actualitzar_estocdecomponents
End Sub

Private Sub bassignarllauna_Click()
End Sub

Private Sub bcontrolestocminim_Click()
   comprovarestocminimdellaunes
  
  'Load formseleccio
  'formseleccio.caption = "Selecciona referencia proveidor"
  'formseleccio.Data1.DatabaseName = camitintes
  'formseleccio.Data1.RecordSource = "SELECT llaunesdecadatintaiestocminim.familiestintes.descripcio AS Familia, llaunesdecadatintaiestocminim.subfamiliestintes.descripcio AS Subfamilia, llaunesdecadatintaiestocminim.familiescolors.descripcio AS [Familia color], llaunesdecadatintaiestocminim.subfamiliescolors.descripcio AS [Subfamilia color], First(llaunesdecadatintaiestocminim.estocminim) AS [Estoc mínim], Sum(llaunesdecadatintaiestocminim.estocactual) AS [Estoc actual] From llaunesdecadatintaiestocminim GROUP BY llaunesdecadatintaiestocminim.familiestintes.descripcio, llaunesdecadatintaiestocminim.subfamiliestintes.descripcio, llaunesdecadatintaiestocminim.familiescolors.descripcio, llaunesdecadatintaiestocminim.subfamiliescolors.descripcio;"
  'formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).width = 2500
 ' formseleccio.DBGrid2.Columns(1).width = 2500
 ' formseleccio.DBGrid2.Columns(2).width = 1500
 ' formseleccio.DBGrid2.Columns(3).width = 1500
 ' formseleccio.DBGrid2.Columns(4).width = 900
 ' formseleccio.DBGrid2.Columns(5).width = 900
 'formseleccio.width = 11000
 ' formseleccio.Show 1
'
'  Unload formseleccio
  
End Sub
Function numcol(nom As String) As Byte
  Dim i As Integer
   numcol = 0
   For i = 0 To reixacomandes.Cols - 1
     If reixacomandes.TextMatrix(0, i) = nom Then numcol = i
   Next i
   
End Function

Private Sub bdescatalogar_Click()
   descatalogartinta
End Sub
Sub descatalogartinta()
    Dim resp As String
    Dim des As Boolean
    resp = UCase(InputBox("Amb aquesta opcio actives o descatalogues aquesta tinta (Descatalogar)" + Chr(10) + " Escriu activar o descatalogar.", "Descatalogar o Activar la tinta"))
    If resp = "DESCATALOGAR" Then
       des = True
         Else
            If resp = "ACTIVAR" Then
               des = False
              Else: Exit Sub
            End If
    End If
    If tintes.Recordset.EditMode > 0 Then
        tintes.Recordset!descatalogat = des
        
          Else
            dbtintes.Execute "update  tintes set descatalogat=" + IIf(des, "True", "False") + " where codi='" + atrim(ccoditinta) + "'"
            tintes.Recordset.Move 0
    End If
    posarcosesdescatalogat
End Sub

Private Sub bimportarllaunes_Click()
   Dim rstllaunes As Recordset
   Dim rsti As Recordset
   Dim rsttinta As Recordset
   Dim rstllaunesnoves As Recordset
   Dim rstformula As Recordset
   Dim rstrefproveidor As Recordset
   Dim rstb As Recordset
   Dim r As String
   
   If Not existeix("c:\ordprog.ini") Then Exit Sub
   If MsgBox("Segur que vols importar les llaunes?", vbCritical + vbYesNo, "Atenció") = vbNo Then Exit Sub
   Set rstllaunes = dbtintes.OpenRecordset("select * from temporaltoteslesllaunes")
   While Not rstllaunes.EOF
      Set rsti = dbtintes.OpenRecordset("select numllauna from llaunes where numllauna='" + atrim(rstllaunes!numlata) + "'")
      If rsti.EOF Then
        Set rsti = dbtintes.OpenRecordset("select * from temp_tintes_disponibles where despan='" + atrim(rstllaunes!despan) + "' AND desfam='" + atrim(rstllaunes!desfam) + "'")
        If rsti.EOF Then MsgBox "Tinta no trobada"
        If Not rsti.EOF Then
         '  GoTo cont
          If atrim(rsti!coditinta) <> "" Then
           Set rsttinta = dbtintes.OpenRecordset("select idtinta from tintes where codi='" + atrim(rsti!coditinta) + "'")
           Set rstformula = dbtintes.OpenRecordset("select codiformula from formules where codiformula='" + atrim(rsti!codiformula) + "'")
           If Not rsttinta.EOF Then
                Set rstrefproveidor = dbtintes.OpenRecordset("select * from tintesreferencies where idtinta=" + atrim(rsttinta!idtinta))
                If rstrefproveidor.EOF Then MsgBox "la tinta " + atrim(rsttinta!descripcio) + " no te referencia associada": GoTo cont
                dbtintes.Execute "insert into llaunes (numllauna,idtinta,situacio,activa,id_refproveidor) values ('" + atrim(rstllaunes!numlata) + "'," + atrim(rsttinta!idtinta) + ",'" + atrim(rstllaunes!situacio) + "',true," + atrim(rstrefproveidor!id) + ")"
                Set rstllaunesnoves = dbtintes.OpenRecordset("select id,numllauna from llaunes where numllauna='" + atrim(rstllaunes!numlata) + "'")
                If rstllaunesnoves.EOF Then MsgBox "Error al crear la llauna " + atrim(rstllaunes!numlata)
                If Not rstformula.EOF Then
                    Set rstb = dbtintes.OpenRecordset("select * from tintesformules where idtinta=" + atrim(rsttinta!idtinta) + " and numformula='" + atrim(rstformula!codiformula) + "'")
                    If rstb.EOF Then
                       dbtintes.Execute "insert into tintesformules (idtinta,numformula) values (" + atrim(rsttinta!idtinta) + ",'" + atrim(rstformula!codiformula) + "')"
                    End If
                    r = possarformulapredeterminada(rsttinta!idtinta)
                End If
                dbtintes.Execute "insert into historiallauna (idnumllauna,data,numrecarrega,tipusmoviment,formula,kg) values (" + atrim(rstllaunesnoves!id) + ",#01/01/2017#,1,'C','" + atrim(rstformula!codiformula) + "'," + passaradecimalpunt(cadbl(rstllaunes!kgpan)) + ")"
                Set rstb = dbtintes.OpenRecordset("select * from historiallauna where idnumllauna=" + atrim(rstllaunesnoves!id) + " order by id desc")
                If Not rstb.EOF Then dbtintes.Execute ("insert into historiallaunalots (idhistoria,idcomponent,numlotbase,tanx100tinta,kgtinta) values (" + atrim(rstb!id) + ",0,'" + atrim(rstllaunes!lofab) + "',0," + passaradecimalpunt(cadbl(rstllaunes!kgpan)) + ")")
                calcularkgdisponiblesllauna atrim(rstllaunesnoves!numllauna)
                   Else: MsgBox "La llauna " + atrim(rstllaunes!numlata) + " no te codidetinta"
           End If
cont:
          End If
            Else: MsgBox "relacio no trobada"
         End If
       End If
       
       Me.caption = "Llauna Nº " + atrim(rstllaunes.AbsolutePosition)
     rstllaunes.MoveNext
   Wend
   MsgBox "Acabat"
End Sub
Function possarformulapredeterminada(ByVal idtinta As Long) As String
  Dim rstf As Recordset
  Set rstf = dbtintes.OpenRecordset("select * from tintesformules where idtinta=" + atrim(idtinta) + " order by predeterminada")
  If Not rstf.EOF Then
     If Not rstf!predeterminada Then
         rstf.Edit
         rstf!predeterminada = True
         rstf.Update
     End If
     possarformulapredeterminada = atrim(rstf!numformula)
  End If
End Function

Private Sub boto_albarans_rebuts_Click()
   
  
End Sub

Sub possar_totselsvidonsalataula(vwhere As String)
   Dim rst As Recordset
   Dim rstr As Recordset
   Dim vsubconsulta As String
   Set rst = dbtintes.OpenRecordset("SELECT distinct  tipusbidons.nominterndelbido,tipusbidons.id FROM (tintesreferencies INNER JOIN tintes ON tintesreferencies.idtinta = tintes.idtinta) INNER JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id where " + vwhere)
   With tintes.Recordset
   While Not rst.EOF
     Set rstr = dbtintes.OpenRecordset("select * from estocsminims where descripciobido='" + atrim(rst!nominterndelbido) + "' and " + vwhere)
     'Clipboard.SetText "select * from estocsminims where descripciobido='" + atrim(cadbl(rst!nominterndelbido)) + "' and " + vwhere
     If rstr.EOF Then
         If InStr(1, vwhere, "codi") > 0 Then
'         Clipboard.SetText "insert into estocsminims (idbido,descripciobido,codi) values (" + atrim(cadbl(rst!id_bido)) + ",'" + treure_apostruf(atrim(rst!nombido)) + "'," + atrim(cadbl(!codi)) + ")"
             dbtintes.Execute "insert into estocsminims (descripciobido,codi,idbido) values ('" + treure_apostruf(atrim(rst!nominterndelbido)) + "'," + atrim(cadbl(!codi)) + "," + atrim(cadbl(rst!id)) + ")"
             Else
               vsubconsulta = "'" + treure_apostruf(atrim(rst!nominterndelbido)) + "'," + atrim(cadbl(!idfamilia)) + "," + atrim(cadbl(!idsubfamilia)) + ", " + atrim(cadbl(!idfamcolor)) + "," + atrim(cadbl(!idsubfamcolor)) + "," + atrim(cadbl(rst!id))
               dbtintes.Execute "insert into estocsminims (descripciobido,idfamilia,idsubfamilia,idfamcolor,idsubfamcolor,idbido) values (" + vsubconsulta + ")"
               'MsgBox "insert into estocsminims (idbido,descripciobido,idfamilia,idsubfamilia,idfamcolor,idsubfamcolor) values (" + vsubconsulta + ")"
         End If
     End If
     rst.MoveNext
   Wend
   End With
   Set rst = Nothing
   Set rstr = Nothing
End Sub
Private Sub botoestocminim_Click()
   Dim vem As String
   Dim ved As String
   Dim vnombidoestocminim As String
   Dim vcodibidoestocminim As Double
   Dim vwhere As String
   vwhere = " codi='" + atrim(tintes.Recordset!codi) + "'"
   'If UCase(Mid(crefcolor + "  ", 1, 2)) = "P-" Then
   '     vwhere = " codi='" + atrim(tintes.Recordset!codi) + "'"
   '       Else
   '         With tintes.Recordset
   '         vwhere = " idfamilia=" + atrim(cadbl(!idfamilia)) + " and idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + "and idfamcolor= " + atrim(cadbl(!idfamcolor)) + " and idsubfamcolor=" + atrim(cadbl(!idsubfamcolor))
   '         End With
   'End If
   possar_totselsvidonsalataula vwhere
   actualitzar_estoc_llaunes vwhere
   vnombidoestocminim = escullir_estocminim_bido(vwhere, vcodibidoestocminim)
   If vnombidoestocminim <> "" Then
     vem = formseleccio.Data1.Recordset![mínim]
     ved = formseleccio.Data1.Recordset![desitjat]
     vem = InputBox("Entra l'estoc mínim que vols tenir per aquestes families de tinta i color.", "Estoc mínim", vem)
     If Not IsNumeric(vem) Then GoTo fi
     ved = InputBox("Entra l'estoc que vols tenir quan es compri.", "Estoc desitjat", ved)
     If Not IsNumeric(ved) Then GoTo fi
     If cadbl(vem) > cadbl(ved) Then MsgBox "L'estoc desitjat no pot ser inferior al estoc mínim", vbCritical, "Error": GoTo fi
     formseleccio.Data1.Recordset.Edit
     formseleccio.Data1.Recordset![mínim] = vem
     formseleccio.Data1.Recordset![desitjat] = ved
     formseleccio.Data1.Recordset.Update
   End If
fi:
   Unload formseleccio
   dbtintes.Execute "delete * from estocsminims where estocminim<1"
   
End Sub
Sub actualitzar_estoc_llaunes(Optional vwhere As String)
   Dim rstr As Recordset
   ratoli "espera"
   Set rstr = dbtintes.OpenRecordset("select * from estocsminims " + IIf(vwhere <> "", "where " + vwhere, ""))
   While Not rstr.EOF
      rstr.Edit
      rstr!estocactual = calcular_estoc_delatinta(rstr)
      rstr.Update
      rstr.MoveNext
   Wend
   Set rstr = Nothing
   ratoli "normal"
End Sub
Function calcular_estoc_delatinta(rst As Recordset) As Double
   Dim rstestoc As Recordset
   Dim vsubconsulta As String
   Dim vwhere As String
   'aquesta funcio també s'utilitza a enviar mail servidor i a manteniment tintes
     ' QUALSEVOL CANVI S'HA D'APLICAR A ALS DOS
   With rst
   If cadbl(!codi) > 0 Then
      vwhere = "codi='" + atrim(rst!codi) + "'"
        Else
         vwhere = " (idfamilia=" + atrim(cadbl(!idfamilia)) + " and idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + "and idfamcolor= " + atrim(cadbl(!idfamcolor)) + " and idsubfamcolor=" + atrim(cadbl(!idsubfamcolor)) + ") "
   End If
   End With
   vsubconsulta = "select idtinta from tintes where " + vwhere
   Set rstestoc = dbtintes.OpenRecordset("SELECT Count(*) AS Tllaunes, Sum(Llaunes.capacitatactual) AS SumaDecapacitatactual, tipusbidons.capacitat FROM Llaunes LEFT JOIN (tipusbidons RIGHT JOIN tintesreferencies ON tipusbidons.id = tintesreferencies.id_bido) ON Llaunes.id_refproveidor = tintesreferencies.id  Where (((Llaunes.capacitatactual) > 0.9) And ((Llaunes.activa) = True)) " + IIf(atrim(rst!descripciobido) <> "", " and tipusbidons.nominterndelbido='" + atrim(rst!descripciobido) + "'", "") + " and Llaunes.idtinta in (" + vsubconsulta + ")  GROUP BY  tipusbidons.capacitat;")
   'Clipboard.SetText "SELECT Count(*) AS Tllaunes, Sum(Llaunes.capacitatactual) AS SumaDecapacitatactual, tipusbidons.capacitat FROM Llaunes LEFT JOIN (tipusbidons RIGHT JOIN tintesreferencies ON tipusbidons.id = tintesreferencies.id_bido) ON Llaunes.id_refproveidor = tintesreferencies.id  Where (((Llaunes.capacitatactual) > 0.9) And ((Llaunes.activa) = True))  and tipusbidons.nominterndelbido='" + atrim(rst!descripciobido) + "' and Llaunes.idtinta in (" + vsubconsulta + ")  GROUP BY  tipusbidons.capacitat;"
   
   calcular_estoc_delatinta = cadbl(rstestoc!SumaDecapacitatactual)
   
End Function
Function escullir_estocminim_bido(vwhere As String, vcodi As Double) As String
  Unload formseleccio
  Load formseleccio
  formseleccio.width = 6000
  formseleccio.caption = "Escullir el bidó que vols possar estoc mínim. (En Kg)"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "SELECT estocsminims.idbido, estocsminims.descripciobido as [Nom bidó], estocsminims.estocminim as [Mínim], estocsminims.estocdesitjat as [Desitjat], estocsminims.estocactual as [Actual] FROM estocsminims where" + vwhere + " order by descripciobido"
  formseleccio.refrescar
  formseleccio.caption = "Escullir el bidó que vols possar estoc mínim. (En Kg)"
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 3000
  formseleccio.DBGrid2.Columns(2).width = 700
  formseleccio.DBGrid2.Columns(3).width = 1000
  formseleccio.DBGrid2.Columns(4).width = 700
  formseleccio.width = 7000
  formseleccio.Show 1
  If seleccioret = 1 Then
   If formseleccio.Data1.Recordset.EOF Then GoTo fi
       escullir_estocminim_bido = atrim(formseleccio.Data1.Recordset![Nom bidó])
       vcodi = cadbl(formseleccio.Data1.Recordset!idbido)
       
  End If
fi:
  


End Function

Private Sub brecalcularpesllaunes_Click()
   If MsgBox("Segur que vols recalcular el pes de totes les llaunes?", vbInformation + vbYesNo, "Atenció") = vbYes Then
      If Not datadellaunes.Recordset.EOF Then datadellaunes.Recordset.MoveFirst
      While Not datadellaunes.Recordset.EOF
         calcularkgdisponiblesllauna datadellaunes.Recordset!numllauna
         datadellaunes.Recordset.MoveNext
      Wend
      MsgBox "Procés acabat."
   End If
End Sub

Private Sub buscador_Change()
  filtrartintes
End Sub


Sub filtrar_estocs()
  Dim vprimeraparaula As String
  Dim vsegonaparaula As String
  Dim vterceraparaula As String
  Dim vcampfiltrar As String
  Dim v As String
  Dim vordre As String
  If etfiltrarperestoc.tag = "" Then Exit Sub
  v = treure_apostruf(buscador_estoc(0)) + " "
  vordre = Mid(buscador_estoc(0) + "  ", 1, 2)
  If vordre = "<<" Or vordre = ">>" Then
     v = Mid(v, 3)
     If vordre = "<<" Then ordreestoc = " order by " + etfiltrarperestoc.tag + " DESC"
     If vordre = ">>" Then ordreestoc = " order by " + etfiltrarperestoc.tag + " ASC"
  End If
  If (Mid(v + "  ", 1, 1) = "<" Or Mid(v + "  ", 1, 1) = ">") Then v = Mid(v, 2) + " "
  vcampfiltrar = etfiltrarperestoc.tag
  vprimeraparaula = atrim(Mid(v, 1, InStr(1, v, " ")))
  vsegonaparaula = atrim(Mid(v, InStr(1, v, vprimeraparaula) + Len(vprimeraparaula), InStr(InStr(1, v, vprimeraparaula) + Len(vprimeraparaula), v, " ")))
  vterceraparaula = atrim(Mid(v, InStr(1, v, vsegonaparaula) + Len(vsegonaparaula)))
  buscador_estoc(0).tag = " where " + vcampfiltrar + " like '*" + vprimeraparaula + "*' " + IIf(vsegonaparaula <> "", " and " + vcampfiltrar + " like '*" + vsegonaparaula + "*'", "") + IIf(vterceraparaula <> "", " and " + vcampfiltrar + " like '*" + vterceraparaula + "*'", "")
  If ordreestoc = "" Then ordreestoc = " order by descripcio"
  poblar_reixa_estoc
  carregar_amples_reixa_estoc
End Sub


Private Sub buscador_GotFocus()
  etajudabusqueda.visible = True
End Sub

Private Sub buscador_LostFocus()
  etajudabusqueda.visible = False
End Sub

Private Sub cdensitat_GotFocus(Index As Integer)
 
End Sub

Private Sub cdensitat_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 
End Sub

Private Sub cfamiliacolor_DropDown()
   escullir_familiacolor
   If csubfamiliacolor = "" Then csubfamiliacolor_DropDown
End Sub

Private Sub cfamiliacolor_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cfiltrecolorstreballs_Change()
   'filtrarimportaciocolorstreballs
   'buscador = cfiltrecolorstreballs
End Sub

Sub creartinta_anterior() 'ara no utilitzo aquesta manera
 Dim coditintanou As String
  Dim coditintacreat As String
  Dim coditintavell As String
  Dim tintanova As Boolean
  If tintes.Recordset.EditMode > 0 Then
      If Not correctestotselscamps Then
        MsgBox "Error en alguns camps, omple els que falten i torna-ho a provar.", vbCritical, "Error"
        Exit Sub
      End If
      
      If tintes.Recordset.EditMode = 2 Then
          coditintanou = crear_coditinta(True) 'ccoditinta
          tintanova = True
            Else:
               coditintavell = ccoditinta
               coditintanou = crear_coditinta(False, coditintavell)
      End If
      
      If nohiharelacio(atrim(ccoditinta)) Then
          ccoditinta = coditintanou
         Else
            If MsgBox("Aquesta tinta ja te una relació feta." + Chr(10) + "Vols modificarla igualment?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
            
      End If
      If atrim(ccoditinta) = "" Then ccoditinta = coditintanou
      If atrim(ccoditinta) = "" Then tintes.Recordset.CancelUpdate: Exit Sub
      descripciotinta = crear_descripciotinta
      coditintacreat = coditintanou
      If ccoditinta = "" Then MsgBox "Error al generar el codi de tinta, falta algun valor per generar-lo", vbCritical, "Error": Exit Sub
      
      tintes.Recordset.Update
  End If
  tintes.RecordSource = "tintes_tot"
  tintes.Refresh
  tintes.Recordset.FindFirst "codi='" + coditintacreat + "'"
  framedadestintes.Enabled = False
  If tintanova Then crear_referencia
End Sub
Sub crear_referencia()
   formrefproveidors.Show
   formrefproveidors.crear_referencia
End Sub
Function nohiharelacio(codi As String) As Boolean
   Dim rst As Recordset
   If codi = "" Then nohiharelacio = True: Exit Function
   Set rst = dbtintes.OpenRecordset("select * from TEMP_TINTES_DISPONIBLES where coditinta='" + codi + "'")
   If rst.EOF Then nohiharelacio = True
   
   Set rst = Nothing
End Function
Function crear_descripciotinta() As String
   Dim rst As Recordset
   crear_descripciotinta = "-"
   Set rst = dbtintes.OpenRecordset("select * from subfamiliestintes where codi=" + atrim(cadbl(tintes.Recordset!idsubfamilia)))
   If rst.EOF Then GoTo fi
   crear_descripciotinta = IIf(Len(atrim(cfamiliacolor)) > 2, atrim(cfamiliacolor) + " ", "") + IIf(Len(atrim(crefcolor)) > 1, atrim(crefcolor) + " ", "") + IIf(Len(atrim(nomserie)) > 1, atrim(nomserie) + " ", "") + atrim(rst!Alias) + IIf(atrim(csubfamiliacolor) <> "-", " " + atrim(csubfamiliacolor), "")
fi:
   If crear_descripciotinta = "-" Then crear_descripciotinta = ""
   Set rst = Nothing
End Function
Function correctestotselscamps() As Boolean
   correctestotselscamps = False
   'If descripciotinta = "" Then Exit Function
   If crefcolor = "" Then Exit Function
  ' If nomproveidor = "" Then Exit Function
   If nomserie = "" Then Exit Function
   If nomfamilia = "" Then Exit Function
   If csubfamilia = "" Then Exit Function
   If cfamiliacolor = "" Then Exit Function
   If csubfamiliacolor = "" Then Exit Function
   
   correctestotselscamps = True
   
End Function
Function crear_coditinta(avisarsirepetit As Boolean, Optional codivell As String) As String
   Dim rst As Recordset
   Dim aliasproveidor As String
   Dim tipusfamilia As String
   Dim refcolor As String
   Dim numerosecuencia As Byte
   'buscar alias proveidor
   Set rst = dbcomandes.OpenRecordset("select * from proveidors where codi=" + atrim(tintes.Recordset!codiproveidor))
   If rst.EOF Then Exit Function
   aliasproveidor = atrim(rst!aliastintes)
   If aliasproveidor = "" Then MsgBox "El proveidor no te un alias assignat, primer assigna-li i despres guarda la tinta.", vbCritical, "Error": Exit Function
   
   'buscar tipusfamilia
   Set rst = dbtintes.OpenRecordset("select tipusfamilia from familiestintes where codi=" + atrim(tintes.Recordset!idfamilia))
   If rst.EOF Then Exit Function
   tipusfamilia = atrim(rst!tipusfamilia)
   If tipusfamilia = "" Then MsgBox "No hi ha tipus de familia assignat a la familia de tintes, primer assigna-li i despres guarda la tinta.", vbCritical, "Error": Exit Function
   
   'buscar refcolor
   refcolor = crefcolor
   If atrim(refcolor) = "" Then MsgBox "No hi ha referència de color entrada, primer assigna-li i despres guarda la tinta.", vbCritical, "Error": Exit Function
   'crear el codi
   crear_coditinta = atrim(aliasproveidor) + atrim(tipusfamilia) + atrim(refcolor)
   
   'buscar numerosecuencia
   numerosecuencia = 1
   If crear_coditinta <> codivell Then
    Set rst = dbtintes.OpenRecordset("select * from tintes where codi like '" + crear_coditinta + "*'")
    If Not rst.EOF Then
       rst.MoveLast
       numerosecuencia = rst.RecordCount + 1
    End If
   End If
   If numerosecuencia > 1 And avisarsirepetit Then
     If MsgBox("El codi de tinta " + ccoditinta + " ja existeix, vols crear un altre codi amb aquests valors?" + Chr(10) + "Potser es la mateixa tinta amb RF", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then
       crear_coditinta = ""
       Exit Function
     End If
   End If
   crear_coditinta = crear_coditinta + IIf(numerosecuencia > 1, "_" + atrim(numerosecuencia), "")
   If Len(crear_coditinta) > 30 Then
     MsgBox "Error al crear el codi de tinta, la longitud del codi es massa llarga.", vbCritical, "Atenció"
     crear_coditinta = ""
   End If
   
End Function



Private Sub checkactives_Click()
 filtrarllaunes
End Sub

Private Sub checkimpresores_Click()
  filtrarllaunes
End Sub

Private Sub checknoinplacsa_Click()

End Sub

Private Sub checkrevisades_Click()

End Sub

Private Sub checktotes_Click()

End Sub

Private Sub comboselimp_Change()

End Sub

Private Sub checktots_Click()
  actualitzar_llista_albarans
End Sub

Private Sub cnominplacsa_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
End Sub

Private Sub combosionookcarta_Click()
   'filtrartintes
   filtrar_tintes
End Sub
Sub vrefrescarnomtintadeclixes(vcoditinta As String, vnovadesc As String)
  If vcoditinta = "" Then Exit Sub
  dbclixes.Execute "update tintes set color='" + vnovadesc + "' where coditinta='" + atrim(vcoditinta) + "'"
End Sub
Private Sub Command1_Click()
  Dim coditintacreat As String
  Dim tintanova As Boolean
  tintanova = IIf(tintes.Recordset.EditMode = 2, True, False)
  If tintes.Recordset.EditMode = 0 Then Exit Sub
  descripciotinta = crear_descripciotinta
  If descripciotinta.tag <> descripciotinta And Not tintanova Then vrefrescarnomtintadeclixes atrim(tintes.Recordset!codi), atrim(descripciotinta)
  
  
  If descripciotinta = "" Then GoTo cancelar
  coditintacreat = ccoditinta
  If ccoditinta = "" Then MsgBox "Error al generar el codi de tinta, falta algun valor per generar-lo", vbCritical, "Error": Exit Sub
  guardar_observacio_tinta
  tintes.Recordset.Update
  dbtintes.Execute "delete * from tintes where descripcio='' or descripcio=null"
acabar:
  tintes.RecordSource = "tintes_tot"
  tintes.Refresh
  tintes.Recordset.FindFirst "codi='" + coditintacreat + "'"
  If tintanova Then crear_referencia_inplacsa tintes.Recordset
  framedadestintes.Enabled = False
  Exit Sub
cancelar:
  tintes.Recordset.CancelUpdate
  GoTo acabar
End Sub
Sub guardar_observacio_tinta()
  dbtintes.Execute "delete * from tintes_observacions where idtinta=" + atrim(tintes.Recordset!idtinta)
  If atrim(cobservacions) <> "" Then dbtintes.Execute "insert into tintes_observacions (idtinta,observacio) values (" + atrim(tintes.Recordset!idtinta) + ",'" + atrim(treure_apostruf(cobservacions)) + "')"
End Sub
Sub crear_referencia_inplacsa(rsttinta As Recordset)
   dbtintes.Execute "Insert into tintesreferencies (idtinta,referencia,id_bido,codiproveidor,nomproveidor,predeterminada) values ('" + atrim(rsttinta!idtinta) + "','" + atrim(rsttinta!codi) + "',18,580,'INPLACSA',true)"
End Sub

Private Sub Command10_Click()
 Dim rst As Recordset
 Dim vid As Long
 Dim viderror As Long
 'vid = 10610
 Set rst = dbtintes.OpenRecordset("select * from zCopia_Llaunes ") 'where id>=1575139743 order by id asc")
 While Not rst.EOF
    vid = rst!id
    viderror = rst!id2
    'dbtintes.Execute "update zCopia_Llaunes set id=" + atrim(vid) + " where id=" + atrim(viderror)
    dbtintes.Execute "update historiallauna set idnumllauna=" + atrim(vid) + " where idnumllauna=" + atrim(viderror)
    'vid = vid + 1
    rst.MoveNext
 Wend
 MsgBox "ara"
  
End Sub

Private Sub Command11_Click()
  Dim v As String
  MsgBox "Aquesta funció nomes la poden fer els ordinador que previament " + Chr(10) + " s'hi hagi possat el driver ODBC de sqlserver per connectar amb INKMAKER", vbInformation, "Atenció"
  'If MsgBox("Segur que vols actualitzar les formules?" + Chr(10) + "Aquest procés tarda uns minuts.", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  v = UCase(InputBox("Vols actualitzar totes les formules o les ultimes 10." + Chr(10) + "Escriu [T] per totes." + Chr(10) + "O escriu el codi de formula que vols actualitzar", "Actualitzar formules", 10))
  If (v <> "T" And Len(v) < 5) And cadbl(atrim(v)) = 0 Then Exit Sub
  gravant = True
  actualitzarformules v
  'tirarllistatdiferencies
  MsgBox "Actualització de formules feta.", vbInformation, "Actualització"
 ' formokformules.Show 1
End Sub
Sub comprovardiferenciesformules()
   Dim contador As Byte
   Dim longitud As Byte
   Dim rstfink As DAO.Recordset
   Dim rstf As Recordset
   Dim rsttest As Recordset
   dbtintes.Execute "delete * from test_actualitzacioformules where ok=false"
   Set rsttest = dbtintes.OpenRecordset("select * from test_actualitzacioformules")
   
   ratoli "espera"
   
   formtintes.Enabled = False
   Set wsODBC = CreateWorkspace("", "tintes", "", dbUseODBC)
   Set conODBC = wsODBC.OpenConnection("connexiosql", dbDriverNoPrompt, , "ODBC;DATABASE=InkmakerDB;UID=sa;PWD=Mak2008;DSN=tintes")
   Set rstfink = conODBC.OpenRecordset("select * from dbo.tblFormula ", dbOpenSnapshot)
   If rstfink.EOF Then MsgBox "No s'ha trobat cap formula al INKMAKER.", vbCritical, "Atenció": GoTo fi
   Set rstf = dbtintes.OpenRecordset("select * from formules")
   rstfink.MoveLast
   rstfink.MoveFirst
   dbtintes.Execute "update  formules set actualitzada=false "
   While Not rstfink.EOF
      rstf.FindFirst "codiformula='" + atrim(rstfink!code) + "'"
      If rstf.NoMatch Then
         test_crearformulanova rstf, rstfink, conODBC, rsttest
           Else
             'comprovar si els valors de la k tinc entrada es correcte... actualitzarla
             test_comprovarformula rstf, rstfink, conODBC, rsttest
      End If
      rstfink.MoveNext
      ensenyarprogresactualitzacio rstfink
   Wend
   conODBC.Close
   test_borrarformuleseliminadesdeinkmaker rsttest
fi:
   Set rstfink = Nothing
   Set rstf = Nothing
   dataformules.Refresh
   ratoli "normal"
   formtintes.Enabled = True
End Sub
Sub actualitzarformules(Optional v As String)
   Dim contador As Byte
   Dim longitud As Byte
   Dim rstfink As DAO.Recordset
   Dim rstf As Recordset
   Dim rsttest As Recordset
   Dim vcont As Double
   Dim vwhere As String
   ratoli "espera"
   If Len(v) > 4 Then
      vwhere = " where Code='" + v + "' "
       Else
        vcont = cadbl(atrim(v))
        If vcont = 0 Then v = "T"
   End If
   Set rsttest = dbtintes.OpenRecordset("select * from test_actualitzacioformules")
   If rsttest.EOF Then
      crearliniablanc rsttest
      Set rsttest = dbtintes.OpenRecordset("select * from test_actualitzacioformules")
   End If
   formtintes.Enabled = False
   Set wsODBC = CreateWorkspace("", "tintes", "", dbUseODBC)
   Set conODBC = wsODBC.OpenConnection("connexiosql", dbDriverNoPrompt, , "ODBC;DATABASE=InkmakerDB;UID=sa;PWD=Mak2008;DSN=tintes")
   Set rstfink = conODBC.OpenRecordset("select * from dbo.tblFormula " + vwhere + " order by IDFormula desc", dbOpenSnapshot)
   If rstfink.EOF Then MsgBox "No s'ha trobat la formula a INKMAKER.", vbCritical, "Atenció": GoTo fi
   Set rstf = dbtintes.OpenRecordset("select * from formules")
   rstfink.MoveLast
   rstfink.MoveFirst
   dbtintes.Execute "update  formules set actualitzada=false "
   actualitzar_components_inkmaker
   While Not rstfink.EOF And (vcont >= 0 Or v = "T")
      rstf.FindFirst "codiformula='" + atrim(rstfink!code) + "'"
      If rstf.NoMatch Then
         crearformulanova rstf, rstfink, conODBC, rsttest
           Else
             'comprovar si els valors de la k tinc entrada es correcte... actualitzarla
             comprovarformula rstf, rstfink, conODBC, rsttest
      End If
      rstfink.MoveNext
      vcont = vcont - 1
      ensenyarprogresactualitzacio rstfink
   Wend
   conODBC.Close
   If v = "T" Then borrarformuleseliminadesdeinkmaker rsttest
   
   'a la linia seguent borro tots els components de la formula que ja no existeixen a inkmaker
   dbtintes.Execute "delete * from detallformules where actualitzada=false"
   
   
   dbtintes.Execute "delete * from test_actualitzacioformules where ok=true"
fi:
   Set rstfink = Nothing
   Set rstf = Nothing
   dataformules.Refresh
   ratoli "normal"
   formtintes.Enabled = True
   frameactualitzacio.visible = False
End Sub
Sub test_borrarformuleseliminadesdeinkmaker(rsttest As Recordset)
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from formules where not actualitzada")
   While Not rst.EOF
       guardartest rsttest, atrim(rst!codiformula), "eliminar", atrim(rst!descripcioformula), ""
       rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub borrarformuleseliminadesdeinkmaker(rsttest As Recordset)
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from formules where not actualitzada")
   While Not rst.EOF
       guardartest rsttest, atrim(rst!codiformula), "eliminar", atrim(rst!descripcioformula), ""
       If rsttest!ok Then
          dbtintes.Execute "delete * from detallformules where idformula=" + atrim(rst!idformula)
          rst.Delete
       End If
       rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub ensenyarprogresactualitzacio(rstfink As Recordset)
   frameactualitzacio.visible = True
   msgactualitzacio = "Formula: " + atrim(rstfink.AbsolutePosition) + " de " + atrim(rstfink.RecordCount)
   If rstfink.EOF Then frameactualitzacio.visible = False
   DoEvents
End Sub
Sub crearliniablanc(rsttest As Recordset)
     rsttest.AddNew
     rsttest!formula = " "
     rsttest!accio = " "
     rsttest!valorantic = ""
     rsttest!valornou = ""
     rsttest.Update
End Sub
Sub guardartest(rsttest As Recordset, codiformula As String, accio As String, valorantic As String, valornou As String)
     If rsttest.EOF Then crearliniablanc rsttest
     rsttest.FindFirst "formula='" + codiformula + "' and accio='" + accio + "' and valorantic='" + valorantic + "'"
     If Not rsttest.NoMatch Then
        If rsttest!valorantic <> valorantic Then GoTo cont
        If rsttest!valornou <> valornou Then GoTo cont
        Exit Sub
     End If
cont:
     rsttest.AddNew
     rsttest!formula = codiformula
     rsttest!accio = accio
     rsttest!valorantic = valorantic
     rsttest!valornou = Mid(valornou, 1, rsttest.Fields("valornou").Size)
     rsttest!ok = gravant
     rsttest.Update
     rsttest.MoveLast
End Sub
Sub test_comprovarformula(rstf As Recordset, rstink As Recordset, conODBC As DAO.Connection, rsttest As Recordset)
   Dim lagran As Long
   Dim rst As Recordset
   Dim rstdetall As Recordset
   Dim hiharf As Boolean
   lagran = rstf!idformula
   rstf.Edit
   rstf!actualitzada = True
   rstf.Update
   If rstf!descripcioformula <> rstink!Description Then
        guardartest rsttest, rstf!codiformula, "mod_descripcio", rstf!descripcioformula, rstink!Description
   End If
   If rstf!series <> rstink!series Then
        guardartest rsttest, rstf!codiformula, "mod_series", rstf!series, rstink!series
   End If
   If atrim(rstf!datacreacio) <> Format(rstink!creationdateandtime, "dd/mm/yyyy") Then
        guardartest rsttest, rstf!codiformula, "mod_data", atrim(rstf!datacreacio), Format(rstink!creationdateandtime, "dd/mm/yy")
   End If
   If rstf!notes <> atrim(rstink!notes) Then
        guardartest rsttest, rstf!codiformula, "mod_notes", atrim(rstf!notes), atrim(rstink!notes)
   End If
   'Set rst = conODBC.OpenRecordset("SELECT Code, Description, DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente WHERE (((dbo.tblFormula.Code)=[Formula que vols buscar]));")
   Set rst = conODBC.OpenRecordset("SELECT Code, dbo.tblformula.Description,IdComponente,DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente where dbo.tblFormulaDetail.formulation=0 and dbo.tblformula.code='" + atrim(rstink!code) + "'")
   Set rstdetall = dbtintes.OpenRecordset("select * from detallformules where idformula=" + atrim(lagran))
   While Not rst.EOF
      'crearelcomponentsical rst, conODBC
      rstdetall.FindFirst "idcomponente=" + atrim(cadbl(rst!idcomponente))
      If InStr(1, atrim(rst!DescComponente), " RF") Then hiharf = True
      If rstdetall.NoMatch Then
         'component nou a la formula
            guardartest rsttest, rstf!codiformula, "mod_componentnou", atrim(rst!DescComponente), ""
          Else
             'comparar percentatge
             If Redondejar(rstdetall![%decomponent], 0) <> Redondejar(cadbl(rst![%decomponent]), 0) Then
                guardartest rsttest, rstf!codiformula, "mod_component", atrim(Redondejar(rstdetall![%decomponent], 0)) + "%  " + atrim(rst!DescComponente), atrim(Redondejar(rst![%decomponent], 0)) + "%  " + atrim(rst!DescComponente)
             End If
      End If
      rst.MoveNext
   Wend
   'comprovo que totes les tintes rf tinguin el component rf
   If Mid(atrim(rstf!codiformula), Len(atrim(rstf!codiformula)) - 1, 2) = "RF" And Not hiharf Then
      guardartest rsttest, rstf!codiformula, "er_formularf", atrim(rstink!Description), ""
   End If
End Sub
Sub comprovarformula(rstf As Recordset, rstink As Recordset, conODBC As DAO.Connection, rsttest As Recordset)
   Dim lagran As Long
   Dim rst As Recordset
   Dim hiharf As Boolean
   lagran = rstf!idformula
   rstf.Edit
   rstf!actualitzada = True
   If rstf!descripcioformula <> rstink!Description Then
        guardartest rsttest, rstf!codiformula, "mod_descripcio", rstf!descripcioformula, rstink!Description
        If rsttest!ok Then rstf!descripcioformula = rstink!Description
   End If
   If rstf!series <> rstink!series Then
        guardartest rsttest, rstf!codiformula, "mod_series", rstf!series, rstink!series
        If rsttest!ok Then rstf!series = rstink!series
   End If
   If atrim(rstf!datacreacio) <> Format(rstink!creationdateandtime, "dd/mm/yyyy") Then
        guardartest rsttest, rstf!codiformula, "mod_data", atrim(rstf!datacreacio), Format(rstink!creationdateandtime, "dd/mm/yy")
        If rsttest!ok Then rstf!datacreacio = Format(rstink!creationdateandtime, "dd/mm/yy")
   End If
   If rstf!notes <> atrim(rstink!notes) Then
        guardartest rsttest, rstf!codiformula, "mod_notes", atrim(rstf!notes), atrim(rstink!notes)
        If rsttest!ok Then rstf!notes = atrim(rstink!notes)
   End If
   If rstf!descripcioformula = "" Then rstf!descripcioformula = " "
   rstf.Update
   
   
   'Set rst = conODBC.OpenRecordset("SELECT Code, Description, DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente WHERE (((dbo.tblFormula.Code)=[Formula que vols buscar]));")
   Set rst = conODBC.OpenRecordset("SELECT Code, dbo.tblformula.Description,IdComponente,DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente where dbo.tblFormulaDetail.formulation=0 and dbo.tblformula.code='" + atrim(rstink!code) + "'")
  ' dbtintes.Execute "delete * from detallformules where idformula=" + atrim(lagran)
   Set rstdetall = dbtintes.OpenRecordset("select * from detallformules where idformula=" + atrim(lagran))
   dbtintes.Execute "update detallformules set actualitzada=false where idformula=" + atrim(lagran)
   While Not rst.EOF
      crearelcomponentsical rst, conODBC
      rstdetall.FindFirst "idcomponente=" + atrim(cadbl(rst!idcomponente))
      If InStr(1, atrim(rst!DescComponente), " RF") Then hiharf = True
      If rstdetall.NoMatch Then
         'component nou a la formula
            guardartest rsttest, rstf!codiformula, "mod_componentnou", atrim(rst!DescComponente) + "   " + atrim(rst![%decomponent]) + "%", ""
          Else
             'comparar percentatge
             If Redondejar(rstdetall![%decomponent], 4) <> Redondejar(rst![%decomponent], 4) Then
                guardartest rsttest, rstf!codiformula, "mod_component", atrim(Redondejar(rstdetall![%decomponent], 1)) + "%  " + atrim(rst!DescComponente), atrim(Redondejar(rst![%decomponent], 1)) + "%  " + atrim(rst!DescComponente)
             End If
             rstdetall.Edit
             rstdetall.actualitzada = True
             rstdetall.Update
      End If
      If rsttest!ok Then
         If rsttest!accio = "mod_component" Or Not rstdetall.NoMatch Then rstdetall.Delete
         dbtintes.Execute "insert into detallformules (idformula,idcomponente,[%decomponent],actualitzada) values (" + atrim(lagran) + ",'" + treure_apostruf(rst![idcomponente]) + "'," + passaradecimalpunt(rst![%decomponent]) + ",true)"
      End If
      rst.MoveNext
   Wend
   'comprovo que totes les tintes rf tinguin el component rf
   If Mid(atrim(rstf!codiformula), Len(atrim(rstf!codiformula)) - 1, 2) = "RF" And Not hiharf Then
      guardartest rsttest, rstf!codiformula, "er_formularf", atrim(rstink!Description), ""
   End If
   
End Sub
Sub test_crearformulanova(rstf As Recordset, rstink As Recordset, conODBC As DAO.Connection, rsttest As Recordset)
   Dim lagran As Long
   Dim rst As Recordset
   guardartest rsttest, atrim(rstink!code), "nova", atrim(rstink!Description), ""
   
   'Set rst = conODBC.OpenRecordset("SELECT Code, Description, DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente WHERE (((dbo.tblFormula.Code)=[Formula que vols buscar]));")
   
   
End Sub

Sub crearformulanova(rstf As Recordset, rstink As Recordset, conODBC As DAO.Connection, rsttest As Recordset)
   Dim lagran As Long
   Dim rst As Recordset
   guardartest rsttest, atrim(rstink!code), "nova", atrim(rstink!Description), ""
   If Not rsttest!ok Then Exit Sub
   Set rst = dbtintes.OpenRecordset("select max(idformula) as lagran from formules ")
   If rst.EOF Then
      lagran = 0
    Else: lagran = cadbl(rst!lagran)
   End If
   lagran = lagran + 1
   rstf.AddNew
   rstf!idformula = lagran
   rstf!codiformula = rstink!code
   rstf!descripcioformula = rstink!Description
   rstf!series = rstink!series
   rstf!datacreacio = Format(rstink!creationdateandtime, "dd/mm/yy")
   rstf!notes = atrim(rstink!notes)
   rstf!actualitzada = True
   rstf.Update
   'Set rst = conODBC.OpenRecordset("SELECT Code, Description, DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente WHERE (((dbo.tblFormula.Code)=[Formula que vols buscar]));")
   Set rst = conODBC.OpenRecordset("SELECT Code, dbo.tblformula.Description,IdComponente,DescComponente, [Quantity]/10 AS [%decomponent] FROM (dbo.tblFormula INNER JOIN dbo.tblFormulaDetail ON dbo.tblFormula.IDFormula = dbo.tblFormulaDetail.IDFormula) INNER JOIN dbo.tblComponenti ON dbo.tblFormulaDetail.IDComponent = dbo.tblComponenti.IdComponente where dbo.tblFormulaDetail.formulation=0 and dbo.tblformula.code='" + atrim(rstink!code) + "'")
   While Not rst.EOF
      crearelcomponentsical rst, conODBC
      'poso el camp d'actualitzada a True perquè sino al final elimina tot lo que no està actualitzat
      dbtintes.Execute "insert into detallformules (idformula,idcomponente,[%decomponent],actualitzada) values (" + atrim(lagran) + ",'" + treure_apostruf(rst![idcomponente]) + "'," + passaradecimalpunt(rst![%decomponent]) + ",true)"
      rst.MoveNext
   Wend
   
End Sub
Sub crearelcomponentsical(rst As Recordset, conODBC As DAO.Connection)
    Dim rstc As Recordset
    Dim rsti As Recordset
    Set rstc = dbtintes.OpenRecordset("select * from componentsbase where idcomponent=" + atrim(rst!idcomponente))
    If rstc.EOF Then
        'com que no existeix
       'crear el component
       
       Set rsti = conODBC.OpenRecordset("select * from dbo.tblcomponenti where idcomponente=" + atrim(cadbl(rst!idcomponente)))
       If Not rsti.EOF Then
        MsgBox "El Component " + atrim(rsti!DescComponente) + " no existeix a la taula de components d'INPLACSA, procedirem a crear-lo...", vbInformation, "Atenció"
        rstc.AddNew
        rstc!idcomponent = rsti!idcomponente
        rstc!codicomponent = atrim(rsti!codcomponente)
        rstc!nomcomponent = treure_apostruf(atrim(rsti!DescComponente))
        rstc.Update
       End If
    End If
    Set rstc = Nothing
    Set rsti = Nothing
End Sub



Private Sub Command13_Click()
   If Not reixacomandes.RowIsVisible(reixacomandes.Row) Then Exit Sub
   comandapreparada.Show 1
   possarcolorbotócomandaacabadaicombinacio reixacomandes.TextMatrix(reixacomandes.Row, 0)
End Sub

Private Sub Command14_Click()
   historialdimpresiodunacomanda
End Sub

Private Sub Command15_Click()
  Dim rstf As Recordset
 Load formseleccio
  formseleccio.width = 8800
  formseleccio.caption = "Escullir Nom de tinta intern."
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "SELECT distinct nominplacsa from tintes WHERE nominplacsa<>'';"
  formseleccio.cmissatge.tag = "0"
  formseleccio.Command3.tag = "filtre"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(0).width = 7500
  'formseleccio.DBGrid2.Columns(1).width = 4000
  'formseleccio.DBGrid2.Columns(2).width = 1700
  formseleccio.Show 1
  If seleccioret = 1 Then
   If formseleccio.Data1.Recordset.EOF Then GoTo fi
   cnominplacsa = atrim(formseleccio.Data1.Recordset!nominplacsa)
  End If
fi:
  Unload formseleccio
End Sub
Sub ensenyar_llegenda()
  Dim msg As String
   msg = "Numero treball:" + Chr(10)
 msg = msg + "   Groc: dos o mes numeros de treball iguals" + Chr(10)
 msg = msg + "   Vermell / negatiu: comanda no activada" + Chr(10)
 msg = msg + "   blau: comanda no activada pero diseny fotograbador confirmat" + Chr(10)
' msg = msg + Chr(10)
 msg = msg + "Gestionat:" + Chr(10)
 msg = msg + "   N (negra) : comanda no revisada i no preparada" + Chr(10)
 msg = msg + "   S (verd) : comanda revisada i NO preparada" + Chr(10)
 msg = msg + "   P (marro) comanda revisada i tintes preparades a palet P" + Chr(10)
 msg = msg + "   M (groc) comanda revisada i tintes a maquina " + Chr(10)
 msg = msg + "   F (negre)  comanda revisada, falta formular pantones" + Chr(10)
 msg = msg + "   C ( vermell) comanda revisada pendent de compra"
' msg = msg + Chr(10)
 msg = msg + "Nova/Repetida/Modificada" + Chr(10)
' msg = msg + Chr(10)
 msg = msg + "Tipus impresio -PET el material de la comanda es PET oju amb el tipus de tinta." + Chr(10)
 msg = msg + "         +DOY (Ès una bossa DOYPACK)" + Chr(10)
 msg = msg + "         +ABL comanda laminada amb material ALOX" + Chr(10)
 msg = msg + "         (Vermell) Canvis de material Vs la impresió anterior." + Chr(10)
 msg = msg + "Ès a muntadora:" + Chr(10)
 msg = msg + "    S,N,R,Ra (Si,No,Reclamada,Reclamada i reactivada) " + Chr(10)
 msg = msg + "         casella vermella i el simbol ! darrera, comanda amb Standby." + Chr(10)
 msg = msg + "         casella verda comanda Reclamada però reactivada." + Chr(10)
 msg = msg + "Comanda:" + Chr(10)
 msg = msg + "   Si la comanda te Call-off  estarà de color blau clar." + Chr(10)
 msg = msg + "   Si la comanda es importancia 4 estarà de color Fucsia." + Chr(10)
 msg = msg + "CdL (Codi de linia):" + Chr(10)
 msg = msg + "       Si es verd, la comanda pot sortir." + vbNewLine
 msg = msg + "       Si té el simbol de negatiu(-) davant i està vermell es que el número es d'una altra versió però no la seva." + Chr(10)
 msg = msg + "       Si ès negatiu amb un doble clic a sobre es pot ACCEPTAR aquest codi o ELIMINARLO." + Chr(10)
 msg = msg + "       Si es TARONJA es que hi ha altres comandes entrades amb la mateixa CdL i alguna programada." + Chr(10)
 msg = msg + "Doble clic a Treball: Veure data entrega clixes." + Chr(10)
 msg = msg + "Doble clic a Secció: Veure data prevista d'entrega comanda." + Chr(10)


Load avis
 avis.missatge = msg
 'MsgBox msg
 avis.Show
 avis.Left = (Screen.width / 2) - (avis.width / 2)
 avis.Top = (Screen.Height / 2) - (avis.Height / 2)
End Sub

Private Sub Command16_Click()
  combosionookcarta.Text = ""
  For i = 0 To filtretinta.Count - 1
     filtretinta(i).Text = ""
  Next i
  colocarfiltretinta
  tintes.RecordSource = "tintes_tot"
  tintes.Refresh
End Sub

Private Sub Command17_Click()
  Dim vnumtreball As String
  Dim vnummodificacio As String
  Dim vnumc As Double
  If reixacomandes.Row > 0 Then vnumc = reixacomandes.TextMatrix(reixacomandes.Row, col)
  vnumc = cadbl(InputBox("Entra la comanda o el treball que vols veure.", "Modificar tintes", vnumc))
  If vnumc = 0 Then Exit Sub
  If vnumc < 100000 Then
    vnummodificacio = cadbl(InputBox("Entra la versió del treball que vols veure.", "Modificar tintes", 1))
    If vnummodificacio = 0 Then Exit Sub
    vnumtreball = vnumc
    vnumc = 0
  End If
  If vnumc > 0 Then
    Set rst = dbcomandes.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vnumc))
    If rst.EOF Then MsgBox "Comanda no trobada": Exit Sub
    vnumtreball = cadbl(rst!numtreball)
    vnummodificacio = cadbl(rst!numordremodificacio)
  End If
  
  Shell "\\serverprodu\dades\progcomandes\aplicacio\clixesnous.exe " + "comandes.ini ''  modificartintes " + vnumtreball + " " + vnummodificacio + " ''", vbNormalFocus
End Sub

Private Sub Command18_Click()
 Dim i As Integer
  Command56_Click
  For i = 0 To filtre.Count - 1
     If filtre(i).Text = "NºTreball" Then filtre(i).SetFocus: filtre(i).Text = "TreballsBlaus": filtre_LostFocus i
  Next i
 
End Sub

Private Sub Command19_Click()
   Shell "\\serverprodu\dades\progcomandes\aplicacio\baixesimpresoramaquina.exe ORDREIMPRESSIO '9' TINTES", vbNormalFocus
End Sub
Sub tirarllistatdiferencies()
    Dim rst As Recordset
    Set rsttest = dbtintes.OpenRecordset("select * from test_actualitzacioformules where formula<>' '")
    If rsttest.EOF Then MsgBox "No hi han diferencies entre Inkmaker i la Base de dades d'Inplacsa", vbInformation, "Diferències": Exit Sub
    llistat.ReportFileName = llegir_ini("General", "rutallistats", "comandes.ini") + "llistatdiferenciesformules.rpt"
    llistat.Destination = crptToWindow
    llistat.CopiesToPrinter = 1
    llistat.DataFiles(0) = camitintes
    'llistat.SelectionFormula = "{TEMP_TINTES_DISPONIBLES.coditinta}<>'' and not {TEMP_TINTES_DISPONIBLES.comprovat}"
    llistat.DiscardSavedData = True
    For i = 0 To 20
     llistat.Formulas(i) = ""
    Next i
    'llistat.Formulas(1) = "dataimpresio='" + Format(Now, "long date", vbMonday) + " " + Format(Now, "hh:nn") + "'"
    'llistat.Formulas(0) = "comanda='" + Format(numcomanda, "#,##0") + "'"
    'llistat.Formulas(2) = "nomdelclient='" + nomdelclient + "'"
    
    DoEvents
    'If existeix("c:\ordprog.ini") Then llistat.Destination = crptToWindow
    llistat.Action = 1
End Sub
Private Sub Command2_Click()
  filtrartintes
End Sub
Sub filtrartintes()
  Dim vprimeraparaula As String
  Dim vsegonaparaula As String
  Dim vterceraparaula As String
  Dim vcampfiltrar As String
  Dim v As String
  Dim vokcarta As String
  vokcarta = IIf(combosionookcarta = "Sí", "okcarta=true and ", IIf(combosionookcarta = "No", "okcarta=false and ", ""))
  vcampfiltrar = etfiltrarper.tag
  v = treure_apostruf(buscador) + " "
  vprimeraparaula = atrim(Mid(v, 1, InStr(1, v, " ")))
  vsegonaparaula = atrim(Mid(v, InStr(1, v, vprimeraparaula) + Len(vprimeraparaula), InStr(InStr(1, v, vprimeraparaula) + Len(vprimeraparaula), v, " ")))
  vterceraparaula = atrim(Mid(v, InStr(1, v, vsegonaparaula) + Len(vsegonaparaula)))
  tintes.RecordSource = "select * from tintes_tot where " + vokcarta + vcampfiltrar + " like '*" + vprimeraparaula + "*' " + IIf(vsegonaparaula <> "", " and " + vcampfiltrar + " like '*" + vsegonaparaula + "*'", "") + IIf(vterceraparaula <> "", " and " + vcampfiltrar + " like '*" + vterceraparaula + "*'", "") + " order by " + vcampfiltrar
  tintes.Refresh
End Sub


Sub comprovar_lesllaunesdelsdosificadors()
   Dim rst As Recordset
   Dim rstc As Recordset
   Dim vllauna As String
   vmsgcontroldosificadors = ""
   Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
   Set rst = dbtintes.OpenRecordset("SELECT * fROM Componentsbase")
   While Not rst.EOF
      Set rstc = dbtintes.OpenRecordset("select * from detallnumeroslotsbase where idcomponent=" + atrim(rst!idcomponent) + " order by data desc")
      If Not rstc.EOF Then
         vllauna = atrim(rstc!numerodelot)
         rstc.MoveNext
         If Not rstc.EOF Then comprovar_llaunacoincideixambdosificador atrim(rstc!numerodelot), vllauna, rst!nomcomponent, dbtintes
      End If
      rst.MoveNext
   Wend
   Set rst = Nothing
   If vmsgcontroldosificadors <> "" Then
      vmsgcontroldosificadors = vmsgcontroldosificadors + "  ARREGLA ELS PROBLEMES I TORNA A VERIFICAR-HO"
      MsgBox vmsgcontroldosificadors, vbCritical, "Error al dosificador"
      enviaremailgeneric "controlestoctintes", "ERRORS LLAUNES AL FER CANVI DE LOTS ALS DOSIFICADORS " + Format(Now, "dd/mm/yy"), vmsgcontroldosificadors
        Else:
           MsgBox "Tot correcte en els dosificadors."
           enviaremailgeneric "controlestoctintes", "OK AL FER EL CANVI DE LOTS ALS DOSIFICADORS", "No hi ha hagut cap error al fer la verificació de CANVI DE LOTS AL DOSIFICADOR."
  End If
  ' Set dbtintes = Nothing
End Sub
Sub comprovar_llaunacoincideixambdosificador(vllaunavella As String, vllaunanova As String, vdescripciodelcomponent As String, dbtintes As Database)
   Dim rstlln As Recordset
   Dim rstllv As Recordset
   If Mid(vllaunavella + " ", 1, 1) <> "A" Or Mid(vllaunanova + " ", 1, 1) <> "A" Then Exit Sub
   Set rstllv = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, tintes.idfamilia, tintes.idsubfamilia, tintes.idfamcolor, tintes.idsubfamcolor FROM Llaunes INNER JOIN tintes ON Llaunes.idtinta = tintes.idtinta where numllauna='" + vllaunavella + "'")
   Set rstlln = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, tintes.idfamilia, tintes.idsubfamilia, tintes.idfamcolor, tintes.idsubfamcolor FROM Llaunes INNER JOIN tintes ON Llaunes.idtinta = tintes.idtinta where numllauna='" + vllaunanova + "'")
   If rstlln.EOF Or rstllv.EOF Then vmsgcontroldosificadors = vmsgcontroldosificadors + "Error de llauna al dosificador de " + atrim(vdescripciodelcomponent) + vbNewLine + "        La llauna nova " + atrim(vllaunanova) + " o l'anterior " + atrim(vllaunavella) + " no existeixen a la base de dades." + vbNewLine + vbNewLine: GoTo fi
   If rstllv!idfamilia <> rstlln!idfamilia Or rstllv!idsubfamilia <> rstlln!idsubfamilia Or rstllv!idfamcolor <> rstlln!idfamcolor Or rstllv!idsubfamcolor <> rstlln!idsubfamcolor Then
        vmsgcontroldosificadors = vmsgcontroldosificadors + "Error de llauna al dosificador de " + atrim(vdescripciodelcomponent) + vbNewLine + "      La llauna " + atrim(vllaunanova) + " del dosificador " + atrim(vdescripciodelcomponent) + " no correspont amb l'anterior " + atrim(vllaunavella) + vbNewLine + vbNewLine
   End If
fi:
End Sub

Private Sub Command21_Click()
  If tintes.Recordset.EditMode <> 0 Then MsgBox "Primer guarda els canvis de la tinta abans d'entrar.", vbCritical, "Atenció": Exit Sub
  Unload formtintessemblants
  Load formtintessemblants
  formtintessemblants.tag = atrim(tintes.Recordset!codi)
  formtintessemblants.Show 1
  
'   filtrarimportaciocolorstreballs False, True
End Sub

Private Sub Command22_Click()
'   filtrarimportaciocolorstreballs True, True
  Dim vllauna  As String
  Dim rstll As Recordset
  vllauna = InputBox("Entra o escaneja la llauna on vols afegir_hi lots", "Afegir lots a una llauna")
  Set rstll = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(vllauna) + "'")
  If rstll.EOF Then MsgBox "Aquesta llauna no existeix.", vbCritical, "Error": GoTo fi
  If Not rstll!activa Then MsgBox "Aquesta llauna no està activa.", vbCritical, "Error": GoTo fi
  colocarsealallauna atrim(rstll!numllauna)
  pestanyes.Tab = 1
  demanarmeslots
fi:
  Set rstll = Nothing
End Sub

Private Sub Command23_Click()
  fcoditintallauna = ""
  fnumllauna = ""
  fdesctintallauna = ""
  fformulallauna = ""
  fsituaciollauna = ""
  checkactives.Value = 1
  checkimpresores.Value = 0
End Sub

Private Sub Command24_Click()
  barrejardosllaunes
End Sub

Private Sub Command25_Click()
  If datadellaunes.Recordset.EOF Then Exit Sub
  colocarsealallauna atrim(datadellaunes.Recordset!numllauna)
End Sub
Sub colocarsealallauna(vnumllauna As String)
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("select idtinta from llaunes where numllauna='" + atrim(vnumllauna) + "'")
  If rst.EOF Then Exit Sub
  tintes.RecordSource = "select * from tintes_tot where idtinta=" + atrim(rst!idtinta)
  tintes.Refresh
  If tintes.Recordset.EOF Then
     tintes.RecordSource = "select * from tintes_tot "
     tintes.Refresh
  End If
  datallaunes.Recordset.FindFirst "numllauna='" + atrim(vnumllauna) + "'"
  If datallaunes.Recordset.NoMatch Then MsgBox "No s'ha trobat aquesta llauna.": Exit Sub
  pestanyes.Tab = 0
End Sub
Private Sub Command26_Click()
   Dim v As String
   v = "1"
   While v <> ""
       Command23_Click
       filtrarllaunes
      ferelretornmanual v
   Wend
End Sub
Sub ferelretornmanual(ByRef v As String)
  Dim nllauna As String
  Dim pesnet As Double
  Dim codifamcolor As Double
  Dim vX As String
  Dim rst As Recordset
  
  nllauna = InputBox("Entra el numero de la llauna que vols retornar" + Chr(10) + "No escriguis res per deixar de retornar llaunes", "Num. Llauna")
  If nllauna = "" Then v = "": Exit Sub
  
  Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, recuperadorsdecontenidors.nomcomercial FROM Llaunes LEFT JOIN recuperadorsdecontenidors ON Llaunes.idproveidorrecuperador = recuperadorsdecontenidors.Id WHERE (((Llaunes.numllauna)='" + nllauna + "'));")
  If rst.EOF Then MsgBox "llauna no trobada": Exit Sub
  If atrim(rst!nomcomercial) = "INPLACSA" Then
      If noespotcanviarelrecuperador(nllauna) Then GoTo fi
      ferelretornAREC nllauna ': GoTo fi
  End If
  datadellaunes.Recordset.FindFirst "numllauna='" + atrim(nllauna) + "'"
  If datadellaunes.Recordset.NoMatch Then MsgBox "llauna no trobada": Exit Sub
  vX = InputBox("Entra el pes de la tinta que queda a la llauna " + atrim(datadellaunes.Recordset!numllauna), "Retorn de tinta")
  If StrPtr(vX) = 0 Then Exit Sub
  pesnet = cadbl(passaradecimal(vX))
  If pesnet >= 0 Then
        ferelretorndetinta atrim(datadellaunes.Recordset!numllauna), cadbl(pesnet), True
        If pesnet = 0 Then
           If cadbl(datadellaunes.Recordset!capacitat) < 180 Then
              If Not hihaunacomprapendent(datadellaunes.Recordset!codi) Then fercompraEstocminim datadellaunes.Recordset
               Else: mirarsidemanarpercomprar datadellaunes.Recordset
           End If
        End If
  End If
fi:
  Set rst = Nothing
End Sub
Function hihaunacomprapendent(vcoditinta As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from comprespendents where not demanat and coditinta='" + atrim(vcoditinta) + "'")
   If Not rst.EOF Then
       hihaunacomprapendent = True
      Else:
         Set rst = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, capcalera.data, capcalera.dataentrega, capcalera.nomprovcomercial, liniescompra.codimaterial, liniescompra.nommaterial, liniescompra.quantitatkg, liniescompra.kgentregats,liniescompra.idliniacompra,liniescompra.totentregat FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE codimaterial=" + atrim(vcoditinta) + " and (((capcalera.numcomanda)>0) AND ((liniescompra.tipusmaterialcomprat)='T')) and (data>dateadd('d',-100,now)) AND liniescompra.totentregat=False order by dataentrega;") 'and not totentregat
         If Not rst.EOF Then hihaunacomprapendent = True
   End If
   Set rst = Nothing
End Function
Function buscar_kg_per_bido(vrst As Recordset) As Double
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from tipusbidons where id=" + atrim(cadbl(vrst!id_bido)))
   If Not rst.EOF Then buscar_kg_per_bido = cadbl(rst!capacitat)
   Set rst = Nothing
End Function

Function fercompraEstocminim(rst As Recordset) As Boolean
   Dim vq As Double
   Dim vobs As String
   Dim rst2 As Recordset
   Dim vllaunesdetinta As Double
   Dim vllaunesdetintaquefalta As Double
   Dim vkgperllauna As Double
   
   If rst.EOF Then Exit Function
   Set rst2 = dbtintes.OpenRecordset("select * from estocsminims where codi='" + atrim(rst!codi) + "'")
   If rst2.EOF Then GoTo fi
   vkgperllauna = buscar_kg_per_bido(rst)
   fercompraEstocminim = True 'si arriba aqui es que hi ha estoc posat i retorno SI per saltar la compra al sortir
   vllaunesdetinta = calcular_LLAUNES_senseSITigualaDOS(rst!idtinta)
   If rst2!estocminim >= vllaunesdetinta Then
        'fer la compra de estocdesitjat-vkgdetinta
        'de la referencia rst!referencia
        vllaunesdetintaquefalta = rst2!estocdesitjat - vllaunesdetinta
        If MsgBox("S'ha de fer compra per complir l'estoc minim." + vbNewLine + "VOLS FER LA COMPRA DE " + atrim(vllaunesdetintaquefalta) + " Llaunes (" + atrim(vllaunesdetintaquefalta * vkgperllauna) + "Kg)?", vbExclamation + vbDefaultButton2 + vbYesNo, "COMPRE PER ESTOC MINIM") = vbYes Then
            vobs = "COMPRA FETA AUTOMATICAMENT PER ESTOC MINIM."
            dbtintes.Execute "insert into comprespendents (descripcio,referencia,coditinta,quantitat,observacio) values ('" + treure_apostrof(rst!descripcio) + "','" + atrim(rst!referencia) + "','" + atrim(cadbl(rst!codi)) + "'," + atrim(vllaunesdetintaquefalta * vkgperllauna) + ",'" + treure_apostruf(vobs) + "')"
            enviarlescompresaldepartamentdecompres
        End If
   End If
fi:
   Set rst2 = Nothing
End Function
Function calcular_LLAUNES_senseSITigualaDOS(vidtinta As String) As Double
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.idtinta, count(Llaunes.capacitatactual) AS totallaunes From Llaunes Where situacio<>'DOS' AND llaunes.capacitatactual>0 and Llaunes.activa = True and Llaunes.idtinta=" + atrim(vidtinta) + " GROUP BY Llaunes.idtinta;")
   'Set rst2 = dbtintes.OpenRecordset("SELECT count(Llaunes.idtinta) as Tllaunes from Llaunes Where llaunes.capacitatactual>0.9 and Llaunes.activa = True and Llaunes.idtinta=" + atrim(vidtinta) + " GROUP BY Llaunes.idtinta;")
   If Not rst.EOF Then calcular_LLAUNES_senseSITigualaDOS = Redondejar(rst!totallaunes, 1)
   Set rst = Nothing
End Function
Sub ferelretornAREC(vllauna As String)
  Dim vX As String
  Dim vidproveidorrecuperador As Long
  vX = InputBox("Entra el pes de la tinta que hi ha al contenidor " + vllauna, "Retorn de tinta")
  If StrPtr(vX) = 0 Then GoTo fi
  ferelretorndetinta vllauna, cadbl(vX), True
  escullir_proveidorrecuperador vidproveidorrecuperador
  If vidproveidorrecuperador = 0 Then
     If MsgBox("No has escullit cap RECUPERADOR de contenidors, VOLS BORRAR AQUEST QUE HI HA ASSIGNAT?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo fi
  End If
  dbtintes.Execute "update llaunes set idproveidorrecuperador=" + atrim(cadbl(vidproveidorrecuperador)) + " where numllauna='" + vllauna + "'"
  'canvi de situacio
  dbtintes.Execute "update  llaunes set situacio='MAG' where numllauna='" + vllauna + "'"
  dbtintes.Execute "insert into historialsituacions (data,situacio,numllauna) values (now,'SALA','" + atrim(vllauna) + "')"
  
fi:
  
End Sub
Sub mirarsidemanarpercomprar(vrst As Recordset)
   Dim vq As Double
   Dim vobs As String
   Dim vdataexecucio As String
   If cadbl(vrst!capacitat) < 180 Then Exit Sub
   If MsgBox("Vols comprar mes d'aquesta tinta?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
   vq = cadbl(InputBox("Quantes referencies vols comprar?", "Compres", 1))
   If vq = 0 Then Exit Sub
   vobs = InputBox("Vols escriure una observació per aquesta compra?", "Observació adicinal")
   vdataexecucio = InputBox("SI VOLS ENVIAR LA COMPRA UN ALTRA DIA QUE NO SIGUI INMEDIATAMENT ESCRIU LA DATA." + vbNewLine + "Ex: 25/" + Format(Now, "mm/yy"), "ENVIAMENT DIFERIT DE LA COMPRA")
   If Not IsDate(vdataexecucio) Then
         dbtintes.Execute "insert into comprespendents (descripcio,referencia,coditinta,quantitat,observacio) values ('" + treure_apostrof(vrst!descripcio) + "','" + atrim(vrst!referencia) + "','" + atrim(cadbl(vrst!codi)) + "'," + atrim(vq) + ",'" + treure_apostruf(vobs) + "')"
         enviarlescompresaldepartamentdecompres
           Else: dbtintes.Execute "insert into compresdiferides (dataexecuciocompra,descripcio,referencia,coditinta,quantitat,observacio) values (#" + Format(vdataexecucio, "mm/dd/yy") + "#,'" + treure_apostrof(vrst!descripcio) + "','" + atrim(vrst!referencia) + "','" + atrim(cadbl(vrst!codi)) + "'," + atrim(vq) + ",'" + treure_apostruf(vobs) + "')"
   End If
End Sub
Sub enviar_compres_programades()
   Dim rst As Recordset
   Dim vvalues As String
   Set rst = dbtintes.OpenRecordset("Select * from compresdiferides where format(dataexecuciocompra,'mm/dd/yy')=format(now,'mm/dd/yy')")
   While Not rst.EOF
      vvalues = "('" + treure_apostrof(rst!descripcio) + "','" + atrim(rst!referencia) + "','" + atrim(cadbl(rst!coditinta)) + "'," + atrim(cadbl(rst!quantitat)) + ",'" + treure_apostruf(atrim(rst!observacio)) + "')"
      dbtintes.Execute "insert into comprespendents (descripcio,referencia,coditinta,quantitat,observacio) values " + vvalues
      dbtintes.Execute "update compresdiferides set dataexecuciocompra=null where id=" + atrim(rst!id)
      rst.MoveNext
   Wend
   If vvalues <> "" Then enviarlescompresaldepartamentdecompres 'si hi ha alguna cosa a vvalues es que almenys he afegit una compra i envio a compres
   dbtintes.Execute "delete * from compresdiferides where dataexecuciocompra=null"
   Set rst = Nothing
End Sub
Function eselcolorcorrecte(nllauna As String, codifamcolor As Double) As Boolean
   Dim rstcolor As Recordset
   Set rstcolor = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, familiescolors.codi FROM familiescolors RIGHT JOIN (tintes LEFT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta) ON familiescolors.codi = tintes.idfamcolor where llaunes.numllauna='" + atrim(nllauna) + "';")
   If Not rstcolor.EOF Then
         If rstcolor!codi <> codifamcolor Then MsgBox "El color escullit no coincideix amb el de la llauna": Exit Function
       Else: MsgBox "No he trobat el color de la tinta relacionada amb aquesta llauna": Exit Function
   End If
   eselcolorcorrecte = True
End Function
Function demanacolorllauna(numllauna As String) As Double
  Load formseleccio
  formseleccio.width = formseleccio.width
  formseleccio.caption = "Escull el color de la llauna"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "SELECT codi, descripcio FROM familiescolors order by descripcio"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 1600
  'formseleccio.DBGrid2.Columns(2).Width = 800
  formseleccio.Show 1
  If seleccioret = 1 Then
   demanacolorllauna = formseleccio.Data1.Recordset!codi
  End If
  Unload formseleccio
  
End Function
Function pesbascula() As Double
Static buffer As String
Static nobascula As Boolean
If Not MSComm1.PortOpen Then
  MSComm1.CommPort = 1
 ' 9600 baudios, sin paridad, 7 bits de datos y 1 bit de parada.
  MSComm1.Settings = "9600,n,8,1"
 ' If nummaq = 1 Then MSComm1.Settings = "2400,n,8,1"
 ' Indicar al control que lea todo el búfer al usar Input.
  MSComm1.InputLen = 0
 
  MSComm1.RTSEnable = True 'Por si necesitas habilitar el RTS
 
 'Abrir Puertos
 On Error GoTo nopossarpes
  MSComm1.PortOpen = True
End If
 i = 0
 buffer = buffer & MSComm1.Input
 If Len(buffer) > 20 Then
   If InStr(1, buffer, "-") Then buffer = "0"
   If InStr(1, buffer, Chr$(13)) > 0 Then buffer = Mid(buffer, InStr(1, buffer, "+") + 1, InStr(1, buffer, Chr$(13)))
   'If InStr(1, buffer, ".") > 0 Then buffer = Mid(buffer, 1, InStr(1, buffer, ".") - 1) + "," + Mid(buffer, InStr(1, buffer, ".") + 1)
   pesbascula = buffer
   buffer = ""
 End If
 Exit Function
nopossarpes:
   pesbascula = 0
End Function

Private Sub Command27_Click()
   If datahistoria.Recordset.EOF Then MsgBox "No has sel.leccionat cap historia per eliminar.", vbInformation, "Atenció": Exit Sub
   If datahistoria.Recordset!tipusmoviment = "C" Then MsgBox "No es pot eliminar una carga." + Chr(10) + "Fes un retorn si vols deixar la llauna buida.", vbCritical + vbOKOnly, "Atenció": Exit Sub
   If MsgBox("Eliminar part de l'historia d'una llauna pot suposar canvis en la seva capacitat." + Chr(10) + "Segur que vols continuar?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   datahistoria.Recordset.Delete
   datahistoria.Refresh
   calcularkgdisponiblesllauna datallaunes.Recordset!numllauna
   datallaunes.Recordset.Move 0
   
End Sub

Private Sub Command28_Click()
  If datallaunes.Recordset.EOF Then MsgBox "Primer escull una llauna.", vbInformation, "Atenció": Exit Sub
  'ensenyarlotsbase datallaunes.Recordset!numllauna
  ensenyarlotsbase_totals datallaunes.Recordset!numllauna
  
End Sub

Sub ensenyarlotsbase(ByVal nllauna As String, Optional vnumlotbase As String, Optional noensenyarmissatge As Boolean)
   Dim rsthistoria As Recordset
   Dim rst As Recordset
   Dim vhistoria As String
   Dim idsdecarga As String
   Dim vdata As String
   Dim vdatatrobada As Date
   Dim vnumlotbuscat As String
   Dim vkgnumlot As Double
   vdata = Date
   If datarecarregues.Recordset.RecordCount > 1 Then
      vdata = InputBox("Entra la data de la comanda.", "Buscar data del lot", Date)
      If Not IsDate(vdata) Then MsgBox "Data erronea", vbCritical, "Error": Exit Sub
   End If
   'Set rsthistoria = dbtintes.OpenRecordset("SELECT llaunes.id,Llaunes.numllauna, historiallauna.data, historiallauna.id,historiallauna.tipusmoviment,historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes RIGHT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) RIGHT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) INNER JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent where (numllauna='" + atrim(nllauna) + "')  ORDER BY historiallauna.data DESC;")
   Set rsthistoria = dbtintes.OpenRecordset("SELECT Llaunes.id, Llaunes.numllauna, historiallauna.data, historiallauna.id, historiallauna.numrecarrega,historiallauna.tipusmoviment, historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) LEFT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) LEFT JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent Where (numllauna='" + atrim(nllauna) + "') ORDER BY historiallauna.data DESC,historiallaunalots.idcomponent ,numrecarrega DESC;")
   While Not rsthistoria.EOF
      If DateDiff("d", vdata, rsthistoria!Data) <= 0 And rsthistoria!tipusmoviment = "C" Then
        vnumlotbuscat = ""
        vkgnumlot = 0
        If (vdatatrobada) = "0:00:00" Then vdatatrobada = rsthistoria!Data
        If vdatatrobada <> rsthistoria!Data Then GoTo tornarhi
        vnumlotbuscat = numdelotbaseonumlotdellauna(atrim(rsthistoria!numlotbase))
        calcularkgdisponiblesllauna substituir(atrim(vnumlotbuscat), " (Sense Lot s´ha de buscar manualment la llauna)", ""), vkgnumlot, True
        If rsthistoria!tipusmoviment = "C" Then vhistoria = vhistoria + "(" + atrim(rsthistoria!numrecarrega) + ") " + atrim(rsthistoria!nomcomponent) + "______" + IIf(vkgnumlot > 0, atrim(vkgnumlot) + "Kg", "") + "____ (" + atrim(rsthistoria!numlotbase) + ") " + " ->    " + vnumlotbuscat + "   " + buscaralbproveidor(vnumlotbuscat) + Chr(13) + Chr(10)
        If atrim(rsthistoria!nomcomponent) = "NUMERO DE LOT MANUAL" And atrim(rsthistoria!numlotbase) <> "0" And atrim(rsthistoria!numlotbase) <> "" Then vnumlotbase = atrim(rsthistoria!numlotbase)
        
        'If InStr(1, idsdecarga, atrim(rsthistoria![historiallauna.id])) = 0 Then idsdecarga = idsdecarga + IIf(idsdecarga = "", "", " ,") + atrim(rsthistoria![historiallauna.id])
       End If
       rsthistoria.MoveNext
   Wend
tornarhi:
 
  
  If Not noensenyarmissatge Then
     'MsgBox vhistoria, vbInformation, "Lots Base de la Llauna"
     'Clipboard.Clear
    ' Clipboard.SetText vhistoria
     Shell "notepad.exe", vbNormalFocus
     wait 1
     SendKeys "^V"
     wait 2
    ' Clipboard.Clear
 End If

End Sub
Function buscaralbproveidor(valbprov As String) As String
  Dim rst As Recordset
  Set rst = dbcompres.OpenRecordset("select numalbaraprov,data from albaransbip where numlotproveidor ='" + atrim(valbprov) + "'")
  If Not rst.EOF Then buscaralbproveidor = "Alb.Prov: " + atrim(rst!numalbaraprov) + " " + atrim(rst!Data)
  Set rst = Nothing
End Function

Sub ensenyarlotsbase_totals(ByVal nllauna As String, Optional vnumlotbase As String, Optional noensenyarmissatge As Boolean, Optional noautocarregarlots As Boolean)
   Dim rsthistoria As Recordset
   Dim rst As Recordset
   Dim vnumlot As String
   Dim vlotsillauna As String
   Dim vhistoria As String
   Dim idsdecarga As String
   Dim vnumllaunarelacionada As String
   Dim vdesctintallaunarelacionada As String
   'Set rsthistoria = dbtintes.OpenRecordset("SELECT llaunes.id,Llaunes.numllauna, historiallauna.data, historiallauna.id,historiallauna.tipusmoviment,historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes RIGHT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) RIGHT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) INNER JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent where (numllauna='" + atrim(nllauna) + "')  ORDER BY historiallauna.data DESC;")
   Set rsthistoria = dbtintes.OpenRecordset("SELECT Llaunes.id, Llaunes.numllauna, historiallauna.data, historiallauna.id, historiallauna.numrecarrega,historiallauna.tipusmoviment, historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) LEFT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) LEFT JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent Where (numllauna='" + atrim(nllauna) + "') ORDER BY historiallauna.data DESC,historiallaunalots.idcomponent ,numrecarrega DESC;")
   While Not rsthistoria.EOF
      If InStr(1, vhistoria, atrim(rsthistoria!numlotbase)) > 0 Then GoTo proxim
       vnumlot = numdelotbaseonumlotdellauna(atrim(rsthistoria!numlotbase), True, vnumllaunarelacionada)
       If vnumllaunarelacionada = "" Then
          vdesctintallaunarelacionada = atrim(rsthistoria!nomcomponent)
           Else: vdesctintallaunarelacionada = nomdelatinta(vnumllaunarelacionada)  'buscar desc de la tinta
       End If
       vlotsillauna = "(" + atrim(rsthistoria!numrecarrega) + ") " + vdesctintallaunarelacionada + "__________" + " ->    " + numdelotbaseonumlotdellauna(atrim(rsthistoria!numlotbase), True)
       If noesunloterroni(atrim(rsthistoria!numlotbase), atrim(rsthistoria!numllauna)) Then
            If (rsthistoria!tipusmoviment = "C" Or rsthistoria!tipusmoviment = "K") And InStr(1, atrim(rsthistoria!nomcomponent), "TARONJA REC") = 0 Then
              If InStr(1, vhistoria, vnumlot) = 0 Then vhistoria = vhistoria + vlotsillauna + Chr(13) + Chr(10)
            End If
            If atrim(rsthistoria!nomcomponent) = "NUMERO DE LOT MANUAL" And atrim(rsthistoria!numlotbase) <> "" Then vnumlotbase = atrim(rsthistoria!numlotbase)
            If InStr(1, idsdecarga, atrim(rsthistoria![historiallauna.id])) = 0 Then idsdecarga = idsdecarga + IIf(idsdecarga = "", "", " ,") + atrim(rsthistoria![historiallauna.id])
       End If
proxim:
       rsthistoria.MoveNext
   Wend
tornarhi:
   vhistoria = vhistoria + Chr(13) + Chr(10)
   If idsdecarga = "" Then GoTo fi
   'Set rsthistoria = dbtintes.OpenRecordset("SELECT llaunes.id,Llaunes.numllauna, historiallauna.data, historiallauna.id,historiallauna.tipusmoviment,historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes RIGHT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) RIGHT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) INNER JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent where (historiallauna.idhistoriabarreja in(" + atrim(idsdecarga) + "))  ORDER BY historiallauna.data DESC;")
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.id, Llaunes.numllauna, historiallauna.data, historiallauna.id, historiallauna.tipusmoviment, historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) LEFT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) LEFT JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent Where (historiallauna.idhistoriabarreja in(" + atrim(idsdecarga) + ")) ORDER BY historiallauna.data DESC,historiallaunalots.idcomponent ,numrecarrega DESC;")
   idsdecarga = ""
   While Not rst.EOF
        Set rsthistoria = dbtintes.OpenRecordset("SELECT Llaunes.id, Llaunes.numllauna, historiallauna.data, historiallauna.id, historiallauna.tipusmoviment, historiallauna.numrecarrega,historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) LEFT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) LEFT JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent Where (numllauna='" + atrim(rst!numllauna) + "') ORDER BY historiallauna.data DESC;")
        While Not rsthistoria.EOF
            vnumlot = numdelotbaseonumlotdellauna(atrim(rsthistoria!numlotbase), , vnumllaunarelacionada)
            If vnumllaunarelacionada = "" Then
              vdesctintallaunarelacionada = atrim(rsthistoria!nomcomponent)
                Else: vdesctintallaunarelacionada = nomdelatinta(vnumllaunarelacionada)  'buscar desc de la tinta
            End If
            vlotsillauna = "(" + atrim(rsthistoria!numrecarrega) + ") " + vdesctintallaunarelacionada + "__________" + " ->    " + vnumlot + " (" + atrim(rsthistoria!numllauna) + ")"
            If noesunloterroni(atrim(rsthistoria!numlotbase), atrim(rsthistoria!numllauna)) Then
                    If (rsthistoria!tipusmoviment = "C" Or rsthistoria!tipusmoviment = "K") And InStr(1, atrim(rsthistoria!nomcomponent), "TARONJA REC") = 0 Then vhistoria = vhistoria + vlotsillauna + Chr(13) + Chr(10)
                    If atrim(rsthistoria!nomcomponent) = "NUMERO DE LOT MANUAL" And atrim(rsthistoria!numlotbase) <> "" Then vnumlotbase = atrim(rsthistoria!numlotbase)
                    If InStr(1, idsdecarga, atrim(rsthistoria![historiallauna.id])) = 0 Then idsdecarga = idsdecarga + IIf(idsdecarga = "", "", " ,") + atrim(rsthistoria![historiallauna.id])
            End If
            rsthistoria.MoveNext
        Wend
        rst.MoveNext
   Wend
   If idsdecarga <> "" Then GoTo tornarhi
fi:
  
 
  If Not noensenyarmissatge Then
      'MsgBox vhistoria, vbInformation, "Lots Base de la Llauna"
      vhistoria = nomllaunaidata(nllauna) + vbNewLine + "Numeros de lots:" + vbNewLine + vhistoria
      Clipboard.Clear
      Clipboard.SetText vhistoria
    
      Shell "notepad.exe", vbNormalFocus
    'Send the keys CTRL+V To Notepad (i.e the window that has focus)
      SendKeys "^V"
      wait 2
      Clipboard.Clear
 End If

End Sub
Function nomllaunaidata(nllauna As String) As String
   nomllaunaidata = "Número de llauna:  " + atrim(nllauna) + vbNewLine
   nomllaunaidata = nomllaunaidata + "Nom de la tinta: " + nomdelatinta(nllauna) + vbNewLine
   nomllaunaidata = nomllaunaidata + "Data de creació de la llauna: " + datacreaciollauna(nllauna) + vbNewLine
End Function
Function datacreaciollauna(vllauna As String) As String
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, historiallauna.data FROM historiallauna LEFT JOIN Llaunes ON historiallauna.idnumllauna = Llaunes.id Where (((Llaunes.numllauna) = '" + vllauna + "') And ((historiallauna.tipusmoviment) = 'C')) ORDER BY historiallauna.data;")
   If Not rst.EOF Then datacreaciollauna = atrim(rst!Data)
   Set rst = Nothing
End Function
Function nomdelatinta(vllauna As String) As String
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from dadesllaunestotes where numllauna='" + atrim(vllauna) + "'", , ReadOnly)
   If Not rst.EOF Then
       nomdelatinta = atrim(rst!descripcio)
   End If
   Set rst = Nothing
End Function
Function noesunloterroni(vlot As String, vllauna As String) As Boolean
   Dim vX As String
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from lotserronis")
   If rst.EOF Then noesunloterroni = True
   While Not rst.EOF
      vX = UCase(atrim(rst!numerolot))
      If vX <> "" Then
       If InStr(1, UCase(atrim(vlot)), vX) = 0 And InStr(1, UCase(atrim(vlot)), vX) = 0 Then
         noesunloterroni = True
           Else: noesunloterroni = False: rst.MoveLast
       End If
      End If
      rst.MoveNext
   Wend
   Set rst = Nothing
End Function
Function numdelotbaseonumlotdellauna(vnumlotbase As String, Optional vnumllaunarelacionada As Boolean, Optional vnumllauna As String) As String
  Dim vultimnumlot As String
  Dim vnumlot As String
  vnumlotbase = UCase(vnumlotbase)
  If Mid(vnumlotbase, 1, 1) = "A" And (Len(vnumlotbase) > 4 And Len(vnumlotbase) < 7) Then
    vultimnumlot = buscarlot0delalluna(vnumlotbase)
    If vultimnumlot = "0" Then vultimnumlot = vnumlotbase + " (Sense Lot s´ha de buscar manualment la llauna)"
    While Mid(vultimnumlot, 1, 1) = "A" And (Len(vultimnumlot) > 4 And Len(vultimnumlot) < 7) And vultimnumlot <> vnumlotbase
     
     vnumlot = buscarlot0delalluna(vultimnumlot)
     If vnumlot <> "0" And vnumlot <> "" Then
          vultimnumlot = vnumlot
            Else: vultimnumlot = vnumlotbase + " (Sense Lot s´ha de buscar manualment la llauna)"
     End If
    Wend
     numdelotbaseonumlotdellauna = vultimnumlot + IIf(vnumllaunarelacionada, "[" + vnumlotbase + "]", "") '+ " (" + atrim(vnumlotbase) + ")"
     vnumllauna = vnumlotbase
       Else: numdelotbaseonumlotdellauna = vnumlotbase
  End If
End Function
Function numdelotbaseonumlotdellauna_total(vnumlotbase As String) As String
  vnumlotbase = UCase(vnumlotbase)
  If Mid(vnumlotbase, 1, 1) = "A" And (Len(vnumlotbase) > 4 And Len(vnumlotbase) < 7) Then
     numdelotbaseonumlotdellauna_total = buscarlot0delalluna(vnumlotbase) + " (" + atrim(vnumlotbase) + ")"
       Else: numdelotbaseonumlotdellauna_total = vnumlotbase
  End If
End Function

Function buscarlot0delalluna(vnumlotbase As String) As String
   Dim rsthistoria As Recordset
   Dim rst As Recordset
   Dim vhistoria As String
   Dim idsdecarga As String
   vhistoria = vnumlotbase
   'Set rsthistoria = dbtintes.OpenRecordset("SELECT llaunes.id,Llaunes.numllauna, historiallauna.data, historiallauna.id,historiallauna.tipusmoviment,historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes RIGHT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) RIGHT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) INNER JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent where (numllauna='" + atrim(nllauna) + "')  ORDER BY historiallauna.data DESC;")
   Set rsthistoria = dbtintes.OpenRecordset("SELECT Llaunes.id, Llaunes.numllauna, historiallauna.data, historiallauna.id, historiallauna.tipusmoviment, historiallauna.idhistoriabarreja, historiallaunalots.*, Componentsbase.nomcomponent FROM ((Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna) LEFT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria) LEFT JOIN Componentsbase ON historiallaunalots.idcomponent = Componentsbase.idcomponent Where (numllauna='" + atrim(vnumlotbase) + "') ORDER BY historiallauna.data DESC;")
   If rsthistoria.EOF Then vhistoria = "0"
   While Not rsthistoria.EOF
       If rsthistoria!tipusmoviment = "C" And rsthistoria!idcomponent = 0 Then vhistoria = atrim(rsthistoria!numlotbase)
       rsthistoria.MoveNext
   Wend
   buscarlot0delalluna = vhistoria
   Set rsthistoria = Nothing
End Function
Function numllaunadelid(id As Long) As String
   Dim rst As Recordset
   numllaunadelid = 0
   Set rst = dbtintes.OpenRecordset("select numllauna from llaunes where id=" + atrim(id))
   If Not rst.EOF Then numllaunadelid = rst!numllauna
End Function

Private Sub Command29_Click()
  If tintes.Recordset.EditMode <> 0 Then MsgBox "Primer guarda els canvis de la tinta abans d'entrar referències.", vbCritical, "Atenció": Exit Sub
  formrefproveidors.Show 1
End Sub

 Function triartinta() As Long
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  
  sql = "SELECT  idtinta,codi,descripcio,referenciacolor from tintes_tot "
  were = " order by descripcio"
  Load formseleccio
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = sql + were
  formseleccio.width = 13000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.Show 1
  If seleccioret = 1 Then
    triartinta = atrim(formseleccio.Data1.Recordset!idtinta)
  End If
  If seleccioret = 9 Then
    triartinta = 0
  End If
  Unload formseleccio
End Function

Private Sub Command30_Click()
'  If datallaunes.Recordset.EOF Then MsgBox "Primer escull una llauna.", vbInformation, "Atenció": Exit Sub
'  ensenyarsituacions datallaunes.Recordset!numllauna
End Sub
Sub ensenyarsituacions(nllauna As String)
   Dim rsthistoria As Recordset
   Dim vhistoria As String
   Set rsthistoria = dbtintes.OpenRecordset("SELECT data,situacio from historialsituacions where numllauna='" + nllauna + "' ORDER BY data desc")
   While Not rsthistoria.EOF
       vhistoria = vhistoria + atrim(rsthistoria!Data) + " - > " + atrim(rsthistoria!situacio) + Chr(10)
       rsthistoria.MoveNext
   Wend
   MsgBox vhistoria, vbInformation, "Historia situacions de la llauna"

End Sub

Private Sub Command31_Click()
  Dim nll As String
  If datallaunes.Recordset.EOF Then Exit Sub
  nll = datallaunes.Recordset!numllauna
  calcularkgdisponiblesllauna datallaunes.Recordset!numllauna
  datallaunes.Refresh
  datallaunes.Recordset.FindFirst "numllauna='" + atrim(nll) + "'"
End Sub

Private Sub Command32_Click()
   reconvertirllauna atrim(datadellaunes.Recordset!numllauna)
End Sub
Function buscarlaultimasituacio(vidtinta As Long) As String
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from llaunes where situacio<>'SALA' and situacio<>'IMP' and idtinta=" + atrim(vidtinta))
   If Not rst.EOF Then buscarlaultimasituacio = atrim(rst!situacio)
   If buscarlaultimasituacio = "" Then buscarlaultimasituacio = "IMP"
End Function
Sub triartintarefproveidor(rsttintes As Recordset)
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  Unload formseleccio
  
  sql = "SELECT  idrefproveidor,codi,descripcio,refproveidor,nominterndelbido from tintes_tot "
  were = " order by descripcio"
  Load formseleccio
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = sql + were
  formseleccio.width = 13000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.cmissatge.tag = 2
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 600
  formseleccio.DBGrid2.Columns(2).width = 4000
  formseleccio.DBGrid2.Columns(3).width = 2500
  formseleccio.DBGrid2.Columns(4).width = 2500
  formseleccio.Show 1
  If seleccioret = 1 Then
    Set rsttintes = dbtintes.OpenRecordset("select * from tintes_tot where idrefproveidor=" + atrim(formseleccio.Data1.Recordset!idrefproveidor))
  End If
  If seleccioret = 9 Then
    'posso rsttintes un valor que no existeixi per deixarlo en EOF
    Set rsttintes = dbtintes.OpenRecordset("select * from tintes_tot where codi=-9999")
  End If
  Unload formseleccio
End Sub
Function crearlanovallaunaareconvertir() As String
   Dim rstllauna As Recordset
   Dim numnovallauna As Double
   Dim rsttintes As Recordset
   
   triartintarefproveidor rsttintes
   If rsttintes.EOF Then Exit Function
   Set rstllauna = dbtintes.OpenRecordset("select numllauna from contadors")
   numnovallauna = rstllauna!numllauna + 1
   dbtintes.Execute "update contadors set numllauna=[numllauna]+1"
   Set rstllauna = dbtintes.OpenRecordset("select * from llaunes")
   rstllauna.AddNew
   rstllauna!numllauna = "A" + atrim(numnovallauna)
   crearlanovallaunaareconvertir = rstllauna!numllauna
   rstllauna!idtinta = rsttintes!idtinta
   rstllauna!id_refproveidor = rsttintes!idrefproveidor
   rstllauna!situacio = buscarlaultimasituacio(rsttintes!idtinta)
   rstllauna!activa = True
   rstllauna.Update
End Function
Sub reconvertirllauna(numllaunaareconvertir As String)
  Dim numllaunaperbuidar As String
  Dim numllaunaaonbuidar As String
  Dim idhistoriabarreja As Long
  Dim desctintes As String
  Dim rstll As Recordset
  Dim rsthistoria As Recordset
  Dim rstlldesti As Recordset
  Dim rstformula As Recordset
  
  
  If atrim(numllaunaareconvertir) = "" Then Exit Sub
  Set rstll = dbtintes.OpenRecordset("SELECT Llaunes.*, tintes.codi,tintes.descripcio FROM tintes LEFT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta where numllauna='" + atrim(numllaunaareconvertir) + "'")
  If rstll.EOF Then MsgBox "Aquesta llauna no existeix.", vbCritical, "Atenció": Exit Sub
  If Not rstll!activa Then MsgBox "Aquesta llauna no està activa no pots convertir-la", vbCritical, "Atenció": Exit Sub
  If MsgBox("Segur que vols convertir la llauna " + numllaunaareconvertir + "-" + UCase(atrim(rstll!descripcio)) + " a un altra tipus de tinta?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  Set rstformula = dbtintes.OpenRecordset("Select * from tintesformules where idtinta=" + atrim(rstll!codi) + " order by predeterminada")
  numllaunaaonbuidar = crearlanovallaunaareconvertir
  If atrim(numllaunaaonbuidar) = "" Then Exit Sub
  Set rstlldesti = dbtintes.OpenRecordset("SELECT Llaunes.*, tintes.descripcio FROM tintes LEFT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta where numllauna='" + atrim(numllaunaaonbuidar) + "'")
  If rstlldesti.EOF Then MsgBox "Aquesta llauna no existeix.", vbCritical, "Atenció": Exit Sub
  If Not rstlldesti!activa Then MsgBox "Aquesta llauna no està activa no pots reconvertir-la", vbCritical, "Atenció": Exit Sub
  desctintes = Chr(10) + "Origen: " + UCase(rstll!numllauna) + " --> " + UCase(rstll!descripcio) + Chr(10)
  desctintes = desctintes + "Destí: " + UCase(rstlldesti!numllauna) + " --> " + UCase(rstlldesti!descripcio)
  
  
  'creo l'historia de la llauna nova i buido la anterior
    calcularkgdisponiblesllauna rstll!numllauna
    Set rsthistoria = dbtintes.OpenRecordset("select * from historiallauna")
    rsthistoria.AddNew
    rsthistoria!idnumllauna = rstlldesti!id
    rsthistoria!numrecarrega = numproximarecarrega(numllaunaaonbuidar, True)
    rsthistoria!Data = Now
    rsthistoria!tipusmoviment = "K"
    rsthistoria!formula = IIf(Not rstformula.EOF, rstformula!numformula, "")
    rsthistoria!kg = rstll!capacitatactual
    idhistoriabarreja = rsthistoria!id
    rsthistoria.Update
    rsthistoria.AddNew
    rsthistoria!idnumllauna = rstll!id
    rsthistoria!numrecarrega = numproximarecarrega(numllaunaareconvertir, True)
    rsthistoria!Data = Now
    rsthistoria!tipusmoviment = "V"
    rsthistoria!formula = ""
    rsthistoria!idhistoriabarreja = idhistoriabarreja
    rsthistoria!kg = rstll!capacitatactual
    rsthistoria.Update
    calcularkgdisponiblesllauna rstll!numllauna
    calcularkgdisponiblesllauna rstlldesti!numllauna

'copio les carectaristiques de la llauna anterior
  rstlldesti.Edit
  rstlldesti!idmaterialcontenidor = rstll!idmaterialcontenidor
  rstlldesti!idmaterialcontenidor = rstll!idproveidorrecuperador
  rstlldesti!preuxrkilo = rstll!preuxrkilo
  rstlldesti.Update
'passo la llauna anterior a inactiva
  rstll.Edit
  rstll!activa = False
  rstll!situacio = "REC"
  rstll.Update
  MsgBox desctintes, vbInformation, "Llaunes convertides"
  If MsgBox("Vols imprimir la nova etiqueta?", vbInformation + vbDefaultButton1 + vbYesNo, "Imprimir") = vbYes Then
      imprimir_etiqueta atrim(rstlldesti!numllauna)
  End If
  Set rsthistoria = Nothing
  Set rstll = Nothing
  Set rstlldesti = Nothing
End Sub

Private Sub Command33_Click()
   formsituacio.Show 1
End Sub

Private Sub Command34_Click()
  If datatintesformules.Recordset.EOF Then Exit Sub
  If MsgBox("Segur que vols eliminar aquesta relació amb la formula?", vbCritical + vbYesNo, "Atenció") = vbYes Then
      datatintesformules.Recordset.Delete
      datatintesformules.Refresh
  End If
End Sub

Private Sub Command35_Click()
  Dim formulaactual As String
  If datatintesformules.Recordset.EOF Then MsgBox "Primer escull quina vols que sigui la predeterminada.", vbCritical, "Error": Exit Sub
  formulaactual = atrim(datatintesformules.Recordset!numformula)
  datatintesformules.Refresh
  While Not datatintesformules.Recordset.EOF
     datatintesformules.Recordset.Edit
     If formulaactual <> atrim(datatintesformules.Recordset!numformula) Then
        datatintesformules.Recordset!predeterminada = False
           Else: datatintesformules.Recordset!predeterminada = True
     End If
     datatintesformules.Recordset.Update
     datatintesformules.Recordset.MoveNext
  Wend
End Sub
Public Function ArchivoEnUso(ByVal sFileName As String) As Boolean
    Dim filenum As Integer, errnum As Integer
    
    
    On Error Resume Next ' Turn error checking off.
    filenum = FreeFile() ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open sFileName For Input Lock Read As #filenum
    Close filenum ' Close the file.
    errnum = err ' Save the error number that occurred.
    On Error GoTo 0 ' Turn error checking back on.
    
    ' Check to see which error occurred.
    Select Case errnum
    
    ' No error occurred.
    ' File is NOT already open by another user.
    Case 0
    ArchivoEnUso = False
    
    ' Error number for «Permission Denied.»
    ' File is already opened by another user.
    Case 70
    ArchivoEnUso = True
    
    ' Another error occurred.
    Case Else
    Error errnum
End Select
End Function
Sub convertir_reixa_a_CSV()
   Dim vnomfitxer As String
   Dim vcol As Long
   Dim vrow As Double
   Dim vlinia As String
   Dim vcapcalera As String
   
   vnomfitxer = "c:\temp\~Seleccio_Exportat.csv"
   If existeix(vnomfitxer) Then
      If ArchivoEnUso(vnomfitxer) Then MsgBox "Hi ha el fitxer d 'exportació CSV obert, tanca'l primer.": Exit Sub
      Kill vnomfitxer
   End If
   Open "c:\temp\~Seleccio_Exportat.csv" For Output As #1
   formseleccio.Data1.Recordset.MoveFirst
   While Not formseleccio.Data1.Recordset.EOF
        vlinia = "": vcapcalera = ""
        For vcol = 0 To formseleccio.DBGrid2.Columns.Count - 1
           If formseleccio.DBGrid2.Columns(vcol).visible Then
                 If formseleccio.DBGrid2.Columns(vcol).width > 100 Then
                       vlinia = vlinia + IIf(vlinia <> "", ";" + atrim(formseleccio.DBGrid2.Columns(vcol)), atrim(formseleccio.DBGrid2.Columns(vcol)))
                       vcapcalera = vcapcalera + IIf(vcapcalera <> "", ";" + atrim(formseleccio.DBGrid2.Columns(vcol).caption), atrim(formseleccio.DBGrid2.Columns(vcol).caption))
                 End If
           End If
        Next vcol
        If formseleccio.DBGrid2.Row = 0 Then Print #1, vcapcalera
        Print #1, vlinia
        formseleccio.Data1.Recordset.MoveNext
   Wend
   Close #1
   If existeix(vnomfitxer) Then obrir_document vnomfitxer
End Sub

Private Sub Command36_Click()
  Dim vsql As String
  Dim vnohiha As Byte
  vsql = "SELECT impresorespantones.comanda From impresorespantones "
  vsql = vsql + " WHERE (((impresorespantones.lot1) Like '*AAAAA*')) OR (((impresorespantones.lot2) Like '*AAAAA*')) OR (((impresorespantones.lot3) Like '*AAAAA*')) OR (((impresorespantones.lot4) Like '*AAAAA*')) OR (((impresorespantones.lot5) Like '*AAAAA*')) OR (((impresorespantones.lot6) Like '*AAAAA*')) OR (((impresorespantones.lot7) Like '*AAAAA*')) OR (((impresorespantones.lot8) Like '*AAAAA*'));"

  If datallaunes.Recordset.EOF Then MsgBox "Has d'escullir una llauna primer.", vbCritical, "Error": Exit Sub
  vsql = substituir(vsql, "AAAAA", atrim(datallaunes.Recordset!numllauna))
  
  Load formseleccio
  formseleccio.caption = "Selecciona una formula"
  formseleccio.Data1.DatabaseName = rutadelfitxer(camitintes) + "baixes.mdb"
  formseleccio.Data1.RecordSource = "SELECT impresores_llaunesgastades.comanda FROM impresores_llaunesgastades where numllauna='" + datallaunes.Recordset!numllauna + "'"
  formseleccio.refrescar
  If formseleccio.Data1.Recordset.EOF Then vnohiha = 1: GoTo nohiha
  formseleccio.DBGrid2.Columns(0).visible = True
  formseleccio.DBGrid2.Columns(0).width = 1500
  formseleccio.bimprimir.visible = True
  formseleccio.width = 6000
  formseleccio.Show 1
  If seleccioret = 9 Then convertir_reixa_a_CSV
  Unload formseleccio
  Exit Sub
nohiha:
  Load formseleccio
  formseleccio.caption = "Selecciona una formula"
  formseleccio.Data1.DatabaseName = rutadelfitxer(camitintes) + "baixes.mdb"
  formseleccio.Data1.RecordSource = vsql
  formseleccio.refrescar
  If formseleccio.Data1.Recordset.EOF Then vnohiha = 2: GoTo nohiha2
  formseleccio.DBGrid2.Columns(0).visible = True
  formseleccio.DBGrid2.Columns(0).width = 1500
  formseleccio.width = 6000
  formseleccio.Show 1
  Unload formseleccio
  Exit Sub
  
nohiha2:
  If vnohiha = 2 Then MsgBox "No s'ha utilitzat per cap comanda."
  Unload formseleccio
  
End Sub


Private Sub Command37_Click()
   Dim vformula As String
   vformula = escullirformula
   If vformula <> "" Then
      dbtintes.Execute "insert into tintesformules (idtinta,numformula) values (" + atrim(tintes.Recordset!idtinta) + ",'" + atrim(vformula) + "')"
      datatintesformules.Refresh
      MsgBox "Relació feta correctament."
'      imprimirinformetinta vformula, tintes.Recordset!idtinta
      
   End If
End Sub
Sub imprimirinformetinta(codiformula As String, idtinta As Long)
' Dim rst As Recordset
  
  Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Dim camp As TextObject
  Dim f  As OLEObject
  Dim rstf As Recordset
  Dim rstt As Recordset
  Set rstt = dbtintes.OpenRecordset("SELECT  idtinta,codi,descripcio,referenciacolor from tintes where idtinta=" + atrim(idtinta))
  If rstt.EOF Then Exit Sub
  Set rstf = dbtintes.OpenRecordset("select codiformula,descripcioformula,series,datacreacio,notes from formules where codiformula='" + atrim(codiformula) + "'")
  If rstf.EOF Then Exit Sub
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "verificaciorelaciotintainkmaker.rpt", 1)

  oreport.FormulaFields.GetItemByName("descripcioformula").Text = "'" + atrim(rstf!codiformula) + " ---->   " + atrim(rstf!descripcioformula) + "'"
  oreport.FormulaFields.GetItemByName("descripcio tinta").Text = "'" + atrim(rstt!descripcio) + "'"
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
Function escullirformula() As String
  Static ultimcodi As String
  Load formseleccio
  formseleccio.caption = "Selecciona una formula"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "select codiformula,descripcioformula,series,datacreacio,notes from formules order by descripcioformula"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = True
  formseleccio.DBGrid2.Columns(1).width = 4500
  formseleccio.DBGrid2.Columns(2).width = 800
  formseleccio.width = 10000
  formseleccio.Show 1
  If seleccioret = 1 Then
   escullirformula = atrim(formseleccio.Data1.Recordset!codiformula)
  End If
  Unload formseleccio
  
End Function


Sub actualitzar_components_inkmaker()
    Dim rstc As Recordset
    Dim rsti As Recordset
    Dim vids As String
    Set rsti = conODBC.OpenRecordset("select * from dbo.tblcomponenti ")
    vids = "0"
    While Not rsti.EOF
       Set rstc = dbtintes.OpenRecordset("select * from componentsbase where idcomponent=" + atrim(rsti!idcomponente))
       If rstc.EOF Then
             rstc.AddNew
               Else: rstc.Edit
       End If
        vids = vids + IIf(vids <> "", ",", "") + atrim(rsti!idcomponente)
        rstc!idcomponent = rsti!idcomponente
        rstc!numdosificador = rsti!vasca
        rstc!codicomponent = atrim(rsti!codcomponente)
        rstc!nomcomponent = Mid(treure_apostruf(atrim(rsti!DescComponente)), 1, 40)
        rstc.Update
        rsti.MoveNext
   Wend
   dbtintes.Execute "delete * from componentsbase where idcomponent not in (" + vids + ")"
    Set rstc = Nothing
    Set rsti = Nothing
End Sub
Sub buscar_tinta(vid_tinta As Integer)
 tintes.RecordSource = "select * from tintes_tot "
 tintes.Refresh
 tintes.Recordset.FindFirst "codi='" + atrim(vid_tinta) + "'"
  pestanyes.Tab = 0
End Sub

Private Sub Command38_Click()
   actualitzar_llista_albarans
End Sub
Sub actualitzar_llista_albarans()
  'Dim rstcompres As Recordset
  Dim rstlinies As Recordset
  Dim rstbido As Recordset
  Dim vreferencia As String
  Dim vcont As Byte
  llistaalbarans.Clear
  ' Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
  Set rstcompres = dbcompres.OpenRecordset("select * from albaransbip where " + IIf(checktots.Value <> 1 And Checkultims30.Value <> 1, " not llaunescreades and", "") + " cint(article)>1999 " + IIf(checktots.Value = 1 Or Checkultims30.Value = 1, " AND DATA<=NOW ", "") + " order by data DESC")
  While Not rstcompres.EOF
      Set rstlinies = dbcompres.OpenRecordset("select * from liniesdescripcio where idliniacompra=" + atrim(cadbl(rstcompres!idliniacompra)) + " and descripcio like 'Ref:*'")
      vreferencia = ""
      If Not rstlinies.EOF Then
        vreferencia = substituir(atrim(rstlinies!descripcio), "Ref: ", " ")
        Set rstbido = dbtintes.OpenRecordset("SELECT tintesreferencies.referencia, tipusbidons.nombido, tipusbidons.capacitat, tipusbidons.litrescompres FROM tintesreferencies LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id where tintesreferencies.referencia='" + atrim(vreferencia) + "'")
        If rstbido.EOF Then vreferencia = ""
      End If
      If jashaescanejatlalbara(atrim(rstcompres!numalbaraprov)) Then
        If vreferencia <> "" Then
            llistaalbarans.AddItem justificar(atrim(rstcompres!numalbaraprov), 15, "E") + justificar(atrim(rstcompres!descripcio), 50, "E") + justificar(atrim(Redondejar(cadbl(rstbido!capacitat))), 5, "D") + justificar(atrim(Redondejar(cadbl(rstcompres!quantitat), 0)), 7, "D") + " Kg"
            llistaalbarans.ItemData(llistaalbarans.NewIndex) = rstcompres!id
              Else:
                llistaalbarans.AddItem justificar(atrim(rstcompres!numalbaraprov), 15, "E") + justificar(atrim(rstcompres!descripcio), 50, "E") + "** Error ref. proveïdor **"
                llistaalbarans.ItemData(llistaalbarans.NewIndex) = rstcompres!id
        End If
      End If
      rstcompres.MoveNext
      If checktots.Value = 1 Then vcont = vcont + 1
      If Checkultims30.Value = 1 Then vcont = vcont + 1
      If Checkultims30.Value = 1 And vcont > 30 Then
           rstcompres.MoveLast: rstcompres.MoveNext
      End If
      If vcont > 200 Then rstcompres.MoveLast: rstcompres.MoveNext
  Wend
  Set rstcompres = Nothing
  Set rstbido = Nothing
  Set rstlinies = Nothing
  'Set dbcompres = Nothing
End Sub
Function jashaescanejatlalbara(vnumalb As String) As Boolean
  Dim rst As Recordset
  Set rst = dbcomandes.OpenRecordset("select * from registre_escanejades_expedicions where tipus='ALB' AND nomfitxer like '" + vnumalb + "*'")
  If Not rst.EOF Then jashaescanejatlalbara = True
  Set rst = Nothing
  
End Function
Function justificar(v As String, longitut As Integer, DoE As String) As String
    v = Mid(v, 1, longitut)
    If DoE = "E" Then
       v = v + Space(longitut - Len(v))
      Else: v = Space(longitut - Len(v)) + v
    End If
    justificar = v
End Function

Private Sub Command39_Click()
  If Command39.tag = "1" Then Exit Sub
   Command39.tag = "1"
  If llistaalbarans.ListIndex > -1 Then
    localitzar_idbip_i_tipusdevido llistaalbarans.ItemData(llistaalbarans.ListIndex)
  End If
  actualitzar_llista_albarans
  Command39.tag = ""
End Sub
Function buscar_la_tinta(vcoditinta As String) As Boolean
  tintes.RecordSource = "select * from tintes_tot"
  tintes.Refresh
  tintes.Recordset.FindFirst "codi='" + atrim(vcoditinta) + "'"
  If tintes.Recordset.NoMatch Then
     buscar_la_tinta = False
       Else: buscar_la_tinta = True
  End If
End Function
Sub localitzar_idbip_i_tipusdevido(vidbip As Long)
   Dim rstcompres As Recordset
  Dim rstlinies As Recordset
  Dim rstbido As Recordset
  Dim vreferencia As String
  actualitzarcarguescomponents
 ' Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
  Set rstcompres = dbcompres.OpenRecordset("select * from albaransbip where id=" + atrim(vidbip))
  If Not rstcompres.EOF Then
      If Not buscar_la_tinta(atrim(rstcompres!article)) Then MsgBox "No trobo la tinta comprada.", vbCritical, "Atenció": Exit Sub
      Set rstlinies = dbcompres.OpenRecordset("select * from liniesdescripcio where idliniacompra=" + atrim(cadbl(rstcompres!idliniacompra)) + " and descripcio like 'Ref:*'")
      vreferencia = ""
      If Not rstlinies.EOF Then
        vreferencia = substituir(atrim(rstlinies!descripcio), "Ref: ", " ")
        Set rstbido = dbtintes.OpenRecordset("SELECT tintesreferencies.referencia,tintesreferencies.id as idreferencia, tipusbidons.nombido, tipusbidons.capacitat, tipusbidons.litrescompres FROM tintesreferencies LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id where tintesreferencies.referencia='" + atrim(vreferencia) + "'")
        If rstbido.EOF Then
           vreferencia = ""
            Else
              crear_llaunes rstbido, rstcompres
        End If
      End If
      If vreferencia <> "" Then
          Else: MsgBox "Error amb la referencia del proveïdor." + Chr(10) + "No localitzo aquesta referència.", vbCritical, "Error"
      End If
  End If
  Set rstcompres = Nothing
  Set rstbido = Nothing
  Set rstlinies = Nothing
'  Set dbcompres = Nothing
End Sub
Function demanar_situacio_llauna() As String
  formescullirsituaciollaunes.Show 1
  demanar_situacio_llauna = atrim(formescullirsituaciollaunes.combosituacio)
  Unload formescullirsituaciollaunes
End Function
Sub crear_llaunes(rstbido As Recordset, rstcompres As Recordset)
  Dim vnumllaunes As Double
  Dim vnumllaunesusuari As Double
  Dim vlitrescompra As Double
  Dim vkgllauna As String
  Dim vtotalkgcreats As Double
  Dim vsituacio As String
  Dim vnllauna As String
  Dim vllaunescreades(100) As String * 10
  Dim i As Byte
  vsituacio = demanar_situacio_llauna
  If vsituacio = "" Then MsgBox "Si no esculls una situació per les llaunes no es crearan", vbCritical, "Error": Exit Sub
  vlitrescompra = cadbl(rstbido!litrescompres)
  If vlitrescompra = 0 Then vlitrescompra = cadbl(rstbido!capacitat)
  vlitrescompra = cadbl(InputBox("De quants Kg son les llaunes?", "Crear Llaunes", vlitrescompra))
  If vlitrescompra > 0 Then vnumllaunes = atrim(Redondejar(cadbl(rstcompres!quantitat) / vlitrescompra, 0))
  vnumllaunesusuari = cadbl(InputBox("Quantes llaunes vols crear d'aquesta compra?" + Chr(10) + Chr(10) + atrim(rstcompres!descripcio), "Crear Llaunes", vnumllaunes))
  If vnumllaunesusuari = 0 Then Exit Sub
  If vnumllaunesusuari < vnumllaunes Then MsgBox "No pots fer menys llaunes que els kilos comprats.", vbCritical, "Error": GoTo fi
  If vnumllaunesusuari > vnumllaunes + 2 Then If MsgBox("El numero de llaunes que vols crear es mes de dues llaunes de les suggerides," + Chr(10) + "ès correcte aixó?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo fi
  ratoli "espera"
  vnumllaunes = vnumllaunesusuari
  If vnumllaunes = 1 Then vlitrescompra = cadbl(rstcompres!quantitat)
  vkgllauna = vlitrescompra
  vlaresta = ((cadbl(rstcompres!quantitat) - vtotalkgcreats) - vkgllauna)
  If vlaresta < 3 And vlaresta > 0 Then vkgllauna = vkgllauna + vlaresta
  For i = 1 To vnumllaunes
       novallauna atrim(rstbido!idreferencia), vkgllauna, atrim(rstcompres!numlotproveidor), vsituacio, 1, vnllauna, cadbl(rstcompres!vidmaterialcontenidor), cadbl(rstcompres!idproveidorrecuperador), atrim(rstcompres!vmatriculacontenidor), atrim(rstcompres!numalbaraprov)
       etcreantllaunes.caption = "Creant la llauna " + vnllauna + "       " + atrim(i) + "/" + atrim(vnumllaunes) + " Llaunes"
       'esperarquelallaunaestiguicreada vnllauna
       'wait 2
       'imprimir_etiqueta_llauna vnllauna
       vllaunescreades(i) = vnllauna
       vtotalkgcreats = vtotalkgcreats + vkgllauna
       If (vtotalkgcreats + vkgllauna) > cadbl(rstcompres!quantitat) Then vkgllauna = cadbl(rstcompres!quantitat) - vtotalkgcreats
       vlaresta = ((cadbl(rstcompres!quantitat) - vtotalkgcreats) - vkgllauna)
       If vlaresta < 3 And vlaresta > 0 Then vkgllauna = vkgllauna + vlaresta
  Next i
  'imprimir les llaunes creades
  For i = 1 To vnumllaunes
     etcreantllaunes.caption = "Imprimint la llauna " + vnllauna + "       " + atrim(i) + "/" + atrim(vnumllaunes) + " Llaunes"
     DoEvents
     imprimir_etiqueta vllaunescreades(i)
     'imprimir_etiqueta_llauna vllaunescreades(i)
  Next i
  rstcompres.Edit
  rstcompres.llaunescreades = True
  rstcompres.Update
  ratoli "normal"
  etcreantllaunes.caption = ""
fi:
End Sub
Sub esperarquelallaunaestiguicreada(vnllauna As String)
   Dim rst As Recordset
   Dim vhora As Date
   vhora = Now
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, Llaunes.capacitatactual, historiallauna.formula, Llaunes.situacio, historiallauna.data, Llaunes.preuxrkilo FROM Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna Where (((Llaunes.numllauna) = '" + vnllauna + "') And ((historiallauna.tipusmoviment) = 'C'))")
   While rst.EOF And DateDiff("s", vhora, Now) < 10
     Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, Llaunes.capacitatactual, historiallauna.formula, Llaunes.situacio, historiallauna.data, Llaunes.preuxrkilo FROM Llaunes LEFT JOIN historiallauna ON Llaunes.id = historiallauna.idnumllauna Where (((Llaunes.numllauna) = '" + vnllauna + "') And ((historiallauna.tipusmoviment) = 'C'))")
   Wend
   If rst.EOF Then vnllauna = ""
End Sub
Sub imprimir_etiqueta_llauna(vnllauna As String)
   Dim vhorainici As Date
   If vnllauna = "" Then Exit Sub
   escriure_ini "Tintes", "imprimir_etiqueta", vnllauna, fitxerini
   Shell llegir_ini("General", "rutallistats", fitxerini) + "etiquetes tintes.exe"
   vhorainici = Now
   While Now <= DateAdd("s", 8, vhorainici)
     If llegir_ini("Tintes", "imprimir_etiqueta", fitxerini) = "" Then GoTo fi
     DoEvents
   Wend
   If Now > DateAdd("s", 8, vhorainici) Then
      MsgBox "Hi ha hagut algun error al imprimir_l'etiqueta o tarda molt a imprimir-la." + Chr(10) + "Nomes es pot imprimir etiquetes desde l'ordinador de Inkmaker" + Chr(10) + "i el programa d'etiquetes obert.", vbCritical, "Error"
      escriure_ini "Tintes", "imprimir_etiqueta", "", fitxerini
   End If
fi:
End Sub

Private Sub Command4_Click()
  Dim numllauna As String
  If Not datallaunes.Recordset.EOF Then numllauna = atrim(datallaunes.Recordset!numllauna)
  editallauna numllauna
  'If numllauna <> "" Then datallaunes.Recordset.FindFirst "numllauna='" + numllauna + "'"
End Sub

Private Sub Command40_Click()
 If Command40.tag = atrim(datallaunes.Recordset!numllauna) Then
      If MsgBox("Aquesta llauna ja la has imprès ara mateix, vols tornar-hi?", vbCritical + vbYesNo + vbDefaultButton2, "Repeticio") = vbNo Then GoTo fi
 End If
 Command40.tag = atrim(datallaunes.Recordset!numllauna)
 imprimir_etiqueta datallaunes.Recordset!numllauna
fi:
End Sub

Private Sub Command41_Click()
   'Dim dbcompres As Database
   Dim vnumliniacompra As Double
   If llistaalbarans.ListIndex > -1 Then
      If UCase(InputBox("Escriu [eliminar] per treure aquesta linia del llistat.", "Treure tinta sense crear llaunes")) <> "ELIMINAR" Then Exit Sub
      vnumliniacompra = cadbl(llistaalbarans.ItemData(llistaalbarans.ListIndex))
    '  Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
      dbcompres.Execute "update albaransbip set llaunescreades=true where id=" + atrim(vnumliniacompra)
      actualitzar_llista_albarans
        Else: MsgBox "Escull primer la linia que vols utilitzar.", vbExclamation, "Atenció"
   End If
   'Set dbcompres = Nothing
End Sub
Function calcular_kg_ultimacomanda(vidtreball As Double, vordremodificacio As Double, vcolor As String) As Double
   Dim rst As Recordset
   Dim vcantitatex As Double
   Dim i As Byte
   Set rst = dbcomandes.OpenRecordset("select comanda,cantitatex,proximaseccio from comandes where numtreball=" + atrim(vidtreball) + " and numordremodificacio=" + atrim(vordremodificacio) + " and (proximaseccio<>'E' and proximaseccio<>'I') order by comanda desc")
   If rst.EOF Then GoTo fi
   vcantitatex = cadbl(rst!cantitatex)
   Set rst = dbbaixes.OpenRecordset("select * from impresorespantones where comanda=" + atrim(rst!comanda))
   If rst.EOF Then GoTo fi
   For i = 1 To 8
        If rst.Fields("pantone" + atrim(i)) = vcolor Then calcular_kg_ultimacomanda = cadbl(rst.Fields("kg" + atrim(i))) / cadbl(vcantitatex)
   Next i
fi:
   Set rst = Nothing
End Function
Function calcular_kgassignatsaaltrescomandes(Optional rsttinter As Recordset, Optional vcoditinta As String, Optional comandesinplicades As String) As Double
  Dim rst As Recordset
  Dim rstt As Recordset
  Dim vconsumteoric As Double
  Dim vultimid As Double
  If atrim(vcoditinta) = "" Then vcoditinta = atrim(cadbl(rsttinter!coditinta))
  Set rst = dbtintes.OpenRecordset("SELECT comandes.comanda,comandes.cantitatex,comandes.numtreball, comandes.numordremodificacio, Tintesclixesnous.coditinta, Tintesclixesnous.id_tinter FROM (comandesrevisadesatintes INNER JOIN comandes ON comandesrevisadesatintes.comanda = comandes.comanda) INNER JOIN Tintesclixesnous ON (comandes.numordremodificacio = Tintesclixesnous.ordremodificacio) AND (comandes.numtreball = Tintesclixesnous.id_treball) WHERE (((comandes.proximaseccio)='E' Or (comandes.proximaseccio)='I') AND ((Tintesclixesnous.coditinta)='" + vcoditinta + "')) AND ((comandesrevisadesatintes.estatgestio)<>'N') order by id_tinter;")
'  Clipboard.Clear
'  Clipboard.SetText "SELECT comandes.comanda,comandes.cantitatex,comandes.numtreball, comandes.numordremodificacio, Tintesclixesnous.coditinta, Tintesclixesnous.id_tinter FROM (comandesrevisadesatintes INNER JOIN comandes ON comandesrevisadesatintes.comanda = comandes.comanda) INNER JOIN Tintesclixesnous ON (comandes.numordremodificacio = Tintesclixesnous.ordremodificacio) AND (comandes.numtreball = Tintesclixesnous.id_treball) WHERE (((comandes.proximaseccio)='E' Or (comandes.proximaseccio)='I') AND ((Tintesclixesnous.coditinta)='" + vcoditinta + "')) AND ((comandesrevisadesatintes.estatgestio)<>'N');"
  While Not rst.EOF
    If vultimid <> rst!id_tinter Then Set rstt = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(rst!id_tinter), , ReadOnly)
    vultimid = rst!id_tinter
    comandesinplicades = comandesinplicades + IIf(comandesinplicades <> "", ",", "") + atrim(rst!comanda)
    If Not rstt.EOF Then vconsumteoric = vconsumteoric + (cadbl(rst![cantitatex]) * calcular_kgmetreteoric(rstt))
    rst.MoveNext
  Wend
 ' MsgBox comandesinplicades
  Set rst = Nothing
  Set rstt = Nothing
  calcular_kgassignatsaaltrescomandes = vconsumteoric
End Function
Sub actualitza_llistatintes(vcomanda As Double, Optional vxls As Boolean, Optional vllistabones As String, Optional vmaquina As String)
  Dim rstc As Recordset
  Dim rstc1 As Recordset
  Dim rstestoc As Recordset
  Dim vlinia As String
  Dim vmuntats As Boolean
  Dim vkgtintaxrmetre As Double
  Dim vidtreball As Double
  Dim vcantitatex As Double
  Dim vmodificacio As Double
  Dim vestatclixe As String
  Dim vconsumteoric As Double
  Dim vetconsumteoric As String
  Dim vestatgestio As String
  Dim vhihaalgunaextensio As Boolean
  Dim vteextensiofeta As Boolean
  Dim vprimercaracter As String
  Dim vhihaprimar As Boolean
  Dim vhihatintesalternatives As Boolean
  
  Command13.Enabled = False
  If InStr(1, vmaquina, "FW") > 0 Then
      vmaquina = "FW"
        Else:
          If InStr(1, vmaquina, "F2") > 0 Then
              vmaquina = "F2"
                Else: vmaquina = ""
          End If
  End If
  Set rstc = dbcomandes.OpenRecordset("select numtreball,numordremodificacio,cantitatex from comandes where comanda=" + atrim(vcomanda), , ReadOnly)
  If rstc.EOF Then Exit Sub
  
  vidtreball = cadbl(rstc!numtreball)
  vmodificacio = cadbl(rstc!numordremodificacio)
  vcantitatex = cadbl(rstc!cantitatex)
  vestatclixe = estatclixemod(vidtreball, vmodificacio)
  vestatgestio = reixacomandes.TextMatrix(reixacomandes.Row, numcol("Gestionat?"))
  llistatintes.BackColor = &HEEE4D7
  cobsoperari.WhatsThisHelpID = vcomanda
  cobsoperari.tag = vidtreball
  If vestatclixe <> "POLIMERS O CLIXES" And vestatclixe <> "CLIXES ENTRATS" And vestatclixe <> "CLIXES REBUTS" Then llistatintes.BackColor = &H8080FF  'GoTo fi
  
   
  Set rstc1 = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(vidtreball) + " and (ordremodificacio=" + atrim(vmodificacio) + " or ordremodificacio=-" + atrim(vmodificacio) + ") order by ordretinter")
  ratoli "espera"
  Command13.tag = atrim(vcomanda)
  Command13.Enabled = True
  While Not rstc1.EOF
      vetconsumteoric = ""
         'miro primer si hi ha tintes alternatives--- si ja n'hi ha un no ho comprovo mes
      If Not vhihatintesalternatives Then
          Set rstc = dbclixes.OpenRecordset("select * from tintes_alternatives where id_tinter=" + atrim(IIf(cadbl(rstc1!tinterlinkambid_treball) > 0, rstc1!tinterlinkambid_treball, rstc1!id_tinter)))
          If Not rstc.EOF Then vhihatintesalternatives = True
      End If
      Set rstc = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(IIf(cadbl(rstc1!tinterlinkambid_treball) > 0, rstc1!tinterlinkambid_treball, rstc1!id_tinter)))
      If rstc.EOF Then GoTo cont
      vlinia = ""
      ' no ensenyar primars If InStr(1, atrim(rstc!color), "PRIMAR") > 0 Then GoTo cont
      vprimercaracter = " "
      If InStr(1, atrim(rstc!color), "PRIMAR") > 0 Then vprimercaracter = ".": vhihaprimar = True
      If atrim(rstc!color) = "" Then GoTo cont 'treure els que no tenen color
      'veure els blanc o no veurels If cadbl(rstc!coditinta) = 2044 Or cadbl(rstc!coditinta) = 3433 Then GoTo cont
'      vkgtintaxrmetre = calcular_kg_ultimacomanda(vidtreball, vmodificacio, rstc!color)
      If Check1(2).Value <> 0 Then vkgtintaxrmetre = calcular_kgassignatsaaltrescomandes(rstc)
      vconsumteoric = (vcantitatex * calcular_kgmetreteoric(rstc))
      If cadbl(vconsumteoric) > 0 Then
          vconsumteoric = vconsumteoric + 10
          If vkgtintaxrmetre > vconsumteoric Then vkgtintaxrmetre = vkgtintaxrmetre - vconsumteoric 'resto el consum d'aquesta comanda del total de les comandes
          vetconsumteoric = "/" + atrim(Redondejar(cadbl(vconsumteoric), 0)) + "K"
            Else: vetconsumteoric = ""
      End If
      vlinia = vlinia + vprimercaracter + atrim(rstc!ordretinter) + " "
      vlinia = vlinia + justificar(atrim(rstc!coditinta), 5, "E")
      vlinia = vlinia + justificar(IIf(cadbl(rstc!anilox) > 0, atrim(rstc!anilox), ""), 4, "E")
      vteextensiofeta = mirarsihihaextensio(rstc!id_tinter)
      vlinia = vlinia + justificar(IIf(vteextensiofeta, "*" + atrim(LCase(rstc!color)), atrim(rstc!color)), 35, "E")
      If vteextensiofeta Then vhihaalgunaextensio = True
      vlinia = vlinia + justificar(buscarllaunadisponible(atrim(rstc!coditinta)), 15, "E")
      'vlinia = vlinia + justificar(atrim(Redondejar(vkgtintaxrmetre * vcantitatex, 0)) + "K" + vetconsumteoric, 7, "D")
      vlinia = vlinia + justificar(atrim(Redondejar(vkgtintaxrmetre, 0)) + "K" + vetconsumteoric, 12, "D")
      vlinia = vlinia + justificar(Format(kg_estoc_familia(cadbl(rstc!coditinta)), "#,##0") + " Kg", 10, "D")
      llistatintes.AddItem vlinia
      llistatintes.ItemData(llistatintes.NewIndex) = cadbl(rstc!id_tinter)
      If vllistabones <> "" Then vxls = IIf(InStr(1, vllistabones, "#" + atrim(llistatintes.NewIndex) + "#") = 0, False, True)
      If vllistabones <> "" And vxls Then llistatintes.Selected(llistatintes.NewIndex) = True
      'If vxls Then Print #1, atrim(vmaquina) + "-" + atrim(vcomanda) + ";" + atrim(rstc!coditinta) + ";" + IIf(cadbl(rstc!anilox) > 0, atrim(rstc!anilox), "") + ";" + atrim(rstc!color) + ";" + buscarllaunadisponible(atrim(rstc!coditinta)) + ";" + atrim(Redondejar(vkgtintaxrmetre * vcantitatex, 0)) + ";" + atrim(Redondejar(vconsumteoric, 0))
      If vxls Then Print #1, atrim(vmaquina) + "-" + atrim(vcomanda) + ";" + atrim(rstc!coditinta) + ";" + IIf(cadbl(rstc!anilox) > 0, atrim(rstc!anilox), "") + ";" + atrim(rstc!color) + ";" + buscarllaunadisponible(atrim(rstc!coditinta)) + ";" + atrim(Redondejar(vconsumteoric, 0))
      
cont:
      rstc1.MoveNext
  Wend
  If vhihatintesalternatives Then
        balternatives.BackColor = &HC78DFA
         Else: balternatives.BackColor = &H8000000F
  End If
  If vhihaprimar Then llistatintes.AddItem "=- - - - - - - - - - - - - - - - -  P R I M A R S - - - - - - - - - - - - - - - - ="
  llistatintes.ForeColor = IIf(vhihaalgunaextensio, &HFF0000, QBColor(0))
  cobsoperari = posarobservaciooperari(vcomanda)
  carregar_observacio_tintes vidtreball, vmodificacio
  mirar_cuatricumia
fi:
  ratoli "normal"
  Set rstc = Nothing
  Set rstc1 = Nothing
  
End Sub
Function posarobservaciooperari(vnumc As Double) As String
  Dim rst As Recordset
  cobsoperari = ""
  Set rst = dbtintes.OpenRecordset("select observacions from comandesrevisadesatintes where comanda=" + atrim(vnumc))
  If Not rst.EOF Then posarobservaciooperari = atrim(rst!observacions)
  Set rst = Nothing
End Function
Function calcular_kgmetreteoric(rsttinter As Recordset) As Double
   Dim vresultat As Double
   Dim vvolum As Double
   Dim vaporte As Double
   Dim rst As Recordset
   Dim rsta As Recordset
   Dim vample As Double
   'amplelamina
   Set rst = dbclixes.OpenRecordset("select amplelamina,bandes from modificacions where id_treball=" + atrim(rsttinter!id_treball) + " and ordre=" + atrim(rsttinter!ordremodificacio))
   If rst.EOF Then Exit Function
   vample = (cadbl(rst!amplelamina) * 10) * cadbl(rst!bandes)
   Set rsta = dbcomandes.OpenRecordset("select volum as volummesgran from aniloxos where lineatura=" + atrim(cadbl(rsttinter!anilox)) + " order by volum Desc")
   'vvolum = IIf(cadbl(rsttinter!volum) > 0, cadbl(rsttinter!volum), 20)
   vvolum = IIf(rsta.EOF, 0, cadbl(rsta!volummesgran))
   vaporte = 30 'no se quin valor es aquest es per defecte
   vresultat = (vaporte * vvolum) / 100000
   vresultat = ((cadbl(rsttinter!tanx100cobertura) / 100) * (vample / 1000)) * vresultat
   vresultat = vresultat * 0.95
   calcular_kgmetreteoric = vresultat
End Function
Function estatclixemod(ByVal ntreball As Double, ByVal ordrem As Double, Optional vdataclixes As Date) As String
  Dim rst As Recordset
  vdataclixes = 0
  If ordrem = 0 Then ordrem = 1
  Set rst = dbclixes.OpenRecordset("SELECT clixes_modifi.id_estatclixe,Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, CLIXES_MODIFI.data_prevista,Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE Clixes_modifi.id_treball=" + atrim(ntreball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordrem) + " AND clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(ntreball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordrem) + ");")
  ' CONSULTA ULTIMA MODIFICACIO AMB DATA FI  VINCULANT "SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE (((Clixes_modifi.id_treball)=" + atrim(id_treball) + ") AND ((Clixes_modifi.ordremodificacio)=" + atrim(ordremodificacio) + ") AND ((Clixes_estats.vinculant)=True and isdate(clixes_modifi.data_fi)) and clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + "));"
  ' CONSULTA ULTIMA MODIFICACIO AMB DATA FI SENSE VINCULANT "SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE (((Clixes_modifi.id_treball)=" + atrim(id_treball) + ") AND ((Clixes_modifi.ordremodificacio)=" + atrim(ordremodificacio) + ") AND (isdate(clixes_modifi.data_fi)) and clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + "));"
  If Not rst.EOF Then
     If rst!id_estatclixe = 15 Or rst!id_estatclixe = 22 Then vdataclixes = Format(rst!data_prevista, "dd/mm/yy")
     estatclixemod = atrim(rst!descrip)
       Else: estatclixemod = ""
  End If
End Function
Function kg_estoc_familia(vcoditinta As Double) As Double
   Dim rst As Recordset
   Dim rstllaunes As Recordset
   If vcoditinta < 1 Then Exit Function
   Set rst = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(vcoditinta) + "'")
   If rst.EOF Then MsgBox "El codi de tinta " + atrim(vcoditinta) + " no l'he trobada a la base de dades", vbCritical, "Error": Exit Function
   If InStr(1, rst!referenciacolor, "P-") = 0 Then
    With rst
     'vsubconsulta = "select idtinta from tintes where idfamilia=" + atrim(cadbl(!idfamilia)) + " and idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + "and idfamcolor= " + atrim(cadbl(!idfamcolor)) + " and idsubfamcolor=" + atrim(cadbl(!idsubfamcolor))
     'Set rstllaunes = dbtintes.OpenRecordset("select sum(capacitatactual) as kgestoc from llaunes where idtinta in (" + vsubconsulta + ") and activa=true and capacitatactual>0")
     'SELECT Sum(llaunes.capacitatactual) AS SumaDecapacitatactual FROM llaunes INNER JOIN tintes ON llaunes.idtinta = tintes.idtinta WHERE (((llaunes.activa)=True) AND ((llaunes.capacitatactual)>0) AND
     vsubconsulta = "(idfamilia=" + atrim(cadbl(!idfamilia)) + " and idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + "and idfamcolor= " + atrim(cadbl(!idfamcolor)) + " and idsubfamcolor=" + atrim(cadbl(!idsubfamcolor)) + "));"
    ' Clipboard.Clear
    ' Clipboard.SetText "SELECT Sum(llaunes.capacitatactual) AS kgestoc FROM llaunes INNER JOIN tintes ON llaunes.idtinta = tintes.idtinta WHERE (((llaunes.activa)=True) AND ((llaunes.capacitatactual)>0) AND " + vsubconsulta
     Set rstllaunes = dbtintes.OpenRecordset("SELECT Sum(llaunes.capacitatactual) AS kgestoc FROM llaunes INNER JOIN tintes ON llaunes.idtinta = tintes.idtinta WHERE mid(referenciacolor,1,2)<>'P-' and (((llaunes.activa)=True) AND ((llaunes.capacitatactual)>0) AND " + vsubconsulta)
     kg_estoc_familia = cadbl(rstllaunes!kgestoc)
    End With
     Else
       vsubconsulta = "select idtinta from tintes where codi='" + atrim(vcoditinta) + "'"
       Set rstllaunes = dbtintes.OpenRecordset("select sum(capacitatactual) as kgestoc from llaunes where idtinta in (" + vsubconsulta + ") and activa=true and capacitatactual>0")
       kg_estoc_familia = cadbl(rstllaunes!kgestoc)
   End If

End Function
Function mirarsihihaextensio(vidtinter As Long, Optional vnumextensio As String) As Boolean
  Dim rst As Recordset
  Dim rstc As Recordset
  Set rst = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(vidtinter))
  If Not rst.EOF Then
      Set rstc = dbtintes.OpenRecordset("select * from extensions_treballsrelacionats where numtreball=" + atrim(rst!id_treball) + " and numordremodificacio=" + atrim(rst!ordremodificacio) + " and coditinta=" + atrim(cadbl(rst!coditinta)))
      If Not rstc.EOF Then
         mirarsihihaextensio = True
         vnumextensio = atrim(rstc!codiextensio)
      End If
  End If
  Set rst = Nothing
  Set rstc = Nothing
End Function
Function buscarllaunaassignada(vcoditinta As String) As String
  Dim rst As Recordset
  Dim vnumc As Double
  vnumc = reixacomandes.TextMatrix(reixacomandes.Row, numcol("Comanda"))
  Set rst = dbtintes.OpenRecordset("select * from assignaciollaunesacomandes where comanda=" + atrim(vnumc) + " and coditinta=" + atrim(cadbl(vcoditinta)))
  If Not rst.EOF Then buscarllaunaassignada = "->" + atrim(rst!numllauna)
End Function
Function buscarllaunadisponible(vcoditinta As String) As String
  Dim rst As Recordset
  Dim vsql As String
  buscarllaunadisponible = buscarllaunaassignada(vcoditinta)
  If buscarllaunadisponible <> "" Then Exit Function
  vsql = "SELECT tintes.codi, Llaunes.id, Llaunes.numllauna, Llaunes.situacio, Llaunes.activa,Llaunes.aimpresores FROM tintes INNER JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta Where (((tintes.codi) = '"
  vsql = vsql + vcoditinta + "') And ((Llaunes.activa) = True)) ORDER BY Llaunes.aimpresores asc;"
  Set rst = dbtintes.OpenRecordset(vsql)
  If Not rst.EOF Then
     rst.MoveLast
     rst.MoveFirst
     buscarllaunadisponible = "#" + atrim(rst.RecordCount) + IIf(rst!aimpresores, "*", " ") + atrim(rst!numllauna) + " (" + atrim(rst!situacio) + ")"
       Else
         Set rst = dbtintes.OpenRecordset("select idtinta from tintes where codi='" + atrim(vcoditinta) + "'")
         If Not rst.EOF Then
           Set rst = dbtintes.OpenRecordset("select * from tintesformules where idtinta=" + atrim(rst!idtinta))
           If rst.EOF Then buscarllaunadisponible = "No Formula"
         End If
  End If
  Set rst = Nothing
End Function

Private Sub Command42_Click()
'   framebotons.visible = True
'   llistacomandes.BackColor = Command42.BackColor
'   actualitza_llistacomandes IIf(checktotes.Value = 0, "muntats", "comanda"), IIf(checkrevisades.Value = 1, True, False), IIf(checknoinplacsa.Value = 1, True, False)
  Set dbstocks = OpenDatabase(rutadelfitxer(cami) + "palets.mdb")
  ratoli "espera"
  refrescar_dades_comandesactives
  poblar_reixa_comandes
  carregar_amples_reixa
  ratoli "normal"
End Sub
Function canvidematerialcomandaanterior(vidtreball As Double, vordre As Double, vcomanda As Double, vcomanda2 As Double, vcomanda3 As Double) As Boolean
  Dim vsql As String
  Dim rst As Recordset
  Dim rst2 As Recordset
   'If vcomanda = 214069 Then Stop
  'vsql = "SELECT impressorestot.comanda, comandes.numtreball, comandes.numordremodificacio, impressorestot.dataimpressio FROM comandes INNER JOIN impressorestot ON comandes.comanda = impressorestot.comanda "
  'vsql = vsql + " Where (((comandes.numtreball) = " + Trim(vidtreball) + ") And ((comandes.numordremodificacio) = " + Trim(vordre) + ")) ORDER BY impressorestot.dataimpressio DESC;"
  vsql = "SELECT impressorestot.comanda, comandes.numtreball, comandes.numordremodificacio, impressorestot.dataimpressio, Mid([refinplacsa],1,2) AS Expr1 "
  vsql = vsql + " FROM (comandes RIGHT JOIN impressorestot ON comandes.comanda = impressorestot.comanda) LEFT JOIN comandes_extres ON comandes.comanda = comandes_extres.comanda"
  vsql = vsql + " WHERE comandes.numtreball=" + Trim(vidtreball) + "  AND Mid([refinplacsa],1,2)<>'PR'"
  vsql = vsql + " ORDER BY impressorestot.dataimpressio DESC;"
  Set rst = dbbaixes.OpenRecordset(vsql)
  'Clipboard.Clear
'  Clipboard.SetText vsql
  If Not rst.EOF Then
      Set rst = dbbaixes.OpenRecordset("select comanda,linkcomanda1,linkcomanda2 from comandes where comanda=" + atrim(rst!comanda))
      Set rst2 = dbbaixes.OpenRecordset("select materialex from comandes where comanda=" + atrim(rst!comanda) + IIf(rst!linkcomanda1 <> 0, " or comanda=" + atrim(rst!linkcomanda1), "") + IIf(rst!linkcomanda2 > 0, " or comanda=" + atrim(rst!linkcomanda2), "") + " order by materialex asc")
      Set rst = dbbaixes.OpenRecordset("select materialex from comandes where comanda=" + atrim(vcomanda) + IIf(vcomanda2 <> 0, " or comanda=" + atrim(vcomanda2), "") + IIf(vcomanda3 > 0, " or comanda=" + atrim(vcomanda3), "") + " order by materialex asc")
      If rst!materialex <> rst2!materialex Then canvidematerialcomandaanterior = True
      If vcomanda2 > 0 Then
         rst.MoveNext: rst2.MoveNext
         If rst!materialex <> rst2!materialex Then canvidematerialcomandaanterior = True
      End If
      If vcomanda3 > 0 Then
         rst.MoveNext: rst2.MoveNext
         If rst!materialex <> cadbl(rst2!materialex) Then canvidematerialcomandaanterior = True
      End If
  End If
  Set rst = Nothing
  Set rst2 = Nothing
End Function
Function dataarribadamaterial(numc As Long) As Date
  Dim v As String
 ' Dim rstt As Recordset
 ' Dim rst
 ' If numc = 0 Then Exit Function
 ' Set rstt = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, capcalera.dataentrega as dataent, liniescompra.totentregat as entregat, comandesxlinia.numcomanda FROM (capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra) RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra WHERE (((comandesxlinia.numcomanda)=" + atrim(numc) + "));", dbOpenSnapshot, dbReadOnly)
 ' If Not rstt.EOF Then
  '   dataarribadamaterial = Format(rstt!dataent, "dd/mm/yy")
  '   If cabool(rstt!entregat) Then dataarribadamaterial = 0
  'End If
  v = estatdelmaterial(numc)
  If v = "AP" Then dataarribadamaterial = DateAdd("d", 1, Now)
  If v = "AE" Or v = "A" Then dataarribadamaterial = 0
  If IsDate(v) Then dataarribadamaterial = v
End Function
Function estatdelmaterial(numc As Long) As String
  Dim rstt As Recordset
  If numc = 0 Then Exit Function
  Set rstt = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, capcalera.dataentrega as dataent, liniescompra.totentregat as entregat, comandesxlinia.numcomanda FROM (capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra) RIGHT JOIN comandesxlinia ON liniescompra.idliniacompra = comandesxlinia.idliniacompra WHERE (((comandesxlinia.numcomanda)=" + atrim(numc) + "));", dbOpenSnapshot, dbReadOnly)
  If Not rstt.EOF Then
     estatdelmaterial = Format(rstt!dataent, "dd/mm/yy")
     If DateDiff("d", rstt!dataent, Now) >= 0 And Not cabool(rstt!entregat) Then
           estatdelmaterial = Format(DateAdd("d", 99, Now), "dd/mm/yy")
     End If
     If cabool(rstt!entregat) Then estatdelmaterial = "E" + Format(rstt!dataent, "dd/mm/yy")
  End If
  Set rstt = dbstocks.OpenRecordset("select * from percomandaoclient where  numcomanda=" + atrim(numc))
  If Not rstt.EOF Then estatdelmaterial = "R"
  Set rstt = dbstocks.OpenRecordset("select * from parcials where  comanda='" + atrim(numc) + "'")
  If Not rstt.EOF Then estatdelmaterial = "A" + IIf(estatdelmaterial <> "", IIf(Mid(estatdelmaterial + " ", 1, 1) = "E", "E", ""), "")
 ' Set rstt = dbcomandes.OpenRecordset("select assignarstock,materialexacte from comandes_extres where comanda=" + atrim(numc))
  Set rstt = dbcomandes.OpenRecordset("SELECT comandes_extres.assignarstock, comandes_extres.materialexacte, comandes.proximaseccio FROM comandes_extres INNER JOIN comandes ON comandes_extres.comanda = comandes.comanda where comandes_extres.comanda=" + atrim(numc), dbOpenSnapshot, dbReadOnly)
  If Not rstt.EOF Then
     If rstt!assignarstock Then estatdelmaterial = "A"
     If cabool(rstt!materialexacte) And atrim(rstt!proximaseccio) = "E" And estatdelmaterial = "A" Then estatdelmaterial = "ESP"
  End If
  Set rstt = Nothing
End Function

Function cabool(valor As Variant) As Boolean
  If IsNull(valor) Then valor = False
  If valor = "" Then valor = False
  If valor = "Sí" Or valor = "S" Then valor = True
  If valor = "No" Or valor = "N" Then valor = False
  If valor Then
    cabool = True
   Else: cabool = False
  End If
End Function
Function algunmaterial_èsALOX(vnumc1 As Double, vnumc2 As Double, vnumc3 As Double) As Boolean
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("SELECT comandes.comanda, subfamiliesmaterials.descripcio FROM comandes LEFT JOIN (materials LEFT JOIN subfamiliesmaterials ON materials.subfamilia = subfamiliesmaterials.codi) ON comandes.materialex = materials.codi WHERE comandes.comanda=" + atrim(vnumc1) + " or comandes.comanda=" + atrim(vnumc2) + " or comandes.comanda=" + atrim(vnumc3))
   algunmaterial_èsALOX = False
   While Not rst.EOF
     If InStr(1, atrim(rst!descripcio), "ALOX") > 0 Then algunmaterial_èsALOX = True
     rst.MoveNext
   Wend
   Set rst = Nothing
End Function
Sub refrescar_dades_comandesactives()
   Dim rstcomandesactives As Recordset
   Dim rstc As Recordset
   Dim rstp As Recordset
   Dim rstpa As Recordset
   Dim rstmas As Recordset
   Dim vestatclixe As String
   Dim rstextra As Recordset
   Dim dbplanificacio As Database
   Dim dbplanificacioalicia As Database
   Dim vdataclixes As Date
   Dim vdataarribadamaterial As Date
   Dim vdies As Double
   Dim valgunateextensio As Boolean
   Dim vmaquina As String
   Dim rstmat As Recordset
   Dim vhoraultimaactualitzacio As String
   Dim vdiesextres As Double
   Dim v As String
   Dim vX As Double
   Dim vsql As String
   
   vhoraultimaactualitzacio = llegir_ini("Tintes", "horaultimaactualitzaciocomandestintes", rutadelfitxer(cami) + "valorsprograma.ini")
   If vhoraultimaactualitzacio = "{[}]" Then vhoraultimaactualitzacio = ""
   If vhoraultimaactualitzacio <> "" Then
        If DateDiff("n", vhoraultimaactualitzacio, Now) < 2 Then
            MsgBox "Ja s'està actualitzant en un altra ordinador, espera una estona i torna-ho a provar.", vbCritical, "Error"
           Exit Sub
        End If
   End If
   'actualitzo la hora que he entrat
   escriure_ini "Tintes", "horaultimaactualitzaciocomandestintes", Trim(Now), rutadelfitxer(cami) + "valorsprograma.ini"
   
   dbtintes.Execute "delete * from comandesactives"
   dbtintes.Execute "DELETE comandesrevisadesatintes.*, comandes.proximaseccio FROM comandesrevisadesatintes INNER JOIN comandes ON comandesrevisadesatintes.comanda = comandes.comanda WHERE (((comandes.proximaseccio)<>'E' And (comandes.proximaseccio)<>'I'));"
   Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificaciooperaris.mdb")
   Set dbplanificacioalicia = OpenDatabase(rutadelfitxer(cami) + "planificacio.mdb")
   Set rstcomandesactives = dbtintes.OpenRecordset("select * from comandesactives")
   Set rstextra = dbcomandes.OpenRecordset("Select * from comandes_extres")
   Set rstc = dbcomandes.OpenRecordset("select * from comandes where numtreball<>null and numtreball<>0 and  (proximaseccio='E' or proximaseccio='I')")
   vsql = "SELECT materials.descripcio as descripcio,materials.familia,materials.subfamilia, tractamentcares.descripcio AS Tcara1, tractamentcares_1.descripcio AS Tcara2 FROM (materials LEFT JOIN tractamentcares ON materials.codidescmatcara1 = tractamentcares.codi) LEFT JOIN tractamentcares AS tractamentcares_1 ON materials.codidescmatcara2 = tractamentcares_1.codi "

   With rstcomandesactives
   While Not rstc.EOF
      rstextra.FindFirst "comanda=" + atrim(rstc!comanda)
      valgunateextensio = False
      If Not rstc.EOF Then Set rstmat = dbcomandes.OpenRecordset(vsql + " where materials.codi=" + atrim(cadbl(rstc!materialex)))
      .AddNew
      vmaquina = ""
      !estaamuntadora = estaamuntadora(rstc!comanda, vmaquina)
      'Set rstp = dbplanificacio.OpenRecordset("select ordre,maquina from planificacioimp where comanda=" + atrim(rstc!comanda))
     ' Set rstpa = dbplanificacioalicia.OpenRecordset("select ordre,maquina from planificacioimp where comanda=" + atrim(rstc!comanda))
     ' If Not rstp.EOF Then
     '       !ordremaquina = IIf(rstp!ordre > 0 And rstp!ordre < 999, atrim(rstp!ordre), Null)
     '       If Not rstpa.EOF And IsNull(!ordremaquina) Then !ordremaquina = rstpa!ordre
     '       If cadbl(!ordremaquina) = 0 Then !ordremaquina = 999
     ' End If
      !comanda = rstc!comanda
      !seccioactual = rstc!proximaseccio
      !novaorepetida = atrim(rstc!impressio) 'IIf(rstc.impressio = "N" Or rstc.impressio = "M", "N", " ")
      !tipusimpresio = IIf(atrim(rstc!formaimp) = "T", "Transp", "Normal")
      If algunmaterial_èsALOX(rstc!comanda, rstc!linkcomanda1, rstc!linkcomanda2) Then !tipusimpresio = !tipusimpresio + " -ALOX"
      If !novaorepetida = "R" Or !novaorepetida = "M" Then !canvidematerialsR = canvidematerialcomandaanterior(rstc!numtreball, rstc!numordremodificacio, rstc!comanda, rstc!linkcomanda1, rstc!linkcomanda2)
      If Not rstmat.EOF Then
          If InStr(1, rstmat!descripcio, "PVDC") > 0 Then
              !tipusimpresio = !tipusimpresio + "-PVDC"
               Else
                If Mid(rstmat!descripcio, 1, 4) = "PET " Then
                    !tipusimpresio = !tipusimpresio + "-PET"
                    If (rstmat!TCARA1 Like "CORONA*" And rstmat!Tcara2 Like "PLAIN*") Or (rstmat!TCARA1 Like "PLAIN*" And rstmat!Tcara2 Like "CORONA*") Then !tipusimpresio = !tipusimpresio + "Corona"
                End If
                If Mid(rstmat!descripcio, 1, 4) = "PEAD" Then !tipusimpresio = !tipusimpresio + "-PEAD"
                
          End If
          If InStr(1, UCase(rstmat!descripcio), "ANTIVAHO") > 0 Then !tipusimpresio = !tipusimpresio + "-NOVAHO"
      End If
      
      If (cadbl(rstmat!familia) = 500 And cadbl(rstmat!subfamilia) = 63) Or (cadbl(rstmat!familia) = 576 And cadbl(rstmat!subfamilia) = 290) Or cadbl(rstc!materialex) = 1258 Or cadbl(rstc!materialex) = 1254 Or cadbl(rstc!materialex) = 1038 Then
         !tipusimpresio = !tipusimpresio + " -ABL"
      End If
      If cadbl(rstc!materialex) = 1266 Then
         !tipusimpresio = !tipusimpresio + " -EST"
      End If
      If atrim(rstextra!est_o_past) <> "" Then !tipusimpresio = !tipusimpresio + " +" + atrim(rstextra!est_o_past)

      !estamuntada = IIf(estamuntada(rstc!comanda) = "S", True, False)
      !numtreball = rstc!numtreball
      !versiotreball = rstc!numordremodificacio
      !metres = rstc!cantitatex
      !codiclient = cadbl(rstc!client)
      !nomclient = buscarnomclient(rstc!client)
      !marcailinia = atrim(rstc!marcailinia)
      '!estaamuntadora = IIf(estaamuntadora(rstc!comanda) = "S", True, False)
      !tetintesforainplacsa = mirarsitetintaforainplacsa(rstc!numtreball, rstc!numordremodificacio, valgunateextensio)
      !tealgunateextensio = IIf(valgunateextensio, "S", " ")
      vestatclixe = estatclixemod(cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), vdataclixes)
      If vestatclixe <> "CLIXES ENTRATS" And vestatclixe <> "CLIXES REBUTS" Or !novaorepetida = "F" Or IsNull(rstc!dataactivacio) Then !numtreball = !numtreball * -1
      'And vestatclixe <> "POLIMERS O CLIXES"
      !estatclixe = vestatclixe
      !tecalloff = mirarsitecalloff(rstc!comanda)
      !tedoypack = IIf(InStr(1, UCase(rstc!obssol1), "DOYPAC") > 0, True, False)
      If !tedoypack Then !tipusimpresio = !tipusimpresio + "+DOY"
      'If rstc!comanda = 219223 Then Stop
      vdataarribadamaterial = dataarribadamaterial(rstc!comanda)
      vX = DateDiff("d", Now, vdataarribadamaterial)
      vdies = DateDiff("d", Now, vdataclixes)
      If vdies < vX Then vdies = vX
      If vX > 99 Or vX < 0 Then vX = 0
      If vX > vdies Then vdies = vX
     
      v = buscardataentregacomanda(rstc!comanda)
      If InStr(1, v, "@") Then
           v = substituir(atrim(v), "@", "")
           !tipusimpresio = "@" + !tipusimpresio
      End If
      !dataexpedicio = IIf(v = "Sense Data", "01/01/9999", v)
      !reprint = mirarisiesreimpres(rstc!numtreball, rstc!numordremodificacio)
     ' If rstc!comanda = 214138 Then Stop
      !ordremaquina = diesquefalten(!dataexpedicio, rstc!comanda) + IIf(vdies > 0, vdies / 100, 0)
      If !ordremaquina > 200 Then !ordremaquina = 999 + IIf(vdies > 0, vdies / 100, 0)
      If !ordremaquina = 999 And !seccioactual = "I" Then !ordremaquina = 0.01
      .Update
      rstc.MoveNext
   Wend
   End With
   wait 1  'poso un temps d'espera perquè avegades fa abans l'update que no pas apareixen els datos a la taula
   'actualitzo el valor de codidelinia d'impresió
   dbtintes.Execute "UPDATE comandesactives LEFT JOIN Modificacions ON (comandesactives.versiotreball = Modificacions.ordre) AND (comandesactives.numtreball = Modificacions.id_treball) SET comandesactives.CodiLinia = Format([modificacions].[codidelinia],'000')+'#'+Trim([modificacions].[codideliniav]) WHERE (((Modificacions.codidelinia)>0 And (Modificacions.codidelinia) Is Not Null));"
   buscoelscodisdeliniaqueencarahohison
   'actualitzao el camp de estatdelagestió
   dbtintes.Execute "UPDATE comandesactives LEFT JOIN comandesrevisadesatintes ON comandesactives.comanda = comandesrevisadesatintes.comanda SET comandesactives.gestionat = IIf([comandesrevisadesatintes].[estatgestio] Is Not Null,[comandesrevisadesatintes].[estatgestio],'N');"
   
   'actualiatzo el camp de tintesrevisadesnovesomodificacdes
   dbtintes.Execute "UPDATE comandesactives LEFT JOIN comandesrevisadesatintes ON comandesactives.comanda = comandesrevisadesatintes.comanda SET comandesactives.revisatnovamodificada = comandesrevisadesatintes.revisatnovamodificada;"
   
   'actualiatzo el camp de combinacio de tintes feta
   dbtintes.Execute "UPDATE comandesactives LEFT JOIN comandesrevisadesatintes ON comandesactives.comanda = comandesrevisadesatintes.comanda SET comandesactives.combinaciollaunesfeta = comandesrevisadesatintes.combinaciollaunesfeta;"
   
   'poso el valor horaultimaactualitzacio a sensedata
   escriure_ini "Tintes", "horaultimaactualitzaciocomandestintes", "", rutadelfitxer(cami) + "valorsprograma.ini"
   
   Set dbplanificacio = Nothing
   Set dbplanificacioalicia = Nothing
   
End Sub
Function mirarisiesreimpres(vnumtreball As Double, vordre As Double) As Boolean
    Dim rst As Recordset
    Set rst = dbclixes.OpenRecordset("select reimpres from modificacions where id_treball=" + atrim(vnumtreball) + " and ordre=" + atrim(vordre))
    If Not rst.EOF Then mirarisiesreimpres = rst!reimpres
    Set rst = Nothing
End Function
Function diesquefalten(vdataexpedicio As Date, vnumc As Double) As Double
     Dim vdiesextres As Double
     vdiesextres = buscar_dies_extres(vnumc)
     diesquefalten = DateDiff("d", Now, vdataexpedicio)
     diesquefalten = diesquefalten - vdiesextres
    
     If diesquefalten < vdiesextres Then
          diesquefalten = (vdiesextres - diesquefalten) * -1
     End If
End Function
Function buscar_dies_extres(vnumc As Double) As Double
    Dim rstc As Recordset
    Dim rstclixes As Recordset
    Dim rstplan As Recordset
    Dim rstproducte As Recordset
    Dim vruta As String
    Set rstc = dbcomandes.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc))
    If rstc.EOF Then GoTo fi
    Set rstclixes = dbclixes.OpenRecordset("select sistemadimpresio,bandes,reimpres from modificacions where id_treball=" + atrim(cadbl(rstc!numtreball)) + " and ordre=" + atrim(rstc!numordremodificacio))
    If rstclixes.EOF Then GoTo fi
    Set rstproducte = dbcomandes.OpenRecordset("select * from productes where codi='" + atrim(rstc!producte) + "'")
    If rstproducte.EOF Then GoTo fi
    vruta = rstproducte!ruta
    If vnumc = 214533 Then Stop
    Set rstplan = dbplanificacioalicia.OpenRecordset("select * from planificaciosol where comanda=" + atrim(vnumc))
    buscar_dies_extres = IIf(rstclixes!reimpres = True, 2, 0)
    If rstplan.EOF And InStr(1, vruta, "S") > 0 Then buscar_dies_extres = 5
    buscar_dies_extres = IIf(atrim(rstc!microperforat) <> "" And atrim(rstc!microperforat) <> "N", 7, buscar_dies_extres)
       'microperforat a Rebobinadora     7 dies menys
       'Reprint                          2 dies menys
       'Soladores a Inplacsa             5 dies menys
fi:
    Set rst = Nothing
    Set rstclixes = Nothing
    Set rstplan = Nothing
    Set rstproducte = Nothing
End Function
Sub buscoelscodisdeliniaqueencarahohison()
   Dim rst As Recordset
   Dim rst2 As Recordset
   Dim vnumtreball As Double
   Set rst = dbtintes.OpenRecordset("select * from comandesactives where codilinia='' or codilinia=null")
   While Not rst.EOF
      vnumtreball = IIf(rst!estatclixe = "POLIMERS O CLIXES" And rst!numtreball < 0, rst!numtreball * -1, rst!numtreball)
      Set rst2 = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(vnumtreball) + " and codidelinia<>null and codidelinia>0 order by ordre desc")
      If Not rst2.EOF Then
         rst.Edit: rst!CodiLinia = IIf(rst!numtreball < 0 And vnumtreball > 0, "", "-") + Format(rst2!codidelinia, "000") + "#" + atrim(rst2!codideliniav): rst.Update
      End If
      rst.MoveNext
   Wend
   Set rst = Nothing
   Set rst2 = Nothing
End Sub
Function mirarsitecalloff(vnumc As Double) As Boolean
    Dim rst As Recordset
    Set rst = dbcomandes.OpenRecordset("select comanda from calloffs_detall where comanda=" + atrim(vnumc))
    If Not rst.EOF Then mirarsitecalloff = True
    Set rst = Nothing
End Function
Function mirarsitetintaforainplacsa(numtreball As Double, ordre As Double, valgunateextensio As Boolean) As Boolean
    Dim rsttintes As Recordset
    Dim rstc1 As Recordset
    Dim rstcolor As Recordset
    Dim vidtinta As String
    mirarsitetintaforainplacsa = False
    Set rstc1 = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(numtreball) + " and ordremodificacio=" + atrim(ordre) + " order by ordretinter")
    While Not rstc1.EOF
       Set rsttintes = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(IIf(cadbl(rstc1!tinterlinkambid_treball) > 0, rstc1!tinterlinkambid_treball, rstc1!id_tinter)))
       If InStr(1, atrim(rsttintes!color), "PRIMAR") > 0 Or atrim(rsttintes!color) = "" Then GoTo proximatinta
       If cadbl(rsttintes!coditinta) = 2044 Or cadbl(rsttintes!coditinta) = 3433 Then GoTo proximatinta
       Set rstcolor = dbtintes.OpenRecordset("select idtinta from tintes where codi='" + atrim(rsttintes!coditinta) + "'")
       If Not rstcolor.EOF Then
                If mirarsihihaextensio(rsttintes!id_tinter) = True Then valgunateextensio = True
                If mirarsitetintaforainplacsa = False Then
                    vidtinta = atrim(rstcolor!idtinta)
                    Set rstcolor = dbtintes.OpenRecordset("select * from tintesreferencies where idtinta=" + vidtinta + " and nomproveidor<>'INPLACSA'")
                    If Not rstcolor.EOF Then mirarsitetintaforainplacsa = True ': GoTo fi
                    Set rstcolor = dbtintes.OpenRecordset("select * from tintesformules where idtinta=" + vidtinta)
                    If rstcolor.EOF Then mirarsitetintaforainplacsa = True ': GoTo fi
                End If
       End If
proximatinta:
       rstc1.MoveNext
    Wend
fi:
     Set rsttintes = Nothing
     Set rstcolor = Nothing
     Set rstc1 = Nothing
End Function
Sub netejar_reixa_comandes()
   Dim rst As Recordset
   Dim i As Byte
   Dim col As Byte
   Set rst = dbtintes.OpenRecordset("select * from comandesactives")
   reixacomandes.Rows = 1
   reixacomandes.Cols = 1
   col = 0
   For i = 0 To rst.Fields.Count - 1
     If valorpropietat(rst.Fields(i), "Caption") <> "" Then
      reixacomandes.Cols = col + 1
      reixacomandes.col = col
      reixacomandes.Text = valorpropietat(rst.Fields(i), "Caption")
      If filtre.Count <= col Then Load filtre(col)
      If Screen.ActiveControl.Name <> "filtre" Then
       filtre(col).DataField = rst.Fields(i).Name
       filtre(col).Text = valorpropietat(rst.Fields(i), "Caption")
      End If
      col = col + 1
     End If
   Next i
   
End Sub
Function valorpropietat(rst As Field, v As String) As String
  Dim i As Byte
  For i = 0 To rst.Properties.Count - 1
      If rst.Properties(i).Name = v Then valorpropietat = rst.Properties(i)
  Next i
End Function
Sub poblar_reixa_comandes(Optional velordre As String)
   Dim rst As Recordset
   Dim fila As Integer
   Dim i As Byte
   Dim col As Integer
   Dim w As String
   Dim vordre As String
   Dim vsql As String
   
   If InStr(1, werescomandes, "codilinia") > 0 Then vordre = " order by codilinia"
   If InStr(1, werescomandes, "comanda in (") > 0 Then vordre = " "
   If velordre <> "" Then vordre = velordre
   'If InStr(1, werescomandes, "estamuntada=False") > 0 Then vordre = " order by dataexpedicio"
   reixacomandes.Redraw = False
   netejar_reixa_comandes
   ettotalcomandes.caption = ""
   w = werescomandes + IIf(werescomandes <> "" And Command63.tag <> "", " and " + Command63.tag, Command63.tag)
   Set rst = dbtintes.OpenRecordset("select * from comandesactives " + IIf(w <> "", " where " + w + IIf(vordre <> "", vordre, ""), IIf(vordre <> "", vordre, " order by ordremaquina")))
   If rst.EOF Then GoTo fi
   Set vrstCloneComandes = dbtintes.OpenRecordset("select * from comandesactives  order by ordremaquina")
   rst.MoveLast
   rst.MoveFirst
   If Not rst.EOF Then ettotalcomandes.caption = "Registres: " + atrim(rst.RecordCount)
   vsql = "SELECT comandesactives.CodiLinia, comandesactives.seccioactual, comandesactives.tecalloff, comandesactives.codiclient From comandesactives WHERE (((comandesactives.tecalloff)<>False) AND ((comandesactives.codiclient)=6841) AND ((Mid([CodiLinia],1,3)) In (SELECT Mid([CodiLinia],1,3) AS Expr1 From comandesactives GROUP BY Mid([CodiLinia],1,3) HAVING (((Mid([CodiLinia],1,3)) Is Not Null And (Mid([CodiLinia],1,3))>'0') AND ((Count(comandesactives.comanda))>1));))) OR (((comandesactives.codiclient)<>6841) AND ((Mid([CodiLinia],1,3)) In (SELECT Mid([CodiLinia],1,3) AS Expr1 From comandesactives GROUP BY Mid([CodiLinia],1,3) HAVING (((Mid([CodiLinia],1,3)) Is Not Null And (Mid([CodiLinia],1,3))>'0') AND ((Count(comandesactives.comanda))>1));))) ORDER BY comandesactives.seccioactual;"
   'If Not rst.EOF Then Set rstCdL = dbtintes.OpenRecordset(vsql)
   If Not rst.EOF Then Set rstCdL = dbtintes.OpenRecordset("SELECT comandesactives.CodiLinia, comandesactives.seccioactual,tecalloff,codiclient From comandesactives ORDER BY comandesactives.seccioactual")
   Set rstCdLestats = dbtintes.OpenRecordset("select * from EstatsCdL")
      
   fila = 1
   While Not rst.EOF
      col = 0
      reixacomandes.Rows = fila + 1
      For i = 0 To rst.Fields.Count - 1
        If valorpropietat(rst.Fields(i), "Caption") <> "" Then
            possar_el_valor_alareixacomandes fila, col, rst.Fields(i), rst
            col = col + 1
        End If
        
      Next i
      fila = fila + 1
      rst.MoveNext
   Wend
fi:
   reixacomandes.Redraw = True
   Set rst = Nothing
   If reixacomandes.Rows > 1 Then
    reixacomandes.Row = 1
    reixacomandes.col = 0
    reixacomandes.ColSel = reixacomandes.Cols - 1
    reixacomandes.RowSel = 1
   End If
   reixacomandes_SelChange
   Set vrstCloneComandes = Nothing
'carregar_liniadelareixaseleccionada
End Sub
Function nhihandosiguals(vnumtreball As Double, vnumordre As Double) As Boolean
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("select count(*) as contador from comandesactives where numtreball=" + atrim(vnumtreball) + " and versiotreball=" + atrim(vnumordre) + " group by numtreball,versiotreball")
  
  If rst!contador > 1 Then nhihandosiguals = True
End Function
Sub possar_el_valor_alareixacomandes(fila As Integer, col As Integer, vcamp As Field, vrst As Recordset)
  Dim v As String
  Dim vcolor As Double
  Dim vestat As String
  Dim vcanvicol As Double
  vcanvicol = -1
  If vcamp.Type = 1 Then v = IIf(vcamp.Value, "S", "N")
  If vcamp.Type = 4 Or vcamp.Type = 7 Then v = cadbl(vcamp.Value)
  If vcamp.Type = 10 Then v = atrim(vcamp.Value)
  If v = "" Then v = atrim(vcamp.Value)
  If vcamp.Name = "ordremaquina" Then If vcamp.Value - Int(vcamp.Value) > 0 Then v = Format(vcamp.Value, "0.00")
  reixacomandes.TextMatrix(fila, col) = v
  If vcamp.Name = "numtreball" Then If cadbl(vcamp.Value) < 0 Then vcolor = QBColor(12)
  If vcamp.Name = "numtreball" Then If nhihandosiguals(cadbl(v), vrst.Fields("versiotreball")) Then vcolor = QBColor(14)
  If vcamp.Name = "numtreball" Then If vrst.Fields("estatclixe") = "POLIMERS O CLIXES" Then vcolor = &HF3B378  'blau 'GoTo fi
  If vcamp.Name = "numtreball" Then
'    If atrim(vrst!maquina) = "FS" Then vcolor = QBColor(5)
  End If
  If vcamp.Name = "tipusimpresio" Then
       If Mid(atrim(vcamp.Value) + " ", 1, 1) = "@" Then
            vcanvicol = 0
            vcolor = &HFF80FF    'si hi ha ! al valor de comandes es que importancia de planificacio es 4
            reixacomandes.TextMatrix(fila, col) = Mid(vcamp.Value, 2)
       End If
  End If
  If vcamp.Name = "comanda" Then If vrst!tecalloff Then vcolor = QBColor(11)        'si te calloff posar el numero de comanda amb blau clar
  If vcamp.Name = "metres" Then If cadbl(vcamp.Value) = 0 Then vcolor = QBColor(12)
  If vcamp.Name = "CodiLinia" Then
      If Mid(atrim(vcamp.Value) + " ", 1, 1) = "-" Then
           vcolor = QBColor(12)
         Else
            If vrst.Fields("seccioactual") = "I" And atrim(vcamp.Value) <> "" Then
                 ' If vrst!comanda = 213991 Then Stop
                  If vrst!codiclient <> 6841 Or vrst.Fields("tecalloff") = True Then
                   'vrstCloneComandes.FindFirst "CodiLinia like '" + Mid(atrim(vcamp.Value), 1, 3) + "*'"
                   'If Not vrst.EOF Then
                      vrstCloneComandes.FindFirst "comanda<>" + atrim(vrst!comanda) + " and CodiLinia like '" + Mid(atrim(vcamp.Value), 1, 3) + "*'"
                      If Not vrstCloneComandes.NoMatch Then vcolor = QBColor(10)
                   'End If
                  End If
                 Else
                    If Not rstCdL.EOF And atrim(vcamp.Value) <> "" Then
                          rstCdL.FindFirst "Codilinia like '" + Mid(atrim(vcamp.Value), 1, 3) + "*' and seccioactual='E'"
                          If Not rstCdL.NoMatch Then
                             rstCdL.FindFirst "Codilinia like '" + Mid(atrim(vcamp.Value), 1, 3) + "*' and seccioactual='I'"
                             If Not rstCdL.NoMatch Then
                                 'If rstCdL!codiclient <> 6841 Then
                                 If vrst!codiclient <> 6841 Then
                                    vcolor = QBColor(15) ' &H80C0FF 'taronja
                                     Else
                                        If vrst.Fields("tecalloff") = True Then
                                                vcolor = QBColor(15) '&H80C0FF  'taronja
                                            Else: vcolor = QBColor(15) 'blanc
                                        End If
                                 End If
                                 If vrst!seccioactual = "I" And vcolor = &H80C0FF Then vcolor = QBColor(10)  'verd
                             End If
                          End If
                          If cadbl(vrst!metres) = 0 Or vrst!numtreball < 0 Then vcolor = QBColor(15) 'Blanc
                    End If
           End If
      End If
    '  If vcolor = &H80C0FF Then
      rstCdLestats.FindFirst "comanda=" + atrim(vrst!comanda)
      If Not rstCdLestats.NoMatch Then
        vestat = atrim(rstCdLestats!estat)
        If vestat = "A" Then vcolor = QBColor(10)  'verd
        If vestat = "R" Then
             If vrst!seccioactual <> "E" Then
                  rstCdLestats.Delete
                   Else: vcolor = QBColor(12)   'vermell
             End If
        End If
      End If
    '  End If
  End If
  If vcamp.Name = "novaorepetida" Then
    If atrim(vrst!revisatnovamodificada) = "S" Then
       vcolor = &HC0FFC0 'taronja
        Else: vcolor = QBColor(15) 'blanc
    End If
  End If
  If vcamp.Name = "tipusimpresio" Then If vrst!canvidematerialsR Then vcolor = QBColor(12)
  If vcamp.Name = "estaamuntadora" Then
     If InStr(1, vcamp.Value, "!") > 0 Then vcolor = QBColor(12)
     If InStr(1, vcamp.Value, "Ra") > 0 Then vcolor = &HC0FFC0
     If Mid(atrim(vcamp.Value) + " ", 1, 1) = "@" Then
            vcolor = &HFF80FF    'si hi ha @ al valor de comandes es que importancia de planificacio es 4
            reixacomandes.TextMatrix(fila, col) = Mid(vcamp.Value, 2)
     End If
  End If
  If vcamp.Name = "gestionat" Then
      If vcamp.Value = "S" Then vcolor = &HC0FFC0
      If vcamp.Value = "C" Then vcolor = QBColor(12)
      If vcamp.Value = "M" Then vcolor = QBColor(14)
      If vcamp.Value = "P" Then vcolor = &H80C0FF
      If vcamp.Value = "N" Then vcolor = 0
      If vrst!combinaciollaunesfeta Then vcolor = &HC78DFA
  End If
  If vcolor > 0 Then
     reixacomandes.col = IIf(vcanvicol = -1, col, vcanvicol)
     reixacomandes.Row = fila
     reixacomandes.CellBackColor = vcolor
  End If
End Sub
Sub guardar_amples_reixa()
Dim j As Integer
If iniconfigreixa <> "" Then
  For j = 0 To reixacomandes.Cols - 1
   escriure_ini "AmplesReixaComandesActives", UCase(reixacomandes.TextMatrix(0, j)), atrim(reixacomandes.ColWidth(j)), iniconfigreixa
 Next j
End If
End Sub
Sub carregar_amples_reixa()
 Dim ample As String
 Dim X As Long
 Dim j As Integer
 If iniconfigreixa <> "" Then ' existeix("c:\windows\" + iniconfigreixa) Then
 
  X = reixacomandes.Left + 35
  For j = 0 To reixacomandes.Cols - 1
   ample = llegir_ini("AmplesReixaComandesActives", UCase(reixacomandes.TextMatrix(0, j)), iniconfigreixa)
   If ample = "{[}]" Then ample = 1000
   reixacomandes.ColWidth(j) = cadbl(ample)
    If X < reixacomandes.width Then
     filtre(j).Left = X
     filtre(j).width = cadbl(ample)
     filtre(j).visible = True
     filtre(j).ForeColor = &H808080
      Else: If filtre.Count < j - 1 Then filtre(j).visible = False
    End If
    X = X + cadbl(ample)
 Next j
End If
filtre(0).width = filtre(0).width - 50
filtre(0).Left = filtre(0).Left + 50
End Sub
Sub actualitza_llistacomandes(vordre As String, Optional nocomprovarcomandesrevisades As Boolean, Optional veurenomescomandestintesnoinplacsa As Boolean)
  Dim rstc As Recordset
  Dim rsttintes As Recordset
  Dim rstcolor As Recordset
  Dim vlinia As String
  Dim vmuntats As Boolean
  Dim dbplanificacio As Database
  Dim vordreimaquina As String
  Dim rstp As Recordset
  Dim rstcr As Recordset
  Dim rstcomandesextres As Recordset
  etactualitzant.visible = True: DoEvents
  If vordre = "muntats" Then vmuntats = True: vordre = "numtreball"
  Set dbplanificacio = OpenDatabase(rutadelfitxer(cami) + "planificaciooperaris.mdb")
  llistacomandes.Clear
  llistatintes.Clear
  Set rstcomandesextres = dbcomandes.OpenRecordset("select  activadaatintes,comanda from comandes_extres")
  If vordre = "inactives" Then
      Set rstc = dbcomandes.OpenRecordset("select * from comandes where numtreball<>null and numtreball<>0 and  proximaseccio='E'")
     Else: Set rstc = dbcomandes.OpenRecordset("select * from comandes where numtreball<>null and numtreball<>0 and  (proximaseccio='E' or proximaseccio='I') order by " + vordre)
  End If
  
  While Not rstc.EOF
      If vordre = "inactives" Then
         rstcomandesextres.FindFirst "comanda=" + atrim(rstc!comanda)
         If rstcomandesextres.NoMatch Then GoTo cont
         If rstcomandesextres!activadaatintes = True Then GoTo cont
           Else
            If rstc!proximaseccio <> "I" Then
              rstcomandesextres.FindFirst "comanda=" + atrim(rstc!comanda)
              If rstcomandesextres.NoMatch Then GoTo cont
              If rstcomandesextres!activadaatintes = False Then GoTo cont
            End If
      End If
      If veurenomescomandestintesnoinplacsa Then
        Set rsttintes = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(rstc!numtreball) + " and ordremodificacio=" + atrim(rstc!numordremodificacio) + " order by ordretinter")
        While Not rsttintes.EOF
           If InStr(1, atrim(rsttintes!color), "PRIMAR") > 0 Or atrim(rsttintes!color) = "" Then GoTo proximatinta
           Set rstcolor = dbtintes.OpenRecordset("select idtinta from tintes where codi='" + atrim(rsttintes!coditinta) + "'")
           If Not rstcolor.EOF Then
                Set rstcolor = dbtintes.OpenRecordset("select * from tintesreferencies where idtinta=" + atrim(rstcolor!idtinta) + " and nomproveidor<>'INPLACSA'")
                If Not rstcolor.EOF Then GoTo afegirtinta
           End If
proximatinta:
           rsttintes.MoveNext
        Wend
        GoTo cont
      End If
afegirtinta:
      Set rstcr = dbtintes.OpenRecordset("select comanda from comandesrevisadesatintes where comanda=" + atrim(rstc!comanda))
      If Not rstcr.EOF Then If Not nocomprovarcomandesrevisades Then GoTo cont
      
      Set rstp = dbplanificacio.OpenRecordset("select ordre,maquina from planificacioimp where comanda=" + atrim(rstc!comanda))
      If Not rstp.EOF Then
            vordreimaquina = IIf(rstp!ordre > 0 And rstp!ordre < 999, atrim(rstp!ordre), "___") + IIf(rstp!maquina = 7, "FW", IIf(rstp!maquina = 9, "F2", ""))
            If comboselimp = "La FW" And rstp!maquina <> 7 Then GoTo cont
            If comboselimp = "La F2" And rstp!maquina <> 9 Then GoTo cont
            If vordreimaquina = "___" Then vordreimaquina = "N/S      "
         Else: vordreimaquina = "N/S      "
      End If
      vlinia = "" + justificar(vordreimaquina, 7, "D") + " "
      vlinia = vlinia + justificar(atrim(rstc!comanda) + IIf(rstc.impressio = "N" Or rstc.impressio = "M", "N", " ") + IIf(Not rstcr.EOF, "*", " "), 8, "E")
      vlinia = vlinia + justificar(estamuntada(rstc!comanda), 2, "D")
      vlinia = vlinia + justificar(atrim(rstc!numtreball) + "/" + atrim(rstc!numordremodificacio), 9, "D")
      vlinia = vlinia + justificar(Format(rstc!cantitatex, "#,##0") + " Mts  ", 15, "D")
      vlinia = vlinia + justificar(buscarnomclient(rstc!client) + "  ", 38, "E")
      vlinia = vlinia + justificar(atrim(rstc!marcailinia), 50, "E")
      If vmuntats Then If estaamuntadora(rstc!comanda) <> "S" Then GoTo cont
      llistacomandes.AddItem vlinia
      llistacomandes.ItemData(llistacomandes.NewIndex) = rstc!comanda
cont:
      rstc.MoveNext
  Wend
  etactualitzant.visible = False
  Set rstc = Nothing
  Set dbplanificacio = Nothing
  Set rstp = Nothing
  Set rstcr = Nothing
  Set rstcolor = Nothing
  Set rsttintes = Nothing
End Sub
Function buscarnomclient(vcodiclient As Long) As String
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select nom from clients where codi=" + atrim(vcodiclient))
   If Not rst.EOF Then buscarnomclient = atrim(rst!nom)
   Set rst = Nothing
End Function
Function estaamuntadora(numc As Double, Optional vmaquina As String) As String
   Dim rst As Recordset
   estaamuntadora = "N"
   Set rst = dbbaixes.OpenRecordset("SELECT comanda,nummaquina  FROM muntadora_ordremuntatge where comanda=" + atrim(cadbl(numc)) + ";")
   If Not rst.EOF Then estaamuntadora = "S": vmaquina = rst!nummaquina: GoTo fi
   Set rst = dbbaixes.OpenRecordset("select * from planificacio_reclamades where numcomanda=" + atrim(numc))
   If Not rst.EOF Then estaamuntadora = "R" + IIf(rst!reactivada, "a", "")
   Set rst = dbcomandes.OpenRecordset("select passaraimpresores from comandes_extres where comanda=" + atrim(numc), , ReadOnly)
   If Not rst.EOF Then If rst!passaraimpresores = 0 Then estaamuntadora = estaamuntadora + "!"
   If Len(estaamuntadora) = 1 Then estaamuntadora = estaamuntadora + "-"
fi:
   Set rst = Nothing
End Function
Function estamuntada(numc As Double) As String
   Dim rst As Recordset
   estamuntada = "N"
   Set rst = dbbaixes.OpenRecordset("SELECT muntadoratot.comanda  FROM comandes INNER JOIN muntadoratot ON comandes.comanda = muntadoratot.comanda WHERE (((muntadoratot.acabada)=True) AND ((comandes.proximaseccio)='I') and muntadoratot.comanda=" + atrim(cadbl(numc)) + ");")
   If Not rst.EOF Then estamuntada = "S"
   Set rst = Nothing
End Function


Private Sub Command43_Click()
  If Command43.caption = "Fes clic per comfirmar" Then filtrar_formules_ambcomponents: Exit Sub
  Command43.caption = "Fes clic per comfirmar"
  Check1(3).Value = 1
  'checknomesseleccionats.visible = True
  Command43.BackColor = &HFF00&
  carregar_componentalallista
  dllistadecomponents.visible = True
  dllistadecomponents.SetFocus
  
End Sub
Sub carregar_componentalallista()
  Dim rst As Recordset
  dbtintes.Execute "update componentsbase set esbase=' ' where esbase is null" 'trec el valor null de esbase per si de cas
  Set rst = dbtintes.OpenRecordset("select * from componentsbase")
  dllistadecomponents.Clear
  While Not rst.EOF
    If Check1(3).Value = 1 And atrim(rst!esbase) = "" Then GoTo proxima
    dllistadecomponents.AddItem atrim(rst!nomcomponent)
    dllistadecomponents.ItemData(dllistadecomponents.NewIndex) = cadbl(rst!idcomponent)
proxima:
    rst.MoveNext
  Wend
End Sub
Sub amagarllistacomponents()
dllistadecomponents.visible = False
checknomesseleccionats.visible = False
  Command43.BackColor = &H8000000F
  Command43.caption = "Filtrar formula amb component concret"
End Sub
Function convertirentanpercent(v, camp) As String
   Dim vcont As Double
   Dim v1 As Double
   Dim v2 As Double
   vcodi = atrim(v + "%")
   While vcont < Len(vcodi)
      If Mid(vcodi, vcont + 1, 1) = "-" Then
        vcodi = Mid(vcodi, 1, Len(vcodi) - 1)
        v1 = cadbl(Mid(vcodi, 1, vcont))
        If Len(vcodi) >= vcont + 2 Then
           v2 = cadbl(Mid(vcodi, vcont + 2))
        End If
        GoTo sortir
          Else
            If Mid(vcodi, vcont + 1, 1) = "%" Then v1 = cadbl(Mid(vcodi, 1, vcont)): GoTo sortir
      End If
      vcont = vcont + 1
   Wend
sortir:
  If v1 > 0 And v2 > 0 Then convertirentanpercent = IIf(convertirentanpercent <> "", " and ", "") + "(idcomponente=" + atrim(camp) + " and [%decomponent] between " + atrim(passaradecimalpunt(atrim(v1))) + " and " + atrim(passaradecimalpunt(atrim(v2))) + ")"
  If v1 > 0 And v2 = 0 Then convertirentanpercent = IIf(convertirentanpercent <> "", " and ", "") + "(idcomponente=" + atrim(camp) + " and [%decomponent] = " + atrim(passaradecimalpunt(atrim(v1))) + ")"

End Function
Sub filtrar_formules_bases(vfiltrar As String)
 Dim vsql As String
 Dim rst As Recordset
 Dim rst2 As Recordset
 Dim vsql2 As String
 
 vsql = " idformula in (SELECT detallformules.IDFormula From detallformules WHERE detallformules.IdComponente In (" + vfiltrar + ") GROUP BY detallformules.IDFormula HAVING Count(*)>=0;)"
' " idformula in (SELECT detallformules.IDFormula From detallformules WHERE detallformules.IdComponente not In (" + vfiltrar + ") GROUP BY detallformules.IDFormula HAVING Count(*)>=0;)"
 Set rst = dbtintes.OpenRecordset("select * FROM formules RIGHT JOIN FormulesAmbLlaunesactives ON formules.codiformula = FormulesAmbLlaunesactives.numformula where " + vsql) '+ " AND ((FormulesAmbLlaunesactives.CuentaDeid)>0);")
 While Not rst.EOF
    
    vsql2 = "SELECT esbase,detallformules.IDFormula, detallformules.IdComponente, Componentsbase.nomcomponent FROM detallformules LEFT JOIN Componentsbase ON detallformules.IdComponente = Componentsbase.idcomponent WHERE detallformules.IDFormula=" + atrim(rst!idformula) + " AND [esbase]<>' ' AND detallformules.IdComponente Not In (" + vfiltrar + ");"
   ' Clipboard.Clear
  ' Clipboard.SetText vsql2
    Set rst2 = dbtintes.OpenRecordset(vsql2)
    If rst2.EOF Then vFiltreSI = vFiltreSI + IIf(vFiltreSI = "", "", ",") + atrim(rst!idformula)
    '     vFiltreSI = vFiltreSI + IIf(vFiltreSI = "", "", ",") + atrim(rst!idformula)
    rst.MoveNext
 Wend
 'MsgBox vfiltreSI
 If vFiltreSI <> "" Then
  dataformules.RecordSource = "select * FROM formules LEFT JOIN FormulesAmbLlaunesactives ON formules.codiformula = FormulesAmbLlaunesactives.numformula where idformula in (" + vFiltreSI + ")"
   Else: dataformules.RecordSource = "select * FROM formules LEFT JOIN FormulesAmbLlaunesactives ON formules.codiformula = FormulesAmbLlaunesactives.numformula where idformula=-1"
 End If
 dataformules.Refresh
 reixaformules.Refresh
End Sub
Sub filtrar_formules_ambcomponents()
    Dim vfiltrepercentatge As String
    Dim vfiltrar As String
    Dim vcount As Integer
    Dim vvalortanpercent As String
    
    amagarllistacomponents
    For i = 0 To dllistadecomponents.ListCount - 1
      If dllistadecomponents.Selected(i) Then
        'vfiltrar = vfiltrar + IIf(vfiltrar <> "", " and ", "") + "DetallFormules.IdComponente = " + atrim(dllistadecomponents.ItemData(i))
        vfiltrar = vfiltrar + IIf(vfiltrar <> "", ",", "") + atrim(dllistadecomponents.ItemData(i))
    ''    vvalortanpercent = InputBox("Entra el tan% que vols pel component " + dllistadecomponents.List(i) + Chr(10) + "Ex: 5-10", "Criteris")
    ''    vfiltrepercentatge = vfiltrepercentatge + IIf(vfiltrepercentatge <> "", " AND ", "") + convertirentanpercent(vvalortanpercent, dllistadecomponents.ItemData(i))
        vcount = vcount + 1
      End If
    Next i
    If vfiltrar = "" Then Exit Sub
    If Check1(3).Value = 1 Then filtrar_formules_bases vfiltrar: GoTo fi
    If checknomesseleccionats = 1 Then
       'vfiltrar = " idformula in (SELECT detallformules.IDFormula From detallformules WHERE detallformules.IdComponente In (" + vfiltrar + ") GROUP BY detallformules.IDFormula HAVING Count(*)=" + atrim(vcount) + ";)"
       vfiltrar = "idformula in(SELECT detallformules1.IDFormula FROM detallformules AS detallformules1 WHERE " + vfiltrepercentatge + IIf(vfiltrepercentatge <> "", " and ", "") + "(((detallformules1.IdComponente) In (" + vfiltrar + ")) AND (((SELECT Count(*) From detallformules Where detallformules.IDformula = detallformules1.IDformula GROUP BY detallformules.IDFormula))=" + atrim(vcount) + ")) GROUP BY detallformules1.IDFormula HAVING (((Count(*))=" + atrim(vcount) + "));)"
          Else:
            If Check1(3).Value = 1 Then vcount = 1
            vfiltrar = " idformula in (SELECT detallformules.IDFormula From detallformules WHERE " + vfiltrepercentatge + IIf(vfiltrepercentatge <> "", " and ", "") + " detallformules.IdComponente In (" + vfiltrar + ") GROUP BY detallformules.IDFormula HAVING Count(*)>=" + atrim(vcount) + ";)"
    End If
 ' Clipboard.Clear
 ' Clipboard.SetText dataformules.RecordSource
    dataformules.RecordSource = "select * FROM formules LEFT JOIN FormulesAmbLlaunesactives ON formules.codiformula = FormulesAmbLlaunesactives.numformula where " + vfiltrar + " AND ((FormulesAmbLlaunesactives.CuentaDeid)>0);"
    dataformules.Refresh
    reixaformules.Refresh
fi:
End Sub
Private Sub Command44_Click()
Dim vsubconsulta As String
  If tintes.Recordset.EOF Then Exit Sub
  With tintes.Recordset
  If Mid(!referenciacolor + "  ", 1, 2) = "P-" Then Exit Sub
  vsubconsulta = " mid(referenciacolor+'  ',1,2)<>'P-' and  tintes.idfamilia=" + atrim(cadbl(!idfamilia)) + " and tintes.idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + " and tintes.idfamcolor=" + atrim(cadbl(!idfamcolor)) + " and tintes.idsubfamcolor=" + atrim(cadbl(!idsubfamcolor))
  'vsubconsulta = " tintes.idfamilia=" + atrim(cadbl(!idfamilia)) + " and tintes.idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + " and tintes.idfamcolor=" + atrim(cadbl(!idfamcolor)) + " and tintes.idsubfamcolor=" + atrim(cadbl(!idsubfamcolor))
  tintes.RecordSource = "select * from tintes_tot where " + vsubconsulta + " order by descripcio"
  tintes.Refresh
  End With
End Sub

Private Sub Command45_Click()
 Dim rst As Recordset
 carregar_reixaformulacio
 Set rst = dbclixes.OpenRecordset("select coditinta from tintes where id_tinter=" + atrim(llistatintes.ItemData(llistatintes.ListIndex)))
 If rst.EOF Then Exit Sub
 tintes.RecordSource = "tintes_tot"
 tintes.Refresh
 tintes.Recordset.FindFirst "codi='" + atrim(rst!coditinta) + "'"
 If Not tintes.Recordset.NoMatch Then
    Set rst = dbtintes.OpenRecordset("select * from tintesformules where idtinta=" + atrim(tintes.Recordset!idtinta) + " order by predeterminada")
    If Not rst.EOF Then
       Set rst = dbtintes.OpenRecordset("select * from formules where codiformula='" + atrim(rst!numformula) + "'")
       If Not rst.EOF Then
          escullir_formula_semblant_filtre atrim(rst!codiformula), atrim(rst!descripcioformula)
          pestanyes.Tab = 2
          DoEvents
          pestanyesforumes.Tab = 2
       End If
    End If
 End If
 Set rst = Nothing
End Sub
Sub possar_lesseleccionades(vseleccionades As String)
  Dim i As Integer
  vseleccionades = ""
  For i = 0 To llistatintes.ListCount - 1
     If llistatintes.Selected(i) Then vseleccionades = vseleccionades + "#" + atrim(i) + "#"
  Next i
End Sub
Private Sub Command46_Click()
  generar_fitxer_csv_o_refrescar
End Sub
Sub generar_fitxer_csv_o_refrescar(Optional vrefrescar As Boolean)
 Dim i As Integer
   Dim col As Integer
   Dim vseleccionades As String
   possar_lesseleccionades vseleccionades
   llistatintes.Clear
   On Error GoTo errorcrearfitxer
   Open "c:\temp\~exportaciotintes.csv" For Append As #1
   On Error GoTo 0
   'For i = 0 To llistacomandes.ListCount - 1
   '   If llistacomandes.Selected(i) Then
    For col = reixacomandes.Row To reixacomandes.RowSel
          actualitza_llistatintes reixacomandes.TextMatrix(col, 0), IIf(vrefrescar, False, True), vseleccionades, reixacomandes.TextMatrix(col, 3)
    Next col
   '   End If
   'Next i
   Close #1
   If Not vrefrescar And existeix("c:\temp\~exportaciotintes.csv") Then obrir_document "c:\temp\~exportaciotintes.csv"
Exit Sub
errorcrearfitxer:
   MsgBox "Error al crear el fitxer d'Excel, mira que no el tinguis obert i torna-ho a provar", vbCritical, "Error"
End Sub

Private Sub Command47_Click()
   Dim v As String
   Dim vnum As String
   Dim i As Integer
   Dim j As Integer
   Dim vnotrobats As String
   v = InputBox("Entra les comandes que vols buscar." + Chr(10) + "Ex:  175500 / 152000 141000", "Buscar comanda")
   For i = 0 To llistacomandes.ListCount - 1
      llistacomandes.Selected(i) = False
   Next i
   i = 1
   While i < Len(v) + 1
     If IsNumeric(Mid(v, i, 1)) Or Mid(v, i, 1) = "." Then
        vnum = vnum + Mid(v, i, 1)
          Else
buscar:
            For j = 0 To llistacomandes.ListCount - 1
              If llistacomandes.ItemData(j) = cadbl(vnum) Then llistacomandes.Selected(j) = True: GoTo cont
            Next j
            vnotrobats = vnotrobats + " " + vnum
cont:
            vnum = ""
     End If
     i = i + 1
     If i = Len(v) + 1 Then
        GoTo buscar
     End If
   Wend
   If vnotrobats <> "" Then MsgBox "Comandes que no he trobat a la llista:" + Chr(10) + vnotrobats, vbCritical, "No trobades"
End Sub

Private Sub Command48_Click()
  Dim rst As Recordset
  Dim rstc As Recordset
  For i = 0 To llistatintes.ListCount - 1
      If llistatintes.Selected(i) Then
          Set rst = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(llistatintes.ItemData(i)))
          If Not rst.EOF Then
             passar_extensio_a_feta cadbl(rst!id_treball), cadbl(rst!ordremodificacio), cadbl(rst!coditinta), atrim(rst!color)
          End If
      End If
  Next i
  Set rstc = Nothing
  Set rst = Nothing
  generar_fitxer_csv_o_refrescar True
End Sub
Sub passar_extensio_a_feta(vnumtreball As Double, vmodificacio As Double, vcoditinta As Double, vnomcolor As String)
   Dim rstc As Recordset
'    Set rstc = dbtintes.OpenRecordset("select * from extensions where idtreball=" + atrim(vnumtreball) + " and coditinta=" + atrim(vcoditinta))
'    If rstc.EOF Then
'       dbtintes.Execute "insert into extensions (idtreball,ordremodificacio,coditinta) values (" + atrim(vnumtreball) + "," + atrim(vmodificacio) + "," + atrim(vcoditinta) + ")"
'         Else
'           If MsgBox("La tinta " + atrim(vnomcolor) + " ja te una extensió assignada per aquest treball versió " + atrim(rstc!ordremodificacio) + Chr(10) + "Vols eliminar-la?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
'              dbtintes.Execute "delete * from extensions where idtreball=" + atrim(vnumtreball) + " and ordremodificacio=" + atrim(vmodificacio) + " and coditinta=" + atrim(vcoditinta)
'           End If
'    End If
End Sub
Private Sub Command49_Click()
   crear_temporal_comandesactives
End Sub

Sub crear_temporal_comandesactives()
On Error GoTo errorcrearfitxer
   Open "c:\temp\~exportaciotintes.csv" For Output As #1
   Print #1, "Comanda;Tinta;Anilox;Nom de la tinta;Llauna(Sit);Kg teorics"
   Close #1
   Exit Sub
errorcrearfitxer:
   MsgBox "Error al crear el fitxer d'Excel de comandes actives, mira que no el tinguis obert i torna-ho a provar", vbCritical, "Error"
End Sub
Private Sub Command5_Click()
  If datahistoria.Recordset.RecordCount > 0 Then
     If UCase(InputBox("No pots borrar aquesta llauna si ja te historia." + Chr(10) + "Entra la contrasenya per eliminar-la", "Error")) = "INPLACSA" Then GoTo eliminarla
     Exit Sub
       Else:
eliminarla:
         If MsgBox("Segur que vols borrar aquesta llauna?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
           datahistoria.Refresh
           While Not datahistoria.Recordset.EOF
                dbtintes.Execute "delete * from historiallaunalots where idhistoria=" + atrim(datahistoria.Recordset!id)
                datahistoria.Recordset.Delete
                datahistoria.Recordset.MoveNext
           Wend
           dbtintes.Execute "delete * from historialsituacions where numllauna='" + atrim(datallaunes.Recordset!numllauna) + "'"
           dbtintes.Execute "delete * from llaunes where numllauna='" + atrim(datallaunes.Recordset!numllauna) + "'"
          ' datallaunes.Recordset.Delete
           datallaunes.Refresh
         End If
       
  End If
End Sub

Sub nomdelafamiliadelcomponent(vlot As String, vnomfam1 As String, vnomfam2 As String)
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from llaunes where numllauna='" + atrim(vlot) + "'")
   If Not rst.EOF Then
       Set rst = dbtintes.OpenRecordset("select * from tintes_tot where idtinta=" + atrim(rst!idtinta))
       If Not rst.EOF Then
           vnomfam2 = treure_apostruf(atrim(rst!descripciofam))
           If InStr(atrim(rst!descripciosubfam), " FROTE") > 0 Then vnomfam2 = treure_apostruf(atrim(rst!descripciosubfam))
           vnomfam1 = treure_apostruf(atrim(rst!referenciacolor) + " " + atrim(rst!descripcioserie))
       End If
   End If
   Set rst = Nothing
End Sub
Function buscarlotinjector(vnomcomponent As String) As String
   Dim vhashtag As String
   Dim rst As Recordset
   If InStr(1, vnomcomponent, "#I") > 0 Then
      vhashtag = agafarhashtag(Mid(vnomcomponent, InStr(1, vnomcomponent, "#I")))
      vnomcomponent = substituir(vnomcomponent, "#I" + vhashtag, "")
   End If
   buscarlotinjector = vhashtag
   If buscarlotinjector <> "" Then
     Set rst = dbtintes.OpenRecordset("SELECT* FROM detallnumeroslotsbase where idcomponent in (select idcomponent from componentsbase where numdosificador=" + atrim(buscarlotinjector) + ") order by data desc;")
     If Not rst.EOF Then buscarlotinjector = atrim(rst!numerodelot)
     Set rst = Nothing
   End If
End Function
Function agafarhashtag(vnom As String) As String
   Dim i As Byte
   i = 3
   While IsNumeric(Mid(vnom, i, 1))
      agafarhashtag = agafarhashtag + Mid(vnom, i, 1)
      i = i + 1
   Wend
End Function

Private Sub Command50_Click()
   Dim vnumdosoficador As String
   Dim rst As Recordset
   Dim oapp As CRAXDDRT.Application
   Dim oreport As CRAXDDRT.Report
   Dim vnomfam1 As String
   Dim vnomfam2 As String
   Dim vnomcomponent As String
   Dim vhiharf As Boolean
   Dim vnumllaunarelacionat As String
   Dim vi1 As String
   Dim vi2 As String
   Dim vi3 As String
   Dim vi4 As String
   Dim vtmp1 As String
   Dim vtmp2 As String
   If datacomponents.Recordset.EOF Then Exit Sub
   Set oapp = New CRAXDDRT.Application
   Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "etiquetanumdosificador.rpt", 1)
   If Not datalotsbase.Recordset.EOF Then
      vnumllaunarelacionat = datalotsbase.Recordset!numerodelot
      nomdelafamiliadelcomponent vnumllaunarelacionat, vnomfam1, vnomfam2
   End If
   vnomcomponent = UCase(atrim(datacomponents.Recordset!nomcomponent))
  If vnumllaunarelacionat = "" Then vnumllaunarelacionat = ""
   If InStr(1, vnomcomponent, "#") > 0 And vnumllaunarelacionat = "0" Then
      vi1 = buscarlotinjector(vnomcomponent)
      vi2 = buscarlotinjector(vnomcomponent)
      vi3 = buscarlotinjector(vnomcomponent)
      vi4 = buscarlotinjector(vnomcomponent)
      nomdelafamiliadelcomponent vi1, vtmp1, vtmp2
      If vi1 <> "" Then
        If InStr(atrim(vtmp2), " FROTE") > 0 Then
          vhiharf = True
           Else: vnomfam1 = vtmp1: vnomfam2 = vtmp2
        End If
      End If
      nomdelafamiliadelcomponent vi2, vtmp1, vtmp2
      If vi2 <> "" Then
        If InStr(atrim(vtmp2), " FROTE") > 0 Then
          vhiharf = True
'           Else: vnomfam1 = vtmp1: vnomfam2 = vtmp2
        End If
      End If
      nomdelafamiliadelcomponent vi3, vtmp1, vtmp2
      If vi3 <> "" Then
        If InStr(atrim(vtmp2), " FROTE") > 0 Then
          vhiharf = True
 '          Else: vnomfam1 = vtmp1: vnomfam2 = vtmp2
        End If
      End If
      nomdelafamiliadelcomponent vi4, vtmp1, vtmp2
      If vi4 <> "" Then
        If InStr(atrim(vtmp2), " FROTE") > 0 Then
          vhiharf = True
  '         Else: vnomfam1 = vtmp1: vnomfam2 = vtmp2
        End If
      End If
      If vhiharf Then vnomfam2 = "TINTA EXTERIOR"
   End If
   If InStr(1, vnomcomponent, "#") > 0 Then vnomcomponent = atrim(Mid("  " + vnomcomponent, 1, InStr(1, "  " + vnomcomponent, "#") - 1))
   oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "tintes.mdb"
   oreport.FormulaFields.GetItemByName("nomdosificador").Text = "'" + vnomcomponent + "'"
   oreport.FormulaFields.GetItemByName("nomfamilia1").Text = "'" + atrim(vnomfam1) + "'"
   oreport.FormulaFields.GetItemByName("nomfamilia2").Text = "'" + atrim(vnomfam2) + "'"
   If vhiharf Then oreport.Sections("PageHeaderSection1").ReportObjects.Item("nomfamilia11").BackColor = QBColor(10)
   
   vnumdosoficador = "I" + Format(cadbl(datacomponents.Recordset!numdosificador), "000")
'  GENERA EL CODI DE BARRES DEL NUMERO DE TREBALL
   escriure_ini "Tbarcode", "nomfitxer", "c:\temp\~vnumdosificador.bmp", "generartbarcode.ini"
   escriure_ini "Tbarcode", "pixelsample", "2000", "generartbarcode.ini"
   escriure_ini "Tbarcode", "pixelsalt", "500", "generartbarcode.ini"
   escriure_ini "Tbarcode", "text", atrim(vnumdosoficador), "generartbarcode.ini"
   escriure_ini "Tbarcode", "printdatatext", "1", "generartbarcode.ini"
   escriure_ini "Tbarcode", "tipusbarcode", "62", "generartbarcode.ini"
   Shell llegir_ini("General", "rutallistats", "comandes.ini") + "generarimatgedecodidebarres.exe"
   
   Set rst = dbtintes.OpenRecordset("select * from contadors")
   rst.Edit
   r = copiafoto("c:\temp\~vnumdosificador.bmp", rst!imatge)
   rst.Update
   wait 2
   Set rst = Nothing
   oreport.DiscardSavedData
 ' If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
  '  Else
  '    oreport.DisplayProgressDialog = False
 '     oreport.PrintOut False, 1
 ' End If
   
   
End Sub

Private Sub Command51_Click()
     Dim vtreball As Double
     Dim vmodificacio As Double
     Dim rstc As Recordset
     Dim vnomtinta As String
     Dim vcoditinta As Double
     vtreball = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("NºTreball")))
     vmodificacio = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("Versió")))
     vtreball = cadbl(InputBox("Entra el numero de treball que vols escullir la tinta.", "Escull el treball", atrim(vtreball)))
     If vtreball = 0 Then Exit Sub
     vmodificacio = cadbl(InputBox("Entra la versió d'aquest treball", "Escull la versió", atrim(vmodificacio)))
     If vmodificacio = 0 Then Exit Sub
     'Set rstc = dbcomandes.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(vcomanda))
     'If rstc.EOF Then MsgBox "Aquesta comanda no existeix", vbCritical, "Error": Exit Sub
'     triarlatinta vcoditinta, vnomtinta, cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio)
     vnomtinta = "."
     While atrim(vnomtinta) <> ""
       triarlatinta vcoditinta, vnomtinta, vtreball, vmodificacio
       If vnomtinta <> "" Then demanardadesextensio vtreball, vmodificacio, vcoditinta, vnomtinta
     Wend
     'passar_extensio_a_feta vtreball, vmodificacio, vcoditinta, vnomtinta
     
End Sub
Sub demanardadesextensio(vtreball As Double, vmodificacio As Double, vcoditinta As Double, vnomtinta As String)
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("select * from extensions_treballsrelacionats where numtreball=" + atrim(vtreball) + " and numordremodificacio=" + atrim(vmodificacio) + " and coditinta=" + atrim(vcoditinta))
  Unload formextensions
  Load formextensions
  formextensions.etnomtinta = vnomtinta
  formextensions.etnomtinta.tag = vcoditinta
  formextensions.ettreballactual.tag = vtreball
  formextensions.ettreballactual.WhatsThisHelpID = vmodificacio
  formextensions.ettreballactual.caption = "Nº Treball: " + atrim(vtreball) + " / " + atrim(vmodificacio)
  formextensions.carregar_extensio IIf(Not rst.EOF, atrim(rst!codiextensio), "")
  formextensions.Show 1
End Sub
Sub triarlatinta(vcoditinta As Double, vnomtinta As String, vtreball As Double, vmodificacio As Double)
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  vnomtinta = ""
  vcoditinta = 0
  sql = "SELECT coditinta,color From tintes WHERE (((tintes.color)<>'') AND "
  sql = sql + " ((tintes.id_tinter) In (SELECT Tintes.tinterlinkambid_treball From tintes WHERE tintes.tinterlinkambid_treball>0 and  Tintes.id_treball=" + atrim(vtreball) + " AND Tintes.ordremodificacio=" + atrim(vmodificacio) + ") Or "
  sql = sql + " (tintes.id_tinter) In (SELECT Tintes.id_tinter FROM Tintes wHERE  tintes.tinterlinkambid_treball<=0 and Tintes.id_treball=" + atrim(vtreball) + " AND Tintes.ordremodificacio=" + atrim(vmodificacio) + ")));"
  'Clipboard.Clear
  'Clipboard.SetText sql
  'sql = "SELECT coditinta,color from tintes where id_treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vmodificacio) + " and tinterlinkambid_treball=0"
 ' Clipboard.Clear
 ' Clipboard.SetText sql + were
  Unload formseleccio
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "clixesnous.mdb"
  formseleccio.Data1.RecordSource = sql + were
  formseleccio.width = 7000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 600
  formseleccio.DBGrid2.Columns(1).width = 4000
  formseleccio.Show 1
  
  If seleccioret = 1 Then
    On Error GoTo fi
    If Not formseleccio.Data1.Recordset.EOF Then
        vnomtinta = atrim(formseleccio.Data1.Recordset!color)
        vcoditinta = cadbl(formseleccio.Data1.Recordset!coditinta)
    End If
  End If
  
  If seleccioret = 9 Then
    vnomtinta = ""
    vcoditinta = 0
  End If
 '  Data1.Recordset!client = Text2.Text
 '  nomclient.Caption = atrim(formseleccio.Data1.Recordset!nom)
  
 ' End If
fi:
  Unload formseleccio
End Sub


Private Sub Command52_Click()
  Dim vnumc As String
  amagar_boto_imprimircomanda
  If Command52.Enabled = False Then Exit Sub
  escriure_ini "Baixes", "imprimircomanda", "0", "comandes.ini"
  'If llistacomandes.ListIndex < 0 Then Exit Sub
  If reixacomandes.Row > 0 Then vnumc = reixacomandes.TextMatrix(reixacomandes.Row, col)
  vnumc = cadbl(InputBox("Entra la comanda o el treball que vols veure.", "Visualitzar la comanda", vnumc))
  If vnumc = 0 Then Exit Sub
  If vnumc < 100000 Then vnumc = buscarlacomandadeltreball(cadbl(vnumc))
  If vnumc = 0 Then Exit Sub
  veurelacomanda cadbl(vnumc) 'llistacomandes.ItemData(llistacomandes.ListIndex)
  'veureelpdf llistacomandes.ItemData(llistacomandes.ListIndex)
  'wait 3
  'escriure_ini "Baixes", "imprimircomanda", "0", "comandes.ini"
End Sub
Function buscarlacomandadeltreball(vnumtreball As Double)
   Dim rst As Recordset
   Set rst = dbcomandes.OpenRecordset("select comanda from comandes where numtreball=" + atrim(vnumtreball) + " order by comanda desc")
   If Not rst.EOF Then buscarlacomandadeltreball = rst!comanda
End Function

Private Sub Command52_LostFocus()
  escriure_ini "Baixes", "imprimircomanda", "0", "comandes.ini"
End Sub

'Private Sub Command53_Click()
'  framebotons.visible = False
'  llistacomandes.BackColor = Command53.BackColor
'  actualitza_llistacomandes "inactives", , IIf(checknoinplacsa.Value = 1, True, False)
'End Sub

Private Sub Command54_Click()
 Dim i As Integer
   llistatintes.Clear
   For i = 0 To llistacomandes.ListCount - 1
      If llistacomandes.Selected(i) Then
        dbcomandes.Execute "update comandes_extres set activadaatintes=true where comanda=" + atrim(llistacomandes.ItemData(llistacomandes.ListIndex))
      End If
   Next i
End Sub

Private Sub Command55_Click()
    demanarmeslots True
End Sub

Sub demanarmeslots(Optional noensenyarmissatge As Boolean)
    Dim vvalors As String
    Dim vnumerodelot As String
    Dim i As Integer
    Dim vkg As Double
    
    vnumerodelot = " "
    Load formmeslots
    If noensenyarmissatge Then formmeslots.tag = "noautocarregar"
    formmeslots.Show 1
    'If formmeslots.llistallaunes.ListCount = 0 Then Exit Sub
    If Mid(formmeslots.vnumlot.tag + "  ", 1, 1) = "-" Then treure_lot Mid(atrim(formmeslots.vnumlot.tag), 2): GoTo fi
    i = 0
    While i < formmeslots.llistallaunes.ListCount
        'vnumerodelot = atrim(InputBox("Entra el numero de Lot o llauna que vols afegir a aquesta llauna.", "Afegir lot"))
        vnumerodelot = formmeslots.llistallaunes.List(i)
        vkg = cadbl(formmeslots.llistallaunes.ItemData(i)) / 100
        If Mid(vnumerodelot + " ", 1, 1) = "*" Then vnumerodelot = atrim(Mid(vnumerodelot + " ", 2))
        If InStr(1, vnumerodelot, ",") > 0 Or InStr(1, vnumerodelot, " ") > 0 Or InStr(1, vnumerodelot, ";") > 0 Or InStr(1, vnumerodelot, ":") > 0 Then
           MsgBox "Hi han simbols no vàlids en el numero de lot ->  " + vnumerodelot, vbCritical, "Error": vnumerodelot = ""
        End If
        If vnumerodelot = "" Then GoTo fi
        datahistoria.Recordset.FindFirst "tipusmoviment='C' or tipusmoviment='K'"
        If Not datahistoria.Recordset.NoMatch Then
            vvalors = " (" + atrim(datahistoria.Recordset!id) + ",0,'" + vnumerodelot + "',0," + passaradecimalpunt(atrim(vkg)) + ")"
            dbtintes.Execute "insert into historiallaunalots (idhistoria,idcomponent,numlotbase,tanx100tinta,kgtinta) values " + vvalors
            Else: MsgBox "No hi ha una historia per relaciona el Nº de lot", vbCritical, "Error"
        End If
        i = i + 1
    Wend
fi:
    Unload formmeslots
End Sub
Sub treure_lot(vnumlot As String)
   Dim rst As Recordset
   Dim vnumllauna As String
   vnumllauna = atrim(datallaunes.Recordset!numllauna)
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.numllauna, historiallaunalots.numlotbase FROM (historiallauna RIGHT JOIN Llaunes ON historiallauna.idnumllauna = Llaunes.id) LEFT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria WHERE Llaunes.numllauna='" + vnumllauna + "' and idcomponent=0 and numlotbase='" + vnumlot + "'")
   If Not rst.EOF Then
         vnumlotnou = atrim(InputBox("Escriu el nou numero de LOT que substitueix el " + vnumlot, "NMou LOT"))
         If vnumlotnou <> "" Then
              rst.Edit
              rst!numlotbase = atrim(vnumlotnou)
              rst.Update
              MsgBox "El LOT " + vnumlot + " s'ha canviat pel Nº:" + vnumlotnou, vbInformation, "CANVI FET"
         End If
       Else: MsgBox "No he trobat aquest LOT en aquesta llauna."
   End If
   Set rst = Nothing
End Sub
Private Sub Command56_Click()
   werescomandes = ""
   Command63.tag = ""
   Command63.BackColor = &HFFFF&
   Command13.Enabled = False
   reixacomandes.visible = False
   poblar_reixa_comandes
   carregar_amples_reixa
   reixacomandes.visible = True
   If Screen.ActiveControl.Name = "Command56" Then carregar_liniadelareixaseleccionada
End Sub

Private Sub Command57_Click()
  Dim i As Integer
  Command56_Click
  For i = 0 To filtre.Count - 1
     If filtre(i).Text = "Gestionat?" Then filtre(i).Text = "N,S,C,F,P"  ': filtre_LostFocus i
     If filtre(i).Text = "Ès a muntadora" Then filtre(i).Text = "S"  ': filtre_LostFocus i
     If filtre(i).Text = "Muntada" Then filtre(i).Text = "S,N" ': filtre_LostFocus i
  Next i
  filtre(0).SetFocus
  filtre_LostFocus 0
End Sub

Private Sub Command58_Click()
  Dim i As Integer
  Command56_Click
  For i = 0 To filtre.Count - 1
     If filtre(i).Text = "Tintes fora" Then filtre(i).Text = "S":
     If filtre(i).Text = "Ès a muntadora" Then filtre(i).Text = "N"
  Next i
  filtre(0).SetFocus
  filtre_LostFocus 0
End Sub

Private Sub Command59_Click()
'Frame5(1).Top = llistatintes.Top
'Frame5(1).Left = Command59.Left - Frame5(1).width
  Frame5(1).visible = Not Frame5(1).visible
  If Frame5(1).visible Then
      If Command59.BackColor = 65535 Then
          If MsgBox("Vols copiar els valors de la versió anterior?", vbExclamation + vbDefaultButton2 + vbYesNo, "Atencio") = vbYes Then
                mirar_cuatricumia True
          End If
      End If
      escullir_maq_cuatricomia 1, True
      carregar_valors_cuatricomia
      kgxrecuperar(3).SetFocus
  End If
End Sub
Sub carregar_valors_cuatricomia(Optional vnumtreballiversio As String)
    'valorscuatricomia_treball
   Dim rst As Recordset
   Dim vnummaq As Double
   vnummaq = maquinaescullidaPerCuatricomia
   Frame5(1).tag = ""
   Frame5(1).caption = ""
   kgxrecuperar(3) = 0
    kgxrecuperar(4) = 0
    kgxrecuperar(5) = 0
    kgxrecuperar(6) = 0
    kgxrecuperar(19) = 0
    kgxrecuperar(20) = 0
    kgxrecuperar(21) = 0
    kgxrecuperar(22) = 0
   Combo(0).Clear: Combo(1).Clear: Combo(2).Clear: Combo(3).Clear
   If vnumtreballiversio = "" Then
        vnumtreballiversio = Trim(Abs(reixacomandes.TextMatrix(reixacomandes.Row, numcol("NºTreball")))) + "/" + reixacomandes.TextMatrix(reixacomandes.Row, numcol("Versió"))
        vnumtreballiversio = InputBox("Entra el numero de treball i la versió:", "Treball i versió de la cuatricomia", vnumtreballiversio)
   End If
   If vnummaq = 0 Then
    Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where numtreballiversio='" + vnumtreballiversio + "'")
    If Not rst.EOF Then vnummaq = rst!nummaq Else vnummaq = 9
   End If
   Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where nummaq=" + atrim(vnummaq) + " and numtreballiversio='" + vnumtreballiversio + "'")
   If rst.EOF And vnummaq = 1 Then Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where numtreballiversio='" + vnumtreballiversio + "'")
   If InStr(2, vnumtreballiversio, "/") > 0 Then
        Frame5(1).tag = vnumtreballiversio
        Frame5(1).caption = "Nivell cuatricomia " + vnumtreballiversio
        While Not rst.EOF
            possar_color_cuatricomia rst
            rst.MoveNext
        Wend
   End If
   Set rst = Nothing
   llegir_tolerancies_cuatricomia
End Sub

Private Sub Command6_Click()
   novallauna
End Sub
Sub editallauna(Optional numllauna As String)
   Dim nllauna As String
   Dim kgllauna As String
   If datallaunes.Recordset.EOF Or datallaunes.Recordset.BOF Then MsgBox "No hi ha cap llauna triada.", vbCritical, "Error": Exit Sub
  ' nllauna = InputBox("Escriu el nou nom de la llauna.", "Nova llauna", atrim(datallaunes.Recordset!numllauna))
  ' If atrim(nllauna) = "" Then Exit Sub
   'kgllauna = InputBox("De quants Kilos es aquesta llauna?", "Nova llauna", atrim(datallaunes.Recordset!capacitatmaxima))
  ' If cadbl(kgllauna) = 0 Then Exit Sub
   datallaunes.Recordset.Edit
   If datallaunes.Recordset.EditMode > 0 And numllauna <> "" Then
      datallaunes.Recordset.FindFirst "numllauna='" + numllauna + "'"
      datallaunes.Recordset.Edit
   End If
   'datallaunes.Recordset!numllauna = nllauna
   'datallaunes.Recordset!capacitatmaxima = cadbl(kgllauna)
   'datallaunes.Recordset!idtinta = tintes.Recordset!idtinta
   'datallaunes.Recordset.Update
End Sub
Function escullir_referenciaproveidor(idtinta As Double, vcapacitatllauna As Double) As Double
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("SELECT tintesreferencies.*, tipusbidons.* FROM tintesreferencies LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id Where idtinta = " + atrim(idtinta))
   If rst.EOF Then escullir_referenciaproveidor = 0: Exit Function
   rst.MoveLast
   If rst.RecordCount = 1 Then vcapacitatllauna = rst!capacitat: escullir_referenciaproveidor = rst![tintesreferencies.id]: Exit Function
   escullir_referenciaproveidor = esculliralgunareferenciaproveidor(idtinta, vcapacitatllauna)
End Function
Function esculliralgunareferenciaproveidor(idtinta As Double, vcapacitatllauna As Double) As Double
  esculliralgunareferenciaproveidor = 0
  Load formseleccio
  formseleccio.caption = "Selecciona referencia proveidor"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "SELECT tintesreferencies.id, tintesreferencies.referencia, tintesreferencies.nomproveidor, tipusbidons.capacitat, IIf(predeterminada,'*','') AS Predet FROM tintesreferencies INNER JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id "
  formseleccio.Data1.RecordSource = formseleccio.Data1.RecordSource + " WHERE (((tintesreferencies.idtinta)=" + atrim(idtinta) + ")) order by referencia desc;"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 1100
  formseleccio.DBGrid2.Columns(2).width = 3000
  formseleccio.DBGrid2.Columns(3).width = 1300
  formseleccio.DBGrid2.Columns(4).width = 400
  formseleccio.width = 8000
  formseleccio.Show 1
  If seleccioret = 1 Then
   esculliralgunareferenciaproveidor = atrim(formseleccio.Data1.Recordset!id)
   Set rst = dbtintes.OpenRecordset("SELECT tintesreferencies.*, tipusbidons.* FROM tintesreferencies LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id Where tintesreferencies.id = " + atrim(esculliralgunareferenciaproveidor))
   If Not rst.EOF Then vcapacitatllauna = cadbl(rst!capacitat)
   Set rst = Nothing
  End If
  Unload formseleccio
  

End Function
Function comprovareldecimal(v As String) As String
   If elsimboldecimal = "," Then comprovareldecimal = substituir(v, ".", ",")
   If elsimboldecimal = "." Then comprovareldecimal = substituir(v, ",", ".")
End Function
Function elsimboldecimal() As String
    Dim v As Double
    elsimboldecimal = ","
    If InStr(1, Trim(1 / 2), ".") > 0 Then elsimboldecimal = "."
End Function
Function calcular_preu_kg_amblaformula() As Double
    If datatintesformules.Recordset.EOF Then Exit Function
    calcular_preu_kg_amblaformula = saber_preu_kg_tinta_llauna("", atrim(datatintesformules.Recordset!numformula))
End Function
Sub escullir_proveidorrecuperador(vid As Long)
  Load formseleccio
  formseleccio.sortirs.tag = "filtre"
  Set formseleccio.Data1.Recordset = dbcomandes.OpenRecordset("select id,nomcomercial from recuperadorsdecontenidors order by nomcomercial")
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 2000
  formseleccio.width = 5000
  formseleccio.refrescar
  
  formseleccio.Show 1
  If seleccioret = 1 Then
   vid = formseleccio.DBGrid2.Columns("id")
  End If
fi:
  Unload formseleccio
End Sub
Sub escullir_material_contenidor(vidmaterialcontenidor As Long)
    Dim vtipusmaterial As String
      Load formseleccio
      formseleccio.caption = "Escull l'albarà "
      formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
      formseleccio.Data1.RecordSource = "select codi,descripcio from contenidors_material order by descripcio"
      formseleccio.refrescar
      formseleccio.DBGrid2.Columns(1).width = 5800
      formseleccio.DBGrid2.Columns(0).visible = False
      formseleccio.width = 6800
      'formseleccio.DBGrid2.Columns(0).Width = 1000
      'formseleccio.DBGrid2.Columns(1).Width = 3000
      formseleccio.Command2.tag = "0"
      formseleccio.caption = "Escullir tipus de contenidor"
      formseleccio.Show 1
      If seleccioret = 1 Then
           vidmaterialcontenidor = cadbl(formseleccio.Data1.Recordset!codi)
      End If
      Unload formseleccio
     
End Sub
Sub novallauna(Optional vid_refproveidor As String, Optional kgllauna As String, Optional vnumerodelot As String, Optional vsituacio As String, Optional vquantitatllaunes As Byte, Optional nllauna As String, Optional vidmaterialcontenidor As Long, Optional vidproveidorrecuperador As Long, Optional vmatriculacontenidor As String, Optional valbaraproveidor As String)
   Dim rstll As Recordset
   Dim rstllaunaanterior As Recordset
   Dim vvalors As String
   Dim capacitatllauna As Double
   Dim rstllauna As Recordset
   Dim vformula As String
   Dim vpreuperkilo As Double
   Dim i As Integer
   Dim j As Integer
   Dim vlotdirecta As Boolean
   Dim vresp As String
   'nllauna = InputBox("Escriu el nom de la nova llauna.", "Nova llauna")
   'If Len(atrim(nllauna)) > 10 Then MsgBox "El nom de la llauna no pot tenir mes de 10 digits", vbCritical, "Error": Exit Sub
   If atrim(vnumerodelot) <> "" Then vlotdirecta = True
   Set rstllaunaanterior = dbtintes.OpenRecordset("Select * from dadesllaunesrecargues where numllauna='" + atrim(vnumerodelot) + "'")
   If vid_refproveidor <> "" Then GoTo etcrearllauna
   kgllauna = InputBox("De quants Kilos es aquesta llauna?", "Nova llauna")
   kgllauna = comprovareldecimal(kgllauna)
   If cadbl(kgllauna) = 0 Then GoTo fi
   'vsituacio = UCase(InputBox("Entra la situacio de la llauna", "Nova llauna", "IMP"))
   'If vsituacio = "" Then Exit Sub
   vsituacio = demanar_situacio_llauna
   If vsituacio = "" Then MsgBox "Si no esculls una situació per les llaunes no es crearan", vbCritical, "Error": GoTo fi
   vid_refproveidor = escullir_referenciaproveidor(tintes.Recordset!idtinta, capacitatllauna)
   If cadbl(vid_refproveidor) = 0 Then GoTo fi
   If (capacitatllauna + 4) < cadbl(kgllauna) Then MsgBox atrim(cadbl(kgllauna)) + " Kg són massa kilos per aquesta llauna de " + atrim(capacitatllauna) + "Kg.", vbCritical, "Error": Exit Sub
   If capacitatllauna > 300 Then
      escullir_material_contenidor vidmaterialcontenidor
      escullir_proveidorrecuperador vidproveidorrecuperador
      vmatriculacontenidor = InputBox("Escriu la MATRICULA d'aquest contenidor", "Matricula")
      If Len(vmatriculacontenidor) > 50 Then vmatriculacontenidor = Mid(vmatriculacontenidor, 1, 50)
      If vidmaterialcontenidor = 0 Or vidproveidorrecuperador = 0 Then GoTo fi
   End If
   'vnumerodelot = UCase(atrim(InputBox("Entra el numero de LOT aquesta llauna" + Chr(10) + "També pots possar Nº de llauna." + Chr(10) + "Si es un lot manual escriu * davant del numero", "Nova llauna")))
   'If vnumerodelot = "" Then MsgBox "Necessitem numero de lot per donar l'alta la llauna.", vbCritical, "Error": Exit Sub
   formmeslots.Show 1
   If formmeslots.llistallaunes.ListCount = 0 Then MsgBox "Necessitem numero de lot per donar l'alta la llauna.", vbCritical, "Error": GoTo fi
   vnumerodelot = formmeslots.llistallaunes.List(0)
   If Mid(vnumerodelot + " ", 1, 1) <> "*" Then
        Set rstllaunaanterior = dbtintes.OpenRecordset("Select * from dadesllaunesrecargues_totes where numllauna='" + atrim(vnumerodelot) + "'")
        If saber_preu_kg_tinta_llauna(vnumerodelot, "") = 0 Then
            If rstllaunaanterior.EOF Then MsgBox "Amb aquest lot o llauna no trobo cap referencia de compra." + Chr(10) + "Assegura que l'albarà d'entrada estigui entrat.", vbCritical, "Atenció": Exit Sub
        End If
        Else: vnumerodelot = atrim(Mid(vnumerodelot + " ", 2))
   End If
   vresp = InputBox("Quantes llaunes iguals vols donar d'alta?" + Chr(10) + "Màxim 9 llaunes", "X llaunes", 1)
   If StrPtr(vresp) = 0 Then GoTo fi
   vquantitatllaunes = cadbl(vresp)
   If vquantitatllaunes < 1 Or vquantitatllaunes > 9 Then MsgBox "Massa llaunes màxim 9", vbCritical, "Error": GoTo fi
etcrearllauna:
   If rstllaunaanterior.EOF Then
       vpreuperkilo = calcular_preu_kg_tinta(vnumerodelot)
       If vpreuperkilo = 0 Then vpreuperkilo = calcular_preu_kg_amblaformula
         Else: vpreuperkilo = cadbl(rstllaunaanterior!preuxrkilo)
   End If
   For i = 1 To vquantitatllaunes
      'nova  llauna
        Set rstllauna = dbtintes.OpenRecordset("select numllauna from contadors")
        nllauna = "A" + atrim(rstllauna!numllauna + 1)
        dbtintes.Execute "update contadors set numllauna=[numllauna]+1"
        If atrim(nllauna) = "" Then Exit Sub
      'posso dades ala llauna
        Set rstll = dbtintes.OpenRecordset("select * from llaunes")
        rstll.AddNew
        rstll!numllauna = UCase(treure_apostruf(nllauna))
        rstll!capacitatactual = cadbl(kgllauna)
        rstll!idtinta = tintes.Recordset!idtinta
        rstll!activa = True
        rstll!situacio = vsituacio
        rstll!id_refproveidor = vid_refproveidor
        rstll!preuxrkilo = vpreuperkilo
        rstll!idmaterialcontenidor = cadbl(vidmaterialcontenidor)
        rstll!idproveidorrecuperador = cadbl(vidproveidorrecuperador)
        rstll!vmatriculacontenidor = atrim(vmatriculacontenidor)
        rstll!albaraproveidor = Mid(atrim(valbaraproveidor), 1, 30)
        rstll.Update
      'posso la histori i lots
        datallaunes.Refresh
        datallaunes.Recordset.FindFirst "numllauna='" + atrim(nllauna) + "'"
        vvalors = "(" + atrim(datallaunes.Recordset!id) + "," + atrim(numproximarecarrega(nllauna)) + ",now,'C',0,'" + possarformulapredeterminada(tintes.Recordset!idtinta) + "'," + atrim(passaradecimalpunt(kgllauna)) + ",0)"
        dbtintes.Execute "insert into historiallauna (idnumllauna,numrecarrega,data,tipusmoviment,comanda,formula,kg,idhistoriabarreja) values " + vvalors
        calcularkgdisponiblesllauna nllauna
        datallaunes.Refresh
        datallaunes.Recordset.FindFirst "numllauna='" + atrim(nllauna) + "'"
        If vlotdirecta Then
            vvalors = " (" + atrim(datahistoria.Recordset!id) + ",0,'" + vnumerodelot + "',0,0)"
            dbtintes.Execute "insert into historiallaunalots (idhistoria,idcomponent,numlotbase,tanx100tinta,kgtinta) values " + vvalors
          Else
                j = 0
                While j < formmeslots.llistallaunes.ListCount
                    vnumerodelot = treure_apostruf(formmeslots.llistallaunes.List(j))
                    If Mid(vnumerodelot + " ", 1, 1) = "*" Then vnumerodelot = atrim(Mid(vnumerodelot + " ", 2))
                    vvalors = " (" + atrim(datahistoria.Recordset!id) + ",0,'" + vnumerodelot + "',0,0)"
                    dbtintes.Execute "insert into historiallaunalots (idhistoria,idcomponent,numlotbase,tanx100tinta,kgtinta) values " + vvalors
                    j = j + 1
                Wend
        End If
   Next i
fi:
   Set rstll = Nothing
   Set rstllauna = Nothing
   Set rstllaunaanterior = Nothing
   Unload formmeslots
End Sub

Sub filtrarimportacio()
 Dim vfiltrar As String
 Dim vlike As String
 vlike = IIf(semblants.Value = 1, "*", "")
 vfiltrar = " despan like '*" + fdespan + "*' and descol like '*" + fdescol + "*' and desfam like '" + vlike + fdesfam + vlike + "' and refpan like '*" + frefpan + "*'"
 dataimportacio.RecordSource = "select * from TEMP_TINTES_DISPONIBLES where " + vfiltrar + IIf(inclourevinculatstinta.Value = 0, " and (coditinta='' or coditinta = null)", "") + IIf(inclourevinculatsformules.Value = 0, " and (codiformula='' or codiformula = null)", "")
 dataimportacio.Refresh
 If Not dataimportacio.Recordset.EOF Then
   dataimportacio.Recordset.MoveLast
   dataimportacio.Recordset.MoveFirst
 End If
 reixaimportacio.Refresh
 filtreformuladesc = fdespan
 filtreformulaserie = fdesfam
End Sub

Private Sub Command60_Click()
  
End Sub

Private Sub Command61_Click()
  
End Sub

Private Sub Command62_Click()
Dim vnumc As String
  If reixacomandes.Row > 0 Then vnumc = reixacomandes.TextMatrix(reixacomandes.Row, col)
  vnumc = cadbl(InputBox("Entra la comanda que vols veure.", "Visualitzar el PDF", vnumc))
  If vnumc = 0 Then Exit Sub
  'veurelacomanda cadbl(vnumc) 'llistacomandes.ItemData(llistacomandes.ListIndex)
  veureelpdf cadbl(vnumc)
  'wait 3
  'escriure_ini "Baixes", "imprimircomanda", "0", "comandes.ini"
End Sub

Private Sub Command63_Click()
  Dim i As Integer
  If Command63.BackColor = &HFFFF& Then
    Command63.BackColor = QBColor(13)
    Command63.tag = "tecalloff=true "
     Else
       Command63.BackColor = &HFFFF&
       Command63.tag = ""
  End If
   werescomandes = werescomandes + " "
   filtre(0).SetFocus
   filtre_LostFocus (0)
   'poblar_reixa_comandes
   'carregar_amples_reixa
 
 
End Sub

Private Sub Command64_Click()
'   If crear_temporal_compresactives Then
'     possar_les_compres_actives_al_csv
'     Close #1
'     If existeix("c:\temp\~compresactives.csv") Then
'        obrir_document "c:\temp\~compresactives.csv"
'
'     End If
'   End If
 reixacomandes.visible = False
   poblar_reixa_comandes " order by ordremaquina"
   carregar_amples_reixa
     reixacomandes.visible = True
   carregar_liniadelareixaseleccionada
End Sub
Sub possar_les_compres_actives_al_csv()
   Dim rstcompres As Recordset
   Dim i As Byte
   Dim vlinia As String
   Dim vsql As String
   Set rstcompres = dbcompres.OpenRecordset("SELECT capcalera.numcomanda, capcalera.data, capcalera.dataentrega, capcalera.nomprovcomercial, liniescompra.codimaterial, liniescompra.nommaterial, liniescompra.quantitatkg, liniescompra.kgentregats FROM capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra WHERE (((capcalera.numcomanda)>0) AND ((liniescompra.kgentregats)=0) AND ((liniescompra.totentregat)=False) AND ((liniescompra.tipusmaterialcomprat)='T')) order by dataentrega; ")
   While Not rstcompres.EOF
      vlinia = ""
      For i = 0 To rstcompres.Fields.Count - 1
        If InStr(1, atrim(rstcompres.Fields(i)), "MORCHEM") > 0 Then GoTo proxima
        If InStr(1, atrim(rstcompres.Fields(i)), "HENKEL AG") > 0 Then GoTo proxima
        vlinia = vlinia + substituir(atrim(rstcompres.Fields(i)), ";", ",") + ";"
      Next i
      Print #1, vlinia
proxima:
      rstcompres.MoveNext
   Wend
   Set rstcompres = Nothing
   
End Sub
Function crear_temporal_compresactives() As Boolean
   On Error GoTo errorcrearfitxer
   Open "c:\temp\~compresactives.csv" For Output As #1
   Print #1, "Comanda;Data comanda        ;Data Prevista       ;Nom del proveïdor          ;Codi Tinta;Descripció de la tinta                ;Kg comprats"
   crear_temporal_compresactives = True
   Exit Function
errorcrearfitxer:
   MsgBox "Error al crear el fitxer d'Excel de compres actives, mira que no el tinguis obert i torna-ho a provar", vbCritical, "Error"
End Function

Private Sub Command65_Click()
   If Command65.BackColor = &HC0C0FF Then
      Command65.BackColor = &H8000000F
        Else: Command65.BackColor = &HC0C0FF
   End If
   reixacomandes.SetFocus
   ensenyar_informaciodeltreball
   reixacomandes_SelChange
End Sub
Function maquinaonsimprimiraaquestacomanda(vnumc As Double)
   Dim rst As Recordset
   maquinaonsimprimiraaquestacomanda = "NO ASSIGNADA"
   Set rst = dbbaixes.OpenRecordset("select * from impresores_ordreimpresio where comanda=" + atrim(vnumc))
   If Not rst.EOF Then maquinaonsimprimiraaquestacomanda = atrim(rst!nommaquina)
   Set rst = Nothing
End Function
Sub ensenyar_informaciodeltreball()
   Dim vnumtreball As Double
   Dim vnumordre As Double
   Dim vtmp As String
   Dim rst As Recordset
   Dim rstm As Recordset
   llistatintes.Clear
   llistatintes.tag = ""
   If reixacomandes.Row > 0 Then
     vnumtreball = reixacomandes.TextMatrix(reixacomandes.Row, numcol("NºTreball"))
     vnumordre = reixacomandes.TextMatrix(reixacomandes.Row, numcol("Versió"))
     
   '  vnumtreball = 6047
   '  vnumordre = 2
     If vnumtreball < 0 Then vnumtreball = vnumtreball * -1
     Set rst = dbclixes.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(vnumtreball) + " and ordre=" + atrim(vnumordre) + " order by id", , ReadOnly)
     Set rstm = dbclixes.OpenRecordset("select * from modificacions where id_treball=" + atrim(vnumtreball) + " and ordre=" + atrim(vnumordre) + " order by ordre", , ReadOnly)
     llistatintes.Clear: llistatintes.BackColor = &HFFFF&
     If rstm.EOF Then Exit Sub
     If Not rst.EOF Then llistatintes.AddItem "Observacions Disseny:"
     While Not rst.EOF
         llistatintes.AddItem atrim(rst!observacio)
         rst.MoveNext
     Wend
     Set rst = Nothing
     Set rst = dbbaixes.OpenRecordset("select * from idstreball where id=" + atrim(vnumtreball))
     If Not rst.EOF Then vtmp = atrim(rst!obsidtreball)
     If Not rst.EOF Then If Len(vtmp) > 0 Then llistatintes.BackColor = &HFFFF&: llistatintes.AddItem "Observacions Màquina:": llistatintes.tag = vnumtreball
     While Len(vtmp) > 0
        llistatintes.AddItem Mid(vtmp, 1, 99)
        If Len(vtmp) > 80 Then
             vtmp = Mid(vtmp, 100)
              Else: vtmp = ""
        End If
     Wend
     llistatintes.AddItem Chr(13)
     llistatintes.AddItem atrim(rstm!descripcio) + IIf(atrim(rstm!teXtevalidaciocolors) <> "", "-->VALIDACIÓ COLORS:  " + atrim(rstm!teXtevalidaciocolors), "")
   End If
   Set rst = Nothing
   Set rstm = Nothing
End Sub
Sub assignarllaunaacoditinta()
   Dim vnllauna As String
   Dim rstll As Recordset
   Dim vcoditinta As Double
   Dim vnumc As Double
   Dim vsituacio As String
   Dim vnumtreball As Double
   If llistatintes.ListCount = 0 Or llistatintes.ListIndex = -1 Then MsgBox "Primer escull una tinta", vbCritical, "Error": GoTo fi
   Set rstll = dbclixes.OpenRecordset("select coditinta from tintes where id_tinter=" + atrim(llistatintes.ItemData(llistatintes.ListIndex)))
   If rstll.EOF Then Exit Sub
   vcoditinta = cadbl(rstll!coditinta)
   vnumc = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, col))
   vnumtreball = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("NºTreball")))
   If vnumtreball < 0 Then vnumtreball = vnumtreball * -1
   vnllauna = UCase(InputBox("Entra el Nº de llauna que vols relacionar amb aquesta tinta." + Chr(10) + "ESCRIU [CAP] PER TREURE LA RELACIÓ.", "Relacionar llauna amb comanda"))
   If atrim(vnllauna) = "" Then GoTo fi
   If atrim(vnllauna) = "CAP" Then dbtintes.Execute "delete * from assignaciollaunesacomandes where comanda=" + atrim(vnumc) + " and coditinta=" + atrim(vcoditinta): GoTo cont
   Set rstll = dbtintes.OpenRecordset("select codi from dadesllaunes where numllauna='" + Trim(vnllauna) + "'")
   If Not rstll.EOF Then
       If vcoditinta <> cadbl(rstll!codi) Then MsgBox "No coincideix el codi de tinta amb el de la llauna escullida", vbCritical, "Error": GoTo fi
   End If
   dbtintes.Execute "delete * from assignaciollaunesacomandes where comanda=" + atrim(vnumc) + " and coditinta=" + atrim(vcoditinta)
   dbtintes.Execute "insert into assignaciollaunesacomandes (comanda,coditinta,numllauna) values (" + atrim(vnumc) + "," + atrim(vcoditinta) + ",'" + vnllauna + "')"
cont:
   Set rstll = dbtintes.OpenRecordset("select count(*) as Q from assignaciollaunesacomandes where comanda=" + atrim(vnumc))
   vsituacio = reixacomandes.TextMatrix(reixacomandes.Row, numcol("Gestionat?"))
   If rstll!Q >= llistatintes.ListCount Then
        If vsituacio <> "M" Then canviarestatcomanda_reixacomandes vnumc, vnumtreball, "M"
         Else
           If vsituacio = "M" Then
             MsgBox "La situació anterior de la comanda era de 'M' peró ara no hi ha totes les tintes relacionades, passare la comanda a 'S'.", vbInformation, "Canvi de situació de gestionada"
             canviarestatcomanda_reixacomandes vnumc, vnumtreball, "S"
           End If
   End If
fi:
   Set rstll = Nothing
   reixacomandes_SelChange
End Sub

Sub actualitzar_estocactual()
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from estocsminims")
   While Not rst.EOF
      rst.Edit
      rst!estocactual = calcular_estoc_delatinta(rst)
      rst.Update
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Private Sub Command7_Click()
  Dim rst As Recordset
  Dim vmsg As String
  Dim vnomtinta As String
  Dim vestocnecessari As Double
  Dim vcomandesinplicades As String
  Dim vcoditinta As String
  
  ratoli "espera"
  If MsgBox("Vols actualitzar l'estoc de tintes o no cal?", vbExclamation + vbDefaultButton2 + vbYesNo, "Actualitzar") = vbYes Then actualitzar_estocactual
  'borro els valors de comandesrevisadesatintes que ja estan entregades
  dbtintes.Execute "DELETE comandes.proximaseccio, comandesrevisadesatintes.* FROM comandesrevisadesatintes LEFT JOIN comandes ON comandesrevisadesatintes.comanda = comandes.comanda WHERE (((comandes.proximaseccio)='T' Or (comandes.proximaseccio)='V' Or (comandes.proximaseccio)='P'));"

  dbtintes.Execute "delete * from consultaestocs"
  poblar_reixa_estoc
  carregar_amples_reixa_estoc
  dbtintes.Execute "insert into consultaestocs select * from estocsminims"
  Set rst = dbtintes.OpenRecordset("SELECT * from consultaestocs")
     vcoditinta = InputBox("Escriu el numero de codi de tinta que vols buscar.", "Codi tinta")
  afegir_comandesrevisadesatintes rst, IIf(checknomesfora.Value = 1, True, False), IIf(checkinclou.Value = 1, True, False), vcoditinta
  Set rst = dbtintes.OpenRecordset("SELECT * from consultaestocs")
  While Not rst.EOF
    vnomtinta = crearnomdelatinta(rst)
    If checknomesfora.Value = 1 And rst!codi <> "" Then If Not tintafabricadafora(atrim(rst!codi)) Then rst.Delete: GoTo proxim
    If vnomtinta <> "" Then
        vcomandesinplicades = ""
        rst.Edit
        rst!descripcio = vnomtinta + " "
        'rst!estoccomprat = Redondejar(buscar_comprat(rst), 0)
        
        rst!estocnecessari = Redondejar(rst!estocnecessari, 0)
       ' rst!estocactual = calcular_estoc_delatinta(rst)
        rst!estoctotal = Redondejar(cadbl(rst!estocactual) + cadbl(rst!estoccomprat) - cadbl(rst!estocnecessari), 0)
        
        rst.Update
    End If
proxim:
    rst.MoveNext
  Wend
  dbtintes.Execute "delete * from consultaestocs where descripcio=null or descripcio='' or descripcio=' '"
  poblar_reixa_estoc
  carregar_amples_reixa_estoc
  ratoli "normal"
  Set rst = Nothing
End Sub
Function tintafabricadafora(vcoditinta As String) As Boolean
   Dim rstcolor As Recordset
   Dim rst As Recordset
   Set rst = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(vcoditinta) + "'")
   If rst.EOF Then GoTo fi
   Set rstcolor = dbtintes.OpenRecordset("select * from tintesreferencies where idtinta=" + atrim(rst!idtinta) + " and nomproveidor<>'INPLACSA'")
   If Not rstcolor.EOF Then tintafabricadafora = True
fi:
   Set rstcolor = Nothing
   Set rst = Nothing
End Function
Sub afegir_comandesrevisadesatintes(rstestocs As Recordset, Optional nomesforainplacsa As Boolean, Optional incloutintescomandes As Boolean, Optional vcoditinta As String)
    Dim rst As Recordset
    Dim rstc1 As Recordset
    Dim rsttinta As Recordset
    Dim vcomandesinplicades As String
    Dim calcular_estocnecessari As Double
    Dim vcodisfetsservir As String
    Set rst = dbtintes.OpenRecordset("SELECT comandes.numtreball, comandes.numordremodificacio FROM comandesactives INNER JOIN comandes ON comandesactives.comanda = comandes.comanda WHERE (((comandesactives.gestionat)<>'N'));")
    While Not rst.EOF
       If cadbl(vcoditinta) = 0 Then Set rstc1 = dbclixes.OpenRecordset("select * from tintes where coditinta<>'' and id_treball=" + atrim(rst!numtreball) + " and (ordremodificacio=" + atrim(rst!numordremodificacio) + " or ordremodificacio=-" + atrim(rst!numordremodificacio) + ") order by ordretinter")
       If cadbl(vcoditinta) > 0 Then Set rstc1 = dbclixes.OpenRecordset("select * from tintes where coditinta='" + atrim(cadbl(vcoditinta)) + "' and id_treball=" + atrim(rst!numtreball) + " and (ordremodificacio=" + atrim(rst!numordremodificacio) + " or ordremodificacio=-" + atrim(rst!numordremodificacio) + ") order by ordretinter")
       While Not rstc1.EOF
        Set rsttinta = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(rstc1!coditinta) + "'")
        If rsttinta.EOF Then GoTo proxim
        If InStr(1, vcodisfetsservir, "#" + atrim(rstc1!coditinta) + "#") > 0 Then GoTo proxim
        If Not rstc1.EOF Then
             vcodisfetsservir = vcodisfetsservir + "#" + atrim(rstc1!coditinta) + "#"
             If nomesforainplacsa Then If Not tintafabricadafora(atrim(rstc1!coditinta)) Then GoTo proxim
             If Mid(atrim(rsttinta!referenciacolor) + "  ", 1, 2) = "P-" Then
             rstestocs.FindFirst "codi='" + atrim(rstc1!coditinta) + "'"
             If rstestocs.NoMatch And atrim(rstc1!coditinta) <> "" Then
                With rsttinta
                 rstestocs.FindFirst "(idfamilia=" + atrim(cadbl(!idfamilia)) + " and idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + "and idfamcolor= " + atrim(cadbl(!idfamcolor)) + " and idsubfamcolor=" + atrim(cadbl(!idsubfamcolor)) + ") "
                 If rstestocs.NoMatch And atrim(rstc1!coditinta) <> "" Then
                  If Not incloutintescomandes Then GoTo actualitzar 'salto les que no estan controlades per estoc
                  rstestocs.AddNew: rstestocs!codi = atrim(rstc1!coditinta)
                    Else: If Not rstestocs.NoMatch Then rstestocs.Edit
                 End If
                End With
             End If
               Else
                With rsttinta
                 rstestocs.FindFirst "(idfamilia=" + atrim(cadbl(!idfamilia)) + " and idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + "and idfamcolor= " + atrim(cadbl(!idfamcolor)) + " and idsubfamcolor=" + atrim(cadbl(!idsubfamcolor)) + ") "
                 If rstestocs.NoMatch And atrim(rstc1!coditinta) <> "" Then
                  If Not incloutintescomandes Then GoTo actualitzar 'salto les que no estan controlades per estoc
                  rstestocs.AddNew
                  rstestocs!idfamilia = atrim(!idfamilia)
                  rstestocs!idsubfamilia = atrim(!idsubfamilia)
                  rstestocs!idfamcolor = atrim(!idfamcolor)
                  rstestocs!idsubfamcolor = atrim(!idsubfamcolor)
                    Else: If Not rstestocs.NoMatch Then rstestocs.Edit
                 End If
                End With
                
            End If
       End If
actualitzar:
       If rstestocs.EditMode > 0 Then
            vcomandesinplicades = ""
            calcular_estocnecessari = 0
            calcular_estocnecessari = calcular_kgassignatsaaltrescomandes(, atrim(rstc1!coditinta), vcomandesinplicades)
            If calcular_estocnecessari < 0.5 And calcular_estocnecessari > 0 Then calcular_estocnecessari = 0.5
            
            If atrim(rstestocs!codi) = "" Then
                 rstestocs!estocnecessari = cadbl(rstestocs!estocnecessari) + calcular_estocnecessari
                 If vcomandesinplicades <> "" And atrim(rstestocs!comandesinplicades) <> vcomandesinplicades Then rstestocs!comandesinplicades = atrim(rstestocs!comandesinplicades) + IIf(atrim(rstestocs!comandesinplicades) <> "", ",", "") + vcomandesinplicades
                   Else
                   rstestocs!estocnecessari = calcular_estocnecessari
                   If vcomandesinplicades <> "" Then rstestocs!comandesinplicades = vcomandesinplicades
            End If
            rstestocs.Update
       End If
proxim:
       rstc1.MoveNext
      Wend

       rst.MoveNext
    Wend
    Set rst = Nothing
    Set rstc1 = Nothing
End Sub
Function calcular_estocnecessari(vcoditinta As String, vcomandesinplicades As String) As Double
   If vcoditinta = "" Then calcular_estocnecessari = 0: Exit Function
   calcular_estocnecessari = calcular_kgassignatsaaltrescomandes(, vcoditinta, vcomandesinplicades)
End Function
Function buscar_comprat(rst As Recordset) As Double
   Dim vsql As String
   Dim rstcompres As Recordset
   vsql = "SELECT capcalera.numcomanda, capcalera.data, capcalera.dataentrega, capcalera.nomprovcomercial, liniescompra.codimaterial, liniescompra.nommaterial, liniescompra.quantitatkg, liniescompra.kgentregats, tintes.idfamilia, tintes.idsubfamilia, tintes.idfamcolor, tintes.idsubfamcolor FROM (capcalera RIGHT JOIN liniescompra ON capcalera.id = liniescompra.idcompra) LEFT JOIN tintes ON liniescompra.codimaterial = cdbl(tintes.codi) WHERE (((capcalera.numcomanda)>0) AND ((liniescompra.kgentregats)=0) AND ((liniescompra.totentregat)=False) AND ((liniescompra.tipusmaterialcomprat)='T')) ORDER BY capcalera.dataentrega;"
   Set rstcompres = dbcompres.OpenRecordset(vsql)
   If atrim(rst!codi) <> "" Then
        'sumo per codi de tinta
        While Not rstcompres.EOF
           If atrim(rstcompres!codimaterial) = atrim(rst!codi) Then buscar_comprat = buscar_comprat + cadbl(rstcompres!quantitatkg)
           rstcompres.MoveNext
        Wend
        Else
          'sumo per families de tinta
          While Not rstcompres.EOF
           If cadbl(rstcompres!idfamilia) = cadbl(rst!idfamilia) And cadbl(rstcompres!idsubfamilia) = cadbl(rst!idsubfamilia) And cadbl(rstcompres!idfamcolor) = cadbl(rst!idfamcolor) And cadbl(rstcompres!idsubfamcolor) = cadbl(rst!idsubfamcolor) Then
              buscar_comprat = buscar_comprat + cadbl(rstcompres!quantitatkg)
           End If
           rstcompres.MoveNext
          Wend
           
   End If
End Function
Sub netejar_reixa_estoc()
   Dim rst As Recordset
   Dim i As Byte
   Dim col As Byte
   Set rst = dbtintes.OpenRecordset("select * from consultaestocs")
   reixaestocs.Rows = 1
   reixaestocs.Cols = 1
   col = 0
   For i = 0 To rst.Fields.Count - 1
     If valorpropietat(rst.Fields(i), "Caption") <> "" Then
      reixaestocs.Cols = col + 1
      reixaestocs.col = col
      reixaestocs.Text = valorpropietat(rst.Fields(i), "Caption")
      campsestoc(col) = rst.Fields(i).Name
      'If filtre.Count <= col Then Load filtre(col)
      'If Screen.ActiveControl.Name <> "filtre" Then
      ' filtre(col).DataField = rst.Fields(i).Name
      ' filtre(col).Text = valorpropietat(rst.Fields(i), "Caption")
      'End If
      col = col + 1
     End If
   Next i
   
End Sub

Sub poblar_reixa_estoc()
   Dim rst As Recordset
   Dim fila As Integer
   Dim i As Byte
   Dim col As Integer
   reixaestocs.Redraw = False
   netejar_reixa_estoc
   ettotalestocs.caption = ""
   Set rst = dbtintes.OpenRecordset("select * from consultaestocs " + IIf(buscador_estoc(0).tag = "", "where ", buscador_estoc(0).tag + " and ") + " comandesinplicades<>'' " + ordreestoc)
   If rst.EOF Then GoTo fi
   rst.MoveLast
   rst.MoveFirst
   If Not rst.EOF Then ettotalestocs.caption = "Registres: " + atrim(rst.RecordCount)
   fila = 1
   While Not rst.EOF
      col = 0
      reixaestocs.Rows = fila + 1
      For i = 0 To rst.Fields.Count - 1
        If valorpropietat(rst.Fields(i), "Caption") <> "" Then
            possar_el_valor_alareixaestoc fila, col, rst.Fields(i)
            col = col + 1
        End If
      Next i
      fila = fila + 1
      rst.MoveNext
   Wend
fi:
   reixaestocs.Redraw = True
   Set rst = Nothing
End Sub
Sub possar_el_valor_alareixaestoc(fila As Integer, col As Integer, vcamp As Field)
  Dim v As String
  Dim vcolor As Double
  If vcamp.Type = 1 Then v = IIf(vcamp.Value, "S", "N")
  If vcamp.Type = 4 Or vcamp.Type = 7 Then v = cadbl(vcamp.Value)
  If vcamp.Type = 10 Then v = atrim(vcamp.Value)
  If v = "" Then v = atrim(vcamp.Value)
  reixaestocs.TextMatrix(fila, col) = v
  If vcamp.Name = "numtreball" Then If cadbl(vcamp.Value) < 0 Then vcolor = QBColor(12)
  If vcamp.Name = "metres" Then If cadbl(vcamp.Value) = 0 Then vcolor = QBColor(12)
  If vcamp.Name = "gestionat" Then
      If vcamp.Value = "S" Then vcolor = &HC0FFC0
      If vcamp.Value = "C" Then vcolor = QBColor(12)
      If vcamp.Value = "M" Then vcolor = QBColor(14)
      If vcamp.Value = "N" Then vcolor = 0
  End If
  If vcolor > 0 Then
     reixaestocs.col = col
     reixaestocs.Row = fila
     reixaestocs.CellBackColor = vcolor
  End If
End Sub
Sub guardar_amples_reixa_estoc()
Dim j As Integer
If iniconfigreixa <> "" Then
  For j = 0 To reixaestocs.Cols - 1
   escriure_ini "AmplesReixaestocs", UCase(reixaestocs.TextMatrix(0, j)), atrim(reixaestocs.ColWidth(j)), iniconfigreixa
 Next j
End If
End Sub
Sub carregar_amples_reixa_estoc()
 Dim ample As String
 Dim X As Long
 Dim j As Integer
 If iniconfigreixa <> "" Then ' existeix("c:\windows\" + iniconfigreixa) Then
 If reixaestocs.Cols < 3 Then Exit Sub
  X = reixaestocs.Left + 35
  For j = 0 To reixaestocs.Cols - 1
   ample = llegir_ini("AmplesReixaestocs", UCase(reixaestocs.TextMatrix(0, j)), iniconfigreixa)
   If ample = "{[}]" Then ample = 1000
   reixaestocs.ColWidth(j) = cadbl(ample)
    If X < reixaestocs.width Then
    ' filtre(j).Left = x
    ' filtre(j).width = cadbl(ample)
    ' filtre(j).visible = True
    ' filtre(j).ForeColor = &H808080
    '  Else: If filtre.Count < j - 1 Then filtre(j).visible = False
    End If
    X = X + cadbl(ample)
 Next j
End If
'filtre(0).width = filtre(0).width - 50
'filtre(0).Left = filtre(0).Left + 50
End Sub























Function capdetriatc() As Boolean
  datacolortreball.Refresh
  capdetriatc = True
  While Not datacolortreball.Recordset.EOF
      If datacolortreball.Recordset!seleccionar Then capdetriatc = False
      datacolortreball.Recordset.MoveNext
  Wend
  If capdetriatc Then MsgBox "No hi ha cap registre de colors seleccionat."
End Function
Function capdetriat() As Boolean
  dataimportacio.Refresh
  capdetriat = True
  While Not dataimportacio.Recordset.EOF
      If dataimportacio.Recordset!seleccionar Then capdetriat = False
      dataimportacio.Recordset.MoveNext
  Wend
  If capdetriat Then MsgBox "No hi ha cap registre de llauna seleccionat."
End Function
Private Sub Command8_Click()
     werescomandes = " reprint=true"
     poblar_reixa_comandes
'    Dim vinici As Date
'    Dim vfi As Date
'    Dim valor As String
'    vdiaanterior = Format(DateAdd("d", -1, Now), "dd/mm/yy")
'    While WeekDay(vdiaanterior, vbMonday) = 7 Or WeekDay(vdiaanterior, vbMonday) = 6
'       vdiaanterior = Format(DateAdd("d", -1, vdiaanterior), "dd/mm/yy")
'    Wend
'    valor = InputBox("Entra la data d'inici de la consulta.", "Data inici", vdiaanterior)
'    If Not IsDate(valor) Then MsgBox "Aquesta data no es vàlida", vbCritical, "Error": Exit Sub
'    vinici = valor
'    valor = InputBox("Entra la data de fi de la consulta.", "Data fi", Format(Now, "dd/mm/yy"))
'    If Not IsDate(valor) Then MsgBox "Aquesta data no es vàlida", vbCritical, "Error": Exit Sub
'    vfi = valor
'    llistademodificacions vinici, vfi
End Sub
Sub llistademodificacions(vinici As Date, vfi As Date)
  Dim ventredates As String
  Dim vnomesrevisades As String
  vnomesrevisades = " Mid([numtreball],1,InStr(1,[numtreball],'/')-1)  in (select numtreball from comandesrevisadesatintes where estatgestio='S')"
  ventredates = "data>=#" + Format(vinici, "mm/dd/yy") + "# and data<=#" + Format(vfi, "mm/dd/yy") + "#"
  Unload formseleccio
  Load formseleccio
  'formseleccio.Command3.tag = "filtre"
  formseleccio.sortirs.tag = "filtre"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "tintes.mdb"
  formseleccio.Data1.RecordSource = "select tintes_controlcanvis.Id, tintes_controlcanvis.numtreball, tintes_controlcanvis.data, tintes_controlcanvis.usuari, tintes_controlcanvis.campafectat, tintes_controlcanvis.valoractual AS VALOR_ANTERIOR, tintes_controlcanvis.valoranterior AS VALOR_ACTUAL from tintes_controlcanvis where campafectat like '*: color*' and " + ventredates + " and " + vnomesrevisades
  formseleccio.refrescar
  'If formseleccio.Data1.Recordset.RecordCount = 0 Then MsgBox "No hi ha registres", vbInformation, "Atenció": Exit Sub
  formseleccio.DBGrid2.Columns(1).width = 1000
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(2).NumberFormat = "dd/mm/yy"
  formseleccio.DBGrid2.Columns(3).width = 2000
  formseleccio.DBGrid2.Columns(4).width = 2000
  formseleccio.DBGrid2.Columns(5).width = 4000
  formseleccio.DBGrid2.Columns(6).width = 4000
  formseleccio.DBGrid2.col = 4
  formseleccio.width = formtintes.width
  formseleccio.colocar_botofiltre 4
  DoEvents
  If formseleccio.Data1.Recordset.EOF Then Exit Sub
  formseleccio.Show 1
  Unload formseleccio
End Sub
Function triar_coditinta() As String
  Dim des As Double
  Dim sql As String
  Dim rst As Recordset
  Dim were As String
  Dim nummaq As Byte
  Dim caigudes As Double
  
  sql = "SELECT  idtinta,codi,descripcio,referenciacolor from tintes_tot "
  were = " order by descripcio"
  Load formseleccio
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = sql + were
  formseleccio.width = 13000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.caption = "Escull la tinta que vols buscar"
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.Show 1
  If seleccioret = 1 Then
    triar_coditinta = atrim(formseleccio.Data1.Recordset!codi)
  End If
  If seleccioret = 9 Then
    triar_coditinta = 0
  End If
  Unload formseleccio
End Function
Sub borrar_taula(db As Database, vnomtaula As String)
  On Error GoTo fi
  db.Execute "drop table " + vnomtaula
  Exit Sub
fi:
  If InStr(1, err.Description, "no existe") = 0 Then
     MsgBox "Error borrant la taula " + vnomtaula
  End If
End Sub
Function crear_tmp_filtretintes(vcoditinta As String) As Boolean
  Dim rst As Recordset
  Dim rst2 As Recordset
  borrar_taula dbclixes, "tmp_consultatintesclixes"
  dbclixes.Execute "SELECT Clixes.id_treball AS Treball,modificacions.ordre, Clixes.arxiu AS Arxiu_, ([marca]+' - '+[Linia]) AS Marca_i_Linia, Clixes.estatclixe AS Estat_clixé into tmp_consultatintesclixes FROM Clixes INNER JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball  where clixes.id_treball in (SELECT Tintes.id_treball FROM Tintes LEFT JOIN Tintes AS Tintes_1 ON Tintes.tinterlinkambid_treball = Tintes_1.id_tinter where IIf([tintes].[tinterlinkambid_treball]>0,[tintes_1].[coditinta],[tintes].[coditinta])='" + vcoditinta + "')"
  Set rst = dbclixes.OpenRecordset("select treball,ordre as mordre from tmp_consultatintesclixes ")
  If rst.EOF Then crear_tmp_filtretintes = False: GoTo fi
  crear_tmp_filtretintes = True
  Set rst2 = dbclixes.OpenRecordset("select id_treball,max(ordre)as mordre from modificacions where id_treball in(select treball from tmp_consultatintesclixes) group by id_treball")
  While Not rst.EOF
    rst2.FindFirst "id_treball=" + atrim(rst!treball)
    If Not rst2.NoMatch Then
        If rst2!mordre <> rst!mordre Then rst.Edit: rst!treball = 0: rst.Update
    End If
    rst.MoveNext
  Wend
  dbclixes.Execute "delete * from tmp_consultatintesclixes where treball=0"
fi:
  Set rst = Nothing
  Set rst2 = Nothing
End Function

Private Sub Command9_Click()
  Dim vsubselect As String
  Dim vcoditinta As String
  Unload formseleccio
  vcoditinta = atrim(triar_coditinta)
  If Not crear_tmp_filtretintes(vcoditinta) Then Exit Sub
  
  Unload formseleccio
  Load formseleccio
  formseleccio.bimprimir.visible = True
  'formseleccio.Command3.tag = "filtre"
  formseleccio.sortirs.tag = ""
  'vsubselect = "SELECT Clixes.id_treball, Clixes.arxiu, [marca] & ' - ' & [linia] AS Marca_Linia From Clixes WHERE (((Clixes.databaixaclixe) Is Null)"
  vsubselect = "SELECT * from tmp_consultatintesclixes"
  
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "clixesnous.mdb"
  formseleccio.Data1.RecordSource = vsubselect
'  formseleccio.Data1.RecordSource = "select id_treball,arxiu,Marca_i_Linia from consulta_marcailinia " + IIf(vcoditinta <> "", " where id_treball in (select id_treball from tintes where coditinta='" + vcoditinta + "')", "")
 ' formseleccio.Data1.RecordSource = vsubselect + IIf(vcoditinta <> "", " where clixes.id_treball in (select id_treball from tintes where coditinta='" + vcoditinta + "')", "") + " GROUP BY Clixes.id_treball order by 3"
  'Clipboard.Clear
  'Clipboard.SetText formseleccio.Data1.RecordSource
  'Clipboard.SetText vsubselect + IIf(vcoditinta <> "", " where Treball in (select id_treball from tintes where coditinta='" + vcoditinta + "')", "") + " GROUP BY Clixes.id_treball;"
  
  formseleccio.refrescar
  'If formseleccio.Data1.Recordset.RecordCount = 0 Then MsgBox "No hi ha registres", vbInformation, "Atenció": Exit Sub
  formseleccio.DBGrid2.Columns(0).width = 900
  formseleccio.DBGrid2.Columns(1).width = 900
  formseleccio.DBGrid2.Columns(2).width = 900
  formseleccio.DBGrid2.Columns(3).width = 7000
  formseleccio.DBGrid2.Columns(4).width = 2000
  formseleccio.DBGrid2.col = 2
  formseleccio.cmissatge.tag = "2"
  formseleccio.width = 15000
  formseleccio.colocar_botofiltre 3
  DoEvents
  If formseleccio.Data1.Recordset.EOF Then Exit Sub
  seleccioret = 1
  While seleccioret = 1
  
   formseleccio.Show 1
   If seleccioret = 9 Then imprimir_resultat_marques: GoTo fi
   If seleccioret <> 1 Then GoTo fi
   Set rst = dbclixes.OpenRecordset("select max(ordre) as maxordre from modificacions where id_treball=" + atrim(cadbl(formseleccio.Data1.Recordset!treball)))
'   ColocarEnTop formseleccio, False
   ratoli "espera"
   If Not rst.EOF Then obrir_pdf_treball cadbl(formseleccio.Data1.Recordset!treball), cadbl(rst!maxordre)
   
   wait 2
   ratoli "normal"
   If MsgBox("Vols buscar mes treballs?", vbInformation + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then GoTo fi
  Wend
fi:
  Unload formseleccio
  Set rst = Nothing
  borrar_taula dbclixes, "tmp_consultatintesclixes"
  'Unload formseleccio
End Sub
Sub imprimir_resultat_marques()
   If crear_temporal_resultatmarques Then
     possar_les_marques_al_csv
     Close #1
     If existeix("c:\temp\~marquestemp.csv") Then
        obrir_document "c:\temp\~marquestemp.csv"
     End If
   End If
End Sub
Sub possar_les_marques_al_csv()
   Dim rstcompres As Recordset
   Dim i As Byte
   Dim vlinia As String
   Dim vsql As String
   Set rstcompres = formseleccio.Data1.Recordset
   While Not rstcompres.EOF
      vlinia = ""
      For i = 0 To rstcompres.Fields.Count - 1
        vlinia = vlinia + substituir(atrim(rstcompres.Fields(i)), ";", ",") + ";"
      Next i
      Print #1, vlinia
proxima:
      rstcompres.MoveNext
   Wend
   Set rstcompres = Nothing
   
End Sub
Function crear_temporal_resultatmarques() As Boolean
   On Error GoTo errorcrearfitxer
   Open "c:\temp\~marquestemp.csv" For Output As #1
   Print #1, "Treball; Arxiu   ; Marca i linia                "
   crear_temporal_resultatmarques = True
   Exit Function
errorcrearfitxer:
   MsgBox "Error al crear el fitxer d'Excel de marques, mira que no el tinguis obert i torna-ho a provar", vbCritical, "Error"
End Function

Public Function ColocarEnTop(vform As Form, ByVal fColocarEnTop As Boolean) As Boolean
On Error Resume Next
Dim f As Boolean
'Si la función falla devuelve \"False\"
f = (SetWindowPos(vform.hwnd, IIf(fColocarEnTop = True, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, flags) <> 0)
fEstaEnTop = (fColocarEnTop And (f = True))
ColocarEnTop = f
End Function
Private Sub crefcolor_LostFocus()
    crefcolor = UCase(crefcolor)
End Sub

Private Sub csubfamilia_DropDown()
   escullir_subfamiliatinta
   If cfamiliacolor = "" Then cfamiliacolor_DropDown
End Sub
Sub escullir_subfamiliatinta()
  Static ultimcodi As String
  If cadbl(tintes.Recordset!idfamilia) = 0 Then Exit Sub
  Load formseleccio
  formseleccio.caption = "Selecciona SubFamilia Tinta"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "select codi,descripcio,color from subfamiliestintes where codifam=" + atrim(cadbl(tintes.Recordset!idfamilia)) + " order by descripcio"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 4500
  formseleccio.DBGrid2.Columns(2).width = 800
  If cadbl(ultimcodi) > 0 Then formseleccio.Data1.Recordset.FindFirst "codi=" + atrim(ultimcodi)
  formseleccio.Show 1
  If seleccioret = 1 Then
   csubfamilia = atrim(formseleccio.Data1.Recordset!descripcio)
   tintes.Recordset!idsubfamilia = formseleccio.Data1.Recordset!codi
   ultimcodi = atrim(tintes.Recordset!idsubfamilia)
  End If
  Unload formseleccio
  
End Sub
Private Sub csubfamilia_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub csubfamiliacolor_DropDown()
    escullir_subfamiliacolor
End Sub
Sub escullir_subfamiliacolor()
  Static ultimcodi As String
  Load formseleccio
  formseleccio.caption = "Selecciona SubFamilia Color"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "select codi,descripcio from subfamiliescolors order by descripcio"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 4500
  If cadbl(ultimcodi) > 0 Then formseleccio.Data1.Recordset.FindFirst "codi=" + atrim(ultimcodi)
  
  formseleccio.Show 1
  If seleccioret = 1 Then
   csubfamiliacolor = atrim(formseleccio.Data1.Recordset!descripcio)
   tintes.Recordset!idsubfamcolor = formseleccio.Data1.Recordset!codi
   ultimcodi = atrim(tintes.Recordset!idsubfamcolor)
  End If
  Unload formseleccio
  
End Sub
Private Sub csubfamiliacolor_KeyPress(KeyAscii As Integer)
  KeyPress = 0
End Sub

Private Sub datacolortreball_Reposition()
  If Not datacolortreball.Recordset.EOF Then
    etquancolor = atrim(datacolortreball.Recordset.RecordCount) + " Registres a la reixa."
  End If
End Sub

Private Sub datacomponents_Reposition()
   possar_lotsbase
End Sub
Sub possar_lotsbase()
 If datacomponents.Recordset.EOF Then
   datalotsbase.RecordSource = "SELECT* FROM detallnumeroslotsbase where idcomponent=0;"
   datalotsbase.Refresh
   Exit Sub
 End If
 datalotsbase.RecordSource = "SELECT* FROM detallnumeroslotsbase where idcomponent=" + atrim(datacomponents.Recordset!idcomponent) + " order by data desc;"
 datalotsbase.Refresh
   
End Sub

Private Sub datadellaunes_Reposition()
  If Not datadellaunes.Recordset.EOF Then
     'ettotalllaunes.caption = "Total llaunes: " + atrim(datadellaunes.Recordset.RecordCount)
  '   ettotalllaunes.caption = comptar_llaunes '"Total llaunes: " + atrim(datadellaunes.Recordset.RecordCount)
      Else: ettotalllaunes.caption = ""
  End If
  possar_color_llauna
End Sub
Function comptar_llaunes() As String
  Dim rst As Recordset
  Set rst = datadellaunes.Database.OpenRecordset("select count(numllauna) as tllaunes from dadesllaunes where nombido like'LATA*'" + IIf(datadellaunes.tag <> "", " and " + datadellaunes.tag, ""))
  comptar_llaunes = atrim(cadbl(rst!Tllaunes)) + " Llaunes"
  Set rst = datadellaunes.Database.OpenRecordset("select count(numllauna) as tllaunes from dadesllaunes where nombido like'BIDON*'" + IIf(datadellaunes.tag <> "", " and " + datadellaunes.tag, ""))
  comptar_llaunes = comptar_llaunes + "   " + atrim(cadbl(rst!Tllaunes)) + " Bidons"
  Set rst = datadellaunes.Database.OpenRecordset("select count(numllauna) as tllaunes from dadesllaunes where nombido like'CONTEN*'" + IIf(datadellaunes.tag <> "", " and " + datadellaunes.tag, ""))
  comptar_llaunes = comptar_llaunes + "    " + atrim(cadbl(rst!Tllaunes)) + " Contenidors"
  comptar_llaunes = comptar_llaunes + "    (Actives)"
End Function

Private Sub dataformules_Reposition()
   kgxrecuperar(2) = cadbl(dataformules.Recordset!aniloxformulada)
   posar_detallformula
End Sub
Sub posar_detallformula()
 Dim sumar As Double
 If dataformules.Recordset.EOF Then
   datadetallformules.RecordSource = "SELECT   Componentsbase.codicomponent, nomcomponent, [%decomponent] FROM DetallFormules LEFT JOIN Componentsbase ON DetallFormules.IdComponente = Componentsbase.idcomponent where idformula=0;"
   datadetallformules.Refresh
   Exit Sub
 End If
 datadetallformules.RecordSource = "SELECT   Componentsbase.codicomponent, nomcomponent, [%decomponent] FROM DetallFormules LEFT JOIN Componentsbase ON DetallFormules.IdComponente = Componentsbase.idcomponent where idformula=" + atrim(dataformules.Recordset!idformula) + ";"
 datadetallformules.Refresh
 While Not datadetallformules.Recordset.EOF
   sumar = sumar + cadbl(datadetallformules.Recordset![%decomponent])
   datadetallformules.Recordset.MoveNext
 Wend
 ltotalpercent = "Total: " + atrim(Redondejar(sumar, 0)) + "%"
 
End Sub

Private Sub dataimportacio_Reposition()
  If Not dataimportacio.Recordset.EOF Then
    etquant = atrim(dataimportacio.Recordset.RecordCount) + " Registres a la reixa."
  End If
End Sub

Private Sub datallaunes_Reposition()
   If Not datallaunes.Recordset.EOF Then
    datarecarregues.RecordSource = "select distinct numrecarrega from historiallauna where idnumllauna=" + atrim(datallaunes.Recordset!id) + " order by numrecarrega Desc"
    datarecarregues.Refresh
    If Not datarecarregues.Recordset.EOF Then
        datahistoria.RecordSource = "select * from historiallauna where idnumllauna=" + atrim(datallaunes.Recordset!id) + " order by data Desc"
        datahistoria.Refresh
          Else
            datahistoria.RecordSource = "select * from historiallauna where idnumllauna=-999 order by data Desc"
            datahistoria.Refresh
    End If
      Else
            datahistoria.RecordSource = "select * from historiallauna where idnumllauna=-999 order by data Desc"
            datahistoria.Refresh
            datarecarregues.RecordSource = "select distinct numrecarrega from historiallauna where idnumllauna=-999 order by numrecarrega desc"
            datarecarregues.Refresh
   End If
   If Not tintes.Recordset.EOF Then possarkgtotals cadbl(tintes.Recordset!idtinta)
End Sub
Sub possarkgtotals(vidtinta As Double)
   Dim rst As Recordset
   Dim rst2 As Recordset
   etkgtotals = ""
   Set rst = dbtintes.OpenRecordset("SELECT Llaunes.idtinta, Sum(Llaunes.capacitatactual) AS totalkg From Llaunes Where llaunes.capacitatactual>0 and Llaunes.activa = True and Llaunes.idtinta=" + atrim(vidtinta) + " GROUP BY Llaunes.idtinta;")
   Set rst2 = dbtintes.OpenRecordset("SELECT count(Llaunes.idtinta) as Tllaunes from Llaunes Where llaunes.capacitatactual>0.9 and Llaunes.activa = True and Llaunes.idtinta=" + atrim(vidtinta) + " GROUP BY Llaunes.idtinta;")
   If Not rst.EOF Then etkgtotals = atrim(rst2!Tllaunes) + " Llaunes" + Chr(10) + atrim(Redondejar(rst!totalkg, 1)) + " Kg"
   Set rst2 = Nothing
End Sub
Sub ensenyarkgtotalsxrbido(vidtinta As Double)
   Dim rst As Recordset
   Dim vmsg As String
   Set rst = dbtintes.OpenRecordset("SELECT Count(*) AS Tllaunes, Sum(Llaunes.capacitatactual) AS SumaDecapacitatactual, tipusbidons.capacitat FROM Llaunes LEFT JOIN (tipusbidons RIGHT JOIN tintesreferencies ON tipusbidons.id = tintesreferencies.id_bido) ON Llaunes.id_refproveidor = tintesreferencies.id  Where (((Llaunes.capacitatactual) > 0.9) And ((Llaunes.activa) = True)) and Llaunes.idtinta=" + atrim(vidtinta) + " GROUP BY  tipusbidons.capacitat;")

   'Clipboard.SetText "SELECT Count(*) AS Tllaunes, Sum(Llaunes.capacitatactual) AS SumaDecapacitatactual, tipusbidons.capacitat FROM  Llaunes LEFT JOIN (tipusbidons RIGHT JOIN tintesreferencies ON tipusbidons.id = tintesreferencies.id_bido) ON Llaunes.id_refproveidor = tintesreferencies.idtinta Where (((Llaunes.capacitatactual) > 0.9) And ((Llaunes.activa) = True)) and Llaunes.idtinta=" + atrim(vidtinta) + " GROUP BY  tipusbidons.capacitat;"
   'a = InputBox(a, a, "SELECT Count(*) AS Tllaunes, Sum(Llaunes.capacitatactual) AS SumaDecapacitatactual, tipusbidons.capacitat FROM  Llaunes LEFT JOIN (tipusbidons RIGHT JOIN tintesreferencies ON tipusbidons.id = tintesreferencies.id_bido) ON Llaunes.id_refproveidor = tintesreferencies.idtinta Where (((Llaunes.capacitatactual) > 0.9) And ((Llaunes.activa) = True)) and Llaunes.idtinta=" + atrim(vidtinta) + " GROUP BY  tipusbidons.capacitat;")
   While Not rst.EOF
     vmsg = vmsg + atrim("Bidó de " + atrim(rst!capacitat) + ":         " + atrim(rst!Tllaunes) + " Llaunes   " + atrim(Redondejar(rst!SumaDecapacitatactual, 1)) + " Kg" + Chr(10))
     rst.MoveNext
   Wend
   Set rst = Nothing
   MsgBox vmsg
End Sub

Sub possar_color_llauna()
   Dim rst As Recordset
   Dim vcodipantone As String
   Dim vsql As String
   fmostracolor.BackColor = &H8000000F
   If Not datadellaunes.Recordset.EOF Then
      Set rst = datadellaunes.Database.OpenRecordset("select referenciacolor from tintes where codi='" + atrim(datadellaunes.Recordset!codi) + "'")
      If rst.EOF Then GoTo fi
      If Len(rst!referenciacolor) < 4 Then GoTo fi
      vcodipantone = atrim(cadbl(Mid(rst!referenciacolor, 3)))
      fmostracolor.BackColor = buscar_hex_delpantone(vcodipantone)
   End If
fi:
   Set rst = Nothing
End Sub
Function buscar_hex_delpantone(vcodipantone As String) As Variant
   Dim vsql As String
   Dim rst As Recordset
   vsql = "SELECT Pantones.PANTONENUM, Pantones.Hex From Pantones WHERE Pantones.PANTONENUM='" + treure_apostruf(vcodipantone) + "'"
   Set rst = datallaunes.Database.OpenRecordset(vsql)
   If Not rst.EOF Then buscar_hex_delpantone = rst!Hex
   Set rst = Nothing
   If buscar_hex_delpantone = 0 Then buscar_hex_delpantone = &H8000000F
End Function

Private Sub dllistadecomponents_LostFocus()
  If Screen.ActiveControl.Name <> "Command43" And Screen.ActiveControl.Name <> "checknomesseleccionats" Then amagarllistacomponents
End Sub

Private Sub eliminar_Click()
   If hihareferenciesdaquestatinta(atrim(tintes.Recordset!codi)) Then
        MsgBox "Aquesta tinta ja té alguna Formula, Referència o Llauna associada.", vbCritical, "Atenció": Exit Sub
   End If
   If MsgBox("Segur que vols borrar aquesta tinta?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then Exit Sub
   If datallaunes.Recordset.RecordCount = 0 Then
        dbtintes.Execute "delete * from tintes where idtinta=" + atrim(tintes.Recordset!idtinta)
        tintes.Recordset.Delete
        tintes.Refresh
          Else: MsgBox "No es pot borrar aquesta tinta perquè te llaunes associades.", vbCritical, "Atenció"
   End If
End Sub
Function hihareferenciesdaquestatinta(coditinta As String) As Boolean
   Dim rst As Recordset
   Dim idtinta As Long
   hihareferenciesdaquestatinta = False
   Set rst = dbtintes.OpenRecordset("select * from tintes where codi='" + atrim(coditinta) + "'")
   If rst.EOF Then Exit Function
   idtinta = rst!idtinta
   Set rst = dbtintes.OpenRecordset("select * from tintesformules where idtinta=" + atrim(idtinta))
   If Not rst.EOF Then hihareferenciesdaquestatinta = True: Exit Function
   Set rst = dbtintes.OpenRecordset("select * from tintesreferencies where idtinta=" + atrim(idtinta))
   If Not rst.EOF Then hihareferenciesdaquestatinta = True: Exit Function
   Set rst = dbtintes.OpenRecordset("select * from llaunes where idtinta=" + atrim(idtinta))
   If Not rst.EOF Then hihareferenciesdaquestatinta = True: Exit Function
   Set rst = Nothing
End Function

Private Sub etkgtotals_Click()
  If tintes.Recordset.EOF Then Exit Sub
  ensenyarkgtotalsxrbido tintes.Recordset!idtinta
End Sub

Private Sub fcoditintallauna_Change()
filtrarllaunes
End Sub

Private Sub fdescol_Change()
filtrarimportacio
End Sub

Private Sub fdesctintallauna_Change()
   filtrarllaunes
End Sub

Private Sub fdesfam_Change()
filtrarimportacio
End Sub

Private Sub fdespan_Change()
  filtrarimportacio
End Sub

Private Sub fformulallauna_Change()
filtrarllaunes
End Sub

Private Sub filtre_GotFocus(Index As Integer)
    bxrcontrolagafafocus Index
End Sub
Sub bxrcontrolagafafocus(i As Integer)
  Dim cntrl As Control
  Set cntrl = Screen.ActiveControl
  If cntrl.Text <> "" Then
     If cntrl.Text = reixacomandes.TextMatrix(0, i) Then cntrl.Text = ""
     cntrl.ForeColor = QBColor(0)
   Else:
       cntrl.Text = reixacomandes.TextMatrix(0, i)
       cntrl.ForeColor = &H808080
  End If
End Sub

Private Sub filtre_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then KeyCode = 0: filtre(IIf(Index = 0, 1, 0)).SetFocus
End Sub

Private Sub filtre_LostFocus(Index As Integer)
  ultimwerescomandes = werescomandes
  werescomandes = crearfiltre
  If filtre(Index).Text = "" Then
    filtre(Index).Text = reixacomandes.TextMatrix(0, Index)
    filtre(Index).ForeColor = &H808080
  End If
  If ultimwerescomandes <> werescomandes Then
   reixacomandes.visible = False
   poblar_reixa_comandes
   carregar_amples_reixa
     reixacomandes.visible = True
   carregar_liniadelareixaseleccionada
  End If
End Sub
Function crearfiltre() As String
  Dim i As Integer
  Dim were As String
  Dim w As String
  For i = 0 To filtre.Count - 1
    If filtre(i).Text <> reixacomandes.TextMatrix(0, i) Then
      If reixacomandes.TextMatrix(0, i) = "Comanda" Then filtre(i).Text = convertircampcomandaallista(filtre(i))
      w = crearwere(i)
      If were = "" Then
         were = w
        Else: If w <> "" Then were = were + " and " + w
      End If
    End If
  Next i
  crearfiltre = were
End Function
Function convertircampcomandaallista(ByVal v As String) As String
   Dim i As Integer
   i = 1
   While i < Len(v) + 1
     If IsNumeric(Mid(v, i, 1)) Or Mid(v, i, 1) = "." Then
        vnum = vnum + Mid(v, i, 1)
          Else
buscar:
            convertircampcomandaallista = convertircampcomandaallista + IIf(convertircampcomandaallista <> "", ",", "") + vnum
cont:
            vnum = ""
     End If
     i = i + 1
     If i = Len(v) + 1 Then
        GoTo buscar
     End If
   Wend
End Function
Function crearwere(i As Integer) As String
   Dim w As String
   Dim j As Integer
   Dim rst As Recordset
   Dim vcamp As String
   If filtre(i) = "" Then Exit Function
   Set rst = dbtintes.OpenRecordset("select * from comandesactives")
   vcamp = filtre(i).DataField
   If InStr(1, filtre(i), "TreballsBlaus") > 0 Then
       crearwere = "estatclixe='POLIMERS O CLIXES' "
       GoTo fi
   End If
   If InStr(1, filtre(i), "ComandaFucsia") > 0 Then
       crearwere = "mid(tipusimpresio,1,1) = '@'"
       GoTo fi
   End If
   If vcamp = "CodiLinia" Then
       w = filtre(i)
       If InStr(1, w, "#") = 0 And w <> "-" Then w = Format(w, "000") + "#"
       w = substituir(w, "#", "\-")
       crearwere = "codilinia like '*" + substituir(w, "\-", "[#]") + "*' "
       GoTo fi
   End If
  
   If rst.Fields(vcamp).Type = 8 Then
      If IsDate(filtre(i)) Then
         crearwere = vcamp + "=#" + Format(filtre(i), "mm/dd/yy") + "# "
      End If
      GoTo fi
   End If
   If rst.Fields(vcamp).Type = 1 Then
         crearwere = vcamp + "=" + IIf(UCase(filtre(i)) = "S", "True", "False")
      GoTo fi
   End If
   If rst.Fields(vcamp).Type = 10 Then
       crearwere = possarweres(vcamp, "LIKE", treure_apostruf(filtre(i)))
       GoTo fi
   End If
   If InStr(1, filtre(i), ",") > 0 Then
       crearwere = vcamp + " in (" + atrim(filtre(i)) + ")"
     Else:
        If Not (Mid(filtre(i), 1, 1) = "<" Or Mid(filtre(i), 1, 1) = ">" Or Mid(filtre(i), 1, 1) = "=") Then
           crearwere = vcamp + "=" + passaradecimalpunt(atrim(cadbl(filtre(i))))
             Else: crearwere = vcamp + passaradecimalpunt(atrim(filtre(i)))
        End If
   End If
   
   
fi:
   Set rst = Nothing
End Function
Function possarweres(ByVal camp As String, condicio As String, ByVal filtre As String) As String
  Dim re As String
'camps(j, 1) + " LIKE '*" + treure_apostruf(filtre(i)) + "*'"
  filtre = filtre + ","
  If camp = "nomclient" And cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1)) > 0 Then camp = "codiclient"
  While InStr(1, filtre, ",") > 0 And filtre <> ""
    If camp <> "codiclient" Then
       re = IIf(re <> "", re + " or ", "") + camp + " like '*" + Mid(filtre, 1, InStr(1, filtre, ",") - 1) + "*'"
      Else: re = IIf(re <> "", re + " or ", "") + camp + " =" + atrim(cadbl(Mid(filtre, 1, InStr(1, filtre, ",") - 1))) + ""
    End If
    filtre = Mid(filtre, InStr(1, filtre, ",") + 1)
  Wend
  If re <> "" Then re = "(" + re + ")"
  possarweres = re
End Function

Private Sub filtreformulacodi_Change()
  filtrarformules
End Sub

Private Sub filtreformuladesc_Change()
   filtrarformules
End Sub

Private Sub filtreformulaserie_Change()
  filtrarformules
End Sub

Private Sub filtretinta_LostFocus(Index As Integer)
  filtrar_tintes
End Sub
Sub filtrar_tintes()
   Dim vprimeraparaula As String
  Dim vsegonaparaula As String
  Dim vterceraparaula As String
  Dim vcampfiltrar As String
  Dim v As String
  Dim vweretinta As String
  Dim vokcarta As String
  Dim vordre As String
  vokcarta = IIf(combosionookcarta = "Sí", "okcarta=true ", IIf(combosionookcarta = "No", "okcarta=false ", ""))
  For i = 0 To filtretinta.Count - 1
    If filtretinta(i) = "<<" Or filtretinta(i) = ">>" Then
       vordre = filtretinta(i).DataField + IIf(filtretinta(i) = "<<", " DESC", " ASC")
        Else: If filtretinta(i) <> "" Then vweretinta = vweretinta + IIf(vweretinta <> "", " and ", "") + " (" + crear_weretinta(filtretinta(i).DataField, filtretinta(i).Text) + ") "
    End If
  Next i
  vweretinta = vokcarta + IIf(vweretinta <> "" And vokcarta <> "", " and ", "") + vweretinta
  tintes.RecordSource = "select * from tintes_tot " + IIf(vweretinta <> "", " where " + vweretinta, "") + IIf(vordre <> "", " order by " + vordre, "")
  tintes.Refresh
End Sub
Function crear_weretinta(vcampfiltrar As String, buscador As String) As String
   Dim vprimeraparaula As String
  Dim vsegonaparaula As String
  Dim vterceraparaula As String
  Dim v As String
  v = treure_apostruf(buscador) + " "
  vprimeraparaula = atrim(Mid(v, 1, InStr(1, v, " ")))
  vsegonaparaula = atrim(Mid(v, InStr(1, v, vprimeraparaula) + Len(vprimeraparaula), InStr(InStr(1, v, vprimeraparaula) + Len(vprimeraparaula), v, " ")))
  If InStr(1, v, vsegonaparaula) + Len(vsegonaparaula) > 1 Then vterceraparaula = atrim(Mid(v, InStr(1, v, vsegonaparaula) + Len(vsegonaparaula)))
  crear_weretinta = vcampfiltrar + " like '*" + vprimeraparaula + "*' " + IIf(vsegonaparaula <> "", " and " + vcampfiltrar + " like '*" + vsegonaparaula + "*'", "") + IIf(vterceraparaula <> "", " and " + vcampfiltrar + " like '*" + vterceraparaula + "*'", "")
End Function

Private Sub fnumllauna_Change()
   filtrarllaunes
End Sub

Sub calcular_preukg_llaunesazero(Optional vnumllauna As String)
Dim rst As Recordset
   Dim vnumlotbase As String
   Dim vpreu As Double
   If vnumllauna <> "" Then
      Set rst = dbtintes.OpenRecordset("select * from dadesllaunesrecargues where numllauna='" + atrim(vnumllauna) + "' order by numllauna Desc")
        Else
          Set rst = dbtintes.OpenRecordset("select * from dadesllaunesrecargues where preuxrkilo<1 or preuxrkilo=null order by numllauna Desc")
   End If
   While Not rst.EOF
     vnumlotbase = ""
     ensenyarlotsbase rst!numllauna, vnumlotbase, True
     vpreu = saber_preu_kg_tinta_llauna(vnumlotbase, atrim(rst!formula))
     If vpreu > 0 Then dbtintes.Execute "update llaunes set preuxrkilo=" + passaradecimalpunt(atrim(vpreu)) + " where numllauna='" + atrim(rst!numllauna) + "'"
     rst.MoveNext
     Me.caption = rst.RecordCount - rst.AbsolutePosition
     DoEvents
   Wend
   Me.caption = "Manteniment de Tintes"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = 112 Then
    Command1_Click
  End If
  If KeyCode = 27 Then
    If tintes.Recordset.EditMode = 0 Then Exit Sub
    tintes.Recordset.CancelUpdate
   tintes.RecordSource = "tintes_tot"
  tintes.Refresh
  
  framedadestintes.Enabled = False
  End If
End Sub
Function ObtenerLíneaComando(Optional MaxArgs)
    'Declara las variables.
    Dim c, LíneaComando, LonLínComando, ArgIn, i, NúmArgs
    'Ver si MaxArgs está.
    If IsMissing(MaxArgs) Then MaxArgs = 10
    'Crea una matriz del tamaño correcto.
    ReDim ArgArray(MaxArgs)
    NúmArgs = 0: ArgIn = False
    'Obtiene los argumentos de la línea de comandos.
    LíneaComando = Command()
    LonLínComando = Len(LíneaComando)
    'Recorre la línea de comando carácter a carácter
    'a la vez.

For i = 1 To LonLínComando
        c = Mid(LíneaComando, i, 1)
        'Comprueba espacio o tabulación.
        If (c <> " " And c <> vbTab) Then
            'Ningún espacio o tabulación.
            'Comprueba si está en el argumento.
            If Not ArgIn Then
            'Empieza el nuevo argumento.
            'Comprueba para más argumentos.
                If NúmArgs = MaxArgs Then Exit For
                    NúmArgs = NúmArgs + 1
                    ArgIn = True
                End If
            'Agrega el carácter al argumento actual.

ArgArray(NúmArgs) = ArgArray(NúmArgs) + c
        Else
            'Encontró un espacio o tabulador.
            'Establece ArgIn a False.
            ArgIn = False
        End If
    Next i
    'Redimensiona la matriz lo suficiente para contener los argumentos.
    'ReDim Preserve ArgArray(NúmArgs)
    'Devuelve la matriz en nombre de la función.
    ObtenerLíneaComando = ArgArray()
End Function
Sub cridar_actualitzarformules()
  gravant = True
  actualitzarformules "T"
End Sub
Private Sub Form_Load()
  Dim arguments As Variant
  arguments = ObtenerLíneaComando
  gravant = False
  rellotge1.Enabled = False
  rellotge1.Enabled = True
  fitxerini = "comandes.ini"
  'If App.PrevInstance Then MsgBox "El programa ja està obert.", vbCritical, "Atenció": End
  cami = llegir_ini("General", "cami", fitxerini)
  escriure_ini "Baixes", "imprimircomanda", "0", "comandes.ini"   'aixó es per si es queda obert el programa d'imprimir la comanda
  ruta_relativa_docs = llegir_ini("ruta", "pautacli", rutadelfitxer(cami) + "valorsprograma.ini")
  ruta_documentacio_clixes = llegir_ini("ruta", "ruta_documentacio_clixes", rutadelfitxer(cami) + "valorsprograma.ini")
  iniconfigreixa = "reixamantenimentintes.ini"
  '"c:\misdoc~1\commandes\comandes.mdb"
  If existeix("c:\ordprog.ini") Then
    cami = "\\serverprodu\dades\progcomandes\dades\comandes.mdb"
    brecalcularpesllaunes.visible = True
  End If
  centerscreen Me
  camitintes = rutadelfitxer(cami) + "tintes.mdb"
  Set dbtintes = DBEngine.OpenDatabase(camitintes)
  Set dbcomandes = DBEngine.OpenDatabase(cami)
  Set dbbaixes = DBEngine.OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
  Set dbcompres = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
  Set dbplanificacioalicia = OpenDatabase(rutadelfitxer(cami) + "planificacio.mdb")
  
  datallaunes.DatabaseName = camitintes
  datadellaunes.DatabaseName = camitintes
  datahistoria.DatabaseName = camitintes
  datalotsbase.DatabaseName = camitintes
  datacomponents.DatabaseName = camitintes
  dataformules.DatabaseName = camitintes
  datarecarregues.DatabaseName = camitintes
  datadetallformules.DatabaseName = camitintes
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  connexiosql
  pestanyes.Tab = 0
  tintes.DatabaseName = camitintes
  tintes.RecordSource = "tintes_tot"
  tintes.Refresh
  If Not tintes.Recordset.EOF Then
    tintes.Recordset.MoveLast
    tintes.Recordset.MoveFirst
  End If
  If atrim(arguments(1)) = "actualitzarformules" Then cridar_actualitzarformules: End
  If atrim(arguments(1)) = "llistattoteslesllaunes" Then enviar_llistat_llaunes: End
  If atrim(arguments(1)) = "agrupartreballs" Then
       Formagrupartreballs.Show
       While isloaded("formagrupartreballs")
          DoEvents
       Wend
       End
  End If
  'filtrarimportaciocolorstreballs
 ' Pestanyes.TabVisible(5) = False
 ' Pestanyes.TabVisible(4) = False
  passarllaunesainactives
  'comprovarestocminim
  datadellaunes.RecordSource = "select * from dadesllaunes"
  datadellaunes.Refresh
  ettotalllaunes.caption = comptar_llaunes
  crear_temporal_comandesactives
  Command52.Enabled = True
  'If IsDate(llegir_ini("Tintes", "controlprogramaobert", fitxerini)) Then
  '   If DateDiff("n", llegir_ini("Tintes", "controlprogramaobert", fitxerini), Now) < 2 Then
  '      Command52.Enabled = False
  '       Else: Command52.Enabled = True
  '   End If
  'End If
  amagar_boto_imprimircomanda
  comprovar_llaunes_offline
  dbbaixes.Execute "delete * from planificacio_reclamades where datediff('d',now,datareclamacio)>10"
  Check1(0).tag = "1"
  Check1(1).tag = "1"
  colocarfiltretinta
  
  Exit Sub
fi:
  Me.tag = "Parar"
 
End Sub
Function isloaded(vnomform As String) As Boolean
  Dim f
  For Each f In Forms
   If UCase(f.Name) = UCase(vnomform) Then
         isloaded = True
   End If
  Next
End Function

Sub enviar_llistat_llaunes()
  vllistatllaunesautomatic = True
  If existeix("c:\temp\Llistat_llaunes_primerdemes.pdf") Then Kill "c:\temp\Llistat_llaunes_primerdemes.pdf"
  executar_llistat_llaunes
  
End Sub
Sub comprovar_llaunes_offline()
  Dim v As String
  If existeix("c:\temp\llaunes_offline.txt") Then
    Open "c:\temp\llaunes_offline.txt" For Input As #1
    If EOF(1) Then Close 1: GoTo fi
    Line Input #1, v
    If Len(atrim(v)) > 4 Then
      Close 1
      forminventari.Show 1
    End If
  End If
fi:
  
End Sub
Function crearnomdelatinta(rst As Recordset)
   Dim rstt As Recordset
   Dim vsql As String
   Dim vwhere As String
   If cadbl(rst!codi) > 0 Then
      Set rstt = dbtintes.OpenRecordset("select descripcio from tintes where codi='" + atrim(rst!codi) + "'")
      If Not rstt.EOF Then crearnomdelatinta = atrim(rstt!descripcio)
       Else
         vsql = "SELECT familiestintes.descripcio, subfamiliestintes.descripcio, familiescolors.descripcio, subfamiliescolors.descripcio FROM (((estocsminims INNER JOIN familiestintes ON estocsminims.idfamilia = familiestintes.codi) INNER JOIN subfamiliestintes ON estocsminims.idsubfamilia = subfamiliestintes.codi) INNER JOIN familiescolors ON estocsminims.idfamcolor = familiescolors.codi) INNER JOIN subfamiliescolors ON estocsminims.idsubfamcolor = subfamiliescolors.codi "
         vsql = "SELECT familiestintes.descripcio, subfamiliestintes.descripcio, familiescolors.descripcio, subfamiliescolors.descripcio FROM familiestintes, subfamiliestintes,familiescolors, subfamiliescolors "
         With rst
          vwhere = " where (idfamilia=" + atrim(cadbl(!idfamilia)) + " and idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + "and idfamcolor= " + atrim(cadbl(!idfamcolor)) + " and idsubfamcolor=" + atrim(cadbl(!idsubfamcolor)) + ") "
          'vwhere = " where (familiestintes.codi=" + atrim(cadbl(!idfamilia)) + " and subfamiliestintes.codi=" + atrim(cadbl(!idsubfamilia)) + "and familiescolors.codi= " + atrim(cadbl(!idfamcolor)) + " and subfamiliescolors.codi=" + atrim(cadbl(!idsubfamcolor)) + ") "
         End With
         vsql = "select descripcio from tintes "
          Set rstt = dbtintes.OpenRecordset(vsql + vwhere)
          If Not rstt.EOF Then crearnomdelatinta = "[" + atrim(rstt!descripcio) + "]"
'          If Not rstt.EOF Then
'            crearnomdelatinta = atrim(rstt![familiestintes.descripcio]) + "  " + atrim(rstt![subfamiliestintes.descripcio]) + "  " + atrim(rstt![familiescolors.descripcio]) + "  " + atrim(rstt![subfamiliescolors.descripcio])
'             Else: Stop
 '         End If
   End If
   
   
   Set rstt = Nothing
End Function

Sub comprovarestocminimdellaunes()
  Dim rst As Recordset
  Dim dbtintes As Database
  Dim vmsg As String
  'Set dbtintes = OpenDatabase(rutadelfitxer(cami) + "tintes.mdb")
  actualitzar_estoc_llaunes
  
  Set rst = dbtintes.OpenRecordset("SELECT * from estocsminims")
  vmsg = ""
  While Not rst.EOF
    If cadbl(rst!estocminim) > cadbl(rst!estocactual) Then
      vmsg = vmsg + crearnomdelatinta(rst) + " ---> Actual " + atrim(rst!estocactual) + " Kg / Mínim " + atrim(rst!estocminim) + " Kg (Desitjat " + atrim(cadbl(rst!estocdesitjat)) + " Kg)" + Chr(10)
    End If
    rst.MoveNext
  Wend
  'If vmsg <> "" Then enviaremail "controlestoctintes", "Control estoc mínim de llaunes", vmsg
  MsgBox vmsg
fi:
  Set rst = Nothing
  Set dbtintes = Nothing
End Sub
Sub comprovarestocminim()
  Dim rst As Recordset
  Set rst = dbtintes.OpenRecordset("SELECT llaunesdecadatintaiestocminim.familiestintes.descripcio AS Familia, llaunesdecadatintaiestocminim.subfamiliestintes.descripcio AS Subfamilia, llaunesdecadatintaiestocminim.familiescolors.descripcio AS [Familia color], llaunesdecadatintaiestocminim.subfamiliescolors.descripcio AS [Subfamilia color], First(llaunesdecadatintaiestocminim.estocminim) AS [Estoc mínim], Sum(llaunesdecadatintaiestocminim.estocactual) AS [Estoc actual] From llaunesdecadatintaiestocminim GROUP BY llaunesdecadatintaiestocminim.familiestintes.descripcio, llaunesdecadatintaiestocminim.subfamiliestintes.descripcio, llaunesdecadatintaiestocminim.familiescolors.descripcio, llaunesdecadatintaiestocminim.subfamiliescolors.descripcio ;")
  While Not rst.EOF
    If cadbl(rst![Estoc mínim]) < cadbl(rst![Estoc actual]) Then bcontrolestocminim.BackColor = QBColor(12): GoTo fi
    rst.MoveNext
  Wend
fi:
  Set rst = Nothing
End Sub
Sub passarllaunesainactives()
  If vconnexioodbc Then
     actualitzarcarguescomponents
     datallaunes.Database.Execute ("UPDATE Llaunes LEFT JOIN Recarregarllaunes ON Llaunes.numllauna = Recarregarllaunes.numllauna SET Llaunes.capacitatactual = 0, Llaunes.activa = False WHERE (((Llaunes.capacitatactual)<1) AND ((Llaunes.activa)=True) AND (llaunes.situacio<>'DOS') AND ((Recarregarllaunes.data) Is Null));")
  End If
End Sub
Sub connexiosql()
   On Error GoTo err
   'If existeix("c:\ordprog.ini") Then GoTo err
   Set wsODBC = CreateWorkspace("", "tintes", "", dbUseODBC)
   Set conODBC = wsODBC.OpenConnection("connexiosql", , True, "ODBC;DATABASE=InkmakerDB;UID=sa;PWD=Mak2008;DSN=tintes")
   Set rstfink = conODBC.OpenRecordset("select * from dbo.tblFormula ")
   Label22(38).visible = False
   vconnexioodbc = True
   Exit Sub
err:
   Label22(38).visible = True
   vconnexioodbc = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   escriure_ini "Baixes", "imprimircomanda", "0", "comandes.ini"
   Unload formcanvisanilox
   End
End Sub

Private Sub fotodescatalogar_Click()
    descatalogartinta
End Sub

Private Sub frefpan_Change()
filtrarimportacio
End Sub

Private Sub inclourevinculats_Click()
  
End Sub

Private Sub fsituaciollauna_Change()
filtrarllaunes
End Sub

Private Sub inclourevinculatsformules_Click()
filtrarimportacio
End Sub

Private Sub inclourevinculatstinta_Click()
filtrarimportacio
End Sub

Private Sub inclourevinculatstintac_Click()
'   filtrarimportaciocolorstreballs
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label26_Click()
   actualitza_llistacomandes "muntats"
End Sub

Private Sub llistacomandes_Click()
   Dim i As Integer
   llistatintes.Clear
   For i = 0 To llistacomandes.ListCount - 1
      If llistacomandes.Selected(i) Then
        actualitza_llistatintes llistacomandes.ItemData(i)
      End If
   Next i
End Sub
Sub veureelpdf(ncomanda As String)
  Dim rstc As Recordset
  Set rstc = dbcomandes.OpenRecordset("select numtreball,numordremodificacio from comandes where comanda=" + atrim(cadbl(ncomanda)))
  obrir_pdf_treball cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio)
  
End Sub
Sub obrir_pdf_treball(treball As Double, modificacio As Double)
   Dim generarfitxer_pdf As String
   Dim generarfitxer_pdf_SC As String
   If modificacio = 0 Then modificacio = 1
   'generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "_SC.pdf"
   'If Not existeix(generarfitxer_pdf) Then
   generarfitxer_pdf = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + ".pdf"
   generarfitxer_pdf_SC = ruta_documentacio_clixes + "\" + Format(treball, "00000") + "\pdf" + Format(treball, "00000") + "-" + Format(modificacio, "000") + "_SC.pdf"
   If existeix(generarfitxer_pdf) And existeix(generarfitxer_pdf_SC) Then
      If MsgBox("Hi ha el pdf per capes disponible." + Chr(10) + "VOLS VEURE'L?", vbInformation + vbYesNo + vbDefaultButton2, "ATENCIÓ") = vbYes Then
         generarfitxer_pdf = generarfitxer_pdf_SC
      End If
   End If
   If existeix(generarfitxer_pdf) Then
     obrir_document generarfitxer_pdf
    Else: MsgBox "No he trobat el fitxer" + Chr(10) + generarfitxer_pdf + Chr(10) + " i tampoc el de separació de colors.", vbCritical, "Error"
  End If
End Sub

Private Sub llistacomandes_DblClick()
  If llistacomandes.ListIndex < 0 Then Exit Sub
  veureelpdf llistacomandes.ItemData(llistacomandes.ListIndex)
End Sub

Private Sub llistatintes_Click()
  ensenyar_extensio
  mirar_semblants
End Sub
Sub mirar_semblants()
  Dim vcodi As String
  Dim rst As Recordset
  Dim rstc As Recordset
  Dim vidtinter As String
  vidtinter = llistatintes.ItemData(llistatintes.ListIndex)
  Command45.BackColor = Command14.BackColor
  Command45.tag = ""
  Set rst = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(vidtinter))
  If Not rst.EOF Then
      Set rstc = dbtintes.OpenRecordset("select * from tintes_semblants where coditintarelacio='" + atrim(cadbl(rst!coditinta)) + "'")
      If Not rstc.EOF Then
        Command45.tag = cadbl(rst!coditinta)
        Command45.BackColor = QBColor(12)
      End If
  End If
  Set rst = Nothing
  Set rstc = Nothing
End Sub
Sub ensenyar_extensio()
   Dim vnumext As String
   etextensio = ""
   If mirarsihihaextensio(llistatintes.ItemData(llistatintes.ListIndex), vnumext) Then
     etextensio = "Extensió: " + vnumext
   End If
End Sub
Private Sub llistatintes_DblClick()
  Dim rst As Recordset
  Dim vabans As String
  Dim vdespres As String
  If llistatintes.ListIndex < 0 Then Exit Sub
  If llistatintes.BackColor <> &HFFFF& Then  'si la llista no es groga
     Set rst = dbclixes.OpenRecordset("select coditinta from tintes where id_tinter=" + atrim(llistatintes.ItemData(llistatintes.ListIndex)))
     buscador.Text = ""
     tintes.RecordSource = "select * from tintes_tot " 'where codi='" + atrim(rst!coditinta) + "'"
     tintes.Refresh
     '  Command44_Click
     tintes.Recordset.FindFirst "codi='" + atrim(rst!coditinta) + "'"
     pestanyes.Tab = 0
       Else
         observacio_idtreball cadbl(llistatintes.tag), vabans, vdespres
         If vabans <> vdespres Then enviar_diferencies_per_email cadbl(llistatintes.tag), vabans, vdespres
  End If
  Set rst = Nothing
End Sub
Sub enviar_diferencies_per_email(numtreball As Double, vabans As String, vdespres As String)
  Dim rst As Recordset
  Dim vmsg As String
  Set rst = dbclixes.OpenRecordset("select * from clixes where id_treball=" + atrim(numtreball))
  If Not rst.EOF Then
     vmsg = "Treball: " + atrim(numtreball) + " Linia: " + atrim(rst!marca) + " - " + atrim(rst!linia) + Chr(10) + Chr(10)
     vmsg = vmsg + "Abans: " + Chr(10) + atrim(vabans) + Chr(10) + Chr(10)
     vmsg = vmsg + "Després: " + Chr(10) + atrim(vdespres)
     enviaremailgeneric "impresores@inplacsa.com", "Canvis de les observacions de treball de màquina desde tintes", vmsg
  End If
  Set rst = Nothing
End Sub

Private Sub llistatintes_GotFocus()
etajudadosclics.visible = True
End Sub

Private Sub llistatintes_LostFocus()
etajudadosclics.visible = False
 
End Sub

Private Sub llistatokcartaentredates_Click()
  Dim rst As Recordset
  Dim weredates As String
  Dim vinici As String
  Dim vfi As String
  Dim valor As String
  Dim i As Integer
   Dim col As Integer

   valor = InputBox("Entra la data d'inici de la consulta.", "Data inici", vdiaanterior)
    If Not IsDate(valor) Then MsgBox "Aquesta data no es vàlida", vbCritical, "Error": Exit Sub
    vinici = valor
    valor = InputBox("Entra la data de fi de la consulta.", "Data fi", Format(Now, "dd/mm/yy"))
    If Not IsDate(valor) Then MsgBox "Aquesta data no es vàlida", vbCritical, "Error": Exit Sub
    vfi = valor
    
  weredates = "dataokcarta>=#" + Format(vinici, "mm/dd/yy") + "# and dataokcarta<=#" + Format(vfi, "mm/dd/yy") + "#"
  'Clipboard.Clear
  'Clipboard.SetText "SELECT tintes.descripcio, tintes.referenciacolor, tintes.dataokcarta, tintes.okcarta From tintes WHERE tintes.okcarta=True and " + weredates + " order by dataokcarta;"
  Set rst = dbtintes.OpenRecordset("SELECT tintes.descripcio, tintes.referenciacolor, tintes.dataokcarta, tintes.okcarta From tintes WHERE tintes.okcarta=True and " + weredates + " order by dataokcarta;")
  If rst.EOF Then GoTo fi
   rst.MoveLast
   rst.MoveFirst
   On Error GoTo errorcrearfitxer
   Open "c:\temp\~llistatokcarta.csv" For Output As #1
   On Error GoTo 0
   Print #1, "Total de Ok carta entre el " + vinici + " i el " + vfi + " es de " + atrim(rst.RecordCount)
   Print #1, " "
   Print #1, " "
   Print #1, "Descripcio;Ref_Color;DataOkCarta"
   While Not rst.EOF
      Print #1, atrim(rst!descripcio) + ";" + atrim(rst!referenciacolor) + ";" + atrim(rst!dataokcarta)
      rst.MoveNext
   Wend
   Close #1
   If existeix("c:\temp\~llistatokcarta.csv") Then obrir_document "c:\temp\~llistatokcarta.csv"
fi:
   Set rst = Nothing
Exit Sub
errorcrearfitxer:
   MsgBox "Error al crear el fitxer d'Excel, mira que no el tinguis obert i torna-ho a provar", vbCritical, "Error"
    GoTo fi
End Sub

Private Sub m_families_Click()
  Load formaltafamilies
  formaltafamilies.caption = "Manteniment Families Tintes"
  formaltafamilies.Data1.DatabaseName = camitintes
  formaltafamilies.subfamilies.DatabaseName = camitintes
  formaltafamilies.Data1.tag = "subfamiliestintes"
  formaltafamilies.Data1.RecordSource = "select * from familiestintes"
  formaltafamilies.refrescar
  'formaltafamilies.width = 7000
  formaltafamilies.DBGrid1.Columns(1).width = 3300
  formaltafamilies.DBGrid1.Columns(2).Button = True
  formaltafamilies.DBGrid1.Columns(2).Locked = True
  formaltafamilies.DBGrid1.Refresh
  formaltafamilies.Show
End Sub

Private Sub m_tipusdecontenidorsmaterials_Click()
  Load formaltarep
  formaltarep.caption = "Tipus de contenidors (Material amb que estan fets)"
  formaltarep.width = formaltarep.width * 2
  formaltarep.Data1.DatabaseName = camitintes
  formaltarep.Data1.RecordSource = "select  descripcio as [Descripcio del contenidor] from contenidors_material order by descripcio"
  formaltarep.refrescar
  formaltarep.Data1.tag = "select  descripcio as [Descripcio del contenidor] from contenidors_material order by descripcio"
  'formaltarep.DBGrid1.Columns(0).visible = False
  formaltarep.DBGrid1.Columns(0).width = 6000
  'formaltarep.DBGrid1.Columns(2).width = 3000
  'formaltarep.DBGrid1.Columns(3).width = 1200
  'formaltarep.DBGrid1.Columns(4).width = 1500
  formaltarep.DBGrid1.width = 8000
  formaltarep.width = 8200
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub mcalcularpreukg_Click()
  If MsgBox("segur que vols recalcular tots els preus?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbYes Then
     calcular_preukg_llaunesazero
  End If
End Sub

Private Sub mdetalldelstinters_Click()
 Load formaltarep
  formaltarep.caption = "Detalls dels tinters"
  formaltarep.Data1.DatabaseName = camitintes
  formaltarep.Data1.RecordSource = "select * from detallsdelstinters"
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(0).width = 150 * 20
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub mfamiliacolors_Click()
   Load formaltarep
  formaltarep.caption = "Families Colors"
  formaltarep.Data1.DatabaseName = camitintes
  formaltarep.Data1.RecordSource = "select * from familiescolors"
  formaltarep.Data1.tag = "select * from tintes where idfamcolor="
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(1).width = 150 * 15
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub mllegenda_Click()
ensenyar_llegenda
End Sub

Private Sub mllistatdellaunesxrpalet_Click()
Dim rst As Recordset
  
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatllaunesperpalet.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "tintes.mdb"
  'oreport.RecordSelectionFormula = "{Llaunes.numllauna}='" + atrim(numllauna) + "'"
  'oreport.Sections("D").ReportObjects.Item("serie").BackColor = posarcolorserie(numllauna)
  'oreport.PaperOrientation = crLandscape
  oreport.DiscardSavedData
 ' If existeix("c:\ordprog.ini") Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
  '  Else
  '    oreport.DisplayProgressDialog = False
 '     oreport.PrintOut False, 1
 ' End If
End Sub

Private Sub mllistatllauneaajuntar_Click()
  buscar_llaunes_perajuntar
End Sub
Sub buscar_llaunes_perajuntar()
  Dim rstll As Recordset
  Dim rstll2 As Recordset
  Dim rsttintes As Recordset
  Dim vmsg As String
  Dim vselect As String
  ratoli "espera"
  Set rsttintes = dbtintes.OpenRecordset("select * from tintes")
  While Not rsttintes.EOF
    vselect = "SELECT llaunes.idtinta, Llaunes.numllauna, tintes.codi, Llaunes.capacitatactual, tintes.descripcio, Llaunes.situacio FROM tintes LEFT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta"
    Set rstll = dbtintes.OpenRecordset(vselect + " where capacitatactual>1 and activa and llaunes.idtinta=" + atrim(rsttintes!idtinta))
    Set rstll2 = dbtintes.OpenRecordset(vselect + " where capacitatactual>1 and activa and llaunes.idtinta=" + atrim(rsttintes!idtinta))
    While Not rstll.EOF
      rstll2.FindFirst "numllauna='" + rstll!numllauna + "'"
      While Not rstll2.EOF
        If rstll2!numllauna <> rstll!numllauna Then
          If (rstll!capacitatactual + rstll2!capacitatactual) < 20 Then vmsg = vmsg + atrim(rstll!descripcio) + "--> " + rstll!numllauna + "(" + atrim(rstll!situacio) + ") " + atrim(rstll!capacitatactual) + "Kg" + "  +  " + rstll2!numllauna + "(" + rstll2!situacio + ") " + atrim(rstll2!capacitatactual) + "Kg" + Chr(13) + Chr(10): rstll2.MoveLast
        End If
        rstll2.MoveNext
      Wend
      rstll.MoveNext
    Wend
    rsttintes.MoveNext
  Wend
  ratoli "normal"
  If vmsg = "" Then vmsg = "CAP LLAUNA PER AJUNTAR..."
 ' MsgBox vmsg, vbInformation, "Llistat de llaunes per ajuntar"
  Open "c:\temp\~llistatllaunesperajuntar.txt" For Output As #1
   Print #1, "LLISTAT DE LLAUNES PER AJUNTAR"
   Print #1, "=============================="
   Print #1, ""
   Print #1, vmsg
  Close #1
  'Clipboard.Clear
  'Clipboard.SetText vmsg
  Shell "notepad.exe 'c:\temp\~llistatllaunesperajuntar.txt'", vbNormalFocus
 ' SendKeys "^V"
 ' Clipboard.Clear
  Set rstll = Nothing
  Set rsttintes = Nothing
  
End Sub

Sub executar_llistat_llaunes(Optional vbidonsde20i25 As Boolean, Optional vnomesdinplacsa As Boolean)
Dim rst As Recordset
  Dim vsqlfiltrarvidons As String
  Dim vmaximkg As Double
  Dim vminimkg As Double
   Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.Report
  Set oapp = New CRAXDDRT.Application
  If vbidonsde20i25 Then vsqlfiltarvidons = "({tintesreferencies.id_bido}=12 or {tintesreferencies.id_bido}=18) AND "
  If vnomesdinplacsa Then vsqlfiltarvidons = vsqlfiltarvidons + "({tintesreferencies.nomproveidor}='INPLACSA') and "
  
  If Not vllistatllaunesautomatic Then
    vminimkg = cadbl(InputBox("Entra el MINIM de Kg de les llaunes que vols filtrar.", "Filtre", "2"))
    If vminimkg < 2 Then vminimkg = 2
    vmaximkg = cadbl(InputBox("Entra el MAXIM de Kg de les llaunes que vols filtrar." + Chr(10) + "Zero per no filtrar", "Filtre", "0"))
    If vmaximkg = 0 Then vmaximkg = 1000000000
         Else: vminimkg = 2: vmaximkg = 1000000000
  End If
  
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "llistatllaunesdisponibles.rpt", 1)
  oreport.Database.Tables.Item(1).Location = rutadelfitxer(cami) + "tintes.mdb"
  oreport.RecordSelectionFormula = vsqlfiltarvidons + "{Llaunes.activa} and ({Llaunes.capacitatactual}>" + atrim(vminimkg) + " and {Llaunes.capacitatactual}<" + atrim(vmaximkg) + ")"
  oreport.DiscardSavedData
  If Not vllistatllaunesautomatic Then
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show 1
    Else
       oreport.ExportOptions.DestinationType = crEDTDiskFile
       oreport.ExportOptions.FormatType = crEFTPortableDocFormat
       oreport.ExportOptions.DiskFileName = "c:\temp\Llistat_llaunes_primerdemes.pdf"
       oreport.Export False
  End If
 End Sub

Private Sub mllllaunes20i25_Click()
   Dim vnomesinplacsa As Boolean
   vnomesinplacsa = IIf(MsgBox("Vols veure nomes les de Inplacsa?", vbInformation + vbDefaultButton2 + vbYesNo, "Inplacsa?") = vbYes, True, False)
   executar_llistat_llaunes True, vnomesinplacsa
End Sub

Private Sub mlltoteslesllaunes_Click()
   executar_llistat_llaunes
End Sub

Private Sub modificar_Click()
   Dim vcoditinta As String
   Dim vcodi As String
   'If Not nohiharelacio(atrim(ccoditinta)) Then
   '         vcodi = InputBoxEx("Aquesta tinta ja te una relació feta no pots modificar-la" + Chr(10) + "Entra la contrasenya per modificarla vigilant.", "Modificacio", , , , , , SPassword)
   '         If UCase(vcodi) <> "EDITARTINTA" Then Exit Sub
   '         GoTo edicio
   'End If
   If UCase(InputBoxEx("Modificar la tinta pot portar problemes de seguretat." + Chr(10) + "Entra la contrasenya de modificació", "Atenció", , , , , , SPassword)) <> "INPLACSA" Then Exit Sub
edicio:
   If tintes.Recordset.EditMode = 0 Then
        vcoditinta = tintes.Recordset!codi
        tintes.RecordSource = "tintes_tot"
        tintes.Refresh
        tintes.Recordset.FindFirst "codi='" + vcoditinta + "'"
        If tintes.Recordset.NoMatch Then MsgBox "Error al localitzar la tinta.": Exit Sub
        tintes.Recordset.Edit
        framedadestintes.Enabled = True
        descripciotinta.tag = tintes.Recordset!descripcio
        crefcolor.SetFocus
          Else: MsgBox "Ja estàs editant...", vbCritical, "Error": Exit Sub
    End If
End Sub

Private Sub mregularitzacioinventarillaunes_Click()
   forminventari.Show 1
End Sub

Private Sub mseries_Click()
Load formaltarep
  formaltarep.caption = "Series dels colors"
  formaltarep.Data1.DatabaseName = camitintes
  formaltarep.Data1.RecordSource = "select * from seriescolors"
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(1).width = 150 * 15
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub msituacions_Click()
Load formaltarep
  formaltarep.caption = "Situacions de les llaunes"
  formaltarep.Data1.DatabaseName = camitintes
  formaltarep.Data1.RecordSource = "select * from situacionsllaunes"
  formaltarep.width = formaltarep.width + (formaltarep.width / 1.4)
  formaltarep.refrescar
  
  
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub msubfamcolors_Click()
  Load formaltarep
  formaltarep.caption = "SubFamilies Colors"
  formaltarep.Data1.DatabaseName = camitintes
  formaltarep.Data1.RecordSource = "select * from subfamiliescolors"
  formaltarep.Data1.tag = "select * from tintes where idsubfamcolor="
  formaltarep.refrescar
  formaltarep.DBGrid1.Columns(1).width = 150 * 15
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Sub escullir_familiatinta()
  Static ultimcodi As String
  Load formseleccio
  formseleccio.caption = "Selecciona Familia Tinta"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "select * from familiestintes order by descripcio"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 4500
  formseleccio.DBGrid2.Columns(2).width = 800
  If cadbl(ultimcodi) > 0 Then formseleccio.Data1.Recordset.FindFirst "codi=" + atrim(ultimcodi)
  formseleccio.Show 1
  If seleccioret = 1 Then
   nomfamilia = atrim(formseleccio.Data1.Recordset!descripcio)
   tintes.Recordset!idfamilia = formseleccio.Data1.Recordset!codi
   csubfamilia = ""
   tintes.Recordset!idsubfamilia = 0
   ultimcodi = atrim(tintes.Recordset!idfamilia)
  End If
  Unload formseleccio
  
End Sub

Private Sub mtipusbidons_Click()
Load formaltarep
  formaltarep.caption = "Tipus de Bidons"
  formaltarep.width = formaltarep.width * 2
  formaltarep.Data1.DatabaseName = camitintes
  formaltarep.Data1.RecordSource = "select  id as codi,nominterndelbido as [Nom intern (INPLACSA)],nombido as [Descripcio pel proveïdor],capacitat,litrescompres as [Ltrs_comprar],tara as [Pes_tara] from tipusbidons"
  formaltarep.refrescar
  formaltarep.Data1.tag = "select * from tintesreferencies where id_bido="
  formaltarep.DBGrid1.Columns(0).visible = False
  formaltarep.DBGrid1.Columns(1).width = 3000
  formaltarep.DBGrid1.Columns(2).width = 3000
  formaltarep.DBGrid1.Columns(3).width = 1200
  formaltarep.DBGrid1.Columns(4).width = 1500
  formaltarep.DBGrid1.width = 12000
  formaltarep.width = 12420
  formaltarep.DBGrid1.Refresh
  formaltarep.Show
End Sub

Private Sub netejafiltraimportacio_Click()
  fdespan = ""
  fdescol = ""
  fdesfam = ""
  frefpan = ""
End Sub

Private Sub mtotalKgrec_Click()
  Dim vinici As String
  Dim vfi As String
  Dim rst As Recordset
  vinici = InputBox("Entra la data d'inici de la consulta:", "Inici")
  If Not IsDate(vinici) Then MsgBox "Data no valida": Exit Sub
  vfi = InputBox("Entra la data de fi de la consulta:", "Fi")
  If Not IsDate(vfi) Then MsgBox "Data no valida": Exit Sub
  
  Set rst = dbtintes.OpenRecordset("SELECT Sum(historiallaunalots.kgtinta) AS SumaDeKg FROM historiallauna LEFT JOIN historiallaunalots ON historiallauna.id = historiallaunalots.idhistoria where datamoviment>=#" + Format(vinici, "mm/dd/yy") + "# and datamoviment<=#" + Format(vfi, "mm/dd/yy") + "#")
  MsgBox "El total de recuperat entre " + vinici + " i " + vfi + " es de:  " + atrim(cadbl(rst!SumadeKg)) + "Kg", vbInformation, "Total recuperat entre dates"

  Set rst = Nothing
End Sub

Private Sub mveuredeltes_Click()
  If isloaded("formcanvisanilox") Then Unload formcanvisanilox Else reixa_deltes
End Sub

Private Sub nomfamilia_DropDown()
escullir_familiatinta
If csubfamilia = "" Then csubfamilia_DropDown
End Sub

Private Sub nomfamilia_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub





Private Sub nomproveidor_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub nomserie_DropDown()
   escullir_serie
   If nomfamilia = "" Then nomfamilia_DropDown
End Sub
Sub escullir_serie()
  Static ultimcodi As String
  Load formseleccio
  formseleccio.caption = "Selecciona la Serie"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "select * from seriescolors order by descripcio"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.DBGrid2.Columns(1).width = 4500
  If cadbl(ultimcodi) > 0 Then formseleccio.Data1.Recordset.FindFirst "codi=" + atrim(ultimcodi)
  formseleccio.Show 1
  If seleccioret = 1 Then
   nomserie = atrim(formseleccio.Data1.Recordset!descripcio)
   tintes.Recordset!idserie = formseleccio.Data1.Recordset!codi
   ultimcodi = atrim(tintes.Recordset!idserie)
  End If
  Unload formseleccio
  
End Sub
Private Sub nomserie_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Text2_Change()

End Sub
Sub filtrarllaunes()
 Dim vfiltrar As String
 Dim vlike As String
 
 'vlike = IIf(semblants.Value = 1, "*", "")
' vfiltrar = " numllauna like '*" + fnumllauna + "*' and codi like '*" + fcoditintallauna + "*' and descripcio like '*" + fdesctintallauna + "*' and formula like '*" + fformulallauna + "*' and situacio = '" + fsituaciollauna + "'"
 vfiltrar = " numllauna like '*" + fnumllauna + "*' and codi like '*" + fcoditintallauna + "*' and descripcio like '*" + fdesctintallauna + "*' " + IIf(fsituaciollauna <> "", "and situacio = '" + fsituaciollauna + "'", "")
 datadellaunes.RecordSource = "SELECT Llaunes.numllauna, tintes.codi, Llaunes.capacitatactual, tintes.descripcio, Llaunes.situacio, tintesreferencies.id_bido,tintesreferencies.referencia, tipusbidons.nombido,tipusbidons.capacitat, Llaunes.activa,llaunes.idtinta FROM tipusbidons RIGHT JOIN ((tintes RIGHT JOIN Llaunes ON tintes.idtinta = Llaunes.idtinta) LEFT JOIN tintesreferencies ON Llaunes.id_refproveidor = tintesreferencies.id) ON tipusbidons.id = tintesreferencies.id_bido WHERE " + IIf(checkactives.Value = 1, " Llaunes.activa=True and ", "") + IIf(checkimpresores.Value = 1, " Llaunes.aimpresores=True and ", "") + vfiltrar
 datadellaunes.Refresh
 datadellaunes.tag = vfiltrar
' MsgBox datadellaunes.RecordSource
 If Not datadellaunes.Recordset.EOF Then
   datadellaunes.Recordset.MoveLast
   datadellaunes.Recordset.MoveFirst
 End If
 datadellaunes.Refresh
 ettotalllaunes.caption = comptar_llaunes
End Sub

Sub historialdimpresiodunacomanda()
  Dim numc As String
  Dim rst As Recordset
  Dim numtreball As Double
  'If botohistorial.tag <> "" Then
  '   botohistorial.BackColor = exportarapdf.BackColor
  '   numcomanda = botohistorial.tag
 '    botohistorial.tag = ""
 '    Command4_Click
 '    Exit Sub
 ' End If
  'ratoli "espera"
  
  
  numc = InputBox("Entra la comanda que vols buscar historial." + Chr(10) + "PD. AQUESTA BUSQUEDA POT TRIGAR UNA MICA", "Historial comanda", numcomanda)
  If cadbl(numc) = 0 Then Exit Sub
  
  numtreball = cadbl(buscartreball(numc))
  If numtreball = 0 Then ratoli "normal": MsgBox "Aquesta comanda no te numero de treball.", vbCritical, "Error": Exit Sub
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "baixes.mdb"
  sql = "SELECT impressores.comanda, First(impressores.numeromaquina) AS Imp, format(First(impressores.datainici),'dd/mm/yy') AS Data, First(comandes.numtreball) AS Treball, Last(comandes.numordremodificacio) AS Ordre FROM impressores RIGHT JOIN comandes ON impressores.comanda = comandes.comanda GROUP BY impressores.comanda, impressores.tipus Having (((First(comandes.numtreball)) = " + atrim(numtreball) + ") And ((impressores.tipus) = 'f')) ORDER BY impressores.comanda DESC , Last(comandes.numordremodificacio) DESC;"

  'formseleccio.Data1.RecordSource = "select comanda from comandes where numtreball=" + atrim(cadbl(idtreball)) + " order by comanda Desc"
  formseleccio.Data1.RecordSource = sql
  formseleccio.caption = "Historial"
  
  'formseleccio.Width = 7000
  ratoli "espera"
  formseleccio.refrescar
  ratoli "normal"
  formseleccio.DBGrid2.Columns(0).width = 2220
  formseleccio.DBGrid2.Columns(1).width = 720
  formseleccio.DBGrid2.Columns(2).width = 2370
  formseleccio.DBGrid2.Columns(3).width = 1500
  formseleccio.DBGrid2.Columns(4).width = 800
  formseleccio.width = 10000
'  formseleccio.DBGrid2.Columns(5).width = 800
  formseleccio.Show 1
  If seleccioret = 1 Then
   botohistorial.tag = numcomanda
   botohistorial.BackColor = QBColor(12)
   numcomanda = formseleccio.Data1.Recordset!comanda
   Command4_Click
  End If
  Unload formseleccio
   ratoli "normal"
End Sub
Function buscartreball(numc As String) As String
  Dim rst As Recordset
  buscartreball = 0
  Set rst = dbcomandes.OpenRecordset("select numtreball from comandes where comanda=" + atrim(cadbl(numc)))
  If Not rst.EOF Then buscartreball = atrim(rst!numtreball)
End Function
Private Sub Pestanyes_Click(PreviousTab As Integer)
  colocarfiltretinta
  mveuredeltes.visible = False
  mllegenda.visible = False
  Frame5(6).visible = False
  If pestanyes.caption = "Albarans" Then
     checktots.Value = 0
     Checkultims30.Value = 0
     actualitzar_llista_albarans
  End If
  If pestanyes.caption = "Compres" Then
     actualitzar_llista_compres
  End If
  If pestanyes.caption = "Comandes Actives" Then
   mveuredeltes.visible = True
   If reixacomandes.Rows < 3 Then
     amagar_boto_imprimircomanda
     reixacomandes.visible = False
     poblar_reixa_comandes
     carregar_amples_reixa
     reixacomandes.visible = True
     carregar_liniadelareixaseleccionada
   End If
     mllegenda.visible = True
  End If
  If pestanyes.caption = "Estoc Tintes" And reixaestocs.Rows < 3 Then
     poblar_reixa_estoc
     carregar_amples_reixa_estoc
  End If
  If pestanyes.caption = "Formules" Then pestanyesforumes.Tab = 0
  
End Sub
Sub colocarfiltretinta()
   Dim c As Byte
   c = 0
   For i = 0 To reixatintes.Columns.Count - 1
     If reixatintes.Columns(i).Left > reixatintes.Left + 100 And (reixatintes.Columns(i).Left + reixatintes.Columns(i).width) < (reixatintes.Left + reixatintes.width) Then
       If c >= filtretinta.Count Then Load filtretinta(c)
       filtretinta(c).width = reixatintes.Columns(i).width
       filtretinta(c).Left = reixatintes.Columns(i).Left + 200
       filtretinta(c).visible = True
       filtretinta(c).DataField = reixatintes.Columns(i).DataField
       c = c + 1
     End If
   Next i
End Sub
Sub amagar_boto_imprimircomanda()
   If EstaCorriendo("Etiquetes tintes.exe") Then Command52.Enabled = False
   If existeix("c:\inkmaker") Then Command52.Enabled = False
End Sub
Sub veurelacomanda(numc As Double)
  escriure_ini "Baixes", "imprimircomanda", atrim(numc), "comandes.ini"
  Shell rutadelfitxer(llegir_ini("General", "rutallistats", "comandes.ini")) + "comandes.exe - imprimir", vbNormalFocus
End Sub

Private Sub Pestanyes_DblClick()
 ' MsgBox fercompraEstocminim(datadellaunes.Recordset)
  'Dim rst As Recordset
  'Set rst = dbtintes.OpenRecordset("SELECT extensions.codiextensio, extensions.volum, extensions.anilox From extensions WHERE (((extensions.volum)>0));")
  'While Not rst.EOF
  '  formextensions.actualitzar_dadesextensioalstreballs rst!codiextensio, rst!anilox, rst!volum
  '  rst.MoveNext
  'Wend
 '  Dim rst As Recordset
 '  Dim rstt As Recordset
 '  Set rst = dbtintes.OpenRecordset("select * from llaunes")
 '  rst.MoveLast
 ' rst.MoveFirst
 ' While Not rst.EOF
 '   Set rstt = dbtintes.OpenRecordset("Select * from tintesreferencies where idtinta=" + atrim(cadbl(rst!idtinta)))
 '   If Not rstt.EOF Then
 '      rst.Edit
 '      rst!id_refproveidor = rstt!id
 '      rst.Update
 '   End If
 '
 '  rst.MoveNext
 '  Me.caption = atrim(rst.AbsolutePosition) + " - " + atrim(rst.RecordCount)
 '    DoEvents
 'Wend
  
End Sub

Private Sub reixacolorstreballs_ButtonClick(ByVal ColIndex As Integer)
  Static ultimcodi As String
  Load formseleccio
  formseleccio.caption = "Selecciona detall tinter"
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = "select detall from detallsdelstinters order by detall"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).width = 3500
  If cadbl(ultimcodi) > 0 Then formseleccio.Data1.Recordset.FindFirst "detall='" + atrim(ultimcodi) + "'"
  
  formseleccio.Show 1
  If seleccioret = 1 Then
   reixacolorstreballs.Columns(ColIndex) = atrim(formseleccio.Data1.Recordset!detall)
   ultimcodi = atrim(formseleccio.Data1.Recordset!detall)
   datacolortreball.Recordset.Edit
   datacolortreball.Recordset!seleccionar = True
   datacolortreball.Recordset.Update
  End If
  If seleccioret = 9 Then reixacolorstreballs.Columns(ColIndex) = ""
  Unload formseleccio
  
End Sub

Function llistadetreballsafectats(nom As String) As String
   Dim rst As Recordset
   llistatreballs.Clear
   Set rst = dbclixes.OpenRecordset("select id_treball,ordremodificacio from tintes where color='" + atrim(nom) + "'")
   While Not rst.EOF
      llistatreballs.AddItem atrim(rst!id_treball) + "/" + atrim(rst!ordremodificacio)
      rst.MoveNext
   Wend
End Function

Private Sub reixacolorstreballs_DblClick()
  If reixacolorstreballs.Columns(reixacolorstreballs.col).DataField = "coditinta" Then
      If MsgBox("Vols eliminar aquesta relació?", vbInformation + vbYesNo, "Atenció") = vbYes Then
          datacolortreball.Recordset.Edit
          datacolortreball.Recordset!coditinta = ""
          datacolortreball.Recordset.Update
          datacolortreball.Refresh
      End If
  End If
End Sub

Private Sub reixacolorstreballs_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 If Screen.ActiveControl.Name <> "reixacolorstreballs" Then Exit Sub
If reixacolorstreballs.col = 0 Then reixacolorstreballs.Columns("seleccionar") = IIf(reixacolorstreballs.Columns("seleccionar") = "Sí", "false", "true")
  datacolortreball.Recordset.Move 0
  
  llistadetreballsafectats reixacolorstreballs.Text
End Sub

Private Sub pestanyesforumes_Click(PreviousTab As Integer)
  If pestanyesforumes.Tab = 2 Then carregar_reixaformulacio
  
End Sub
Sub carregar_reixaformulacio()
   wait 1
   
   With reixaformulacio
    If .Rows = 0 Then .Rows = 1
    .Row = 0
    .Cols = 7
    .TextArray(0) = "Components"
    .TextArray(1) = "Formula1 %"
    .TextArray(2) = "Formula2 %"
    .TextArray(3) = "Formula3 %"
    .TextArray(4) = "DiferenciaGrms"
    .TextArray(5) = "Dif_kg_formula"
    .TextArray(6) = "TotalDosificar"
    .ColWidth(0) = 3000
    .ColWidth(1) = 1700
    .ColWidth(2) = 1700
    .ColWidth(3) = 1700
    .ColWidth(4) = 1700
    .ColWidth(5) = 1700
    .ColWidth(6) = 1700
    .Row = 0:  .ColSel = 2:  .col = 2:  .RowSel = .Rows - 1
    .CellBackColor = formulaacomparar.BackColor
    .Row = 0:  .ColSel = 3:  .col = 3:  .RowSel = .Rows - 1
    .CellBackColor = formulaacomparar2.BackColor
    .Row = 0
    .col = 0
   End With
End Sub

Private Sub reixacomandes_Click()
  Dim vnumc As Double
  Dim vcomandes As String
  Dim vVector As Variant
  
  etextensio = ""
  If isloaded("formcanvisanilox") Then
      vnumc = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("Comanda")))
      vcomandes = formcanvisanilox.FramedeltaE.tag
      If InStr(1, vcomandes, atrim(vnumc)) > 0 Then MsgBox "Aquesta comanda ja està afegida a la reixa.", vbInformation, "Atenció": Exit Sub
      vcomandes = atrim(substituir(vcomandes, " ", ""))
      vVector = MySplit(vcomandes, ",")
      If UBound(vVector) >= LBound(vVector) Then
             For i = LBound(vVector) To UBound(vVector)
                If vVector(i) = "0" Then vVector(i) = vnumc: GoTo cont
             Next i
             If UBound(vVector) < 3 Then
                 ReDim Preserve vVector(UBound(vVector) + 1)
                 vVector(UBound(vVector)) = atrim(vnumc)
             End If
               Else
                ReDim Preserve vVector(UBound(vVector) + 1)
                 vVector(UBound(vVector)) = atrim(vnumc)
      End If
cont:
      vcomandes = ""
      For i = LBound(vVector) To UBound(vVector)
                vcomandes = vcomandes + IIf(vcomandes <> "", ",", "") + atrim(cadbl(vVector(i)))
      Next i
      formcanvisanilox.tag = "nomesdelta " + vcomandes
      formcanvisanilox.ensenyar_deltaE
      formcanvisanilox.SetFocus
  End If
End Sub
Private Function MySplit(ByVal sExpression As String, ByVal sDelimiter As String) As Variant
    Dim arrParts() As String
    Dim lIndex As Long
    Dim lCount As Long
    Dim lDelimiterLen As Long

    ' Si la cadena és buida, retorna un array buit
    If Len(sExpression) = 0 Then
        MySplit = Array() ' Crea un array buit
        Exit Function
    End If

    lDelimiterLen = Len(sDelimiter)
    lCount = 0 ' Comptador d'elements a l'array

    ' Inicialitzem l'array amb una mida petita (es redimensionarà)
    ReDim arrParts(0)

    Do While InStr(sExpression, sDelimiter) > 0
        lIndex = InStr(sExpression, sDelimiter)

        ' Afegeix la part abans del delimitador a l'array
        arrParts(lCount) = Left(sExpression, lIndex - 1)
        lCount = lCount + 1

        ' Redimensiona l'array per al següent element
        ReDim Preserve arrParts(lCount)

        ' Elimina la part processada de la cadena
        sExpression = Mid(sExpression, lIndex + lDelimiterLen)
    Loop

    ' Afegeix l'última part de la cadena (o l'única part si no hi havia delimitadors)
    If Len(sExpression) > 0 Or lCount = 0 Then ' Si hi ha text restant o si no hi havia delimitadors al principi
        arrParts(lCount) = sExpression
        lCount = lCount + 1
    End If

    ' Redimensiona l'array a la mida real dels elements
    If lCount > 0 Then
        ReDim Preserve arrParts(lCount - 1)
    Else
        ReDim arrParts(0) ' En cas que la cadena original només fos el delimitador o res
    End If

    MySplit = arrParts
End Function
Function comptarcaracter(vcadena As String, vcaracter As String) As Long
Dim f As Long
Dim vcomptador As Long
For f = 1 To Len(vcadena)
  If Mid$(vcadena, f, 1) = vcaracter Then vcomptador = vcomptador + 1
Next f
comptarcaracter = vcomptador
End Function

Function numfilaonestaelpunter(Y As Single, reixa As MSFlexGrid) As String
   Dim i As Integer
   Dim N As Double
   For i = 0 To reixa.Rows - 1
     If Y > reixa.RowPos(i) Then N = i ' IIf(i = 0, 0, i - 1)
   Next i
   numfilaonestaelpunter = N
End Function

Function numcolumnaonestaelpunter(X As Single, reixa As MSFlexGrid) As String
   Dim i As Byte
   Dim N As Double
   For i = 0 To reixa.Cols - 1
     If X > reixa.ColPos(i) Then N = i ' IIf(i = 0, 0, i - 1)
   Next i
   numcolumnaonestaelpunter = N
End Function
Function buscardataprevistaentregaclixes(vtreball As Double, vordre As Double) As String
    'Dim rst As Recordset
    'Set rst = dbclixes.OpenRecordset("select data_prevista from clixes_modifi where id_treball=" + atrim(vtreball) + " and ordremodificacio=" + atrim(vordre) + " and descripcioestat='POLIMERS O CLIXES'", , ReadOnly)
    'If Not rst.EOF Then buscardataprevistaentregaclixes = atrim(Format(rst!data_prevista, "dd/mm/yy"))
    'Set rst = Nothing
    buscardataprevistaentregaclixes = buscadatadelclixenous(vtreball, vordre)
End Function
Function buscadatadelclixenous(treball As Double, ordremodificacio As Double) As String
  Dim rst As Recordset
  Dim rstv As Recordset
  Dim rstm As Recordset
  If ordremodificacio = 0 Then ordremodificacio = 1
  buscadatadelclixenous = "     "
  Set rst = dbclixes.OpenRecordset("SELECT clixes_modifi.id_estatclixe,Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, CLIXES_MODIFI.data_prevista,Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE Clixes_modifi.id_treball=" + atrim(treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + " AND clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + ");", dbOpenSnapshot, dbReadOnly)
  Set rstm = dbclixes.OpenRecordset("select reimpres from modificacions where id_treball=" + atrim(treball) + " and ordre=" + atrim(ordremodificacio), dbOpenSnapshot, dbReadOnly)
  ' CONSULTA ULTIMA MODIFICACIO AMB DATA FI  VINCULANT "SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE (((Clixes_modifi.id_treball)=" + atrim(id_treball) + ") AND ((Clixes_modifi.ordremodificacio)=" + atrim(ordremodificacio) + ") AND ((Clixes_estats.vinculant)=True and isdate(clixes_modifi.data_fi)) and clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + "));"
  ' CONSULTA ULTIMA MODIFICACIO AMB DATA FI SENSE VINCULANT "SELECT Clixes_modifi.id_treball, Clixes_modifi.ordremodificacio,clixes_modifi.data_fi, Clixes_estats.descripcio as descrip, Clixes_estats.vinculant, Clixes_modifi.ordre FROM Clixes_modifi INNER JOIN Clixes_estats ON Clixes_modifi.id_estatclixe = Clixes_estats.id_estat WHERE (((Clixes_modifi.id_treball)=" + atrim(id_treball) + ") AND ((Clixes_modifi.ordremodificacio)=" + atrim(ordremodificacio) + ") AND (isdate(clixes_modifi.data_fi)) and clixes_modifi.ordre=(select max(ordre) from clixes_modifi WHERE Clixes_modifi.id_treball=" + atrim(id_treball) + " AND Clixes_modifi.ordremodificacio=" + atrim(ordremodificacio) + "));"
  If Not rst.EOF Then
     If rst!id_estatclixe = 8 Then buscadatadelclixenous = "*" + Format(rst!data_fi, "dd/mm/yy")
     If rst!id_estatclixe = 17 Then buscadatadelclixenous = "#NOCOMANDA"
     If rst!id_estatclixe = 15 Or rst!id_estatclixe = 22 Then buscadatadelclixenous = Format(rst!data_prevista, "dd/mm/yy")
     If rst!id_estatclixe = 19 Then buscadatadelclixenous = "!TORNEM"
     If rst!id_estatclixe = 20 Then buscadatadelclixenous = "REBUTS"
     If buscadatadelclixenous = "     " Then buscadatadelclixenous = "NO ES TE CAP DATA PREVISTA."
  End If
  
  Set rstm = Nothing
  Set rst = Nothing
  Set rstv = Nothing
End Function
Function buscardataentregacomanda(vnumc As Double) As String
   Dim rst As Recordset
   Set rst = dbplanificacioalicia.OpenRecordset("select data1,data2,importancia from planificaciototes where comanda=" + atrim(vnumc))
   If Not rst.EOF Then
       buscardataentregacomanda = IIf(IsDate(rst!Data2), atrim(rst!Data2), atrim(rst!Data1))
       If rst!importancia = 4 Then buscardataentregacomanda = "@" + buscardataentregacomanda
   End If
   If atrim(buscardataentregacomanda) = "" Then buscardataentregacomanda = "Sense Data"
   Set rst = Nothing
   
End Function
Private Sub reixacomandes_DblClick()
    Dim vcanvisituacio As String
    Dim vnumc As Double
    Dim col As Integer
    Dim vdataentrega As String
    Dim vnumtreball As Double
    Dim vcolgestionat As Integer
    Dim vresp As String
    Dim vversio As Double
    Dim vcodidelinia As Double
    Dim vcodideliniav As Double
    Dim i As Byte
    Dim v As String
    
'    dbtintes.Execute "update comandes set proximaseccio='T' where comanda=" + atrim(cadbl(reixacomandes.TextMatrix(reixacomandes.Row, col)))
'    Exit Sub
    vnumtreball = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("NºTreball")))
    vversio = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("Versió")))
    If vnumtreball < 0 Then vnumtreball = vnumtreball * -1
    If reixacomandes.TextMatrix(0, vcolreixacomandes) = "NºTreball" Then
        vdataentrega = buscardataprevistaentregaclixes(vnumtreball, cadbl(reixacomandes.TextMatrix(reixacomandes.Row, vcolreixacomandes + 1)))
        If vdataentrega <> "" Then
          MsgBox "El treball " + atrim(vnumtreball) + "/" + reixacomandes.TextMatrix(reixacomandes.Row, vcolreixacomandes + 1) + " té la data prevista d'entrega de clixes: " + Chr(10) + vdataentrega, vbInformation, "Data entrega clixes."
        End If
        Exit Sub
    End If
    If reixacomandes.TextMatrix(0, vcolreixacomandes) = "CdL" Then
        If Mid(reixacomandes.TextMatrix(reixacomandes.Row, vcolreixacomandes) + " ", 1, 1) = "-" Then
            vresp = InputBox("Es correcte aquest CODI DE LINIA per aquest treball?" + vbNewLine + "ESCRIU [CORRECTE] O [ELIMINAR]", "CODI DE LINIA VERSIÓ DIFERENT")
            If UCase(vresp) = "CORRECTE" Then
              vresp = Mid(reixacomandes.TextMatrix(reixacomandes.Row, vcolreixacomandes), 2)
              vcodidelinia = cadbl(Mid(vresp, 1, InStr(1, vresp, "#") - 1))
              vcodideliniav = cadbl(Mid(vresp, InStr(1, vresp, "#") + 1))
              dbtintes.Execute "update modificacions set codidelinia=" + atrim(vcodidelinia) + ",codideliniav=" + atrim(vcodideliniav) + "  where id_treball=" + atrim(vnumtreball) + " and ordre=" + atrim(vversio)
              dbtintes.Execute "update comandesactives set Codilinia='" + vresp + "' where numtreball=" + atrim(vnumtreball) + " and versiotreball=" + atrim(vversio)
              reixacomandes.TextMatrix(reixacomandes.Row, vcolreixacomandes) = vresp
              reixacomandes.col = vcolreixacomandes
              reixacomandes.Row = reixacomandes.Row
              reixacomandes.CellBackColor = 0
              GoTo SORTIRDIRECTAMENT
            End If
            If UCase(vresp) = "ELIMINAR" Then
              vresp = ""
              dbtintes.Execute "update comandesactives set CodiLinia='" + vresp + "' where numtreball=" + atrim(vnumtreball) + " and versiotreball=" + atrim(vversio)
              reixacomandes.TextMatrix(reixacomandes.Row, vcolreixacomandes) = vresp
              reixacomandes.col = vcolreixacomandes
              reixacomandes.Row = reixacomandes.Row
              reixacomandes.CellBackColor = 0
            End If
        End If
        v = reixacomandes.TextMatrix(reixacomandes.Row, vcolreixacomandes)
        Command56_Click
        For i = 0 To filtre.Count - 1
            If filtre(i).Text = "CdL" Then filtre(i).Text = Mid(v, 1, InStr(1, v, "#"))
        Next i
        filtre(0).SetFocus
        filtre_LostFocus 0
SORTIRDIRECTAMENT:
        Exit Sub
    End If
    If reixacomandes.TextMatrix(0, vcolreixacomandes) = "Secció" Then
        vdataentrega = buscardataentregacomanda(cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("Comanda"))))
        If InStr(1, vdataentrega, "@") Then vdataentrega = substituir(vdataentrega, "@", "")
        If vdataentrega <> "" Then
          MsgBox "Data prevista entrega de la comanda: " + vdataentrega, vbInformation, "Data prevista entrega."
        End If
        Exit Sub
    End If
    
    vcanvisituacio = UCase(InputBox("Entra el nou estat de tramitació d'aquesta comanda.  (S,N,C compres,F formula,M màquina,P preparat)", "Canvi d'estat"))
    
    
    If vcanvisituacio = "P" Or vcanvisituacio = "S" Or vcanvisituacio = "N" Or vcanvisituacio = "C" Or vcanvisituacio = "F" Or vcanvisituacio = "M" Then
       If (vcanvisituacio = "P" Or vcanvisituacio = "M") And hihapantonealacomanda(cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("Comanda")))) And reixacomandes.TextMatrix(reixacomandes.Row, numcol("Gestionat?")) <> "P" Then
            Command13_Click   'crido l'entrada de llaunes a la comanda i marxo
            GoTo fi
       End If
       vnumc = 0
       For col = 0 To reixacomandes.Cols - 1
          If reixacomandes.TextMatrix(0, col) = "Comanda" Then vnumc = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, col))
          If reixacomandes.TextMatrix(0, col) = "Gestionat?" Then vcolgestionat = col: reixacomandes.TextMatrix(reixacomandes.Row, col) = vcanvisituacio
       Next col
       canviarestatcomanda_reixacomandes vnumc, vnumtreball, vcanvisituacio
    End If
fi:
End Sub
Function hihapantonealacomanda(vnumc As Double) As Boolean
   Dim i As Long
   For i = 0 To llistatintes.ListCount - 1
        If InStr(1, UCase(llistatintes.List(i)), " P-") > 0 Then hihapantonealacomanda = True: Exit For
   Next i
End Function
Sub canviarestatcomanda_reixacomandes(vnumc As Double, vnumtreball As Double, vestat As String)
    Dim vcolor As Double
    Dim vfila As Double
       vfila = reixacomandes.Row
       dbtintes.Execute "insert into comandesrevisadesatintes (comanda,numtreball,estatgestio) values (" + atrim(vnumc) + "," + atrim(cadbl(vnumtreball)) + ",'N')"
       dbtintes.Execute "update  comandesactives set gestionat='" + vestat + "' where comanda=" + atrim(vnumc)
       dbtintes.Execute "update  comandesrevisadesatintes set estatgestio='" + vestat + "',numtreball=" + atrim(cadbl(vnumtreball)) + " where comanda=" + atrim(vnumc)
       If vestat = "S" Then vcolor = &HC0FFC0
       If vestat = "C" Then vcolor = QBColor(12)
       If vestat = "F" Then vcolor = QBColor(12)
       If vestat = "N" Then vcolor = QBColor(15)
       If vestat = "M" Then vcolor = QBColor(14)
       If vestat = "P" Then vcolor = &H80C0FF
       reixacomandes.col = numcol("Gestionat?")
       reixacomandes.Row = vfila
       reixacomandes.RowSel = vfila
       reixacomandes.CellBackColor = vcolor
       reixacomandes.col = 0
       reixacomandes.ColSel = reixacomandes.Cols - 1
       reixacomandes.RowSel = vfila
End Sub
Sub canviarestatnovamodificades_reixacomandes(vnumc As Double, vnumtreball As Double, vestat As String)
    Dim vcolor As Double
    Dim vfila As Double
       vfila = reixacomandes.Row
       dbtintes.Execute "insert into comandesrevisadesatintes (comanda,numtreball,estatgestio) values (" + atrim(vnumc) + "," + atrim(cadbl(vnumtreball)) + ",'N')"
       dbtintes.Execute "update  comandesactives set revisatnovamodificada='" + vestat + "' where comanda=" + atrim(vnumc)
       dbtintes.Execute "update  comandesrevisadesatintes set revisatnovamodificada='" + vestat + "',numtreball=" + atrim(cadbl(vnumtreball)) + " where comanda=" + atrim(vnumc)
       If vestat = "S" Then vcolor = &HC0FFC0
       If vestat = "N" Then vcolor = QBColor(15)
       reixacomandes.col = numcol("Nova/Repetida")
       reixacomandes.Row = vfila
       reixacomandes.RowSel = vfila
       reixacomandes.CellBackColor = vcolor
       reixacomandes.col = 0
       reixacomandes.ColSel = reixacomandes.Cols - 1
       reixacomandes.RowSel = vfila
End Sub

Private Sub reixacomandes_LostFocus()
   'guardar_amples_reixa
   carregar_amples_reixa
End Sub

Private Sub reixacomandes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 And vcolreixacomandes > 0 Then
     
     'Me.m_opcions_reixa.WindowList = True
     Me.PopupMenu mtintesrevisades
  End If
  vcolreixacomandes = numcolumnaonestaelpunter(X, reixacomandes)
  If vcolreixacomandes = 0 And Button = 2 Then
      ensenyar_materials_comanda cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("Comanda")))
  End If
  'If Y < reixacomandes.RowHeight(0) Then ordenar_reixacomandes cadbl(vcolreixacomandes)
  
End Sub
Sub ensenyar_materials_comanda(vnumc As Double)
   Dim vmaterialscomandes As String
   Dim rst As Recordset
   Dim rstm As Recordset
   Dim vlinkcomanda2 As Double
   Set rstm = dbcomandes.OpenRecordset("select * from [llistat materials xr familia]")
   Set rst = dbcomandes.OpenRecordset("select comanda,linkcomanda1,linkcomanda2,materialex from comandes where comanda=" + atrim(vnumc))
   rstm.FindFirst "codi=" + atrim(rst!materialex)
   If Not rstm.EOF Then vmaterialscomades = vmaterialscomades + atrim(rst!comanda) + "-" + atrim(rstm![familiesmaterials.descripcio]) + "-" + atrim(rstm![subfamiliesmaterials.descripcio]) + "-" + atrim(rstm![familiescolorants.descripcio]) + vbNewLine
   If rst!linkcomanda1 > 0 Then
    Set rst = dbcomandes.OpenRecordset("select comanda,linkcomanda1,linkcomanda2,materialex from comandes where comanda=" + atrim(rst!linkcomanda1))
    rstm.FindFirst "codi=" + atrim(rst!materialex)
    If Not rstm.EOF Then vmaterialscomades = vmaterialscomades + atrim(rst!comanda) + "-" + atrim(rstm![familiesmaterials.descripcio]) + "-" + atrim(rstm![subfamiliesmaterials.descripcio]) + "-" + atrim(rstm![familiescolorants.descripcio]) + vbNewLine
   End If
   If rst!linkcomanda2 > 0 Then
    Set rst = dbcomandes.OpenRecordset("select comanda,linkcomanda1,linkcomanda2,materialex from comandes where comanda=" + atrim(rst!linkcomanda2))
    rstm.FindFirst "codi=" + atrim(rst!materialex)
    If Not rstm.EOF Then vmaterialscomades = vmaterialscomades + atrim(rst!comanda) + "-" + atrim(rstm![familiesmaterials.descripcio]) + "-" + atrim(rstm![subfamiliesmaterials.descripcio]) + "-" + atrim(rstm![familiescolorants.descripcio]) + vbNewLine
   End If
   If vmaterialscomades <> "" Then MsgBox vmaterialscomades, vbInformation, "Materials de la comanda"
End Sub

Private Sub reixacomandes_SelChange()
  On Error GoTo fi
  If Screen.ActiveControl.Name <> "reixacomandes" And reixacomandes.Rows > 1 Then llistatintes.Clear: Exit Sub
  carregar_liniadelareixaseleccionada
  On Error GoTo 0
  mirar_cuatricumia
fi:
End Sub
Sub mirar_cuatricumia(Optional vCopiarUltimaVersio As Boolean)
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim vnumtreballiversio As String
  Dim vnumversio As Double
  Dim vtreball As Double
  Dim vnumcolor As Double
  Dim i As Byte
  If reixacomandes.Rows = 1 Then Exit Sub
  vnumcolor = &HC78DFA
  vnumversio = cadbl(reixacomandes.TextMatrix(reixacomandes.Row, numcol("Versió")))
  vnumtreball = cadbl(Abs(reixacomandes.TextMatrix(reixacomandes.Row, numcol("NºTreball"))))
  vnumtreballiversio = Trim(vnumtreball) + "/" + atrim(vnumversio)
  Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where numtreballiversio='" + atrim(vnumtreball) + "/" + atrim(vnumversio) + "'")
  Command59.BackColor = &H8000000F
  If rst.EOF And vnumversio > 1 Then
       Set rst = dbtintes.OpenRecordset("select * from valorscuatricomia_treball where numtreballiversio='" + atrim(vnumtreball) + "/" + atrim(cadbl(vnumversio) - 1) + "'")
       If Not rst.EOF Then
          vnumcolor = &HFFFF&
       End If
       If vCopiarUltimaVersio Then
            Set rst2 = dbtintes.OpenRecordset("select * from valorscuatricomia_treball")
            While Not rst.EOF
             rst2.AddNew
             For i = 0 To rst.Fields.Count - 1
                  rst2.Fields(i) = rst.Fields(i)
             Next i
             rst2!numtreballiversio = atrim(vnumtreball) + "/" + atrim(cadbl(vnumversio))
             rst2.Update
             rst.MoveNext
            Wend
            vnumcolor = &HC78DFA
            Command59.BackColor = &HC78DFA
       End If
  End If
  If Not rst.EOF Then
       Command59.BackColor = vnumcolor
  End If
  Set rst = Nothing
  Set rst2 = Nothing
End Sub
Sub carregar_liniadelareixaseleccionada()
Dim col As Integer
  Dim vinici As Double
  Dim vfi As Double
  Static ultimrow As Integer
  Static ultimrowsel As Integer
  Command67(8).visible = False
  Command45.BackColor = Command14.BackColor: Command45.tag = ""
  If Command65.BackColor = &HC0C0FF Then ensenyar_informaciodeltreball: Exit Sub
  If ultimrow = reixacomandes.Row And ultimrowsel = reixacomandes.RowSel Then Exit Sub
  llistatintes.Clear
  vinici = reixacomandes.Row
  vfi = reixacomandes.RowSel
  If vinici > vfi Then vfi = vinici: vinici = reixacomandes.RowSel
'  For col = reixacomandes.Row To reixacomandes.RowSel
'  bassignarllauna.Enabled = True
  If (vfi - vinici) > 1 Then bassignarllauna.Enabled = False
  For col = vinici To vfi
    actualitza_llistatintes cadbl(reixacomandes.TextMatrix(col, 0))
    Label22(29) = "NºMaq: " + maquinaonsimprimiraaquestacomanda(cadbl(reixacomandes.TextMatrix(col, 0)))
    If reixacomandes.Rows - 1 >= col Then possarcolorbotócomandaacabadaicombinacio cadbl(reixacomandes.TextMatrix(col, 0))
  Next col
End Sub
Sub carregar_observacio_tintes(vidtreball As Double, vmodificacio As Double)
  Dim rst As Recordset
 
  buscador_estoc(1) = ""
  
  Set rst = dbclixes.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(vidtreball) + " and ordre=" + atrim(900)) ' + vmodificacio))
  If Not rst.EOF Then buscador_estoc(1) = atrim(rst!observacio)
  Set rst = Nothing
End Sub
Sub possarcolorbotócomandaacabadaicombinacio(vnumc As Double)
   Dim rst As Recordset
   If vnumc = 0 Then Exit Sub
   Set rst = dbtintes.OpenRecordset("select datacomandapreparada,combinaciollaunesfeta from comandesrevisadesatintes where comanda=" + atrim(vnumc))
   Command13.BackColor = &H8080FF
   Command67(8).visible = False
   If Not rst.EOF Then
      If IsDate(rst!datacomandapreparada) Then Command13.BackColor = &H6BEBB1
      If rst!combinaciollaunesfeta Then Command67(8).visible = True
   End If
   Set rst = Nothing
End Sub
Private Sub reixacomponents_BeforeDelete(Cancel As Integer)
   If Not datalotsbase.Recordset.EOF Then
      MsgBox "No pots borrar aquest component perquè ja s'ha utilitzat i es perdria l'historia.", vbCritical, "Atenció"
      Cancel = 1
      Exit Sub
   End If
   If MsgBox("Estas segur que vols borrar aquest component?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then
      Cancel = 1
   End If
End Sub

Private Sub reixacomponents_ButtonClick(ByVal ColIndex As Integer)
  Dim sql As String
  Dim coditinta As String
  Dim desctinta As String
  If MsgBox("Segur que vols relacionar aquest surtidor amb una tinta?", vbCritical + vbDefaultButton2 + vbYesNo, "Atenció") = vbNo Then Exit Sub
  sql = "SELECT  idtinta,codi,descripcio,referenciacolor from tintes_tot "
  Load formseleccio
  formseleccio.Data1.DatabaseName = camitintes
  formseleccio.Data1.RecordSource = sql
  formseleccio.width = 13000
  formseleccio.sortirs.tag = "filtre"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).visible = False
  formseleccio.Show 1
  If seleccioret = 0 Then Exit Sub
  If seleccioret = 1 Then
    coditinta = atrim(formseleccio.Data1.Recordset!codi)
    desctinta = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  If seleccioret = 9 Then
    coditinta = 0
    desctinta = ""
  End If
  Unload formseleccio
  datacomponents.Recordset.Edit
  datacomponents.Recordset!coditintarelacionada = atrim(coditinta)
  datacomponents.Recordset!nomtintarelacionada = atrim(desctinta)
  datacomponents.Recordset.Update
End Sub

Private Sub reixacomponents_DblClick()
   Dim vresp As String
   If reixacomponents.Columns(reixacomponents.col).caption = "Base" Then
        vresp = UCase(InputBox("Aquest components es una BASE?" + vbNewLine + "POSSA UNA LLETRA PER RELACIONAR LES BASES ENTRE ELLES.", "Ès Base?"))
        If StrPtr(vresp) = 0 Then Exit Sub
        If vresp = "" Then vresp = " "
        datacomponents.Recordset.Edit
        datacomponents.Recordset!esbase = vresp
        datacomponents.Recordset.Update
   End If
End Sub

Private Sub reixaestocs_DblClick()
   Dim rst  As Recordset
   Dim i As Integer
   If reixaestocs.TextMatrix(0, reixaestocs.col) = "Estoc necessari" Then
         Set rst = dbtintes.OpenRecordset("select * from consultaestocs " + buscador_estoc(0).tag + IIf(buscador_estoc(0).tag <> "", " and ", " where ") + " comandesinplicades<>'' " + ordreestoc)
         If rst.EOF Then GoTo fi
         For i = 1 To reixaestocs.Row - 1
            If Not rst.EOF Then rst.MoveNext
         Next i
         If Not rst.EOF Then
            'MsgBox atrim(rst!comandesinplicades)
            Command56_Click
            For i = 0 To filtre.Count - 1
              If filtre(i).Text = "Comanda" Then filtre(i).SetFocus: filtre(i).Text = atrim(rst!comandesinplicades): filtre_LostFocus i
            Next i
            pestanyes.Tab = 4
         End If
      Else
       ensenyalatintacorresponent
   End If
fi:
 Set rst = Nothing
End Sub
Sub ensenyalatintacorresponent()
    Dim rst As Recordset
     buscador.Text = ""
     vnomtinta = reixaestocs.TextMatrix(reixaestocs.Row, 1)
     If Mid(vnomtinta + " ", 1, 1) = "[" Then
        vnomtinta = Mid(vnomtinta, 2)
        vnomtinta = Mid(vnomtinta, 1, Len(vnomtinta) - 1)
     End If
     tintes.RecordSource = "select * from tintes_tot " 'where codi='" + atrim(rst!coditinta) + "'"
     tintes.Refresh
     tintes.Recordset.FindFirst "descripcio='" + treure_apostruf(vnomtinta) + "'"
     pestanyes.Tab = 0
      Command44_Click
     Set rst = Nothing
End Sub
Private Sub reixaestocs_LostFocus()
   guardar_amples_reixa_estoc
   carregar_amples_reixa_estoc
End Sub

Private Sub reixaestocs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Y < reixaestocs.CellHeight Then
     etfiltrarperestoc = "Filtrar per " + reixaestocs.TextMatrix(0, reixaestocs.col)
     etfiltrarperestoc.tag = campsestoc(reixaestocs.col)
  End If
End Sub

Private Sub reixaformulacio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 And (reixaformulacio.col = 2 Or reixaformulacio.col = 3) Then
       'mcopiar.tag = ""
       'mcopiar.Enabled = False
       'mpegar.Enabled = False
       If mcopiar.tag = "" Then
          mpegar.caption = "Pegar"
          mpegar.Enabled = False
            Else: mpegar.caption = "Pegar (" + mcopiar.tag + ")"
       End If
       Me.PopupMenu m_menucopiarpegar
   End If
End Sub
Private Sub mcopiar_click()
     mcopiar.tag = reixaformulacio.Text
     mpegar.Enabled = True
End Sub
 Private Sub mpegar_click()
     reixaformulacio.Text = mcopiar.tag
     mpegar.caption = "Pegar "
     DoEvents
     mpegar.Enabled = False
 End Sub

Private Sub reixaformules_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rst As Recordset
   Dim rstll As Recordset
   llistallaunesformula.Clear
   If dataformules.Recordset.EOF Then Exit Sub
   Set rst = dbtintes.OpenRecordset("select * from tintes_tot where tintes_tot.idtinta in (select idtinta from tintesformules where numformula='" + atrim((dataformules.Recordset!codiformula)) + "')")
   While Not rst.EOF
     Set rstll = dbtintes.OpenRecordset("select * from llaunes where activa=true and idtinta=" + atrim(cadbl(rst!idtinta)))
     While Not rstll.EOF
       llistallaunesformula.AddItem rstll!numllauna + " --> " + justificar(atrim(rstll!situacio), 4, "D") + " Sit  " + atrim(rstll!capacitatactual) + "Kg"
       rstll.MoveNext
     Wend
     rst.MoveNext
   Wend
   Set rstll = Nothing
   Set rst = Nothing
End Sub

Private Sub reixagrmskilo_DblClick()
  Dim v As String
  If reixagrmskilo.col <> 1 Then Exit Sub
  v = InputBox("Entra el valor de " + atrim(reixagrmskilo.TextMatrix(reixagrmskilo.Row, 0)) + ":", "Entra el valor")
  If StrPtr(v) = 0 Then Exit Sub
  reixagrmskilo.Text = atrim(cadbl(v))
  recalcular_reixagrmskilo
End Sub
Sub recalcular_reixagrmskilo()
  Dim vc As Double
  Dim vB5 As Double
  vB5 = cadbl(reixagrmskilo.TextMatrix(1, 1))
  For i = 2 To 11
    If vB5 > 0 Then
        vc = (cadbl(reixagrmskilo.TextMatrix(i, 1)) / vB5) * 10
        reixagrmskilo.TextMatrix(i, 2) = Redondejar(vc / 100, 2)
    End If
  Next i
End Sub
Private Sub reixahistoria_DblClick()
  Dim v As Double
  Dim vllauna As String
  If reixahistoria.col = 4 Then
     v = cadbl(InputBox("Entra els kilos correctes", "Atenció", reixahistoria.Text))
     If (cadbl(datallaunes.Recordset!capacitat) + 5) < v Then
         MsgBox "Has escrit mes kilos que el que hi cap al bidó" + Chr(10) + "Capacitat llauna: " + atrim(datallaunes.Recordset!capacitat) + " Kg  --->  Valor introduït:  " + atrim(v) + " Kg", vbCritical, "Error": Exit Sub
     End If
     If v >= 0 And v <> cadbl(reixahistoria.Text) Then
        vllauna = datallaunes.Recordset!numllauna
        datahistoria.Recordset.Edit
        datahistoria.Recordset!kg = v
        datahistoria.Recordset.Update
        calcularkgdisponiblesllauna datallaunes.Recordset!numllauna
        datallaunes.Refresh
        datallaunes.Recordset.FindFirst "numllauna='" + atrim(vllauna) + "'"
     End If
  End If
End Sub

Private Sub reixaimportacio_DblClick()
  If reixaimportacio.Columns(reixaimportacio.col).DataField = "coditinta" Then
      If MsgBox("Vols eliminar aquesta relació?", vbInformation + vbYesNo, "Atenció") = vbYes Then
          dataimportacio.Recordset.Edit
          dataimportacio.Recordset!coditinta = ""
          dataimportacio.Recordset.Update
          dataimportacio.Refresh
      End If
  End If
  If reixaimportacio.Columns(reixaimportacio.col).DataField = "codiformula" Then
      If MsgBox("Vols eliminar aquesta relació?", vbInformation + vbYesNo, "Atenció") = vbYes Then
          dataimportacio.Recordset.Edit
          dataimportacio.Recordset!codiformula = ""
          dataimportacio.Recordset.Update
          dataimportacio.Refresh
      End If
  End If
End Sub

Private Sub reixaimportacio_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If Screen.ActiveControl.Name <> "reixaimportacio" Then Exit Sub
If reixaimportacio.col = 0 Then reixaimportacio.Columns("seleccionar") = IIf(reixaimportacio.Columns("seleccionar") = "Sí", "false", "true")
  dataimportacio.Recordset.Move 0
End Sub

Private Sub reixarefproveidor_Click()

End Sub

Private Sub reixallaunes_ButtonClick(ByVal ColIndex As Integer)
Dim vidmaterialcontenidor As Long
Dim vidproveidorrecuperador As Long
Dim vid_refproveidor As Long
Dim capacitatllauna As Double
  Dim vnumllauna As String
  vnumllauna = reixallaunes.Columns("numllauna")
  If datallaunes.Recordset.EditMode = 0 Then MsgBox "Primer has d'editar per modificar algun camp.", vbCritical, "Editar": Exit Sub
  If reixallaunes.Columns(reixallaunes.col).DataField = "referencia" Then
      vid_refproveidor = escullir_referenciaproveidor(tintes.Recordset!idtinta, capacitatllauna)
      If cadbl(vid_refproveidor) = 0 Then GoTo fi
      dbtintes.Execute "update llaunes set id_refproveidor=" + atrim(cadbl(vid_refproveidor)) + " where numllauna='" + vnumllauna + "'"
      If capacitatllauna > 180 Then GoTo demanarmaterialcontenidor
      dbtintes.Execute "update llaunes set idmaterialcontenidor=null,idproveidorrecuperador=null where numllauna='" + vnumllauna + "'"
      datallaunes.Refresh
      datallaunes.Recordset.FindFirst "numllauna='" + numllauna + "'"
  End If
  If reixallaunes.Columns(reixallaunes.col).DataField = "activa" Then
     ' MsgBox datallaunes.Recordset!numllauna
      If Not datallaunes.Recordset.activa Then
        If MsgBox("Vols activar aquesta llauna?", vbInformation + vbDefaultButton2 + vbYesNo, "Activar llauna?") = vbYes Then
           datallaunes.Database.Execute "update llaunes set activa=true where numllauna='" + atrim(datallaunes.Recordset!numllauna) + "'"
           datallaunes.UpdateControls
         End If
      End If
  End If
  If reixallaunes.Columns(reixallaunes.col).DataField = "nomcontenidor" Then
demanarmaterialcontenidor:
      vnumllauna = reixallaunes.Columns("numllauna")
      escullir_material_contenidor vidmaterialcontenidor
      If vidmaterialcontenidor = 0 Then
         If MsgBox("No has escullit cap contenidor, VOLS BORRAR AQUEST QUE HI HA ASSIGNAT?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo fi
      End If
      MsgBox "Ara has d'escullir el Recuperador de contenidors que el recullirà.", vbInformation, "Recuperador"
      escullir_proveidorrecuperador vidproveidorrecuperador
      If vidproveidorrecuperador = 0 Then
         If MsgBox("No has escullit cap RECUPERADOR de contenidors, VOLS BORRAR AQUEST QUE HI HA ASSIGNAT?", vbInformation + vbYesNo + vbDefaultButton2, "Atenció") = vbNo Then GoTo fi
      End If
      dbtintes.Execute "update llaunes set idmaterialcontenidor=" + atrim(cadbl(vidmaterialcontenidor)) + ",idproveidorrecuperador=" + atrim(cadbl(vidproveidorrecuperador)) + " where numllauna='" + vnumllauna + "'"
      datallaunes.Refresh
      datallaunes.Recordset.FindFirst "numllauna='" + numllauna + "'"
  End If
fi:
End Sub

Private Sub reixalotsbase_GotFocus()
   etnumlot.visible = True
End Sub

Private Sub reixalotsbase_LostFocus()
etnumlot.visible = False
End Sub

Private Sub reixalotsbase_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   Dim vnumlot As String
   If datalotsbase.Recordset.EOF Then GoTo fi
   vnumlot = buscarlotinplacsadelallauna(datalotsbase.Recordset!numerodelot)
   If vnumlot = "" Then vnumlot = datalotsbase.Recordset!numerodelot
fi:
   etnumlot.caption = "Nº Lot: " + vnumlot
End Sub

Private Sub reixarecarregues_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not datarecarregues.Recordset.EOF Then
            datahistoria.RecordSource = "select * from historiallauna where idnumllauna=" + atrim(datallaunes.Recordset!id) + " and numrecarrega=" + atrim(datarecarregues.Recordset!numrecarrega) + " order by data Desc"
            datahistoria.Refresh
              Else
                datahistoria.RecordSource = "select * from historiallauna where idnumllauna=-999 order by data Desc"
                datahistoria.Refresh
    End If
End Sub

Private Sub reixatintes_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
  colocarfiltretinta
End Sub

Private Sub reixatintes_GotFocus()
  If tintes.RecordSource = "tintes" Then
    tintes.RecordSource = "tintes_tot"
    tintes.Refresh
  End If
End Sub

Private Sub reixatintes_SelChange(Cancel As Integer)
  If reixatintes.SelStartCol >= 0 Then
     etfiltrarper = "Filtrar per " + reixatintes.Columns(reixatintes.SelStartCol).caption
     etfiltrarper.tag = reixatintes.Columns(reixatintes.SelStartCol).DataField
  End If
End Sub

Private Sub rellotge1_Timer()
    
    estattaula = posarestattaula(tintes.Recordset.EditMode)
    If estattaula = "" Then
       framedadestintes.Enabled = False
         Else: framedadestintes.Enabled = True
    End If
End Sub
Function posarestattaula(estat As Integer) As String
   Static num As Integer
   num = num + 1
   If num = 5 Then num = 1
   If estat = 1 Then posarestattaula = "Editant"
   If estat = 2 Then posarestattaula = "Afegint"
   posarestattaula = posarestattaula + String(num, ".")
   If estat = 0 Then posarestattaula = ""
End Function


'Sub filtrarimportaciocolorstreballs(Optional senseno As Boolean, Optional filtrar As Boolean)
 'Dim vfiltrar As String
 'vfiltrar = crearfiltredesdetexte("nomcolor", cfiltrecolorstreballs)
 'datacolortreball.RecordSource = "select * from tmp_colorsdelstreballs where " + vfiltrar + IIf(inclourevinculatstintac.Value = 0, " and (coditinta='' or coditinta = null)", "")
 'If filtrar Then datacolortreball.RecordSource = datacolortreball.RecordSource + " and seleccionar=" + IIf(senseno, "True", "False")
' datacolortreball.Refresh
' If Not datacolortreball.Recordset.EOF Then
'   datacolortreball.Recordset.MoveLast
'   datacolortreball.Recordset.MoveFirst
' End If
'
' reixacolorstreballs.Refresh'
'
'End Sub
Function crearfiltredesdetexte(camp As String, texte As String) As String
  
  Dim vseleccio As String
  Dim vvalor As String
  Dim vmid As String
  
  vvalor = texte
  While InStr(1, vvalor, ",") And Len(vvalor) > 1
     vmid = Mid(vvalor, 1, InStr(1, vvalor, ",") - 1)
     vvalor = Mid(vvalor, InStr(1, vvalor, ",") + 1)
     vseleccio = vseleccio + IIf(vseleccio <> "", " and ", "") + camp + " like '*" + vmid + "*'"
  Wend
  vseleccio = vseleccio + IIf(vseleccio <> "", " and ", "") + camp + " like '*" + vvalor + "*'"
  crearfiltredesdetexte = vseleccio
End Function

Sub filtrarformules()
 Dim vfiltrar As String
 vfiltrar = " descripcioformula like '*" + filtreformuladesc + "*' and codiformula like '*" + filtreformulacodi + "*' and series like '*" + filtreformulaserie + "*' "
 If Check1(4).Value <> 1 Then
        dataformules.RecordSource = "select * from formules where " + vfiltrar
         Else: dataformules.RecordSource = "SELECT Formules.* FROM Formules RIGHT JOIN FormulesAmbLlaunesactives ON Formules.codiformula = FormulesAmbLlaunesactives.numformula where " + vfiltrar
 End If
 dataformules.Refresh
 reixaformules.Refresh
End Sub

Private Sub semblants_Click()
   filtrarimportacio
End Sub

Private Sub timercontrolalbaransnous_Timer()
    
End Sub


Private Sub timercontrolfocus_Timer()
   vfocusultimcontrol = ""
   timercontrolfocus.Enabled = False
End Sub

Private Sub tintes_Reposition()
   carregar_lookups
   If Not tintes.Recordset.EOF Then
     tintes.caption = "Tinta: " + atrim(tintes.Recordset.AbsolutePosition + 1) + " / " + atrim(tintes.Recordset.RecordCount)
   End If
'   possar_mostra_color_tinta
End Sub
Sub possar_mostra_color_tinta()
   fmostracolortinta.BackColor = &H8000000F
   If Not tintes.Recordset.EOF Then
      If Len(tintes.Recordset!referenciacolor) > 3 Then
         fmostracolortinta.BackColor = buscar_hex_delpantone(Mid(tintes.Recordset!referenciacolor, 3))
      End If
   End If
End Sub
Sub carregar_lookups()
  Dim rst As Recordset
  Dim vsql As String
  Dim vdata As Date

  'If tintes.RecordSource = "tintes" Then Exit Sub
  If tintes.RecordSource = "tintes" Or tintes.Recordset.EOF Then
    descripciotinta.tag = ""
    nomserie = ""
    nomfamilia = ""
    csubfamilia = ""
    cfamiliacolor = ""
    csubfamiliacolor = ""
    datallaunes.RecordSource = "select * from llaunes where idtinta=-9999"
     datallaunes.Refresh
     datatintesformules.RecordSource = "select * from tintesformules where idtinta=-9999"
     datatintesformules.Refresh
     datahistoria.RecordSource = "select * from historiallauna where idnumllauna=-9999"
     datahistoria.Refresh
     'datarefproveidor.RecordSource = "SELECT tintesreferencies.id, tintesreferencies.referencia, tipusbidons.capacitat,tipusbidons.nombido  FROM tintesreferencies LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id where idtinta=-999 order by predeterminada ;"
     'datarefproveidor.Refresh

    Exit Sub
  End If
  nomserie = atrim(tintes.Recordset.Fields(nomserie.DataField))
  nomfamilia = atrim(tintes.Recordset.Fields(nomfamilia.DataField))
  csubfamilia = atrim(tintes.Recordset.Fields(csubfamilia.DataField))
  possarcolorcsubfamilia cadbl(tintes.Recordset!idtinta)
  cfamiliacolor = atrim(tintes.Recordset.Fields(cfamiliacolor.DataField))
  csubfamiliacolor = atrim(tintes.Recordset.Fields(csubfamiliacolor.DataField))

  buscar_estocs_minims
  
  vsql = "SELECT Llaunes.id, Llaunes.numllauna, Llaunes.preuxrkilo, Llaunes.capacitatactual AS kgactuals, comandesrevisadesatintes.estatgestio, [Llaunes].[situacio]+IIf([llaunes].[aimpresores],'*'+Trim(' ' & [comandesrevisadesatintes].[estatgestio]),'') AS situacioiimp, tintesreferencies.referencia, tipusbidons.capacitat, Contenidors_material.descripcio AS nomcontenidor, Llaunes.activa FROM ((assignaciollaunesacomandes RIGHT JOIN ((Llaunes LEFT JOIN tintesreferencies ON Llaunes.id_refproveidor = tintesreferencies.id) LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id) ON assignaciollaunesacomandes.numllauna = Llaunes.numllauna) LEFT JOIN comandesrevisadesatintes ON assignaciollaunesacomandes.comanda = comandesrevisadesatintes.comanda) LEFT JOIN Contenidors_material ON Llaunes.idmaterialcontenidor = Contenidors_material.codi "
  'datallaunes.RecordSource = "SELECT  Llaunes.id,Llaunes.numllauna,llaunes.preuxrkilo, Llaunes.capacitatactual  as kgactuals, Llaunes.situacio+iif(llaunes.aimpresores,'*','') as situacioiimp, tintesreferencies.referencia, tipusbidons.capacitat, Llaunes.activa FROM (Llaunes LEFT JOIN tintesreferencies ON Llaunes.id_refproveidor = tintesreferencies.id) LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id where llaunes.idtinta=" + atrim(cadbl(tintes.Recordset!idtinta)) + " order by activa"
  'datallaunes.RecordSource = "SELECT Llaunes.id, Llaunes.numllauna, Llaunes.preuxrkilo, Llaunes.capacitatactual AS kgactuals, comandesrevisadesatintes.estatgestio, [Llaunes].[situacio]+IIf([llaunes].[aimpresores],'*'+Trim(' ' & [comandesrevisadesatintes].[estatgestio]),'') AS situacioiimp, tintesreferencies.referencia, tipusbidons.capacitat, Llaunes.activa FROM (assignaciollaunesacomandes RIGHT JOIN ((Llaunes LEFT JOIN tintesreferencies ON Llaunes.id_refproveidor = tintesreferencies.id) LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id) ON assignaciollaunesacomandes.numllauna = Llaunes.numllauna) LEFT JOIN comandesrevisadesatintes ON assignaciollaunesacomandes.comanda = comandesrevisadesatintes.comanda Where (((Llaunes.idtinta) = " + atrim(cadbl(tintes.Recordset!idtinta)) + "))ORDER BY Llaunes.activa;"
  datallaunes.RecordSource = vsql + " Where (((Llaunes.idtinta) = " + atrim(cadbl(tintes.Recordset!idtinta)) + "))ORDER BY Llaunes.activa;"
 ' Clipboard.Clear
 ' Clipboard.SetText "SELECT Llaunes.id, Llaunes.numllauna, Llaunes.preuxrkilo, Llaunes.capacitatactual AS kgactuals, comandesrevisadesatintes.estatgestio, [Llaunes].[situacio]+IIf([llaunes].[aimpresores],'*'+Trim(' ' & [comandesrevisadesatintes].[estatgestio]),'') AS situacioiimp, tintesreferencies.referencia, tipusbidons.capacitat, Llaunes.activa FROM (assignaciollaunesacomandes RIGHT JOIN ((Llaunes LEFT JOIN tintesreferencies ON Llaunes.id_refproveidor = tintesreferencies.id) LEFT JOIN tipusbidons ON tintesreferencies.id_bido = tipusbidons.id) ON assignaciollaunesacomandes.numllauna = Llaunes.numllauna) LEFT JOIN comandesrevisadesatintes ON assignaciollaunesacomandes.comanda = comandesrevisadesatintes.comanda Where (((Llaunes.idtinta) = " + atrim(cadbl(tintes.Recordset!idtinta)) + "))ORDER BY Llaunes.activa;"

  datallaunes.Refresh
  datatintesformules.RecordSource = "select * from tintesformules where idtinta=" + atrim(cadbl(tintes.Recordset!idtinta)) + " order by predeterminada"
  datatintesformules.Refresh
  
  'observacions de tinta
  cobservacions = ""
  Set rst = dbtintes.OpenRecordset("select * from tintes_observacions where idtinta=" + atrim(cadbl(tintes.Recordset!idtinta)))
  If Not rst.EOF Then cobservacions = atrim(rst!observacio)
  If UCase(Mid(crefcolor + "  ", 1, 2)) = "P-" Then
    Command44.Enabled = False
     Else: Command44.Enabled = True
  End If

End Sub
Sub buscar_estocs_minims()
  Dim rst As Recordset
  Dim vsubconsulta As String
 With tintes.Recordset
  vsubconsulta = " idfamilia=" + atrim(cadbl(!idfamilia)) + " and idsubfamilia=" + atrim(cadbl(!idsubfamilia)) + " and idfamcolor=" + atrim(cadbl(!idfamcolor)) + " and idsubfamcolor=" + atrim(cadbl(!idsubfamcolor))
  End With
  Set rst = dbtintes.OpenRecordset("select estocminim,estocdesitjat from estocsminims where " + vsubconsulta)
  botoestocminim.BackColor = &H80FF80
  botoestocminim.tag = ""
   botoestocminim.HelpContextID = 0
  botoestocminim.caption = "Estoc mínim"
  If Not rst.EOF Then
    If cadbl(rst!estocminim) > 0 Then
      botoestocminim.BackColor = QBColor(12)
      botoestocminim.tag = cadbl(rst!estocminim)
      botoestocminim.HelpContextID = cadbl(rst!estocdesitjat)
      botoestocminim.caption = "Estoc mínim (" + botoestocminim.tag + ")"
    End If
  End If
End Sub
Sub possarcolorcsubfamilia(idtinta As Long)
  Dim rst As Recordset
  Dim rstcolor As Recordset
  Set rst = dbtintes.OpenRecordset("select * from subfamiliestintes where codi=(select idsubfamilia from tintes where idtinta=" + atrim(cadbl(idtinta)) + ")")
  csubfamilia.BackColor = QBColor(15)
  If Not rst.EOF Then
      Set rstcolor = dbtintes.OpenRecordset("select * from colorsetiquetes where nomcolor='" + atrim(rst!color) + "'")
      If Not rstcolor.EOF Then csubfamilia.BackColor = QBColor(rstcolor!codicolor)
  End If
  Set rst = Nothing
  Set rstcolor = Nothing
End Sub
Sub posarcosesdescatalogat()
    If tintes.Recordset!descatalogat Then
        framedadestintes.BackColor = &HC0C0FF
        framedadestintes.caption = "Manteniment de la Tinta                ***********  DESCATALOGADA   *****************"
          Else
            framedadestintes.BackColor = Frame3.BackColor
            framedadestintes.caption = "Manteniment de la Tinta"
      End If
End Sub

Private Sub treurefiltre_Click()
   filtrarformules
   'dataformules.RecordSource = "select * from formules "
   dataformules.Refresh
   filtreformulacodi = ""
   filtreformuladesc = ""
   filtreformulaserie = ""
End Sub


Sub observacio_idtreball(numid As Integer, vabans As String, vdespres As String)
Dim rst As Recordset
  If numid = 0 Then Exit Sub
  Set rst = dbbaixes.OpenRecordset("select * from idstreball where id=" + atrim(numid))
  If rst.EOF Then rst.AddNew: rst!obsidtreball = " ": rst!id = numid: rst.Update
  Set rst = dbbaixes.OpenRecordset("select * from idstreball where id=" + atrim(numid))

  Load obsidtreball
  obsidtreball.obsid.Text = rst!obsidtreball
  vabans = rst!obsidtreball
  obsidtreball.Show 1
  If atrim(r) <> "" Then
     rst.Edit
     rst!obsidtreball = r
     rst.Update
     vdespres = r
       Else: rst.Delete
  End If
  

End Sub


Sub reixa_deltes()
  Set dbtmpb = dbbaixes
  Unload formcanvisanilox
  Load formcanvisanilox
  formcanvisanilox.FramedeltaE.Top = 0
  formcanvisanilox.Height = formcanvisanilox.FramedeltaE.Height + 100
  formcanvisanilox.tag = "nomesdelta "
  formcanvisanilox.tag = formcanvisanilox.tag + reixacomandesseguents
  formcanvisanilox.Show
  If formtintes.Top + formtintes.Height + 100 + formcanvisanilox.Height < Screen.Height Then
        formcanvisanilox.Top = formtintes.Top + formtintes.Height + 100
        formcanvisanilox.Left = formtintes.Left
         Else
           formcanvisanilox.Left = formtintes.Left + formtintes.width
           formcanvisanilox.Top = formtintes.Top + (formtintes.Height / 2)
  End If
  
End Sub
