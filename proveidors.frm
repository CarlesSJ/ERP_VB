VERSION 5.00
Begin VB.Form proveidors 
   Caption         =   "Manteniment de proveïdors"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   Icon            =   "proveidors.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   360
      Left            =   5985
      Picture         =   "proveidors.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   64
      TabStop         =   0   'False
      ToolTipText     =   "Missatges peu comanda"
      Top             =   900
      Width           =   420
   End
   Begin VB.Frame fdades 
      Caption         =   "Dades Bàsiques"
      Enabled         =   0   'False
      Height          =   4335
      Left            =   60
      TabIndex        =   19
      Top             =   720
      Width           =   6420
      Begin VB.ComboBox codicomptable 
         BackColor       =   &H00FFC0C0&
         DataField       =   "codicomptable"
         DataSource      =   "proveidors"
         Height          =   315
         Left            =   3420
         TabIndex        =   77
         Top             =   195
         Width           =   2325
      End
      Begin VB.Frame framemsg 
         Caption         =   "Missatges peu de comandes de compra"
         Height          =   3825
         Left            =   0
         TabIndex        =   33
         Top             =   15
         Visible         =   0   'False
         Width           =   5745
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Manteniment Peu Comandes"
            Height          =   240
            Left            =   3180
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   120
            Width           =   2520
         End
         Begin VB.CommandButton borrar 
            Height          =   300
            Index           =   9
            Left            =   600
            Picture         =   "proveidors.frx":0B14
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Eliminar peu."
            Top             =   3450
            Width           =   270
         End
         Begin VB.CommandButton borrar 
            Height          =   300
            Index           =   8
            Left            =   600
            Picture         =   "proveidors.frx":109E
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Eliminar peu."
            Top             =   3108
            Width           =   270
         End
         Begin VB.CommandButton borrar 
            Height          =   300
            Index           =   7
            Left            =   600
            Picture         =   "proveidors.frx":1628
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Eliminar peu."
            Top             =   2760
            Width           =   270
         End
         Begin VB.CommandButton borrar 
            Height          =   300
            Index           =   6
            Left            =   600
            Picture         =   "proveidors.frx":1BB2
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Eliminar peu."
            Top             =   2415
            Width           =   270
         End
         Begin VB.CommandButton borrar 
            Height          =   300
            Index           =   5
            Left            =   600
            Picture         =   "proveidors.frx":213C
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Eliminar peu."
            Top             =   2085
            Width           =   270
         End
         Begin VB.CommandButton borrar 
            Height          =   300
            Index           =   4
            Left            =   600
            Picture         =   "proveidors.frx":26C6
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Eliminar peu."
            Top             =   1740
            Width           =   270
         End
         Begin VB.CommandButton borrar 
            Height          =   300
            Index           =   3
            Left            =   600
            Picture         =   "proveidors.frx":2C50
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Eliminar peu."
            Top             =   1395
            Width           =   270
         End
         Begin VB.CommandButton borrar 
            Height          =   300
            Index           =   2
            Left            =   600
            Picture         =   "proveidors.frx":31DA
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Eliminar peu."
            Top             =   1065
            Width           =   270
         End
         Begin VB.CommandButton borrar 
            Height          =   300
            Index           =   1
            Left            =   600
            Picture         =   "proveidors.frx":3764
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Eliminar peu."
            Top             =   720
            Width           =   270
         End
         Begin VB.CommandButton borrar 
            Height          =   300
            Index           =   0
            Left            =   600
            Picture         =   "proveidors.frx":3CEE
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Eliminar peu."
            Top             =   375
            Width           =   270
         End
         Begin VB.CommandButton selmsg 
            Height          =   315
            Index           =   9
            Left            =   45
            Picture         =   "proveidors.frx":4278
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   3435
            Width           =   550
         End
         Begin VB.CommandButton selmsg 
            Height          =   315
            Index           =   8
            Left            =   45
            Picture         =   "proveidors.frx":4802
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   3088
            Width           =   550
         End
         Begin VB.CommandButton selmsg 
            Height          =   315
            Index           =   7
            Left            =   45
            Picture         =   "proveidors.frx":4D8C
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   2747
            Width           =   550
         End
         Begin VB.CommandButton selmsg 
            Height          =   315
            Index           =   6
            Left            =   45
            Picture         =   "proveidors.frx":5316
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   2406
            Width           =   550
         End
         Begin VB.CommandButton selmsg 
            Height          =   315
            Index           =   5
            Left            =   45
            Picture         =   "proveidors.frx":58A0
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   2065
            Width           =   550
         End
         Begin VB.CommandButton selmsg 
            Height          =   315
            Index           =   4
            Left            =   45
            Picture         =   "proveidors.frx":5E2A
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   1724
            Width           =   550
         End
         Begin VB.CommandButton selmsg 
            Height          =   315
            Index           =   3
            Left            =   45
            Picture         =   "proveidors.frx":63B4
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1383
            Width           =   550
         End
         Begin VB.CommandButton selmsg 
            Height          =   315
            Index           =   2
            Left            =   45
            Picture         =   "proveidors.frx":693E
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   1042
            Width           =   550
         End
         Begin VB.CommandButton selmsg 
            Height          =   315
            Index           =   1
            Left            =   45
            Picture         =   "proveidors.frx":6EC8
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   701
            Width           =   550
         End
         Begin VB.CommandButton selmsg 
            Height          =   315
            Index           =   0
            Left            =   45
            Picture         =   "proveidors.frx":7452
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   360
            Width           =   550
         End
         Begin VB.TextBox msg 
            DataField       =   "msg10"
            DataSource      =   "proveidors"
            Height          =   285
            Index           =   9
            Left            =   15
            TabIndex        =   43
            Top             =   3495
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox msg 
            DataField       =   "msg9"
            DataSource      =   "proveidors"
            Height          =   285
            Index           =   8
            Left            =   15
            TabIndex        =   42
            Top             =   3145
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox msg 
            DataField       =   "msg8"
            DataSource      =   "proveidors"
            Height          =   285
            Index           =   7
            Left            =   15
            TabIndex        =   41
            Top             =   2805
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox msg 
            DataField       =   "msg7"
            DataSource      =   "proveidors"
            Height          =   285
            Index           =   6
            Left            =   15
            TabIndex        =   40
            Top             =   2445
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox msg 
            DataField       =   "msg6"
            DataSource      =   "proveidors"
            Height          =   285
            Index           =   5
            Left            =   15
            TabIndex        =   39
            Top             =   2095
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox msg 
            DataField       =   "msg5"
            DataSource      =   "proveidors"
            Height          =   285
            Index           =   4
            Left            =   15
            TabIndex        =   38
            Top             =   1745
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox msg 
            DataField       =   "msg4"
            DataSource      =   "proveidors"
            Height          =   285
            Index           =   3
            Left            =   15
            TabIndex        =   37
            Top             =   1395
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox msg 
            DataField       =   "msg3"
            DataSource      =   "proveidors"
            Height          =   285
            Index           =   2
            Left            =   15
            TabIndex        =   36
            Top             =   1045
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox msg 
            DataField       =   "msg2"
            DataSource      =   "proveidors"
            Height          =   285
            Index           =   1
            Left            =   15
            TabIndex        =   35
            Top             =   690
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox msg 
            DataField       =   "msg1"
            DataSource      =   "proveidors"
            Height          =   285
            Index           =   0
            Left            =   15
            TabIndex        =   34
            Top             =   345
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label desc 
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "msg10"
            Height          =   240
            Index           =   9
            Left            =   900
            TabIndex        =   53
            Top             =   3480
            Width           =   4785
         End
         Begin VB.Label desc 
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "msg9"
            Height          =   240
            Index           =   8
            Left            =   900
            TabIndex        =   52
            Top             =   3135
            Width           =   4785
         End
         Begin VB.Label desc 
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "msg8"
            Height          =   240
            Index           =   7
            Left            =   900
            TabIndex        =   51
            Top             =   2805
            Width           =   4785
         End
         Begin VB.Label desc 
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "msg7"
            Height          =   240
            Index           =   6
            Left            =   900
            TabIndex        =   50
            Top             =   2460
            Width           =   4785
         End
         Begin VB.Label desc 
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "msg6"
            Height          =   240
            Index           =   5
            Left            =   900
            TabIndex        =   49
            Top             =   2115
            Width           =   4785
         End
         Begin VB.Label desc 
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "msg5"
            Height          =   240
            Index           =   4
            Left            =   900
            TabIndex        =   48
            Top             =   1785
            Width           =   4785
         End
         Begin VB.Label desc 
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "msg4"
            Height          =   240
            Index           =   3
            Left            =   900
            TabIndex        =   47
            Top             =   1440
            Width           =   4785
         End
         Begin VB.Label desc 
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "msg3"
            Height          =   240
            Index           =   2
            Left            =   900
            TabIndex        =   46
            Top             =   1095
            Width           =   4785
         End
         Begin VB.Label desc 
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "msg2"
            Height          =   240
            Index           =   1
            Left            =   900
            TabIndex        =   45
            Top             =   765
            Width           =   4785
         End
         Begin VB.Label desc 
            BackStyle       =   0  'Transparent
            Caption         =   "OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"
            DataField       =   "msg1"
            Height          =   240
            Index           =   0
            Left            =   900
            TabIndex        =   44
            Top             =   420
            Width           =   4785
         End
      End
      Begin VB.TextBox Text4 
         DataField       =   "descripciopagament"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   4155
         TabIndex        =   11
         Top             =   3165
         Width           =   1905
      End
      Begin VB.TextBox Text3 
         DataField       =   "formadepagament"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   1140
         TabIndex        =   10
         Top             =   3135
         Width           =   2085
      End
      Begin VB.TextBox Text2 
         DataField       =   "emailcomandes"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   3975
         MaxLength       =   255
         TabIndex        =   9
         Top             =   2625
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         DataField       =   "email"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   1140
         TabIndex        =   8
         Top             =   2655
         Width           =   1800
      End
      Begin VB.TextBox fax 
         DataField       =   "fax"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   3075
         TabIndex        =   7
         Top             =   2220
         Width           =   1545
      End
      Begin VB.TextBox tel 
         DataField       =   "tel"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   1140
         TabIndex        =   6
         Top             =   2235
         Width           =   1440
      End
      Begin VB.TextBox provincia 
         DataField       =   "provinciapais"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   1140
         TabIndex        =   5
         Top             =   1740
         Width           =   4725
      End
      Begin VB.TextBox poblacio 
         DataField       =   "poblacio"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   3405
         TabIndex        =   4
         Top             =   1320
         Width           =   2490
      End
      Begin VB.TextBox codipostal 
         DataField       =   "codipostal"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   1140
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox direccio 
         DataField       =   "direccio"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   1140
         TabIndex        =   2
         Top             =   930
         Width           =   4710
      End
      Begin VB.TextBox nom 
         DataField       =   "nom"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   1140
         TabIndex        =   1
         Top             =   555
         Width           =   4695
      End
      Begin VB.TextBox codi 
         BackColor       =   &H00C0C0C0&
         DataField       =   "codi"
         DataSource      =   "proveidors"
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   210
         Width           =   840
      End
      Begin VB.Label etmissatge 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1995
         TabIndex        =   78
         Top             =   15
         Width           =   75
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi comptable:"
         Height          =   300
         Left            =   2130
         TabIndex        =   75
         Top             =   225
         Width           =   1290
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Desc. Pag:"
         Height          =   300
         Left            =   3300
         TabIndex        =   32
         Top             =   3150
         Width           =   825
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Pag:"
         Height          =   300
         Left            =   180
         TabIndex        =   31
         Top             =   3180
         Width           =   825
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail Comandes:"
         Height          =   420
         Left            =   3090
         TabIndex        =   30
         Top             =   2595
         Width           =   825
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:"
         Height          =   300
         Left            =   180
         TabIndex        =   29
         Top             =   2670
         Width           =   825
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   300
         Left            =   2685
         TabIndex        =   28
         Top             =   2250
         Width           =   450
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Telèfon:"
         Height          =   300
         Left            =   180
         TabIndex        =   27
         Top             =   2220
         Width           =   825
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia / País:"
         Height          =   450
         Left            =   180
         TabIndex        =   26
         Top             =   1725
         Width           =   1080
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Població:"
         Height          =   300
         Left            =   2685
         TabIndex        =   25
         Top             =   1335
         Width           =   795
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi Postal:"
         Height          =   300
         Left            =   180
         TabIndex        =   24
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Direcció:"
         Height          =   300
         Left            =   180
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom:"
         Height          =   300
         Left            =   180
         TabIndex        =   22
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Codi:"
         Height          =   300
         Left            =   180
         TabIndex        =   20
         Top             =   255
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   6420
      Begin VB.CommandButton consultar 
         Height          =   450
         Left            =   5340
         Picture         =   "proveidors.frx":79DC
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda de Registres"
         Top             =   150
         Width           =   450
      End
      Begin VB.CommandButton sortir 
         Height          =   450
         Left            =   5820
         Picture         =   "proveidors.frx":7F66
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Sortir"
         Top             =   150
         Width           =   450
      End
      Begin VB.Data proveidors 
         Caption         =   "Proveidors"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   390
         Left            =   1950
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "proveidors"
         Top             =   210
         Width           =   3360
      End
      Begin VB.CommandButton alta 
         Height          =   360
         Left            =   75
         Picture         =   "proveidors.frx":84F0
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton eliminar 
         Height          =   360
         Left            =   1425
         Picture         =   "proveidors.frx":8A7A
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton gravar 
         Height          =   360
         Left            =   960
         Picture         =   "proveidors.frx":9004
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Guardar Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.CommandButton modificar 
         Height          =   360
         Left            =   525
         Picture         =   "proveidors.frx":958E
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Modificació Registres"
         Top             =   225
         Width           =   420
      End
      Begin VB.Label estattaula 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3810
         TabIndex        =   16
         Top             =   300
         Width           =   105
      End
   End
End
Attribute VB_Name = "proveidors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alta_Click()
alta_registre
framemsg.Visible = False
End Sub
Sub alta_registre()
 If proveidors.Recordset.EditMode = 0 Then
      fdades.Enabled = True
      proveidors.Recordset.AddNew
      DoEvents
        proveidors.Recordset!codiproduccio = cadbl(proveidorsproduccio.proveidorsp.Recordset!codi)
        'busco el mes gran i el poso a codi +1
        Set rsttmp = proveidors.Database.OpenRecordset("select max(codi) as [grancodi] from proveidors_comercial")
        If Not rsttmp.EOF Then
          codi.Text = atrim(cadbl(rsttmp!grancodi) + 1)
              Else: codi.Text = "1"
        End If
        
        codicomptable.SetFocus
     Else: MsgBox "No pots afegir si estàs editant...", vbCritical, "Atenció"
 End If
End Sub

Private Sub borrar_Click(Index As Integer)
  If MsgBox("Segur que vols eliminar aquest peu?", vbCritical + vbYesNo + vbDefaultButton2, "Atenció") = vbYes Then
    desc(Index).Caption = ""
    msg(Index).Text = 0
  End If
End Sub

Sub escullircodicomptable()
   Load formseleccio
  formseleccio.Caption = "Selecciona Codi Comptable"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codisap,nomproveidor,aliastintes from proveidors_codisSAP order by nomproveidor"
  formseleccio.refrescar
  formseleccio.Width = 9500
  formseleccio.Show 1
  If seleccioret = 1 Then
   codicomptable = atrim(formseleccio.Data1.Recordset!codisap)
  End If
  Unload formseleccio
End Sub

Private Sub codicomptable_DropDown()
 escullircodicomptable
End Sub

Private Sub codicomptable_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub codicomptable_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Command1_Click()
  If proveidors.Recordset.EditMode = 2 Then MsgBox "Estas afegint un proveïdor nou... Primer guarda els canvis i despres edita ": Exit Sub
  actualitzamsgs
  framemsg.Visible = Not framemsg.Visible
End Sub

Private Sub Command2_Click()
 missatgespeucomandescompra.Show 1
End Sub

Private Sub consultar_Click()
   Dim b As String
   framemsg.Visible = False
   b = InputBox("Entra la Descripcio del proveidor a buscar o el Codi", "Busqueda")
   If cadbl(b) > 0 Then
     proveidors.RecordSource = "select * from proveidors_comercial where codi=" + atrim(cadbl(b)) + ""
     proveidors.Refresh
     b = ""
      Else
       If b <> "" Then
         'If atrim(b) = "*" Then b = ""
        proveidors.RecordSource = "select * from proveidors_comercial where nom like '*" + b + "*'"
        proveidors.Refresh
       End If
   End If
   If Not proveidors.Recordset.EOF Then proveidors.Recordset.MoveLast: proveidors.Recordset.MoveFirst
End Sub

Private Sub eliminar_Click()
  eliminarproveidor
  framemsg.Visible = False
End Sub
Sub eliminarproveidor()
 On Error GoTo err
  If UCase(InputBox("Segur que vols Eliminar aquest proveidor?" + Chr(10) + Chr(13) + " escriu [ELIMINAR] per confirmar-ho.", "Eliminar proveidor")) = "ELIMINAR" Then
    proveidors.Recordset.Delete
    proveidors.Recordset.MoveNext
    If proveidors.Recordset.EOF Then proveidors.Recordset.MovePrevious
  End If
 Exit Sub
err:
  MsgBox "No s'ha pogut eliminar possiblement perque tingui registres relacionats. O bé no hi ha res per eliminar."
End Sub

Private Sub Form_Activate()
   Dim camp As Control
   Dim nomdelcamp As String
   For Each camp In Me
      nomdelcamp = mirarnomdelcamp(camp)
      If nomdelcamp <> "" Then
        If proveidors.Recordset.Fields(nomdelcamp).Type = 10 And nomdelcamp <> "codicomptable" Then
            camp.MaxLength = proveidors.Recordset.Fields(nomdelcamp).Size
        End If
      End If
   Next
End Sub
Function mirarnomdelcamp(camp As Control) As String
   On Error Resume Next
   mirarnomdelcamp = camp.DataField
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then cancelar_registre
  If KeyCode = 112 Then gravarregistre
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
 Sub cancelar_registre()
   If proveidors.Recordset.EditMode > 0 Then
       proveidors.Recordset.CancelUpdate
       fdades.Enabled = False
   End If
 End Sub
Private Sub Form_Load()
  proveidors.DatabaseName = cami
  proveidors.RecordSource = "select * from proveidors_comercial where codiproduccio=" + atrim(cadbl(proveidorsproduccio.proveidorsp.Tag)) + " order by codi"
  proveidors.Refresh
  Me.Tag = "MANTENIMENT DEL PROVEIDOR " + atrim(proveidorsproduccio.proveidorsp.Recordset!nom)
End Sub

Private Sub Frame2_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub gravar_Click()
  gravarregistre
  framemsg.Visible = False
End Sub
Sub gravarregistre()
If proveidors.EditMode > 0 Then
    proveidors.Recordset.Update
  End If
  fdades.Enabled = False
End Sub
Private Sub modificar_Click()
  If proveidors.Recordset.EditMode = 0 Then
     proveidors.Recordset.Edit
     fdades.Enabled = True
     DoEvents
     codicomptable.SetFocus
  End If
  framemsg.Visible = False
End Sub

Private Sub proveidors_Reposition()
  If Not proveidors.Recordset.EOF Then
   proveidors.Caption = "Proveidors:  " + atrim(cadbl(proveidors.Recordset.AbsolutePosition) + 1) + " de " + atrim(proveidors.Recordset.RecordCount)
   actualitzamsgs
   comprovarsiessap
     Else: proveidors.Caption = "Proveidors"
  End If
End Sub
Sub comprovarsiessap()
   If Not proveidors.Recordset.EOF Then
      If proveidors.Recordset!alta_desde_sap Then
          etmissatge = "Alta del SAP": etmissatge.Visible = True
            Else: etmissatge = "": etmissatge.Visible = False
      End If
   End If
End Sub
Sub actualitzamsgs()
   Dim rstm As Recordset
   Dim dbm As Database
   Set dbm = OpenDatabase(rutadelfitxer(cami) + "compres.mdb")
   For i = 0 To 9
      desc(i) = ""
     If cadbl(msg(i)) > 0 Then
         Set rstm = dbm.OpenRecordset("select descripcio from msgpeucomanda where id=" + atrim(cadbl(msg(i))))
         If Not rstm.EOF Then
             desc(i) = atrim(rstm!descripcio)
         End If
         
     End If
   Next i
   Set rstm = Nothing
   dbm.Close
   Set dbm = Nothing
End Sub

Private Sub selmsg_Click(Index As Integer)
   Load formseleccio
  formseleccio.Caption = "Selecciona Un Missatge"
  formseleccio.Data1.DatabaseName = rutadelfitxer(cami) + "\compres.mdb"
  formseleccio.Data1.RecordSource = "select * from msgpeucomanda"
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 1
  formseleccio.Show 1
  If seleccioret = 1 Then
   desc(Index).Caption = atrim(formseleccio.Data1.Recordset!descripcio)
   msg(Index).Text = cadbl(formseleccio.Data1.Recordset!ID)
  End If
  Unload formseleccio
End Sub

Private Sub sortir_Click()
  Unload Me
End Sub

