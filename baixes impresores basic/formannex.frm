VERSION 5.00
Begin VB.Form formannex 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   11760
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11760
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox cobservacions 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   113
      Top             =   11220
      Width           =   7050
   End
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   6720
      Top             =   1155
   End
   Begin VB.Shape colorstinters1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   255
      Index           =   0
      Left            =   3555
      Shape           =   2  'Oval
      Top             =   6045
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label etpostitcolors 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   180
      TabIndex        =   118
      Top             =   8130
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.Shape colorstinters3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   255
      Index           =   7
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape colorstinters3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   255
      Index           =   6
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape colorstinters3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   255
      Index           =   5
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape colorstinters3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   255
      Index           =   4
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape colorstinters3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   255
      Index           =   3
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape colorstinters3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   255
      Index           =   2
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape colorstinters3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   255
      Index           =   1
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape colorstinters4 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   250
      Index           =   7
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Shape colorstinters4 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   250
      Index           =   6
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Shape colorstinters4 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   250
      Index           =   5
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Shape colorstinters4 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   250
      Index           =   4
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Shape colorstinters4 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   250
      Index           =   3
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Shape colorstinters4 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   250
      Index           =   2
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Shape colorstinters4 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   250
      Index           =   1
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Shape colorstinters3 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   255
      Index           =   0
      Left            =   6645
      Shape           =   2  'Oval
      Top             =   9450
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape colorstinters2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   195
      Index           =   7
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape colorstinters2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   195
      Index           =   6
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape colorstinters2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   195
      Index           =   5
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape colorstinters2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   195
      Index           =   4
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape colorstinters2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   195
      Index           =   3
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape colorstinters2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   195
      Index           =   2
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape colorstinters2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   195
      Index           =   1
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape colorstinters1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   200
      Index           =   7
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.Shape colorstinters1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   200
      Index           =   6
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.Shape colorstinters1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   200
      Index           =   5
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.Shape colorstinters1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   200
      Index           =   4
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.Shape colorstinters1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   200
      Index           =   3
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.Shape colorstinters1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   200
      Index           =   2
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.Shape colorstinters1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   200
      Index           =   1
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.Shape colorstinters4 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   250
      Index           =   0
      Left            =   6600
      Shape           =   2  'Oval
      Top             =   9795
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Shape colorstinters2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      Height          =   250
      Index           =   0
      Left            =   6825
      Shape           =   2  'Oval
      Top             =   9075
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.Label etdataentrega 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2550
      TabIndex        =   117
      Top             =   1740
      Width           =   4620
   End
   Begin VB.Label etadhesiureprint 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   5190
      TabIndex        =   116
      Top             =   2355
      Width           =   1935
   End
   Begin VB.Label ettolerancia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00ED823A&
      Height          =   195
      Left            =   4275
      TabIndex        =   115
      Top             =   3285
      Width           =   2820
   End
   Begin VB.Image ImagePDF 
      Height          =   2925
      Left            =   45
      Stretch         =   -1  'True
      Top             =   8820
      Width           =   6330
   End
   Begin VB.Image logodigimarc 
      Height          =   390
      Left            =   6465
      Picture         =   "formannex.frx":0000
      Stretch         =   -1  'True
      Top             =   2970
      Width           =   660
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Observacions Disseny/Comanda"
      Height          =   285
      Left            =   75
      TabIndex        =   114
      Top             =   10935
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label etfoam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   5355
      TabIndex        =   112
      Top             =   8445
      Width           =   540
   End
   Begin VB.Label etanilox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "420"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   5940
      TabIndex        =   111
      Top             =   8445
      Width           =   330
   End
   Begin VB.Label etvolum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3.4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   6405
      TabIndex        =   110
      Top             =   8445
      Width           =   330
   End
   Begin VB.Label etliniatura 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   6840
      TabIndex        =   109
      Top             =   8445
      Width           =   330
   End
   Begin VB.Label etfoam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   5355
      TabIndex        =   108
      Top             =   8085
      Width           =   540
   End
   Begin VB.Label etanilox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "420"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   5940
      TabIndex        =   107
      Top             =   8085
      Width           =   330
   End
   Begin VB.Label etvolum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3.4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   6405
      TabIndex        =   106
      Top             =   8085
      Width           =   330
   End
   Begin VB.Label etliniatura 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   6840
      TabIndex        =   105
      Top             =   8085
      Width           =   330
   End
   Begin VB.Label etfoam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   5355
      TabIndex        =   104
      Top             =   7770
      Width           =   540
   End
   Begin VB.Label etanilox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "420"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   5940
      TabIndex        =   103
      Top             =   7770
      Width           =   330
   End
   Begin VB.Label etvolum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3.4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   6405
      TabIndex        =   102
      Top             =   7770
      Width           =   330
   End
   Begin VB.Label etliniatura 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   6840
      TabIndex        =   101
      Top             =   7770
      Width           =   330
   End
   Begin VB.Label etfoam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   5355
      TabIndex        =   100
      Top             =   7410
      Width           =   540
   End
   Begin VB.Label etanilox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "420"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   5940
      TabIndex        =   99
      Top             =   7410
      Width           =   330
   End
   Begin VB.Label etvolum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3.4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   6405
      TabIndex        =   98
      Top             =   7410
      Width           =   330
   End
   Begin VB.Label etliniatura 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   6840
      TabIndex        =   97
      Top             =   7410
      Width           =   330
   End
   Begin VB.Label etfoam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   5355
      TabIndex        =   96
      Top             =   7065
      Width           =   540
   End
   Begin VB.Label etanilox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "420"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   5940
      TabIndex        =   95
      Top             =   7065
      Width           =   330
   End
   Begin VB.Label etvolum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3.4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   6405
      TabIndex        =   94
      Top             =   7065
      Width           =   330
   End
   Begin VB.Label etliniatura 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   6840
      TabIndex        =   93
      Top             =   7065
      Width           =   330
   End
   Begin VB.Label etfoam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   5355
      TabIndex        =   92
      Top             =   6705
      Width           =   540
   End
   Begin VB.Label etanilox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "420"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   5940
      TabIndex        =   91
      Top             =   6705
      Width           =   330
   End
   Begin VB.Label etvolum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3.4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   6405
      TabIndex        =   90
      Top             =   6705
      Width           =   330
   End
   Begin VB.Label etliniatura 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   6840
      TabIndex        =   89
      Top             =   6705
      Width           =   330
   End
   Begin VB.Label etfoam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   5355
      TabIndex        =   88
      Top             =   6360
      Width           =   540
   End
   Begin VB.Label etanilox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "420"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   5940
      TabIndex        =   87
      Top             =   6360
      Width           =   330
   End
   Begin VB.Label etvolum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3.4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   6405
      TabIndex        =   86
      Top             =   6360
      Width           =   330
   End
   Begin VB.Label etliniatura 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   6840
      TabIndex        =   85
      Top             =   6360
      Width           =   330
   End
   Begin VB.Label etcomparteix 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NT.5020 XL-35"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   4230
      TabIndex        =   84
      Top             =   8445
      Width           =   1125
   End
   Begin VB.Label etcomparteix 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NT.5020 XL-35"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   4230
      TabIndex        =   83
      Top             =   8115
      Width           =   1125
   End
   Begin VB.Label etcomparteix 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NT.5020 XL-35"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   4230
      TabIndex        =   82
      Top             =   7755
      Width           =   1125
   End
   Begin VB.Label etcomparteix 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NT.5020 XL-35"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   4230
      TabIndex        =   81
      Top             =   7410
      Width           =   1125
   End
   Begin VB.Label etcomparteix 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NT.5020 XL-35"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   4230
      TabIndex        =   80
      Top             =   7065
      Width           =   1125
   End
   Begin VB.Label etcomparteix 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NT.5020 XL-35"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   4230
      TabIndex        =   79
      Top             =   6735
      Width           =   1125
   End
   Begin VB.Label etcomparteix 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NT.5020 XL-35"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4230
      TabIndex        =   78
      Top             =   6375
      Width           =   1125
   End
   Begin VB.Label etacarta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARTA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   7
      Left            =   3870
      TabIndex        =   77
      Top             =   8490
      Width           =   495
   End
   Begin VB.Label etacarta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARTA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   6
      Left            =   3870
      TabIndex        =   76
      Top             =   8145
      Width           =   495
   End
   Begin VB.Label etacarta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARTA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   5
      Left            =   3855
      TabIndex        =   75
      Top             =   7785
      Width           =   495
   End
   Begin VB.Label etacarta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARTA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   4
      Left            =   3855
      TabIndex        =   74
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label etacarta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARTA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   3
      Left            =   3855
      TabIndex        =   73
      Top             =   7095
      Width           =   495
   End
   Begin VB.Label etacarta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARTA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   2
      Left            =   3855
      TabIndex        =   72
      Top             =   6750
      Width           =   495
   End
   Begin VB.Label etacarta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARTA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   1
      Left            =   3840
      TabIndex        =   71
      Top             =   6390
      Width           =   495
   End
   Begin VB.Label etnomtinta 
      BackStyle       =   0  'Transparent
      Caption         =   "NEGRO PRIMAR YUKOFLEX ST(TEXTES)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   225
      TabIndex        =   70
      Top             =   8445
      Width           =   5175
   End
   Begin VB.Label etnomtinta 
      BackStyle       =   0  'Transparent
      Caption         =   "NEGRO PRIMAR YUKOFLEX ST(TEXTES)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   225
      TabIndex        =   69
      Top             =   8100
      Width           =   5070
   End
   Begin VB.Label etnomtinta 
      BackStyle       =   0  'Transparent
      Caption         =   "NEGRO PRIMAR YUKOFLEX ST(TEXTES)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   225
      TabIndex        =   68
      Top             =   7755
      Width           =   5160
   End
   Begin VB.Label etnomtinta 
      BackStyle       =   0  'Transparent
      Caption         =   "NEGRO PRIMAR YUKOFLEX ST(TEXTES)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   225
      TabIndex        =   67
      Top             =   7410
      Width           =   5130
   End
   Begin VB.Label etnomtinta 
      BackStyle       =   0  'Transparent
      Caption         =   "NEGRO PRIMAR YUKOFLEX ST(TEXTES)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   225
      TabIndex        =   66
      Top             =   7065
      Width           =   5205
   End
   Begin VB.Label etnomtinta 
      BackStyle       =   0  'Transparent
      Caption         =   "NEGRO PRIMAR YUKOFLEX ST(TEXTES)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   225
      TabIndex        =   65
      Top             =   6720
      Width           =   5130
   End
   Begin VB.Label etnomtinta 
      BackStyle       =   0  'Transparent
      Caption         =   "NEGRO PRIMAR YUKOFLEX ST(TEXTES)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   225
      TabIndex        =   64
      Top             =   6375
      Width           =   5145
   End
   Begin VB.Label etnomtinta 
      BackStyle       =   0  'Transparent
      Caption         =   "NEGRO PRIMAR YUKOFLEX ST(TEXTES)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   225
      TabIndex        =   54
      Top             =   6030
      Width           =   5160
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "8-"
      Height          =   225
      Left            =   45
      TabIndex        =   63
      Top             =   8475
      Width           =   210
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "7-"
      Height          =   225
      Left            =   45
      TabIndex        =   62
      Top             =   8130
      Width           =   210
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "6-"
      Height          =   225
      Left            =   45
      TabIndex        =   61
      Top             =   7785
      Width           =   210
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "5-"
      Height          =   225
      Left            =   45
      TabIndex        =   60
      Top             =   7440
      Width           =   210
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "4-"
      Height          =   225
      Left            =   45
      TabIndex        =   59
      Top             =   7095
      Width           =   210
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "3-"
      Height          =   225
      Left            =   45
      TabIndex        =   58
      Top             =   6750
      Width           =   210
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "2-"
      Height          =   225
      Left            =   45
      TabIndex        =   57
      Top             =   6405
      Width           =   210
   End
   Begin VB.Label etacarta 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARTA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   0
      Left            =   3840
      TabIndex        =   56
      Top             =   6045
      Width           =   495
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "1-"
      Height          =   225
      Left            =   45
      TabIndex        =   55
      Top             =   6060
      Width           =   210
   End
   Begin VB.Shape Shape22 
      Height          =   360
      Left            =   15
      Top             =   8055
      Width           =   7185
   End
   Begin VB.Shape Shape21 
      Height          =   360
      Left            =   15
      Top             =   8400
      Width           =   7185
   End
   Begin VB.Shape Shape20 
      Height          =   360
      Left            =   15
      Top             =   7710
      Width           =   7185
   End
   Begin VB.Shape Shape19 
      Height          =   360
      Left            =   15
      Top             =   7365
      Width           =   7185
   End
   Begin VB.Shape Shape18 
      Height          =   360
      Left            =   15
      Top             =   6675
      Width           =   7185
   End
   Begin VB.Shape Shape17 
      Height          =   360
      Left            =   15
      Top             =   7020
      Width           =   7185
   End
   Begin VB.Shape Shape16 
      Height          =   360
      Left            =   15
      Top             =   6330
      Width           =   7185
   End
   Begin VB.Shape Shape15 
      Height          =   360
      Left            =   15
      Top             =   5985
      Width           =   7185
   End
   Begin VB.Label etcomparteix 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NT.5020 XL-35"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4230
      TabIndex        =   53
      Top             =   6030
      Width           =   1125
   End
   Begin VB.Label etliniatura 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "53"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6840
      TabIndex        =   52
      Top             =   6000
      Width           =   330
   End
   Begin VB.Label etvolum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3.4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6405
      TabIndex        =   51
      Top             =   6000
      Width           =   330
   End
   Begin VB.Label etanilox 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "420"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5940
      TabIndex        =   50
      Top             =   6000
      Width           =   330
   End
   Begin VB.Label etfoam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5355
      TabIndex        =   49
      Top             =   6000
      Width           =   540
   End
   Begin VB.Shape Shape14 
      Height          =   3105
      Left            =   6810
      Top             =   5655
      Width           =   390
   End
   Begin VB.Shape Shape13 
      Height          =   3105
      Left            =   6345
      Top             =   5655
      Width           =   480
   End
   Begin VB.Shape Shape12 
      Height          =   3105
      Left            =   5865
      Top             =   5655
      Width           =   495
   End
   Begin VB.Shape Shape11 
      Height          =   2805
      Left            =   5355
      Top             =   5955
      Width           =   525
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "F     A     V     L"
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
      Height          =   360
      Left            =   5595
      TabIndex        =   48
      Top             =   5655
      Width           =   1620
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPARTEIX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4260
      TabIndex        =   47
      Top             =   5670
      Width           =   1275
   End
   Begin VB.Label etsistemaimpresio 
      Caption         =   "TRANSPARÈNCIA LAMINAT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1530
      TabIndex        =   46
      Top             =   5670
      Width           =   2700
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   15
      Top             =   5610
      Width           =   7185
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "IMPRESIÓ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   150
      TabIndex        =   45
      Top             =   5655
      Width           =   1755
   End
   Begin VB.Label etquantitat 
      BackStyle       =   0  'Transparent
      Caption         =   "45.500 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   825
      TabIndex        =   44
      Top             =   1995
      Width           =   3960
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Metres:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   43
      Top             =   2145
      Width           =   795
   End
   Begin VB.Label etamplelam 
      BackStyle       =   0  'Transparent
      Caption         =   "84 cm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5715
      TabIndex        =   42
      Top             =   5235
      Width           =   735
   End
   Begin VB.Label etsimulteneitat 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3855
      TabIndex        =   41
      Top             =   5250
      Width           =   240
   End
   Begin VB.Label etrefilat 
      BackStyle       =   0  'Transparent
      Caption         =   "42 cm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   885
      TabIndex        =   40
      Top             =   5220
      Width           =   1350
   End
   Begin VB.Label etmateriallam 
      BackStyle       =   0  'Transparent
      Caption         =   "PET - LACADO QUIM. NO ESTER. - TRANSPARENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1095
      TabIndex        =   39
      Top             =   4950
      Width           =   5550
   End
   Begin VB.Label ettractat 
      BackStyle       =   0  'Transparent
      Caption         =   "1 CARA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5445
      TabIndex        =   38
      Top             =   4260
      Width           =   1350
   End
   Begin VB.Label etespesormaterial 
      BackStyle       =   0  'Transparent
      Caption         =   "20 mc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3345
      TabIndex        =   37
      Top             =   4260
      Width           =   1350
   End
   Begin VB.Label etample 
      BackStyle       =   0  'Transparent
      Caption         =   "87 cm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   810
      TabIndex        =   36
      Top             =   4245
      Width           =   1350
   End
   Begin VB.Label etmaterial 
      BackStyle       =   0  'Transparent
      Caption         =   "PET - LACADO QUIM. NO ESTER. - TRANSPARENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   35
      Top             =   3960
      Width           =   5550
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "AMPLE LAM:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4560
      TabIndex        =   34
      Top             =   5265
      Width           =   1185
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "SIMULTENEITAT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2415
      TabIndex        =   33
      Top             =   5265
      Width           =   1500
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "TRACTAT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4485
      TabIndex        =   32
      Top             =   4290
      Width           =   960
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "ESP:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2895
      TabIndex        =   31
      Top             =   4290
      Width           =   555
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "REFILAT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   30
      Top             =   5250
      Width           =   1065
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "AMPLE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   29
      Top             =   4275
      Width           =   1065
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "MATERIAL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   28
      Top             =   4965
      Width           =   1065
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "MATERIAL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   27
      Top             =   3990
      Width           =   1065
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LAMINADORA / REBOBINADORA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1860
      TabIndex        =   26
      Top             =   4620
      Width           =   4050
   End
   Begin VB.Shape Shape7 
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   15
      Top             =   4575
      Width           =   7170
   End
   Begin VB.Shape Shape6 
      Height          =   990
      Left            =   15
      Top             =   4575
      Width           =   7185
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "MATERIAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3000
      TabIndex        =   25
      Top             =   3615
      Width           =   1725
   End
   Begin VB.Shape Shape5 
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   15
      Top             =   3570
      Width           =   7170
   End
   Begin VB.Shape Shape4 
      Height          =   990
      Left            =   15
      Top             =   3570
      Width           =   7185
   End
   Begin VB.Label etcilindre 
      BackStyle       =   0  'Transparent
      Caption         =   "640 mm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3225
      TabIndex        =   24
      Top             =   3270
      Width           =   870
   End
   Begin VB.Label etdesarroll 
      BackStyle       =   0  'Transparent
      Caption         =   "320 mm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1800
      TabIndex        =   23
      Top             =   3255
      Width           =   1035
   End
   Begin VB.Label etespesorpolimer 
      BackStyle       =   0  'Transparent
      Caption         =   "1.14 mc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   495
      TabIndex        =   22
      Top             =   3240
      Width           =   900
   End
   Begin VB.Label ettipus 
      BackStyle       =   0  'Transparent
      Caption         =   "FLEXO STD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4650
      TabIndex        =   21
      Top             =   2955
      Width           =   1905
   End
   Begin VB.Label etproveidor 
      BackStyle       =   0  'Transparent
      Caption         =   "MILLER GRAPHICS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1140
      TabIndex        =   20
      Top             =   2940
      Width           =   2910
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "CIL:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2790
      TabIndex        =   19
      Top             =   3285
      Width           =   465
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "DES.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1335
      TabIndex        =   18
      Top             =   3285
      Width           =   420
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "ESP:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   17
      Top             =   3270
      Width           =   450
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TIPUS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   16
      Top             =   2985
      Width           =   705
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEÏDOR:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   15
      Top             =   2970
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "POLIMERS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2970
      TabIndex        =   14
      Top             =   2610
      Width           =   1725
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   15
      Top             =   2565
      Width           =   7185
   End
   Begin VB.Shape Shape2 
      Height          =   990
      Left            =   15
      Top             =   2565
      Width           =   7185
   End
   Begin VB.Label etreprint 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Reprint"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5160
      TabIndex        =   13
      Top             =   1995
      Width           =   2040
   End
   Begin VB.Label etcingular 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Cingular Real"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   30
      TabIndex        =   12
      Top             =   1485
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label etclient 
      BackStyle       =   0  'Transparent
      Caption         =   "DUJARDIN FOODS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2775
      TabIndex        =   11
      Top             =   75
      Width           =   3975
   End
   Begin VB.Label etlinia 
      BackStyle       =   0  'Transparent
      Caption         =   "FINDUS - WOK CHINESE 325G FINDUS - WOK CHINESE 325G FINDUS - WOK CHINESE 325G FINDUS - WOK CHINESE 325G FINDUS - WOK CHINESE 325G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   90
      TabIndex        =   10
      Top             =   1125
      Width           =   6555
   End
   Begin VB.Label etnt 
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   2250
      TabIndex        =   9
      Top             =   465
      Width           =   1095
   End
   Begin VB.Label etxl 
      BackStyle       =   0  'Transparent
      Caption         =   "XL-999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   4380
      TabIndex        =   8
      Top             =   450
      Width           =   2580
   End
   Begin VB.Label etcb 
      BackStyle       =   0  'Transparent
      Caption         =   "01234567891123"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   3840
      TabIndex        =   7
      Top             =   780
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CB:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   420
      Left            =   3315
      TabIndex        =   6
      Top             =   765
      Width           =   915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ARXIU:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3330
      TabIndex        =   5
      Top             =   435
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NT:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   420
      Left            =   1740
      TabIndex        =   4
      Top             =   435
      Width           =   555
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Texte:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   420
      Left            =   75
      TabIndex        =   3
      Top             =   750
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENT: "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   420
      Left            =   1560
      TabIndex        =   2
      Top             =   45
      Width           =   1245
   End
   Begin VB.Label ettipusimpresio 
      BackStyle       =   0  'Transparent
      Caption         =   "REPETIDA"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   75
      TabIndex        =   1
      Top             =   450
      Width           =   1665
   End
   Begin VB.Label etcomanda 
      BackStyle       =   0  'Transparent
      Caption         =   "000.000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1980
      Left            =   15
      Top             =   -15
      Width           =   7185
   End
   Begin VB.Shape Shape8 
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   15
      Top             =   5610
      Width           =   1440
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   345
      Left            =   1455
      Top             =   5625
      Width           =   5730
   End
End
Attribute VB_Name = "formannex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formX As Single
    Dim formY As Single

Dim rstcannex As Recordset

Dim rstc2annex As Recordset
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Sub cobservacions_GotFocus()
  cobservacions.Top = cobservacions.Top - (cobservacions.Height * 3)
  cobservacions.Height = cobservacions.Height * 4
End Sub

Private Sub cobservacions_LostFocus()
  cobservacions.Top = 11220
  cobservacions.Height = cobservacions.Height / 4
End Sub

Private Sub etcingular_Click()
  sonar_sirena "intermitent"
End Sub

Sub controlo_si_passoperRodonetesDeColors()
  Dim i As Long
  Dim j As Long
  Dim x As Long
  Dim y As Long
  Dim vControl As Variant
  
'  Form1.caption = atrim(formX) + " - " + atrim(formY) + "      " + colorstinters1(0).tag
  x = formX
  y = formY
  If y > (etnomtinta(0).Top) And y < (etnomtinta(7).Top + etnomtinta(7).Height) Then
   For j = 1 To 4
    Set vControl = Me.Controls("colorstinters" + atrim(j))
    For i = 0 To vControl.Count - 1
        If y > vControl(i).Top And y < (vControl(i).Top + vControl(i).Height) Then
             If x > vControl(i).Left And x < (vControl(i).Left + vControl(i).width) Then
                 If vControl(i).visible Then ensenyarPostitColors vControl, i
                 GoTo fi
             End If
        End If
    Next i
   Next j
  End If
  etpostitcolors.visible = False
fi:

End Sub

Private Sub etreprint_Click()
sonar_sirena "continuu"
End Sub
Sub possartoleranciaample(ntreball As Double, nmodificacio As Double, materialex As Double, vreduccio As String, vnumc As Double)
    Dim tanper100 As Byte
    Dim rst As Recordset
    Dim rstm As Recordset
    Dim msgobservacions As String
    tanper100 = 2
    If vreduccio <> "" Then ettolerancia.caption = "Tolerancia desarroll: Reducció cilindre": Exit Sub
    Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
    Set rst = dbclixes.OpenRecordset("select desarroll from modificacions where id_treball=" + atrim(ntreball) + " and ordre=" + atrim(nmodificacio))
    If rst.EOF Then Exit Sub
    Set rstm = dbtmp.OpenRecordset("SELECT comandes.comanda, familiesmaterials.descripcio FROM comandes INNER JOIN (familiesmaterials INNER JOIN materials ON familiesmaterials.codi = materials.familia) ON comandes.materialex = materials.codi WHERE (((comandes.comanda)=" + atrim(vnumc) + " ));")
    If Not rstm.EOF Then If Mid(atrim(rstm!descripcio), 1, 2) = "PE" And Mid(atrim(rstm!descripcio), 1, 3) <> "PET" Then tanper100 = 4
    ettolerancia.caption = "Tolerancia desarroll: " + atrim(cadbl(rst!desarroll) - tanper100) + " a " + atrim(cadbl(rst!desarroll) + tanper100) + " mm"
    Set rst = Nothing
    Set rstm = Nothing
End Sub



Private Sub Form_Activate()
refrescar_rstannex
End Sub

Private Sub Form_Load()
   borrarcamps
  ' Me.ScaleMode = vbPixels
End Sub
Sub refrescar_rstannex()
'   Me.caption = "Carregant taules..."
'   DoEvents
'   Set rstcannex = dbtmp.OpenRecordset("select * from comandesmesextres where comanda>170000", , ReadOnly)
'   Set rstc2annex = rstcannex.Clone 'dbtmp.OpenRecordset("select * from comandesmesextres", , ReadOnly)
'   Me.caption = ""
End Sub

Sub ensenyarPostitColors(vControl As Variant, vIndex As Long)
  etpostitcolors.caption = atrim(Mid(vControl(vIndex).tag, InStr(1, vControl(vIndex).tag + " ", " ")))
  etpostitcolors.Top = vControl(vIndex).Top
  etpostitcolors.Left = etnomtinta(1).Left + 20
  etpostitcolors.visible = True
End Sub
Private Sub Form_Resize()
'  Me.caption = atrim(Me.Height)

End Sub

Private Sub Timer1_Timer()
 'etacarta.visible = Not etacarta.visible
 If logodigimarc.tag = "1" Then
    logodigimarc.visible = Not logodigimarc.visible
      Else: logodigimarc.visible = False
 End If
 
    Dim p As POINTAPI
    
    ' Obtiene las coordenadas del ratón en la pantalla
    GetCursorPos p

    ' Convierte las coordenadas de pantalla a coordenadas del Form
    ScreenToClient Me.hwnd, p

    formX = p.x * Screen.TwipsPerPixelX
    formY = p.y * Screen.TwipsPerPixelY

controlo_si_passoperRodonetesDeColors

End Sub
Sub borrarcamps()
  Dim i As Byte
  etcomanda = ""
  etclient = ""
  etlinia = ""
  etnt = ""
  etxl = ""
  ettipusimpresio = ""
  etquantitat = ""
  etcb = ""
  etcingular = ""
  etreprint = ""
  etadhesiureprint = ""
  etproveidor = ""
  ettipus = ""
  etespesorpolimer = ""
  etdesarroll = ""
  etcilindre = ""
  etmaterial = ""
  etample = ""
  etespesormaterial = ""
  ettractat = ""
  etmateriallam = ""
  etrefilat = ""
  etsimulteneitat = ""
  etamplelam = ""
  cobservacions = ""
  logodigimarc.tag = ""
  For i = 0 To 7
    etnomtinta(i) = ""
    etacarta(i) = ""
    etcomparteix(i) = ""
    etfoam(i) = ""
    etfoam(i).ToolTipText = ""
    etanilox(i) = ""
    etvolum(i) = ""
    etliniatura(i) = ""
  Next i
  
End Sub
Sub carregarcomanda(vnumc As Double)
   Dim rstc As Recordset
   Dim rstcextra As Recordset
   Dim rstc2 As Recordset
   Dim rstm As Recordset
   Dim rstt As Recordset
   Dim rstcli As Recordset
   Dim rstmat As Recordset
   Dim vnomclient As String
   Dim vnomfitxer As String
   Dim rsttmp As Recordset
   Dim rstplan As Recordset
   
   'netejo els camps
   borrarcamps
   Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "\clixesnous.mdb")
   'busco les dades necessaries
   formannex.BackColor = &H8000000F
   Set rstc = dbtmpb.OpenRecordset("select comandafingerprintoriginal from impressorestot where comanda=" + atrim(vnumc))
   If Not rstc.EOF Then If cadbl(rstc!comandafingerprintoriginal) > 0 Then formannex.BackColor = &HFFFF&
   Set rstc = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(vnumc), , ReadOnly)
   Set rstcextra = dbtmp.OpenRecordset("select * from comandes_Extres where comanda=" + atrim(vnumc), , ReadOnly)
   Set rstc2 = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(rstc!linkcomanda1)), , ReadOnly)
   Set rstcli = dbtmp.OpenRecordset("select * from clients where codi=" + atrim(cadbl(rstc!client)))
   Set rsttmp = dbtmp.OpenRecordset("select marcailinia,materialex,cantitatex,mesuracantex,cmaquina,numtreball,numordremodificacio,codibarras,tubolam,espessor,mesuraesp,comanda,refclient,comandaclient,texteimpressio,direnvio,impressio from comandes where comanda=" + atrim(cadbl(vnumc)))
   If Not rstcli.EOF Then vnomclient = atrim(rstcli!nom)
   Set rstmat = dbtmp.OpenRecordset("select * from [llistat materials]")
   rstmat.FindFirst "codi=" + atrim(cadbl(rstc!materialex))
   Set rstm = dbclixes.OpenRecordset("SELECT Clixes.codidebarres, Clixes.linia, Clixes.marca, Clixes.arxiu,Clixes.ubicacio, Modificacions.*, Fotogravadors.nomfotogravador FROM (Clixes RIGHT JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) LEFT JOIN Fotogravadors ON Modificacions.fotograbador = Fotogravadors.codi where modificacions.id_treball=" + atrim(cadbl(rstc!numtreball)) + " and modificacions.ordre=" + atrim(cadbl(rstc!numordremodificacio)))
   If rstm.EOF Then GoTo fi
   Set rstt = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(rstm!id_treball) + " and ordremodificacio=" + atrim(rstm!ordre))
   If rstt.EOF Then GoTo fi
   Set rstplan = dbtmpb.OpenRecordset("Select * from planificaciototes where comanda=" + atrim(vnumc))
   etdataentrega = ""
   If Not rstplan.EOF Then etdataentrega = "Data entrega: " + atrim(format(rstplan!data2, "dd/mm/yy"))
   
   vnomfitxer = ruta_documentacio_clixes + "\" + format(rstm!id_treball, "00000") + "\pdf" + format(rstm!id_treball, "00000") + "-" + format(rstm!ordre, "000") + ".pdf"
   'posso els valors de les etiquetes
   etcomanda = atrim(vnumc)
  etclient = atrim(vnomclient)
  etlinia = atrim(rstm!marca) + "-" + atrim(rstm!linia)
  etnt = atrim(rstm!id_treball) + "/" + atrim(rstm!ordre)
  etxl = atrim(rstm!arxiu) + IIf(atrim(rstm!ubicacio) <> "", " [" + atrim(rstm!ubicacio) + "]", "")
  etsistemaimpresio = IIf(rstc!formaimp = "T", "TRANSPARENCIA", "NORMAL")
  ettipusimpresio = IIf(atrim(rstc!impressio) = "N", "NOVA", IIf(atrim(rstc!impressio) = "M", "MODIFICADA", IIf(atrim(rstc!impressio) = "R", "REPETIDA", "")))
  etquantitat = atrim(format(rstc!cantitatex, "#,##0")) + "Mts  " + atrim(format(rstc!rebkilos, "#,##0")) + "Kg"
  etcb = atrim(rstm!codidebarres)
  'etcingular = IIf(Form1.mirarsihihaCingularReal(rstm!id_treball, rstm!ordre), "Cingular Real", "")
  etreprint = IIf(IIf(Not IsNull(rstm!reimpres), rstm!reimpres, False), "REPRINT", "")
  etproveidor = atrim(rstm!nomfotogravador)
  ettipus = atrim(rstm!sistemadimpresio)
  etespesorpolimer = atrim(rstm!gruixpolimer) + "µ"
  etdesarroll = atrim(rstm!desarroll) + "mm"
  etcilindre = atrim(rstc!cilindres)
  etmaterial = IIf(rstmat.NoMatch, "", construirnommaterial(rstmat))
  etample = atrim(rstc!ampleesq) + "Cm"
  etespesormaterial = atrim(rstc!espessor) + "µ"
  ettractat = IIf(cadbl(rstc!tractatex) > 0, atrim(rstc!tractatex) + "Cara", "NO")
  If cadbl(rstc!linkcomanda1) > 0 Then
   If Not rstc2.EOF Then
    Label14 = "LAMINADORA"
    rstmat.FindFirst "codi=" + atrim(cadbl(rstc2!materialex))
    etmateriallam = IIf(rstmat.NoMatch, "", construirnommaterial(rstmat))
    etrefilat = atrim(rstc2!ampleutil) + "Cm"
    etsimulteneitat = atrim(rstc2!simulteneitatlam)
    etamplelam = atrim(cadbl(rstc2!ampleutil) * cadbl(rstc2!simulteneitatlam)) + "Cm"
   End If
     Else
       Label14 = "REBOBINADORA"
       etmateriallam = etmaterial
       etrefilat = atrim(rstc!amplereb) + "Cm"
       etsimulteneitat = atrim(rstc!simulteneitatreb)
       etamplelam = etample
  End If
  cobservacions = ""
  possartoleranciaample cadbl(rsttmp!numtreball), cadbl(rsttmp!numordremodificacio), rsttmp!materialex, atrim(rsttmp!cmaquina), vnumc
  If atrim(rstm!digimarc) = "Si" Then logodigimarc.tag = "1"
  possar_tintes cadbl(rstc!numtreball), cadbl(rstc!numordremodificacio), vnumc
  'ruta_documentacio_clixes = "\\ORD_COPIES\DOCUMENTACIOCLIXEssd"
  If existeix(ruta_documentacio_clixes) And existeix(vnomfitxer) Then
     carregar_MINIPDF vnomfitxer
      Else:
       ImagePDF = LoadPicture("")
       If Not existeix(ruta_documentacio_clixes) Then MsgBox "No es pot accedir a la ruta dels PDF. (ORD_COPIES)", vbCritical, "ERROR ACCEDINT AL SERVIDOR DE PDF"
  End If
fi:
   Set rstc = Nothing
   Set rstc2 = Nothing
   Set rstm = Nothing
   Set rstt = Nothing
   Set rstmat = Nothing
End Sub

Sub carregar_MINIPDF(vnomfitxer As String)
  Dim vfitxer As String
  Dim vdatapdf As String
  Set ImagePDF = LoadPicture("")
  Dim vnomfitxerini As String
  
  
  vfitxer = substituir(UCase(vnomfitxer), ".PDF", "_MINI.gif")
  vnomfitxerini = substituir(atrim(vfitxer), ".gif", ".ini")
  vdatapdf = llegir_ini("General", "datapdf", vnomfitxerini)
  If vdatapdf <> FileDateTime(vnomfitxer) Then If existeix(vfitxer) Then Kill vfitxer
  'If Not existeix(vnomfitxer) Then MsgBox "No trobo el PDF d'aquest treball." + vbNewLine + vnomfitxer, vbCritical, "Error": GoTo fi
  If Not existeix(vnomfitxer) Then GoTo fi
  If Not existeix(vfitxer) Then
     If existeix("c:\temp\pdfimpresio.gif") Then Kill "c:\temp\pdfimpresio.gif"
     If existeix("c:\temp\pdfimpresio.gif") Then GoTo fi
     If UCase(App.EXEName) <> "PLANIFICACIO" Then Form1.preparaelPDF vnomfitxer, 0, ""
     If existeix("c:\temp\pdfimpresio.gif") Then FileCopy "c:\temp\pdfimpresio.gif", vfitxer
     escriure_ini "General", "datapdf", FileDateTime(vnomfitxer), vnomfitxerini
     escriure_ini "General", "nompdf", vnomfitxer, vnomfitxerini
  End If
  If existeix(vfitxer) Then Set ImagePDF = LoadPicture(vfitxer)
fi:
End Sub
Sub possar_tintes(vid_treball As Double, vordre As Double, vnumc As Double)
  Dim i As Byte
  Dim rst As Recordset
  Dim rsttintes As Recordset
  Dim rstdatos As Recordset
  Dim rstmuntadora As Recordset
  Dim rstfoam As Recordset
  Dim vmsg As String
  Dim rsttintesobservacions As Recordset
  
  'netejar camps
   For i = 0 To 7
    etnomtinta(i) = ""
    etacarta(i) = ""
    etcomparteix(i) = ""
    etfoam(i) = ""
    etfoam(i).ToolTipText = ""
    etanilox(i) = ""
    etvolum(i) = ""
    etliniatura(i) = ""
    colorstinters1(i).tag = "4000"
    colorstinters1(i).DrawMode = 9: colorstinters1(i).width = 250: colorstinters1(i).Height = 250: colorstinters1(i).Top = etnomtinta(i).Top: colorstinters1(i).visible = False
    colorstinters2(i).DrawMode = 9: colorstinters2(i).width = 250: colorstinters2(i).Height = 250: colorstinters2(i).Top = etnomtinta(i).Top: colorstinters2(i).visible = False
    colorstinters3(i).DrawMode = 9: colorstinters3(i).width = 250: colorstinters3(i).Height = 250: colorstinters3(i).Top = etnomtinta(i).Top: colorstinters3(i).visible = False
    colorstinters4(i).DrawMode = 9: colorstinters4(i).width = 250: colorstinters4(i).Height = 250: colorstinters4(i).Top = etnomtinta(i).Top: colorstinters4(i).visible = False
  Next i
  Set rstfoam = dbtmp.OpenRecordset("select * from adhesiusmuntadora", , ReadOnly)
  Set rsttintes = dbtintes.OpenRecordset("select * from tintes order by codi", , dbReadOnly)
  Set rstmuntadora = dbtmpb.OpenRecordset("select * from muntadorescilindres where numcomanda=" + atrim(vnumc), , ReadOnly)
  Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(vid_treball) + " and ordremodificacio=" + atrim(vordre) + " order by ordretinter ASC")
  While Not rst.EOF
    Set rstdatos = rst
    If cadbl(rst!tinterlinkambid_treball) > 0 Then
       Set rstdatos = dbclixes.OpenRecordset("SELECT Tintes.*, Clixes.arxiu FROM Clixes RIGHT JOIN Tintes ON Clixes.id_treball = Tintes.id_treball where id_tinter=" + atrim(rst!tinterlinkambid_treball))
      etcomparteix(rst!ordretinter - 1) = "NT." + atrim(rstdatos!id_treball) + " " + atrim(rstdatos!arxiu)
    End If
    rsttintes.FindFirst "codi='" + atrim(rstdatos!coditinta) + "'"
    If Not rsttintes.NoMatch Then etnomtinta(rst!ordretinter - 1) = atrim(rsttintes!descripcio)
    etnomtinta(rst!ordretinter - 1) = etnomtinta(rst!ordretinter - 1) + " " + atrim(rstdatos!detalltinter) + posar_extensio(rstdatos)
    etanilox(rst!ordretinter - 1) = atrim(rstdatos!anilox)
    etvolum(rst!ordretinter - 1) = atrim(rstdatos!volum)
    etliniatura(rst!ordretinter - 1) = atrim(rstdatos!aniloxclixe)
    
    If Not rstmuntadora.EOF Then
       rstmuntadora.FindFirst "numcilindre=" + atrim(rst!ordretinter)
       If Not rstmuntadora.NoMatch Then
            rstfoam.FindFirst "codiintern='" + atrim(cadbl(rstmuntadora!idadhesiu)) + "'"
            If Not rstfoam.NoMatch Then
              etfoam(rst!ordretinter - 1) = atrim(rstfoam!inicialsfoam)
              etfoam(rst!ordretinter - 1).ToolTipText = atrim(rstfoam!descripcioinplacsa)
            End If
       End If
    End If
    posarColorsTinters rst!ordretinter - 1, rstdatos
    rst.MoveNext
  Wend
  'Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(vid_treball) + " and ordremodificacio=" + atrim(vordre * -1) + " and color like 'VERNIS*' order by ordretinter ASC")
  Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(vid_treball) + " and ordremodificacio=" + atrim(vordre * -1) + " and color<>'' order by color DesC")
  If Not rst.EOF Then
       etadhesiureprint = atrim(rst!color)
        Else
          Set rst = dbclixes.OpenRecordset("select * from tintes where id_treball=" + atrim(vid_treball) + " and ordremodificacio=" + atrim(vordre * -1) + " order by tinterlinkambid_treball desc")
          If Not rst.EOF Then
            Set rst = dbclixes.OpenRecordset("select * from tintes where id_tinter=" + atrim(rst!tinterlinkambid_treball))
            If Not rst.EOF Then etadhesiureprint = atrim(rst!color)
          End If
  End If
  'posso les observacions de tintes
  vmsg = ""
  Set rsttintesobservacions = dbclixes.OpenRecordset("select * from tintes_observacions where id_treball=" + atrim(vid_treball) + " and ordre=" + atrim(vordre) + " order by id asc")
  While Not rsttintesobservacions.EOF
     If atrim(rsttintesobservacions!observacio) <> "" Then vmsg = vmsg + atrim(rsttintesobservacions!observacio) + Chr(13) + Chr(10)
     rsttintesobservacions.MoveNext
  Wend
  If vmsg <> "" Then
     cobservacions.BackColor = QBColor(10)
     cobservacions = vmsg
     cobservacions.visible = True
      Else
        cobservacions.BackColor = QBColor(15)
        cobservacons = ""
        cobservacions.visible = False
  End If
  
  Set rsttintes = Nothing
  Set rstdatos = Nothing
  Set rstmuntadora = Nothing
  Set rstfoam = Nothing
  Set rst = Nothing
End Sub
Sub posarColorsTinters(vIndex As Long, rst As Recordset)
  Dim vcolor As Double
  Dim rstc As Recordset
  comprovar_colorrodoneta cadbl(rst!coditinta), vIndex  'aquest es el color principal
    'busco els colors alternatius i també poso la rodoneta a cadascú
  Set rstc = dbclixes.OpenRecordset("select * from tintes_alternatives where id_tinter=" + atrim(rst!id_tinter))
  While Not rstc.EOF
     comprovar_colorrodoneta rstc!coditinta, vIndex
     rstc.MoveNext
  Wend
  
End Sub
Function posarcolorTINTA(vcoditinta As String) As Double
  Dim rst As Recordset
  Dim vnumcolor As Double
  posarcolorTINTA = QBColor(15)
  Set rst = dbtintes.OpenRecordset("SELECT colorsetiquetes.codicolor, tintes.codi FROM (subfamiliestintes RIGHT JOIN tintes ON subfamiliestintes.codi = tintes.idsubfamilia) LEFT JOIN colorsetiquetes ON subfamiliestintes.color = colorsetiquetes.nomcolor WHERE (((tintes.codi)='" + atrim(vcoditinta) + "'));")
  If rst.EOF Then Exit Function
  vnumcolor = cadbl(rst!codicolor)
  If vnumcolor = 0 Then vnumcolor = 15
  posarcolorTINTA = QBColor(vnumcolor)
  Set rst = Nothing
  
End Function

Sub comprovar_colorrodoneta(vcoditinta As String, vIndex As Long)
  Dim vControl As Control
  Dim vcolor As Double
  Dim vnomtinta As String
  Dim vsql As String
  Dim rst2 As Recordset
  
  vsql = "SELECT tintes.codi, tintes.descripcio, seriescolors.descripcio AS SERIE, familiescolors.descripcio AS COLOR, subfamiliestintes.descripcio AS SUBFAMILIA FROM ((tintes LEFT JOIN seriescolors ON tintes.idserie = seriescolors.codi) LEFT JOIN subfamiliestintes ON tintes.idsubfamilia = subfamiliestintes.codi) INNER JOIN familiescolors ON tintes.idfamcolor = familiescolors.codi"
  Set rst2 = dbtintes.OpenRecordset(vsql + " where tintes.codi='" + atrim(vcoditinta) + "'")
  If rst2.EOF Then Exit Sub
  vnomtinta = atrim(rst2!descripcio)
  Set vControl = colorstinters1(vIndex)
  If colorstinters1(vIndex).visible Then
        If colorstinters2(vIndex).visible Then
              If colorstinters3(vIndex).visible Then
                      Set vControl = colorstinters4(vIndex)
                   Else: Set vControl = colorstinters3(vIndex)
              End If
           Else: Set vControl = colorstinters2(vIndex)
        End If
  End If
  vcolor = posarcolorTINTA(vcoditinta)
  
  If atrim(rst2!serie) = "YUKOFLEX" Then If vcolor = QBColor(15) Then vcolor = &H808080       'gris
  
  activarCOLOR vcolor, vControl, vIndex, vnomtinta
  
          'activarCOLOR &H808080, vcontrol, vindex    'gris
  'End If
  'If atrim(rst2!serie) = "INNOVA" Then
  '    activarCOLOR vcolor, vcontrol, vindex 'vermell
  'End If
  'If InStr(1, atrim(rst2!serie), "PLATINUM") > 0 Then
  '    activarCOLOR &HF1B75F, vcontrol, vindex     'blau
  'End If
  
  'If atrim(rst2!subfamilia) = "ANTIBLOCKING" Then
  '    activarCOLOR &HFF00FF, vcontrol, vindex     'lila
  'End If
  
End Sub
Sub activarCOLOR(vcolor As Double, vControl As Control, vIndex As Long, vnomtinta As String)
    Dim vnumcontrol As Byte
    Dim vleftcontrol As Double
    'If colorstinters2(vindex).visible = True Then Exit Sub
    vControl.visible = True
    vleftcontrol = cadbl(Mid(colorstinters1(vIndex).tag + " ", 1, InStr(1, colorstinters1(vIndex).tag + " ", " ")))
    vControl.Left = cadbl(vleftcontrol) + vControl.width + 10
    vControl.BackColor = vcolor
    vControl.ZOrder 0
    vControl.tag = vnomtinta
'    If colorstinters2(vindex).visible = False Then
        colorstinters1(vIndex).tag = atrim(vControl.Left) + " " + vControl.tag
        vnumcontrol = cadbl(substituir(vControl.Name, "colorstinters", ""))
        If vnumcontrol < 4 Then
          Set vControl = Me.Controls("colorstinters" + Trim(vnumcontrol + 1))(vIndex) '  colorstinters2(vindex)
        End If
 '   End If
End Sub
Function posar_extensio(rstdatos As Recordset) As String
  Dim rst As Recordset

  Set rst = dbtmpb.OpenRecordset("select codiextensio from extensions_treballsrelacionats where numtreball=" + atrim(rstdatos!id_treball) + " and numordremodificacio=" + atrim(rstdatos!ordremodificacio) + " and coditinta=" + atrim(cadbl(rstdatos!coditinta)), , ReadOnly)
  If Not rst.EOF Then posar_extensio = "E:" + atrim(rst!codiextensio)
  Set rst = Nothing
End Function
Function construirnommaterial(rstmat As Recordset) As String
  construirnommaterial = IIf(atrim(rstmat![familiesmaterials.descripcio]) <> "", atrim(rstmat![familiesmaterials.descripcio]), "")
  'construirnommaterial = construirnommaterial + IIf(atrim(rstmat![familiesmaterials.descripcio]) <> "-", "-" + atrim(rstmat![familiesmaterials.descripcio]), "")
  construirnommaterial = construirnommaterial + IIf(atrim(rstmat![subfamiliesmaterials.descripcio]) <> "-", "-" + atrim(rstmat![subfamiliesmaterials.descripcio]), "")
  construirnommaterial = construirnommaterial + IIf(atrim(rstmat![familiescolorants.descripcio]) <> "-", "-" + atrim(rstmat![familiescolorants.descripcio]), "")
  construirnommaterial = construirnommaterial + IIf(atrim(rstmat![subfamiliescolorants.descripcio]) <> "-", "-" + atrim(rstmat![subfamiliescolorants.descripcio]), "")
  construirnommaterial = construirnommaterial + IIf(atrim(rstmat![familiesaditius.descripcio]) <> "-", "-" + atrim(rstmat![familiesaditius.descripcio]), "")
  construirnommaterial = construirnommaterial + IIf(atrim(rstmat![subfamiliesaditius.descripcio]) <> "-", "-" + atrim(rstmat![subfamiliesaditius.descripcio]), "")
End Function

