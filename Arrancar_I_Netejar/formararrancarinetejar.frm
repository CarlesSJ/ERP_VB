VERSION 5.00
Begin VB.Form form1 
   Caption         =   " Arrancar i Netejar"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   11760
   Icon            =   "formararrancarinetejar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox etnumtreball 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDDECE&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   1260
      TabIndex        =   66
      Top             =   4110
      Width           =   1605
   End
   Begin VB.CommandButton bguardarxl 
      Height          =   585
      Left            =   10380
      Picture         =   "formararrancarinetejar.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Guardar XL a la seva lleixa."
      Top             =   2430
      Width           =   1020
   End
   Begin VB.CommandButton boperari 
      BackColor       =   &H00FDDECE&
      Caption         =   "Escullir Operari"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   645
      Width           =   7395
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informació del treball"
      Height          =   1440
      Left            =   120
      TabIndex        =   0
      Top             =   1635
      Width           =   11595
      Begin VB.Frame FrameCompartits 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   10260
         TabIndex        =   28
         Top             =   180
         Visible         =   0   'False
         Width           =   1245
         Begin VB.Image Imgcompartits 
            Height          =   885
            Left            =   135
            Picture         =   "formararrancarinetejar.frx":0B5E
            Stretch         =   -1  'True
            Top             =   15
            Width           =   885
         End
         Begin VB.Label etcompartits 
            BackStyle       =   0  'Transparent
            Caption         =   "Compartits"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   29
            Top             =   930
            Width           =   1110
         End
      End
      Begin VB.Label etmuntatper 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   5445
         TabIndex        =   73
         Top             =   1155
         Width           =   4710
      End
      Begin VB.Label etguardarxl 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H0000C000&
         Height          =   225
         Left            =   6450
         TabIndex        =   63
         Top             =   1155
         Width           =   3690
      End
      Begin VB.Label etmarcailinia 
         BackStyle       =   0  'Transparent
         Caption         =   "-------------------------------------------------------------"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   255
         TabIndex        =   19
         Top             =   180
         Width           =   10860
      End
      Begin VB.Label etarxiuperdefecte 
         BackStyle       =   0  'Transparent
         Caption         =   "Arxiu: XL-???"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   255
         TabIndex        =   18
         Top             =   720
         Width           =   3900
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   3015
      Width           =   11595
      Begin VB.TextBox etnumtreball 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   2
         Left            =   1170
         TabIndex        =   72
         Top             =   1695
         Width           =   1605
      End
      Begin VB.TextBox etnumtreball 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   3
         Left            =   1140
         TabIndex        =   71
         Top             =   2295
         Width           =   1605
      End
      Begin VB.TextBox etnumtreball 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   4
         Left            =   1155
         TabIndex        =   70
         Top             =   2895
         Width           =   1605
      End
      Begin VB.TextBox etnumtreball 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   5
         Left            =   1140
         TabIndex        =   69
         Top             =   3495
         Width           =   1605
      End
      Begin VB.TextBox etnumtreball 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   6
         Left            =   1155
         TabIndex        =   68
         Top             =   4095
         Width           =   1605
      End
      Begin VB.TextBox etnumtreball 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   7
         Left            =   1140
         TabIndex        =   67
         Top             =   4695
         Width           =   1605
      End
      Begin VB.TextBox etnumtreball 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   0
         Left            =   1170
         TabIndex        =   65
         Text            =   "NT: 12345"
         Top             =   495
         Width           =   1605
      End
      Begin VB.TextBox etneteja 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   7
         Left            =   7860
         TabIndex        =   61
         Top             =   5070
         Width           =   2745
      End
      Begin VB.TextBox etneteja 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   6
         Left            =   7860
         TabIndex        =   60
         Top             =   4470
         Width           =   2745
      End
      Begin VB.TextBox etneteja 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   5
         Left            =   7830
         TabIndex        =   59
         Top             =   3900
         Width           =   2745
      End
      Begin VB.TextBox etneteja 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   4
         Left            =   7845
         TabIndex        =   58
         Top             =   3285
         Width           =   2745
      End
      Begin VB.TextBox etneteja 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   3
         Left            =   7860
         TabIndex        =   57
         Top             =   2655
         Width           =   2745
      End
      Begin VB.TextBox etneteja 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   2
         Left            =   7845
         TabIndex        =   56
         Top             =   2085
         Width           =   2745
      End
      Begin VB.TextBox etneteja 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   1
         Left            =   7845
         TabIndex        =   55
         Top             =   1485
         Width           =   2745
      End
      Begin VB.TextBox etneteja 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   7860
         TabIndex        =   54
         Top             =   885
         Width           =   2745
      End
      Begin VB.CommandButton bnetejar 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   7
         Left            =   10740
         Picture         =   "formararrancarinetejar.frx":0F21
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton bnetejar 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   6
         Left            =   10740
         Picture         =   "formararrancarinetejar.frx":13AB
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4083
         Width           =   615
      End
      Begin VB.CommandButton bnetejar 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   5
         Left            =   10740
         Picture         =   "formararrancarinetejar.frx":1835
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3485
         Width           =   615
      End
      Begin VB.CommandButton bnetejar 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   4
         Left            =   10740
         Picture         =   "formararrancarinetejar.frx":1CBF
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2887
         Width           =   615
      End
      Begin VB.CommandButton bnetejar 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   3
         Left            =   10740
         Picture         =   "formararrancarinetejar.frx":2149
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2289
         Width           =   615
      End
      Begin VB.CommandButton bnetejar 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   2
         Left            =   10740
         Picture         =   "formararrancarinetejar.frx":25D3
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1691
         Width           =   615
      End
      Begin VB.CommandButton bnetejar 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   1
         Left            =   10740
         Picture         =   "formararrancarinetejar.frx":2A5D
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1093
         Width           =   615
      End
      Begin VB.CommandButton bnetejar 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   0
         Left            =   10740
         Picture         =   "formararrancarinetejar.frx":2EE7
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   495
         Width           =   615
      End
      Begin VB.TextBox etestatclixe 
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   7
         Left            =   1215
         TabIndex        =   38
         Top             =   5070
         Width           =   5880
      End
      Begin VB.TextBox etestatclixe 
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   6
         Left            =   1215
         TabIndex        =   37
         Top             =   4470
         Width           =   5880
      End
      Begin VB.TextBox etestatclixe 
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   5
         Left            =   1215
         TabIndex        =   36
         Top             =   3870
         Width           =   5880
      End
      Begin VB.TextBox etestatclixe 
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   4
         Left            =   1215
         TabIndex        =   35
         Top             =   3270
         Width           =   5880
      End
      Begin VB.TextBox etestatclixe 
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   3
         Left            =   1215
         TabIndex        =   34
         Top             =   2685
         Width           =   5880
      End
      Begin VB.TextBox etestatclixe 
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   2
         Left            =   1215
         TabIndex        =   33
         Top             =   2085
         Width           =   5880
      End
      Begin VB.TextBox etestatclixe 
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   1
         Left            =   1215
         TabIndex        =   32
         Top             =   1485
         Width           =   5880
      End
      Begin VB.TextBox etestatclixe 
         BackColor       =   &H00FDDECE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Index           =   0
         Left            =   1215
         TabIndex        =   31
         Top             =   885
         Width           =   5880
      End
      Begin VB.CommandButton binfo 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   7
         Left            =   120
         Picture         =   "formararrancarinetejar.frx":3371
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Informació sobre operari i data"
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton binfo 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   6
         Left            =   120
         Picture         =   "formararrancarinetejar.frx":3669
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Informació sobre operari i data"
         Top             =   4080
         Width           =   615
      End
      Begin VB.CommandButton binfo 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   5
         Left            =   120
         Picture         =   "formararrancarinetejar.frx":3961
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Informació sobre operari i data"
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton binfo 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   4
         Left            =   120
         Picture         =   "formararrancarinetejar.frx":3C59
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Informació sobre operari i data"
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton binfo 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   3
         Left            =   120
         Picture         =   "formararrancarinetejar.frx":3F51
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Informació sobre operari i data"
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton binfo 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   2
         Left            =   120
         Picture         =   "formararrancarinetejar.frx":4249
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Informació sobre operari i data"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton binfo 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   1
         Left            =   120
         Picture         =   "formararrancarinetejar.frx":4541
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Informació sobre operari i data"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton binfo 
         BackColor       =   &H00FDDECE&
         Height          =   600
         Index           =   0
         Left            =   120
         Picture         =   "formararrancarinetejar.frx":4839
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Informació sobre operari i data"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton bcolor 
         BackColor       =   &H00FDDECE&
         Caption         =   "Color 8"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   7
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4680
         Width           =   9540
      End
      Begin VB.CommandButton bcolor 
         BackColor       =   &H00FDDECE&
         Caption         =   "Color 7"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   6
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4080
         Width           =   9540
      End
      Begin VB.CommandButton bcolor 
         BackColor       =   &H00FDDECE&
         Caption         =   "Color 6"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   5
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3480
         Width           =   9540
      End
      Begin VB.CommandButton bcolor 
         BackColor       =   &H00FDDECE&
         Caption         =   "Color 5"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2880
         Width           =   9540
      End
      Begin VB.CommandButton bcolor 
         BackColor       =   &H00FDDECE&
         Caption         =   "Color 4"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   9540
      End
      Begin VB.CommandButton bcolor 
         BackColor       =   &H00FDDECE&
         Caption         =   "Color 3"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   9540
      End
      Begin VB.CommandButton bcolor 
         BackColor       =   &H00FDDECE&
         Caption         =   "Color 2"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   9540
      End
      Begin VB.CommandButton bcolor 
         BackColor       =   &H00FDDECE&
         Caption         =   "Color 1"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   0
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   9540
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Netejar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   270
         Left            =   10650
         TabIndex        =   40
         Top             =   150
         Width           =   930
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Arrancar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Left            =   60
         TabIndex        =   39
         Top             =   150
         Width           =   1080
      End
      Begin VB.Label ettinter 
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00ED823A&
         Height          =   480
         Index           =   0
         Left            =   795
         TabIndex        =   10
         Top             =   660
         Width           =   375
      End
      Begin VB.Label ettinter 
         BackStyle       =   0  'Transparent
         Caption         =   "8."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00ED823A&
         Height          =   480
         Index           =   7
         Left            =   780
         TabIndex        =   17
         Top             =   4890
         Width           =   375
      End
      Begin VB.Label ettinter 
         BackStyle       =   0  'Transparent
         Caption         =   "7."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00ED823A&
         Height          =   480
         Index           =   6
         Left            =   780
         TabIndex        =   16
         Top             =   4290
         Width           =   375
      End
      Begin VB.Label ettinter 
         BackStyle       =   0  'Transparent
         Caption         =   "6."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00ED823A&
         Height          =   480
         Index           =   5
         Left            =   780
         TabIndex        =   15
         Top             =   3675
         Width           =   375
      End
      Begin VB.Label ettinter 
         BackStyle       =   0  'Transparent
         Caption         =   "5."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00ED823A&
         Height          =   480
         Index           =   4
         Left            =   780
         TabIndex        =   14
         Top             =   3075
         Width           =   375
      End
      Begin VB.Label ettinter 
         BackStyle       =   0  'Transparent
         Caption         =   "4."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00ED823A&
         Height          =   480
         Index           =   3
         Left            =   780
         TabIndex        =   13
         Top             =   2475
         Width           =   375
      End
      Begin VB.Label ettinter 
         BackStyle       =   0  'Transparent
         Caption         =   "3."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00ED823A&
         Height          =   480
         Index           =   2
         Left            =   780
         TabIndex        =   12
         Top             =   1875
         Width           =   375
      End
      Begin VB.Label ettinter 
         BackStyle       =   0  'Transparent
         Caption         =   "2."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00ED823A&
         Height          =   480
         Index           =   1
         Left            =   780
         TabIndex        =   11
         Top             =   1260
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1530
      Left            =   105
      TabIndex        =   42
      Top             =   75
      Width           =   11610
      Begin VB.CheckBox checkpendents 
         Caption         =   "Només veure pendents."
         Height          =   225
         Left            =   3285
         TabIndex        =   64
         Top             =   165
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.CommandButton btreball 
         BackColor       =   &H00EAD9CE&
         Caption         =   "Escullir NT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   570
         Width           =   2160
      End
      Begin VB.Label etntescullit 
         Caption         =   "Escullir treball"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   390
         TabIndex        =   46
         Top             =   225
         Width           =   2265
      End
      Begin VB.Label Label2 
         Caption         =   "Escullir OPERARI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   5625
         TabIndex        =   45
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label etcomanda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9900
         TabIndex        =   44
         Top             =   135
         Width           =   1575
      End
   End
   Begin VB.Menu mllistats 
      Caption         =   "Llistats"
      Begin VB.Menu mentredates 
         Caption         =   "Entre dates"
      End
      Begin VB.Menu mntreball 
         Caption         =   "Nº Treball"
      End
      Begin VB.Menu mnumlot 
         Caption         =   "Per Lot"
      End
   End
   Begin VB.Menu mbossaoficines 
      Caption         =   "Bossa a oficines"
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nomoperari As String
Dim numop As Integer

Private Sub bguardarxl_Click()
   If cadbl(etntescullit.Tag) = 0 Then MsgBox "No hi ha cap treball escullit", vbCritical, "Error": Exit Sub
   If Not comprovar_si_ja_esta Then MsgBox "No estan tots els clixes arrancats", vbCritical, "Error"
   carregar_dades , etcomanda
End Sub

Private Sub binfo_Click(Index As Integer)
   Dim vopactual As String
   If bcolor(Index).Caption = "" Then MsgBox "En aquest tinter no hi havia color.", vbCritical, "Error": Exit Sub
   vopactual = etestatclixe(Index)
   If InStr(1, vopactual, "Op:") > 0 Then
     vopactual = Mid(vopactual, InStr(1, vopactual, "Op:"))
     vopactual = Mid(vopactual, 4, InStr(1, vopactual, " -") - 4)
     If cadbl(vopactual) <> numop Then MsgBox "Aquesta tipificació la va posar un altra treballador tu no pots modificar-la", vbCritical, "Error": Exit Sub
   End If
   Unload Formtipificacions
   Formtipificacions.Show 1
   If Formtipificacions.Tag <> "" Then
       'guardar la tipificació
       If etestatclixe(Index) <> "" Then If MsgBox("Aquesta linia ja te tipificació. Vols eliminar-la?", vbExclamation + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Sub
       If Formtipificacions.Tag = "- BORRAR -" Then Formtipificacions.Tag = ""
       dbbaixes.Execute "update muntadorescilindres set oparrencar=" + atrim(numop) + ", dataarrencar=now, detallestatclixe='" + treure_apostruf(Formtipificacions.Tag) + "' where id=" + atrim(bcolor(Index).Tag)
       carregar_dades , cadbl(etcomanda)
       
   End If
End Sub

Private Sub bnetejar_Click(Index As Integer)
   If bcolor(Index).Caption = "" Then MsgBox "En aquest tinter no hi havia color.", vbCritical, "Error": Exit Sub
   If MsgBox("Clixé net?", vbDefaultButton2 + vbYesNo, "Comfirma") = vbNo Then
       dbbaixes.Execute "update muntadorescilindres set opneteja=0, dataneteja=null where id=" + atrim(bcolor(Index).Tag)
         Else
             dbbaixes.Execute "update muntadorescilindres set opneteja=" + atrim(numop) + ", dataneteja=now where id=" + atrim(bcolor(Index).Tag)
   End If
   carregar_dades , cadbl(etcomanda)
   comprovar_si_ja_esta
End Sub
Function comprovar_si_ja_esta() As Boolean
  Dim i As Byte
  Dim vtotfet As Boolean
  vtotfet = True
  For i = 0 To 7
     If bcolor(i).Caption <> "" And etneteja(i) = "" And bcolor(i).Enabled Then vtotfet = False
  Next i
  comprovar_si_ja_esta = vtotfet
  If vtotfet Then
    If MsgBox("Tot fet, vols guardar?", vbInformation + vbDefaultButton2 + vbYesNo, "Guardar a XL") = vbYes Then
         guardar_a_XL
    End If
  End If
End Function
Sub guardar_a_XL()
   Load formguardarXL
   formguardarXL.etoperari = atrim(numop)
   formguardarXL.Show 1
End Sub

Private Sub boperari_Click()
 Dim numoptmp As Integer
 Dim nomoptmp As String
  Load formseleccio
  formseleccio.Data1.DatabaseName = camicomandes
  formseleccio.Data1.RecordSource = "select distinct codi,descripcio from operaris where (maquina='M' or maquina='I') and actiu<>0 order by codi "
  formseleccio.Caption = "Selecció d'Operari"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   numoptmp = cadbl(formseleccio.Data1.Recordset!codi)
   nomoptmp = atrim(formseleccio.Data1.Recordset!descripcio)
  End If
  Unload formseleccio
  If numoptmp <> 0 Then
     nomoperari = nomoptmp
     numop = numoptmp
     boperari.Caption = nomoptmp
     For Each objecte In Me
      If objecte.Name <> "llistat" And objecte.Name <> "llistatbob" And objecte.Name <> "Line1" And objecte.Name <> "comandaacavada" Then
        objecte.Enabled = True
      End If
     Next objecte
      Else: If cadbl(numop) = 0 Then MsgBox "Has d'escullir un operari per treballar": Exit Sub
  End If
   If cadbl(comanda) > 0 Then
      btreball_Click
   End If
   

End Sub

Private Sub btreball_Click()
   Dim vnumtreball As String
   vnumtreball = InputBox("ENTRA EL NÚMERO DE TREBALL QUE VOLS", "ESCULLIR TREBALL")
   If comprovarNT(cadbl(vnumtreball)) And vnumtreball <> "" Then
        carregar_dades cadbl(vnumtreball)
   End If
End Sub
Function escullirunacomanda(vsql As String) As Double
  Load formseleccio
  formseleccio.Data1.DatabaseName = rutadelfitxer(camicomandes) + "baixes.mdb"
  formseleccio.Data1.RecordSource = vsql
  formseleccio.Caption = "Escullir comanda"
  formseleccio.refrescar
'  formseleccio.DBGrid2.Columns(1).Visible = False
  formseleccio.DBGrid2.Columns(4).Visible = False
  formseleccio.DBGrid2.Columns(0).Width = 2200
  formseleccio.DBGrid2.Columns(1).Width = 1800
  formseleccio.DBGrid2.Columns(2).Width = 800
  formseleccio.DBGrid2.Columns(3).Width = 3000
 ' formseleccio.DBGrid2.Columns(4).Width = 1500
  
  formseleccio.Show 1
  If seleccioret = 1 Then
   escullirunacomanda = cadbl(formseleccio.Data1.Recordset!comanda)
  End If
  Unload formseleccio
End Function
Sub netejar_dades(Optional vtreball As Double)
  Dim i As Byte
  etarxiuperdefecte = ""
  etarxiuperdefecte.Tag = ""
  etguardarxl = ""
  etmarcailinia = ""
  etntescullit = ""
  If vtreball <> 0 Then etntescullit.Tag = ""
  For i = 0 To 7
   bcolor(i).Caption = ""
   bcolor(i).Tag = ""
   bcolor(i).BackColor = &HFDDECE
   binfo(i).BackColor = &HFDDECE
   bnetejar(i).BackColor = &HFDDECE
   etnumtreball(i).BackColor = &HFDDECE
   etnumtreball(i) = ""
   etestatclixe(i).BackColor = &HFDDECE
   etestatclixe(i) = ""
   etneteja(i).BackColor = &HFDDECE
   etneteja(i) = ""
   binfo(i).Enabled = True: bcolor(i).Enabled = True: bnetejar(i).Enabled = True
  Next i
  
End Sub
Sub carregar_dades(Optional vnumtreball As Double, Optional vnumc As Double)
  
  Dim rst As Recordset
  Dim vsql As String
  Dim vcolor As Double
  Dim vnomespendents As String
  Dim rsttinters As Recordset
  netejar_dades vnumtreball
  If vnumc = 0 Then
        'vsql = "SELECT distinct muntadorescilindres.numcomanda as [Comanda], muntadorescilindres.oparrencar, muntadorescilindres.oparrencar, comandes.numtreball as [Treball], comandes.numordremodificacio as [Versió] FROM muntadorescilindres INNER JOIN comandes ON muntadorescilindres.numcomanda = comandes.comanda WHERE (((muntadorescilindres.oparrencar)=0) AND ((comandes.numtreball)=" + atrim(vnumtreball) + "));"
        'vsql = "SELECT  distinct muntadorescilindres.numcomanda as [Comanda],comandes.numtreball as [Treball],   comandes.numordremodificacio as [Versió],impressorestot.dataimpressio FROM (muntadorescilindres INNER JOIN comandes ON muntadorescilindres.numcomanda = comandes.comanda) INNER JOIN impressorestot ON comandes.comanda = impressorestot.comanda WHERE (((comandes.numtreball)=" + atrim(vnumtreball) + ")) order by dataimpressio desc;"
        vnomespendents = IIf(checkpendents.Value = 1, "and muntadorescilindres.opendreçar=0", "")
        vsql = "SELECT DISTINCT muntadorescilindres.numcomanda AS Comanda, IIf(Tintes_clixesnous_1.id_treball>0,Tintes_clixesnous_1.id_treball,Tintes_clixesnous.id_treball) AS Treball, tintes_clixesnous.ordremodificacio as [Versió], format(impressorestot.dataimpressio,'dd/mm/yy hh:nn') as DataImp,impressorestot.dataimpressio  FROM ((muntadorescilindres LEFT JOIN Tintes_clixesnous ON muntadorescilindres.id_tinter = Tintes_clixesnous.id_tinter) LEFT JOIN impressorestot ON muntadorescilindres.numcomanda = impressorestot.comanda) LEFT JOIN Tintes_clixesnous AS Tintes_clixesnous_1 ON Tintes_clixesnous.tinterlinkambid_treball = Tintes_clixesnous_1.id_tinter Where "
        vsql = vsql + " (((IIf(Tintes_clixesnous_1.id_treball > 0, Tintes_clixesnous_1.id_treball, Tintes_clixesnous.id_treball)) = " + Trim(cadbl(vnumtreball)) + ")) " + vnomespendents + " ORDER BY impressorestot.dataimpressio DESC;"
'        Clipboard.Clear
'        Clipboard.SetText vsql
        ratoli "espera"
        Set rst = dbbaixes.OpenRecordset(vsql)
        If rst.EOF Then MsgBox "No hi ha cap muntatge amb aquest treball.", vbCritical, "Error": ratoli "normal": Exit Sub
        rst.MoveLast
        ratoli "normal"
        If rst.RecordCount > 1 Then
           vnumc = escullirunacomanda(vsql)
           If vnumc = 0 Then Exit Sub
            Else: vnumc = rst!comanda
        End If
  End If
  etcomanda = atrim(vnumc)
  'si entra per numcomanda busca quin treball es sino no cal
  If vnumtreball = 0 Then
     'Set rst = dbbaixes.OpenRecordset("select numtreball from comandes where comanda=" + atrim(vnumc))
     'If rst.EOF Then Exit Sub
     'vnumtreball = cadbl(rst!numtreball)
     vnumtreball = cadbl(etntescullit.Tag)
  End If
  'busca la marca i linia i arxiu
  Set rst = dbclixes.OpenRecordset("select marca,linia,arxiu from clixes where id_treball=" + atrim(vnumtreball))
  If rst.EOF Then Exit Sub
  etmarcailinia = atrim(rst!marca) + " - " + atrim(rst!linia)
  etarxiuperdefecte = "Arxiu: " + atrim(rst!arxiu)
  etarxiuperdefecte.Tag = atrim(rst!arxiu)
  
  
  
  Set rst = dbbaixes.OpenRecordset("Select * from muntadorescilindres where numcomanda=" + atrim(vnumc) + " order by numcilindre")
  bguardarxl.Tag = ""
  While Not rst.EOF
    Set rsttinters = dbbaixes.OpenRecordset("select * from tintes_clixesnous where id_tinter=" + atrim(cadbl(rst!id_tinter)))
    If rsttinters.EOF Then GoTo proxim
    If rsttinters!clixeosleeve = "Sleeve" Then GoTo proxim
    If cadbl(rsttinters!tinterlinkambid_treball) > 0 Then Set rsttinters = dbbaixes.OpenRecordset("select * from tintes_clixesnous where id_tinter=" + atrim(rsttinters!tinterlinkambid_treball))
    If rsttinters!id_treball = vnumtreball Then
        binfo(rst!numcilindre - 1).Enabled = True: bcolor(rst!numcilindre - 1).Enabled = True: bnetejar(rst!numcilindre - 1).Enabled = True
        If cadbl(rsttinters!tinterlinkambid_treball) < 0 Then  'si hi ha ancora rosa també el marco com a  no guardar-lo a la mateixa bossa
           binfo(rst!numcilindre - 1).Enabled = False: bcolor(rst!numcilindre - 1).Enabled = False: bnetejar(rst!numcilindre - 1).Enabled = False
        End If
         Else
           binfo(rst!numcilindre - 1).Enabled = False: bcolor(rst!numcilindre - 1).Enabled = False: bnetejar(rst!numcilindre - 1).Enabled = False
    End If
    bcolor(rst!numcilindre - 1).Caption = atrim(rst!descripcio)
    bcolor(rst!numcilindre - 1).Tag = atrim(rst!ID)
    'guardo els ids per despres poder guardar la bossa
    If binfo(rst!numcilindre - 1).Enabled Then bguardarxl.Tag = IIf(bguardarxl.Tag <> "", bguardarxl.Tag + ",", "") + atrim(rst!ID)
    If atrim(rst!descripcio) <> "" Then etnumtreball(rst!numcilindre - 1) = Mid(atrim(rst!observacio), 1, InStr(1, atrim(rst!observacio), " "))
    If atrim(rst!detallestatclixe) <> "" Then
       etestatclixe(rst!numcilindre - 1) = "Data: " + atrim(Format(rst!dataarrencar, "dd/mm/yy")) + " Op:" + atrim(rst!oparrencar) + " -" + atrim(rst!detallestatclixe)
       'etnumtreball(rst!numcilindre - 1) = Mid(rst!observacio, 1, InStr(1, rst!observacio, "XL:"))
    End If
    If cadbl(rst!opneteja) > 0 Then etneteja(rst!numcilindre - 1) = "Data: " + atrim(Format(rst!dataneteja, "dd/mm/yy")) + " Op:" + atrim(rst!opneteja)
    'si hi ha un estat del clixe poso verd si no blau
    If atrim(rst!detallestatclixe) = "" Then
        vcolor = &HFDDECE '&HFDDECE blau
         Else:
            vcolor = &H6BEBB1 '&H6BEBB1 verd
    End If
    binfo(rst!numcilindre - 1).BackColor = vcolor
    If cadbl(rst!opneteja) > 0 Then
       vcolor = &HFF00FF
       bnetejar(rst!numcilindre - 1).BackColor = vcolor
    End If
    
    bcolor(rst!numcilindre - 1).BackColor = vcolor
    etestatclixe(rst!numcilindre - 1).BackColor = vcolor
    etneteja(rst!numcilindre - 1).BackColor = vcolor
    etntescullit = "NT: " + Trim(vnumtreball)
    etntescullit.Tag = Trim(vnumtreball)
proxim:
    rst.MoveNext
  Wend
  
  'carrega les linies de muntadora per poder marcar com a arrencades
  Set rst = dbbaixes.OpenRecordset("Select * from muntadorescilindres where id in(" + atrim(bguardarxl.Tag) + ")")
  If Not rst.EOF Then
     If cadbl(atrim(rst!opendreçar)) > 0 Then etguardarxl = "Guardat per: " + atrim(rst!opendreçar) + "   " + Format(rst!dataendreça, "dd/mm/yy")
  End If
  'poso el nom del operari que va fer el muntatge
  etmuntatper = ""
  Set rst = dbbaixes.OpenRecordset("select * from muntadoratot where comanda=" + atrim(vnumc))
  If Not rst.EOF Then
      etmuntatper = "Muntat per:  " + atrim(rst!firma) + " - " + atrim(rst!nomfirma)
  End If
  
End Sub
Function comprovarNT(vnt As Double) As Boolean
   Dim rst As Recordset
   Dim rst2 As Recordset
   comprovarNT = False
   Set rst = dbcomandes.OpenRecordset("select comanda,proximaseccio from comandes where proximaseccio<>'E' and numtreball=" + atrim(vnt), , ReadOnly)
   If rst.EOF Then MsgBox "D'aquest treball no s'ha imprès cap comanda.", vbCritical, "Error": GoTo fi
   While Not rst.EOF And comprovarNT = False
      Set rst2 = dbbaixes.OpenRecordset("select * from impressores where tipus='F' and comanda=" + atrim(rst!comanda))
      If rst2.EOF Then
        comprovarNT = False
        MsgBox "Aquest numero de treball encara no s'ha imprès o s'ha de utilitzar per una altra comanda pendent.", vbCritical, "Atenció"
        GoTo fi
          Else:
            comprovarNT = True
      End If
   Wend
fi:
   Set rst2 = Nothing
   Set rst = Nothing
End Function

Private Sub Form_Load()
  camicomandes = llegir_ini("General", "cami", "comandes.ini")
  cami = llegir_ini("General", "camibaixes", "comandes.ini")
  fitxerini = "comandes.ini"
  If cami = "{[}]" Then
    escriure_ini "General", "camibaixes", InputBox("Entra la ruta de baixes", "Atenció", "y:\comandes\baixes.mdb"), "comandes.ini"
  End If
  If Not existeix("c:\ordprog.ini") Then assignardecimalipunt
  centerscreen Me
  Set dbcomandes = OpenDatabase(cami)
  Set dbclixes = OpenDatabase(rutadelfitxer(cami) + "clixesnous.mdb")
  Set dbbaixes = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb")
  If llegir_ini("Baixes", "programaamaquina", fitxerini) = "1" Then
   Shell ("net time \\serverprodu /set /y")
  End If
  For Each objecte In Me
      If objecte.Name <> "boperari" And objecte.Name <> "Line1" And objecte.Name <> "rellotge" And objecte.Name <> "llistat" And objecte.Name <> "llistatbob" Then
        objecte.Enabled = False
      End If
  Next objecte
 netejar_dades
End Sub

Private Sub mbossaoficines_Click()
 Dim vtreball As String
 Dim resp As Double
 resp = vbYes
 While resp = vbYes
     vtreball = InputBox("Escaneja el treball de la bossa que t'emportes a OFICINA.", "ESCANEJA TREBALL")
     If StrPtr(vtreball) = 0 Then Exit Sub
     If cadbl(vtreball) > 0 Then
         dbclixes.Execute "update  clixes set ubicacio='OFICINA' where id_treball=" + atrim(vtreball)
          MsgBox "Ubicació canviada al treball " + vtreball
          Else: MsgBox "Aixó no es un número de treball vàlid.", vbCritical, "Error"
     End If
     resp = MsgBox("Vols escanejar mes treballs?", vbInformation + vbDefaultButton2 + vbYesNo)
 Wend
End Sub

Private Sub mentredates_Click()
   Dim vsql As String
   Dim vinici As String
   Dim vfi As String
   vinici = InputBox("Escriu la data d'inici de la consulta." + vbNewLine + " Ex: dd/mm/yy", "Inici")
   If Not IsDate(vinici) Then Exit Sub
   vfi = InputBox("Escriu la data de fi de la consulta." + vbNewLine + " Ex: dd/mm/yy", "Fi")
   If Not IsDate(vfi) Then Exit Sub
   
   vsql = "SELECT muntadorescilindres.numcomanda, First(muntadorescilindres.observacio) AS PrimeroDeobservacio, First(muntadorescilindres.dataendreça) AS PrimeroDedataendreça From muntadorescilindres "
   vsql = vsql + " Where (((muntadorescilindres.dataendreça) > #" + Format(vinici, "mm/dd/yy") + "# And (muntadorescilindres.dataendreça) < #" + Format(vfi, "mm/dd/yy") + "#)) GROUP BY muntadorescilindres.numcomanda;"
   ferCSVarrancar vsql
   
End Sub

Private Sub mntreball_Click()
   Dim vsql As String
   Dim vnumtreball As String
   vnumtreball = InputBox("Escriu el numero de treball que vols buscar.", "Treball")
   If cadbl(vnumtreball) = 0 Then Exit Sub
   vsql = "SELECT muntadorescilindres.numcomanda,First(muntadorescilindres.observacio) AS PrimeroDeobservacio, First(muntadorescilindres.dataendreça) AS PrimeroDedataendreça From muntadorescilindres "
   vsql = vsql + " where numcomanda in (select comanda from comandes where numtreball=" + atrim(vnumtreball) + ") GROUP BY muntadorescilindres.numcomanda;"
   ferCSVarrancar vsql
End Sub

Private Sub mnumlot_Click()
   Dim vsql As String
   Dim vnumc As String
   vnumc = InputBox("Escriu el numero de comanda que vols saber.", "Comanda")
   If cadbl(vnumc) = 0 Then Exit Sub
   vsql = "SELECT muntadorescilindres.numcomanda,First(muntadorescilindres.observacio) AS PrimeroDeobservacio, First(muntadorescilindres.dataendreça) AS PrimeroDedataendreça From muntadorescilindres WHERE (((muntadorescilindres.numcomanda) In (select comanda from comandes "
   vsql = vsql + " where numcomanda=" + atrim(vnumc) + "))) GROUP BY muntadorescilindres.numcomanda;"
   ferCSVarrancar vsql

End Sub
Sub ferCSVarrancar(vsql As String)
   Dim rst As Recordset
   Dim vfitxer As String
   If Not existeix("c:\temp") Then MkDir "c:\temp"
   vfitxer = "c:\temp\llistatEndreçar.csv"
   Set rst = dbbaixes.OpenRecordset(vsql)
   If rst.EOF Then GoTo fi
   Open vfitxer For Output As 1
   Print #1, "Comanda;Descripció;Data endreça"
   While Not rst.EOF
     Print #1, atrim(rst!numcomanda) + ";" + atrim(rst!primerodeobservacio) + ";" + atrim(rst!PrimeroDedataendreça)
     rst.MoveNext
   Wend
   Close 1
   If existeix(vfitxer) Then obrir_document vfitxer
fi:
   Set rst = Nothing
End Sub
