VERSION 5.00
Object = "{8C45F041-B87C-11D1-96EF-845C0FC10100}#1.3#0"; "SCROLLBOX.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form formcomandes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manteniment de Comandes"
   ClientHeight    =   7860
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10575
   Icon            =   "comandes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin ScrollBoxCtl.ScrollBox formscrooll 
      Height          =   7005
      Left            =   90
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   750
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   12356
      ScrollBars      =   2
      Caption         =   ""
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Frame areadatos 
         Enabled         =   0   'False
         Height          =   25000
         Left            =   0
         TabIndex        =   9
         Top             =   -105
         Width           =   10065
         Begin VB.Frame imp1 
            Caption         =   "Impressora-1"
            Height          =   3930
            Left            =   45
            TabIndex        =   266
            Top             =   6525
            Width           =   9915
            Begin VB.TextBox Text40 
               DataField       =   "tinta1a"
               DataSource      =   "data1"
               Height          =   285
               Left            =   885
               TabIndex        =   305
               Top             =   195
               Width           =   2910
            End
            Begin VB.TextBox Text46 
               DataField       =   "tinta2a"
               DataSource      =   "data1"
               Height          =   285
               Left            =   885
               TabIndex        =   304
               Top             =   450
               Width           =   2910
            End
            Begin VB.TextBox Text47 
               DataField       =   "tinta3a"
               DataSource      =   "data1"
               Height          =   285
               Left            =   885
               TabIndex        =   303
               Top             =   705
               Width           =   2910
            End
            Begin VB.TextBox Text48 
               DataField       =   "tinta4a"
               DataSource      =   "data1"
               Height          =   285
               Left            =   885
               TabIndex        =   302
               Top             =   960
               Width           =   2910
            End
            Begin VB.TextBox Text49 
               DataField       =   "tinta5a"
               DataSource      =   "data1"
               Height          =   285
               Left            =   885
               TabIndex        =   301
               Top             =   1215
               Width           =   2910
            End
            Begin VB.TextBox Text50 
               DataField       =   "tinta6a"
               DataSource      =   "data1"
               Height          =   285
               Left            =   885
               TabIndex        =   300
               Top             =   1470
               Width           =   2910
            End
            Begin VB.TextBox Text51 
               DataField       =   "lin1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3840
               MaxLength       =   3
               TabIndex        =   299
               Top             =   195
               Width           =   465
            End
            Begin VB.TextBox Text52 
               DataField       =   "lin2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3840
               MaxLength       =   3
               TabIndex        =   298
               Top             =   450
               Width           =   465
            End
            Begin VB.TextBox Text53 
               DataField       =   "lin3"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3840
               MaxLength       =   3
               TabIndex        =   297
               Top             =   705
               Width           =   465
            End
            Begin VB.TextBox Text54 
               DataField       =   "lin4"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3840
               MaxLength       =   3
               TabIndex        =   296
               Top             =   960
               Width           =   465
            End
            Begin VB.TextBox Text55 
               DataField       =   "lin5"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3840
               MaxLength       =   3
               TabIndex        =   295
               Top             =   1215
               Width           =   465
            End
            Begin VB.TextBox Text56 
               DataField       =   "lin6"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3840
               MaxLength       =   3
               TabIndex        =   294
               Top             =   1455
               Width           =   465
            End
            Begin VB.TextBox Text57 
               DataField       =   "tinta1b"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4350
               TabIndex        =   293
               Top             =   195
               Width           =   2970
            End
            Begin VB.TextBox Text58 
               DataField       =   "tinta2b"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4350
               TabIndex        =   292
               Top             =   450
               Width           =   2970
            End
            Begin VB.TextBox Text59 
               DataField       =   "tinta3b"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4350
               TabIndex        =   291
               Top             =   705
               Width           =   2970
            End
            Begin VB.TextBox Text60 
               DataField       =   "tinta4b"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4350
               TabIndex        =   290
               Top             =   960
               Width           =   2970
            End
            Begin VB.TextBox Text61 
               DataField       =   "tinta5b"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4350
               TabIndex        =   289
               Top             =   1215
               Width           =   2970
            End
            Begin VB.TextBox Text62 
               DataField       =   "tinta6b"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4350
               TabIndex        =   288
               Top             =   1470
               Width           =   2970
            End
            Begin VB.TextBox Text63 
               DataField       =   "numerotintes"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8175
               TabIndex        =   287
               Top             =   195
               Width           =   405
            End
            Begin VB.TextBox Text64 
               DataField       =   "impressio"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8865
               TabIndex        =   286
               Top             =   210
               Visible         =   0   'False
               Width           =   405
            End
            Begin VB.TextBox Text65 
               DataField       =   "formaimp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9345
               TabIndex        =   285
               Top             =   225
               Visible         =   0   'False
               Width           =   405
            End
            Begin VB.ComboBox cimpressio 
               Height          =   315
               ItemData        =   "comandes.frx":0442
               Left            =   8175
               List            =   "comandes.frx":044F
               TabIndex        =   284
               Top             =   555
               Width           =   1425
            End
            Begin VB.ComboBox ctipusimp 
               Height          =   315
               ItemData        =   "comandes.frx":046F
               Left            =   8175
               List            =   "comandes.frx":0479
               TabIndex        =   283
               Top             =   885
               Width           =   1440
            End
            Begin VB.TextBox Text66 
               DataField       =   "dessarroll"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8175
               TabIndex        =   282
               Top             =   1230
               Width           =   495
            End
            Begin VB.TextBox Text67 
               DataField       =   "cilindres"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9330
               TabIndex        =   281
               Top             =   1230
               Width           =   495
            End
            Begin VB.TextBox Text68 
               DataField       =   "obert"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8175
               TabIndex        =   280
               Top             =   1530
               Width           =   495
            End
            Begin VB.TextBox Text69 
               DataField       =   "arxiu"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9135
               TabIndex        =   279
               Top             =   1545
               Width           =   690
            End
            Begin VB.TextBox Text70 
               DataField       =   "arxiumontadora"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8595
               TabIndex        =   278
               Top             =   2400
               Width           =   1230
            End
            Begin VB.TextBox Text71 
               DataField       =   "codibarras"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6900
               TabIndex        =   277
               Text            =   "1234567890123"
               Top             =   2400
               Width           =   1335
            End
            Begin VB.TextBox Text72 
               DataField       =   "mtrsminut"
               DataSource      =   "data1"
               Height          =   285
               Left            =   5685
               TabIndex        =   276
               Top             =   2385
               Width           =   630
            End
            Begin VB.TextBox Text73 
               DataField       =   "impressora"
               DataSource      =   "data1"
               Height          =   285
               Left            =   870
               TabIndex        =   275
               Top             =   2370
               Width           =   630
            End
            Begin VB.TextBox Text74 
               DataField       =   "obsimp2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   495
               TabIndex        =   274
               Top             =   2985
               Width           =   7500
            End
            Begin VB.TextBox Text75 
               DataField       =   "obsimp1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   495
               TabIndex        =   273
               Top             =   2745
               Width           =   7500
            End
            Begin VB.TextBox Text76 
               DataField       =   "obsimpgen2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   495
               TabIndex        =   272
               Top             =   3525
               Width           =   7500
            End
            Begin VB.TextBox Text77 
               DataField       =   "obsimpgen1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   495
               TabIndex        =   271
               Top             =   3285
               Width           =   7500
            End
            Begin VB.TextBox Text78 
               DataField       =   "arxiuext"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   8220
               Locked          =   -1  'True
               MouseIcon       =   "comandes.frx":0494
               MousePointer    =   99  'Custom
               TabIndex        =   270
               TabStop         =   0   'False
               Top             =   2955
               Width           =   1290
            End
            Begin VB.CommandButton Command1 
               Height          =   285
               Left            =   9570
               Picture         =   "comandes.frx":08D6
               Style           =   1  'Graphical
               TabIndex        =   269
               Top             =   2955
               Width           =   285
            End
            Begin VB.TextBox Text79 
               DataField       =   "arxiuext"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   8220
               Locked          =   -1  'True
               MouseIcon       =   "comandes.frx":0C9C
               MousePointer    =   99  'Custom
               TabIndex        =   268
               TabStop         =   0   'False
               Top             =   3510
               Width           =   1275
            End
            Begin VB.CommandButton Command3 
               Height          =   285
               Left            =   9555
               Picture         =   "comandes.frx":10DE
               Style           =   1  'Graphical
               TabIndex        =   267
               Top             =   3510
               Width           =   285
            End
            Begin VB.TextBox Text141 
               DataField       =   "tinta7a"
               DataSource      =   "data1"
               Height          =   285
               Left            =   885
               TabIndex        =   333
               Top             =   1725
               Width           =   2910
            End
            Begin VB.TextBox Text140 
               DataField       =   "tinta8a"
               DataSource      =   "data1"
               Height          =   285
               Left            =   885
               TabIndex        =   332
               Top             =   1980
               Width           =   2910
            End
            Begin VB.TextBox Text139 
               DataField       =   "lin7"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3840
               MaxLength       =   3
               TabIndex        =   331
               Top             =   1710
               Width           =   465
            End
            Begin VB.TextBox Text138 
               DataField       =   "lin8"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3840
               MaxLength       =   3
               TabIndex        =   330
               Top             =   1965
               Width           =   465
            End
            Begin VB.TextBox Text137 
               DataField       =   "tinta7b"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4350
               TabIndex        =   329
               Top             =   1725
               Width           =   2970
            End
            Begin VB.TextBox Text136 
               DataField       =   "tinta8b"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4350
               TabIndex        =   328
               Top             =   1980
               Width           =   2970
            End
            Begin VB.Label Label1 
               Caption         =   "1ª Tinta A:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   38
               Left            =   105
               TabIndex        =   327
               Top             =   240
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "2ª Tinta A:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   39
               Left            =   105
               TabIndex        =   326
               Top             =   495
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "3ª Tinta A:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   40
               Left            =   105
               TabIndex        =   325
               Top             =   750
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "4ª Tinta A:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   41
               Left            =   105
               TabIndex        =   324
               Top             =   1005
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "5ª Tinta A:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   42
               Left            =   105
               TabIndex        =   323
               Top             =   1260
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "6ª Tinta A:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   43
               Left            =   105
               TabIndex        =   322
               Top             =   1515
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "NºTinters:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   50
               Left            =   7440
               TabIndex        =   321
               Top             =   255
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "Impressió:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   51
               Left            =   7455
               TabIndex        =   320
               Top             =   570
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "Forma Im:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   52
               Left            =   7455
               TabIndex        =   319
               Top             =   870
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "Desarroll:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   53
               Left            =   7440
               TabIndex        =   318
               Top             =   1260
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "Cilindres:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   54
               Left            =   8685
               TabIndex        =   317
               Top             =   1260
               Width           =   630
            End
            Begin VB.Label Label1 
               Caption         =   "Obert (N/1/2/C)"
               DataSource      =   "data1"
               Height          =   420
               Index           =   55
               Left            =   7440
               TabIndex        =   316
               Top             =   1455
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "Arxiu:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   56
               Left            =   8715
               TabIndex        =   315
               Top             =   1590
               Width           =   540
            End
            Begin VB.Label Label1 
               Caption         =   "A.M:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   57
               Left            =   8265
               TabIndex        =   314
               Top             =   2445
               Width           =   435
            End
            Begin VB.Label Label1 
               Caption         =   "C.Bar.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   58
               Left            =   6375
               TabIndex        =   313
               Top             =   2430
               Width           =   570
            End
            Begin VB.Label Label1 
               Caption         =   "Mtrs/Min.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   59
               Left            =   4950
               TabIndex        =   312
               Top             =   2475
               Width           =   915
            End
            Begin VB.Label nomimpressora 
               Caption         =   "nomimpressora"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   1
               Left            =   1635
               TabIndex        =   311
               Top             =   2445
               Width           =   3225
            End
            Begin VB.Label Label1 
               Caption         =   "Impress.:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   60
               Left            =   120
               TabIndex        =   310
               Top             =   2430
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "Obs. Imp"
               DataSource      =   "data1"
               Height          =   480
               Index           =   61
               Left            =   60
               TabIndex        =   309
               Top             =   2775
               Width           =   510
            End
            Begin VB.Label Label1 
               Caption         =   "Obs. Client"
               DataSource      =   "data1"
               Height          =   480
               Index           =   62
               Left            =   60
               TabIndex        =   308
               Top             =   3315
               Width           =   465
            End
            Begin VB.Label Label1 
               Caption         =   "Arxiu PDF"
               DataSource      =   "clients"
               Height          =   255
               Index           =   63
               Left            =   8610
               TabIndex        =   307
               Top             =   2745
               Width           =   1035
            End
            Begin VB.Label Label1 
               Caption         =   "Arxiu Impressora:"
               DataSource      =   "clients"
               Height          =   255
               Index           =   64
               Left            =   8385
               TabIndex        =   306
               Top             =   3285
               Width           =   1395
            End
            Begin VB.Label Label1 
               Caption         =   "7ª Tinta A:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   45
               Left            =   105
               TabIndex        =   335
               Top             =   1770
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "8ª Tinta A:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   44
               Left            =   105
               TabIndex        =   334
               Top             =   2025
               Width           =   750
            End
         End
         Begin VB.Frame sol 
            Caption         =   "Soldadora"
            Height          =   3255
            Left            =   120
            TabIndex        =   201
            Top             =   17280
            Width           =   9885
            Begin VB.TextBox Text135 
               DataField       =   "troquel"
               DataSource      =   "data1"
               Height          =   285
               Left            =   975
               TabIndex        =   219
               Top             =   1110
               Width           =   630
            End
            Begin VB.TextBox Text134 
               DataField       =   "ansa"
               DataSource      =   "data1"
               Height          =   285
               Left            =   975
               TabIndex        =   220
               Top             =   1410
               Width           =   630
            End
            Begin VB.TextBox Text133 
               DataField       =   "cinta"
               DataSource      =   "data1"
               Height          =   285
               Left            =   975
               TabIndex        =   221
               Top             =   1710
               Width           =   630
            End
            Begin VB.ComboBox Combo11 
               DataField       =   "microperforatsol"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":14A4
               Left            =   8490
               List            =   "comandes.frx":14AE
               TabIndex        =   215
               Top             =   975
               Width           =   585
            End
            Begin VB.TextBox Text132 
               DataField       =   "unitatsxcaixa"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9000
               TabIndex        =   224
               Top             =   2340
               Width           =   675
            End
            Begin VB.TextBox Text131 
               DataField       =   "unitatsxpaquet"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8100
               TabIndex        =   223
               Top             =   2340
               Width           =   675
            End
            Begin VB.TextBox Text130 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6600
               Locked          =   -1  'True
               TabIndex        =   257
               TabStop         =   0   'False
               Top             =   1695
               Width           =   1455
            End
            Begin VB.TextBox Text129 
               DataField       =   "tipusoldadura"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6210
               TabIndex        =   222
               Top             =   1710
               Width           =   330
            End
            Begin VB.TextBox Text128 
               DataField       =   "unitatespsol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8415
               TabIndex        =   254
               Top             =   120
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.TextBox Text127 
               DataSource      =   "data1"
               Height          =   285
               Left            =   7815
               Locked          =   -1  'True
               TabIndex        =   210
               Top             =   390
               Width           =   930
            End
            Begin VB.TextBox Text126 
               DataField       =   "espessorsol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6990
               TabIndex        =   209
               Top             =   405
               Width           =   780
            End
            Begin VB.TextBox Text125 
               DataField       =   "fuellebocasol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6165
               TabIndex        =   208
               Top             =   405
               Width           =   780
            End
            Begin VB.TextBox Text124 
               DataField       =   "fuellebasesol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   5355
               TabIndex        =   207
               Top             =   405
               Width           =   735
            End
            Begin VB.TextBox Text123 
               DataField       =   "solapasol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4545
               TabIndex        =   206
               Top             =   405
               Width           =   780
            End
            Begin VB.TextBox Text122 
               DataField       =   "longitudsol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3735
               TabIndex        =   205
               Top             =   405
               Width           =   735
            End
            Begin VB.TextBox Text121 
               DataField       =   "amplesol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   2100
               TabIndex        =   203
               Top             =   405
               Width           =   735
            End
            Begin VB.ComboBox Combo15 
               DataField       =   "simulteneitatsol"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":14B8
               Left            =   7620
               List            =   "comandes.frx":14CB
               TabIndex        =   214
               Top             =   960
               Width           =   675
            End
            Begin VB.TextBox Text120 
               DataField       =   "soldadora"
               DataSource      =   "data1"
               Height          =   285
               Left            =   960
               TabIndex        =   212
               Top             =   810
               Width           =   630
            End
            Begin VB.TextBox Text119 
               DataField       =   "ampleplegsol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   2910
               TabIndex        =   204
               Top             =   405
               Width           =   780
            End
            Begin VB.ComboBox Combo14 
               DataField       =   "costatobertsol"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":14DE
               Left            =   9120
               List            =   "comandes.frx":14EE
               TabIndex        =   216
               Top             =   975
               Width           =   615
            End
            Begin VB.ComboBox Combo13 
               DataField       =   "microperforatsol"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":14FE
               Left            =   8250
               List            =   "comandes.frx":1508
               TabIndex        =   232
               Top             =   -6285
               Width           =   585
            End
            Begin VB.TextBox Text118 
               DataField       =   "cantitatsol"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8790
               TabIndex        =   211
               Top             =   390
               Width           =   720
            End
            Begin VB.TextBox Text117 
               DataField       =   "numtaladros"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8100
               TabIndex        =   217
               Top             =   1695
               Width           =   675
            End
            Begin VB.TextBox Text116 
               DataField       =   "diametremm"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8985
               TabIndex        =   218
               Top             =   1710
               Width           =   675
            End
            Begin VB.TextBox Text115 
               DataField       =   "tac"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6810
               TabIndex        =   213
               Top             =   975
               Width           =   660
            End
            Begin VB.TextBox Text114 
               DataField       =   "obssol2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   540
               TabIndex        =   226
               Top             =   2265
               Width           =   7500
            End
            Begin VB.TextBox Text113 
               DataField       =   "obssol1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   540
               TabIndex        =   225
               Top             =   2040
               Width           =   7500
            End
            Begin VB.TextBox Text112 
               DataField       =   "obslamgen2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   -7875
               TabIndex        =   231
               Top             =   -10785
               Width           =   6660
            End
            Begin VB.TextBox Text111 
               DataField       =   "arxiusol"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   8085
               Locked          =   -1  'True
               MouseIcon       =   "comandes.frx":1512
               MousePointer    =   99  'Custom
               TabIndex        =   230
               TabStop         =   0   'False
               Top             =   2895
               Width           =   1410
            End
            Begin VB.CommandButton Command7 
               Height          =   285
               Left            =   9525
               Picture         =   "comandes.frx":1954
               Style           =   1  'Graphical
               TabIndex        =   229
               TabStop         =   0   'False
               Top             =   2880
               Width           =   285
            End
            Begin VB.TextBox Text88 
               DataField       =   "obssolgen2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   540
               TabIndex        =   228
               Top             =   2895
               Width           =   7500
            End
            Begin VB.TextBox Text17 
               DataField       =   "obssolgen1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   540
               TabIndex        =   227
               Top             =   2625
               Width           =   7500
            End
            Begin VB.ComboBox Combo10 
               DataField       =   "migelaboratsol"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":1D1A
               Left            =   1245
               List            =   "comandes.frx":1D27
               TabIndex        =   202
               Top             =   375
               Width           =   645
            End
            Begin VB.Label Label1 
               Caption         =   "Troquel:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   128
               Left            =   105
               TabIndex        =   264
               Top             =   1185
               Width           =   705
            End
            Begin VB.Label Label1 
               Caption         =   "Ansa:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   127
               Left            =   105
               TabIndex        =   263
               Top             =   1470
               Width           =   705
            End
            Begin VB.Label Label1 
               Caption         =   "Cinta:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   107
               Left            =   105
               TabIndex        =   262
               Top             =   1785
               Width           =   705
            End
            Begin VB.Label truquel 
               Caption         =   "Truquel"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   1710
               TabIndex        =   261
               Top             =   1185
               Width           =   4500
            End
            Begin VB.Label ansa 
               Caption         =   "ansa"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   1710
               TabIndex        =   260
               Top             =   1485
               Width           =   4500
            End
            Begin VB.Label cinta 
               Caption         =   "Cinta"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   1710
               TabIndex        =   259
               Top             =   1785
               Width           =   4500
            End
            Begin VB.Label Label1 
               Caption         =   "Un. Caixa"
               DataSource      =   "data1"
               Height          =   270
               Index           =   126
               Left            =   8985
               TabIndex        =   258
               Top             =   2085
               Width           =   780
            End
            Begin VB.Label Label1 
               Caption         =   "Tipus Soldadura:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   125
               Left            =   6390
               TabIndex        =   256
               Top             =   1470
               Width           =   1605
            End
            Begin VB.Label Label1 
               Caption         =   "Mesura:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   124
               Left            =   7935
               TabIndex        =   255
               Top             =   165
               Width           =   690
            End
            Begin VB.Label Label1 
               Caption         =   "Espessor:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   123
               Left            =   7050
               TabIndex        =   253
               Top             =   180
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "Fuelle Bo:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   122
               Left            =   6225
               TabIndex        =   252
               Top             =   180
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "Fuelle Ba:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   121
               Left            =   5400
               TabIndex        =   251
               Top             =   180
               Width           =   810
            End
            Begin VB.Label Label1 
               Caption         =   "Solapa:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   120
               Left            =   4695
               TabIndex        =   250
               Top             =   180
               Width           =   630
            End
            Begin VB.Label Label1 
               Caption         =   "Longitud:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   119
               Left            =   3780
               TabIndex        =   249
               Top             =   180
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "B/L/F/BB:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   118
               Left            =   1185
               TabIndex        =   248
               Top             =   150
               Width           =   1005
            End
            Begin VB.Label Label1 
               Caption         =   "Ample:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   117
               Left            =   2250
               TabIndex        =   247
               Top             =   165
               Width           =   555
            End
            Begin VB.Label Label1 
               Caption         =   "Simultaneitat:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   116
               Left            =   7470
               TabIndex        =   246
               Top             =   765
               Width           =   975
            End
            Begin VB.Label nomsoldadora 
               Caption         =   "nomsoldadora"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   1680
               TabIndex        =   245
               Top             =   900
               Width           =   4500
            End
            Begin VB.Label Label1 
               Caption         =   "Soldadora:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   115
               Left            =   120
               TabIndex        =   244
               Top             =   885
               Width           =   1035
            End
            Begin VB.Label Label1 
               Caption         =   "Plegat:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   114
               Left            =   3060
               TabIndex        =   243
               Top             =   180
               Width           =   630
            End
            Begin VB.Label Label1 
               Caption         =   "Quantitat:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   111
               Left            =   8820
               TabIndex        =   240
               Top             =   165
               Width           =   825
            End
            Begin VB.Label Label1 
               Caption         =   "Nº Taladros:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   110
               Left            =   7995
               TabIndex        =   239
               Top             =   1470
               Width           =   930
            End
            Begin VB.Label Label1 
               Caption         =   "Diam. m/m"
               DataSource      =   "data1"
               Height          =   270
               Index           =   109
               Left            =   8910
               TabIndex        =   238
               Top             =   1485
               Width           =   1020
            End
            Begin VB.Label Label1 
               Caption         =   "TAC:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   108
               Left            =   6915
               TabIndex        =   237
               Top             =   765
               Width           =   480
            End
            Begin VB.Label Label1 
               Caption         =   "Un. Paquet:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   106
               Left            =   8085
               TabIndex        =   236
               Top             =   2085
               Width           =   1005
            End
            Begin VB.Label Label1 
               Caption         =   "Obs. Sold"
               DataSource      =   "data1"
               Height          =   480
               Index           =   105
               Left            =   105
               TabIndex        =   235
               Top             =   2100
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Obs. Client"
               DataSource      =   "data1"
               Height          =   525
               Index           =   104
               Left            =   105
               TabIndex        =   234
               Top             =   2685
               Width           =   525
            End
            Begin VB.Label Label1 
               Caption         =   "Arxiu Soldadora:"
               DataSource      =   "clients"
               Height          =   255
               Index           =   103
               Left            =   8205
               TabIndex        =   233
               Top             =   2685
               Width           =   1290
            End
            Begin VB.Label Label1 
               Caption         =   "C. Obert:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   113
               Left            =   9090
               TabIndex        =   242
               Top             =   765
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "MicroP:"
               DataSource      =   "data1"
               Height          =   210
               Index           =   112
               Left            =   8490
               TabIndex        =   241
               Top             =   765
               Width           =   630
            End
         End
         Begin VB.Frame reb 
            Caption         =   "Rebobinadora"
            Height          =   2655
            Left            =   120
            TabIndex        =   161
            Top             =   14580
            Width           =   9885
            Begin VB.ComboBox Combo9 
               DataField       =   "migelaborat"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":1D35
               Left            =   1245
               List            =   "comandes.frx":1D42
               TabIndex        =   162
               Top             =   345
               Width           =   645
            End
            Begin VB.TextBox Text108 
               DataField       =   "obsrebgen1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   585
               TabIndex        =   178
               Top             =   2055
               Width           =   7500
            End
            Begin VB.TextBox Text110 
               DataField       =   "obsrebgen2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   585
               TabIndex        =   179
               Top             =   2325
               Width           =   7500
            End
            Begin VB.CommandButton Command5 
               Height          =   285
               Left            =   9555
               Picture         =   "comandes.frx":1D50
               Style           =   1  'Graphical
               TabIndex        =   198
               TabStop         =   0   'False
               Top             =   1725
               Width           =   285
            End
            Begin VB.TextBox Text109 
               DataField       =   "arxiureb"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   8130
               Locked          =   -1  'True
               MouseIcon       =   "comandes.frx":2116
               MousePointer    =   99  'Custom
               TabIndex        =   175
               TabStop         =   0   'False
               Top             =   1725
               Width           =   1380
            End
            Begin VB.TextBox Text107 
               DataField       =   "obslamgen2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   -7875
               TabIndex        =   195
               Top             =   -10785
               Width           =   6660
            End
            Begin VB.TextBox Text106 
               DataField       =   "obsreb1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   585
               TabIndex        =   176
               Top             =   1470
               Width           =   7500
            End
            Begin VB.TextBox Text105 
               DataField       =   "obsreb2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   585
               TabIndex        =   177
               Top             =   1695
               Width           =   7500
            End
            Begin VB.ComboBox Combo7 
               DataField       =   "etiqintcanutu"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":2558
               Left            =   6255
               List            =   "comandes.frx":2562
               TabIndex        =   173
               Top             =   1065
               Width           =   615
            End
            Begin VB.ComboBox Combo6 
               DataField       =   "etiqintcanutu"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":256C
               Left            =   7950
               List            =   "comandes.frx":2576
               TabIndex        =   174
               Top             =   1050
               Width           =   585
            End
            Begin VB.TextBox Text104 
               DataField       =   "diamextbob"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4410
               TabIndex        =   172
               Top             =   1095
               Width           =   660
            End
            Begin VB.TextBox Text103 
               DataField       =   "mtrslinbob"
               DataSource      =   "data1"
               Height          =   285
               Left            =   2970
               TabIndex        =   171
               Top             =   1095
               Width           =   675
            End
            Begin VB.TextBox Text102 
               DataField       =   "kilosbob"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1260
               TabIndex        =   170
               Top             =   1095
               Width           =   675
            End
            Begin VB.TextBox Text101 
               DataField       =   "tubbase"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9360
               TabIndex        =   169
               Top             =   705
               Width           =   420
            End
            Begin VB.ComboBox Combo5 
               DataField       =   "microperforat"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":2580
               Left            =   7935
               List            =   "comandes.frx":258A
               TabIndex        =   168
               Top             =   705
               Width           =   585
            End
            Begin VB.ComboBox Combo4 
               DataField       =   "caratractada"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":2594
               Left            =   6240
               List            =   "comandes.frx":25A1
               TabIndex        =   167
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox Text100 
               DataField       =   "matintbob"
               DataSource      =   "data1"
               Height          =   285
               Left            =   5145
               TabIndex        =   165
               Top             =   405
               Width           =   930
            End
            Begin VB.TextBox Text99 
               DataField       =   "rebobinadora"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1260
               TabIndex        =   166
               Top             =   720
               Width           =   600
            End
            Begin VB.ComboBox Combo3 
               DataField       =   "simulteneitatreb"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":25AD
               Left            =   4035
               List            =   "comandes.frx":25C0
               TabIndex        =   164
               Top             =   405
               Width           =   675
            End
            Begin VB.TextBox Text98 
               DataField       =   "amplereb"
               DataSource      =   "data1"
               Height          =   285
               Left            =   2475
               TabIndex        =   163
               Top             =   405
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Arxiu Rebobinadora:"
               DataSource      =   "clients"
               Height          =   255
               Index           =   102
               Left            =   8235
               TabIndex        =   199
               Top             =   1485
               Width           =   1710
            End
            Begin VB.Label Label1 
               Caption         =   "Obs. Client"
               DataSource      =   "data1"
               Height          =   480
               Index           =   101
               Left            =   105
               TabIndex        =   197
               Top             =   2115
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Obs. Reb."
               DataSource      =   "data1"
               Height          =   480
               Index           =   100
               Left            =   105
               TabIndex        =   196
               Top             =   1530
               Width           =   465
            End
            Begin VB.Label Label1 
               Caption         =   "Et. Int. Canutu:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   99
               Left            =   5145
               TabIndex        =   194
               Top             =   1140
               Width           =   1170
            End
            Begin VB.Label Label1 
               Caption         =   "Et. Ext. Bob:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   98
               Left            =   6945
               TabIndex        =   193
               Top             =   1140
               Width           =   1050
            End
            Begin VB.Label Label1 
               Caption         =   "Diametre:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   97
               Left            =   3705
               TabIndex        =   192
               Top             =   1125
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Mtrs Bobina:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   96
               Left            =   2040
               TabIndex        =   191
               Top             =   1140
               Width           =   1020
            End
            Begin VB.Label Label1 
               Caption         =   "Kilos Bobina:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   95
               Left            =   120
               TabIndex        =   190
               Top             =   1140
               Width           =   1020
            End
            Begin VB.Label Label1 
               Caption         =   "Tub Base:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   270
               Index           =   94
               Left            =   8580
               TabIndex        =   189
               Top             =   765
               Width           =   825
            End
            Begin VB.Label Label1 
               Caption         =   "Microperforat:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   93
               Left            =   6930
               TabIndex        =   188
               Top             =   795
               Width           =   1170
            End
            Begin VB.Label Label1 
               Caption         =   "Cara Tractada:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   92
               Left            =   5130
               TabIndex        =   187
               Top             =   795
               Width           =   1170
            End
            Begin VB.Label Label1 
               Caption         =   "Lot Mat. Int. Bob."
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   91
               Left            =   5160
               TabIndex        =   186
               Top             =   180
               Width           =   1725
            End
            Begin VB.Label desclot1 
               Caption         =   "descripcio del lot1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   1
               Left            =   6150
               TabIndex        =   185
               Top             =   435
               Width           =   3570
            End
            Begin VB.Label Label1 
               Caption         =   "Rebobinadora:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   90
               Left            =   105
               TabIndex        =   184
               Top             =   795
               Width           =   1125
            End
            Begin VB.Label nomrebobinadora 
               Caption         =   "nomrebobinadora"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   1
               Left            =   1875
               TabIndex        =   183
               Top             =   810
               Width           =   3045
            End
            Begin VB.Label Label1 
               Caption         =   "Simultaneitat:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   89
               Left            =   3885
               TabIndex        =   182
               Top             =   180
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Ample Reb:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   88
               Left            =   2460
               TabIndex        =   181
               Top             =   180
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Tubo o Lam:"
               DataSource      =   "data1"
               Height          =   270
               Index           =   87
               Left            =   1140
               TabIndex        =   180
               Top             =   150
               Width           =   1125
            End
         End
         Begin VB.Frame lam1 
            Caption         =   "Laminadora-1"
            Height          =   4110
            Left            =   90
            TabIndex        =   103
            Top             =   10470
            Width           =   9885
            Begin VB.ComboBox Combo2 
               DataField       =   "simulteneitatlam"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":25D3
               Left            =   8910
               List            =   "comandes.frx":25E6
               TabIndex        =   113
               Top             =   780
               Width           =   675
            End
            Begin VB.TextBox Text97 
               DataField       =   "arxiulaminadora"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   8160
               Locked          =   -1  'True
               MouseIcon       =   "comandes.frx":25F9
               MousePointer    =   99  'Custom
               TabIndex        =   125
               TabStop         =   0   'False
               Top             =   3165
               Width           =   1320
            End
            Begin VB.CommandButton Command4 
               Height          =   285
               Left            =   9525
               Picture         =   "comandes.frx":2A3B
               Style           =   1  'Graphical
               TabIndex        =   116
               TabStop         =   0   'False
               Top             =   3165
               Width           =   285
            End
            Begin VB.TextBox Text96 
               DataField       =   "obslam2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   480
               TabIndex        =   123
               Top             =   3180
               Width           =   7500
            End
            Begin VB.TextBox Text95 
               DataField       =   "obslam1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   480
               TabIndex        =   122
               Top             =   2940
               Width           =   7500
            End
            Begin VB.TextBox Text94 
               DataField       =   "obslamgen2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   480
               TabIndex        =   126
               Top             =   3780
               Width           =   7500
            End
            Begin VB.TextBox Text93 
               DataField       =   "obslamgen1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   480
               TabIndex        =   124
               Top             =   3540
               Width           =   7500
            End
            Begin VB.TextBox Text92 
               DataField       =   "mtr/minrodillocola"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8895
               TabIndex        =   117
               Top             =   1710
               Width           =   645
            End
            Begin VB.TextBox Text90 
               DataField       =   "rodillocola"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8895
               TabIndex        =   115
               Top             =   1410
               Width           =   435
            End
            Begin VB.TextBox grmt2 
               DataField       =   "grmt2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   7050
               TabIndex        =   121
               Top             =   1440
               Width           =   600
            End
            Begin VB.TextBox vadhesiu 
               DataField       =   "tipusadhesiu"
               DataSource      =   "data1"
               Height          =   285
               Left            =   405
               TabIndex        =   151
               Top             =   1305
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.TextBox pes2 
               DataField       =   "pes2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4800
               TabIndex        =   120
               Top             =   1590
               Width           =   600
            End
            Begin VB.TextBox pes1 
               DataField       =   "pes1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   4800
               TabIndex        =   119
               Top             =   1290
               Width           =   600
            End
            Begin MSFlexGridLib.MSFlexGrid reixaconsums 
               Height          =   870
               Left            =   150
               TabIndex        =   150
               TabStop         =   0   'False
               Tag             =   "1"
               Top             =   2025
               Width           =   9645
               _ExtentX        =   17013
               _ExtentY        =   1535
               _Version        =   327680
               Rows            =   3
               Cols            =   16
               FixedCols       =   0
               BackColor       =   16777215
               ForeColor       =   0
               ForeColorFixed  =   16711680
               ForeColorSel    =   0
               AllowBigSelection=   0   'False
               TextStyle       =   3
               ScrollBars      =   0
            End
            Begin VB.TextBox adhesiu 
               Height          =   285
               Left            =   1035
               TabIndex        =   118
               Top             =   1290
               Width           =   3000
            End
            Begin VB.TextBox Text91 
               DataField       =   "camisa"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8895
               TabIndex        =   114
               Top             =   1110
               Width           =   645
            End
            Begin VB.TextBox Text89 
               DataField       =   "amplelaminar"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8910
               TabIndex        =   112
               Top             =   495
               Width           =   885
            End
            Begin VB.TextBox Text87 
               DataField       =   "ampleutil"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8910
               TabIndex        =   111
               Top             =   195
               Width           =   885
            End
            Begin VB.TextBox Text86 
               DataField       =   "tensiototal"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6945
               TabIndex        =   110
               Top             =   795
               Width           =   930
            End
            Begin VB.TextBox Text85 
               DataField       =   "tensiodesb2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6945
               TabIndex        =   109
               Top             =   495
               Width           =   930
            End
            Begin VB.TextBox Text84 
               DataField       =   "tensiodesb1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6945
               TabIndex        =   108
               Top             =   195
               Width           =   930
            End
            Begin VB.TextBox Text83 
               DataField       =   "mtr/minmaquina"
               DataSource      =   "data1"
               Height          =   285
               Left            =   5085
               TabIndex        =   107
               Top             =   795
               Width           =   630
            End
            Begin VB.TextBox Text82 
               DataField       =   "laminadora"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1050
               TabIndex        =   106
               Top             =   795
               Width           =   630
            End
            Begin VB.TextBox Text81 
               DataField       =   "lotmatdesb2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1050
               TabIndex        =   105
               Top             =   495
               Width           =   930
            End
            Begin VB.TextBox Text80 
               DataField       =   "lotmatdesb1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1050
               TabIndex        =   104
               Top             =   210
               Width           =   930
            End
            Begin VB.Label desclot2 
               Caption         =   "descripcio del lot2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   1
               Left            =   2070
               TabIndex        =   160
               Top             =   555
               Width           =   3630
            End
            Begin VB.Label desclot1 
               Caption         =   "descripcio del lot1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   2070
               TabIndex        =   159
               Top             =   210
               Width           =   3690
            End
            Begin VB.Label Label1 
               Caption         =   "Arxiu Laminadora:"
               DataSource      =   "clients"
               Height          =   255
               Index           =   86
               Left            =   8235
               TabIndex        =   158
               Top             =   2955
               Width           =   1365
            End
            Begin VB.Label Label1 
               Caption         =   "Obs. Lam"
               DataSource      =   "data1"
               Height          =   480
               Index           =   85
               Left            =   45
               TabIndex        =   157
               Top             =   2970
               Width           =   435
            End
            Begin VB.Label Label1 
               Caption         =   "Obs.  Client"
               DataSource      =   "data1"
               Height          =   480
               Index           =   84
               Left            =   45
               TabIndex        =   156
               Top             =   3555
               Width           =   510
            End
            Begin VB.Label Label1 
               Caption         =   "Mtrs/Min Rodillo:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   83
               Left            =   7680
               TabIndex        =   155
               Top             =   1770
               Width           =   1350
            End
            Begin VB.Label Label1 
               Caption         =   "Rodillo Cola:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   82
               Left            =   7995
               TabIndex        =   154
               Top             =   1470
               Width           =   960
            End
            Begin VB.Label litres2 
               Alignment       =   2  'Center
               Caption         =   "Litres2"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   225
               Index           =   2
               Left            =   6075
               TabIndex        =   153
               Top             =   1665
               Width           =   390
            End
            Begin VB.Label litres1 
               Alignment       =   2  'Center
               Caption         =   "litres1"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   225
               Index           =   1
               Left            =   6075
               TabIndex        =   152
               Top             =   1350
               Width           =   390
            End
            Begin VB.Label ºC2 
               Alignment       =   2  'Center
               Caption         =   "ºC2"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   270
               Index           =   0
               Left            =   5415
               TabIndex        =   149
               Top             =   1650
               Width           =   390
            End
            Begin VB.Label grcm2 
               Alignment       =   2  'Center
               Caption         =   "Grcm2"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   270
               Index           =   0
               Left            =   4170
               TabIndex        =   148
               Top             =   1650
               Width           =   570
            End
            Begin VB.Label Label1 
               Caption         =   "Cola Gr/mt2:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   81
               Left            =   6915
               TabIndex        =   147
               Top             =   1245
               Width           =   1020
            End
            Begin VB.Label ºC1 
               Alignment       =   2  'Center
               Caption         =   "ºC1"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   225
               Index           =   0
               Left            =   5400
               TabIndex        =   146
               Top             =   1335
               Width           =   390
            End
            Begin VB.Label grcm1 
               Alignment       =   2  'Center
               Caption         =   "Grcm1"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   270
               Index           =   0
               Left            =   4170
               TabIndex        =   145
               Top             =   1335
               Width           =   570
            End
            Begin VB.Label Label1 
               Caption         =   "% Pes"
               DataSource      =   "data1"
               Height          =   255
               Index           =   80
               Left            =   4875
               TabIndex        =   144
               Top             =   1095
               Width           =   540
            End
            Begin VB.Label Label1 
               Caption         =   "%Litres"
               DataSource      =   "data1"
               Height          =   255
               Index           =   79
               Left            =   5985
               TabIndex        =   143
               Top             =   1095
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "ºC"
               DataSource      =   "data1"
               Height          =   255
               Index           =   78
               Left            =   5505
               TabIndex        =   142
               Top             =   1095
               Width           =   255
            End
            Begin VB.Label Label1 
               Caption         =   "Gr.Cm2"
               DataSource      =   "data1"
               Height          =   255
               Index           =   77
               Left            =   4200
               TabIndex        =   141
               Top             =   1095
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Camisa:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   76
               Left            =   8100
               TabIndex        =   140
               Top             =   1170
               Width           =   765
            End
            Begin VB.Label enduridor 
               Caption         =   "DESCRIPCIO DE L'ENDURIDOR"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   270
               Index           =   0
               Left            =   1065
               TabIndex        =   139
               Top             =   1635
               Width           =   3000
            End
            Begin VB.Label Label1 
               Caption         =   "Descripció Adhesiu i Enduridor  (F2)"
               DataSource      =   "data1"
               Height          =   255
               Index           =   75
               Left            =   1260
               TabIndex        =   138
               Top             =   1095
               Width           =   2610
            End
            Begin VB.Label Label1 
               Caption         =   "Ample Lam.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   74
               Left            =   7935
               TabIndex        =   137
               Top             =   540
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Simultaneitat:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   73
               Left            =   7935
               TabIndex        =   136
               Top             =   825
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Ample Útil:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   72
               Left            =   7935
               TabIndex        =   135
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Tensió Total:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   71
               Left            =   5775
               TabIndex        =   134
               Top             =   840
               Width           =   1155
            End
            Begin VB.Label Label1 
               Caption         =   "Tensió Desb. 2:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   70
               Left            =   5775
               TabIndex        =   133
               Top             =   555
               Width           =   1410
            End
            Begin VB.Label Label1 
               Caption         =   "Tensió Desb. 1:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   69
               Left            =   5775
               TabIndex        =   132
               Top             =   255
               Width           =   1410
            End
            Begin VB.Label Label1 
               Caption         =   "Mtrs/Min.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   68
               Left            =   4350
               TabIndex        =   131
               Top             =   855
               Width           =   915
            End
            Begin VB.Label nomlaminadora 
               Caption         =   "nomlaminadora"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   1815
               TabIndex        =   130
               Top             =   870
               Width           =   2505
            End
            Begin VB.Label Label1 
               Caption         =   "Laminadora:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   67
               Left            =   135
               TabIndex        =   129
               Top             =   855
               Width           =   1005
            End
            Begin VB.Label Label1 
               Caption         =   "Lot Desb 2:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   66
               Left            =   135
               TabIndex        =   128
               Top             =   555
               Width           =   900
            End
            Begin VB.Label Label1 
               Caption         =   "Lot Desb 1:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   65
               Left            =   135
               TabIndex        =   127
               Top             =   255
               Width           =   840
            End
         End
         Begin VB.Frame ext 
            Caption         =   "Extrussora"
            Height          =   2940
            Left            =   90
            TabIndex        =   49
            Top             =   3585
            Width           =   9915
            Begin VB.ComboBox Combo8 
               DataField       =   "tubolam"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":2E01
               Left            =   1125
               List            =   "comandes.frx":2E0E
               TabIndex        =   52
               Top             =   180
               Width           =   630
            End
            Begin VB.ComboBox Combo1 
               DataField       =   "simulteneitat"
               DataSource      =   "data1"
               Height          =   315
               ItemData        =   "comandes.frx":2E1C
               Left            =   8460
               List            =   "comandes.frx":2E2F
               TabIndex        =   65
               Top             =   1080
               Width           =   720
            End
            Begin VB.TextBox Text39 
               DataField       =   "refilate"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8475
               TabIndex        =   59
               Top             =   225
               Width           =   930
            End
            Begin VB.TextBox Text38 
               DataField       =   "refilatd"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6825
               TabIndex        =   58
               Top             =   225
               Width           =   855
            End
            Begin VB.TextBox Text42 
               DataField       =   "arxiuext"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   8145
               Locked          =   -1  'True
               MouseIcon       =   "comandes.frx":2E42
               MousePointer    =   99  'Custom
               TabIndex        =   94
               TabStop         =   0   'False
               Top             =   1410
               Width           =   1365
            End
            Begin VB.CommandButton Command2 
               Height          =   315
               Left            =   9570
               Picture         =   "comandes.frx":3284
               Style           =   1  'Graphical
               TabIndex        =   71
               Top             =   1410
               Width           =   315
            End
            Begin VB.TextBox Text37 
               DataField       =   "obsext2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1140
               TabIndex        =   91
               Top             =   1965
               Width           =   7500
            End
            Begin VB.TextBox Text36 
               DataField       =   "obsext1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1140
               TabIndex        =   72
               Top             =   1725
               Width           =   7500
            End
            Begin VB.TextBox Text35 
               DataField       =   "obsextgen2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1140
               TabIndex        =   90
               Top             =   2565
               Width           =   7500
            End
            Begin VB.TextBox Text34 
               DataField       =   "obsextgen1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1140
               TabIndex        =   89
               Top             =   2325
               Width           =   7500
            End
            Begin VB.TextBox Text33 
               DataField       =   "pes1000mtrs"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6825
               TabIndex        =   64
               Top             =   1110
               Width           =   855
            End
            Begin VB.TextBox Text31 
               DataField       =   "mesuracantex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8085
               TabIndex        =   87
               Top             =   840
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.TextBox Text30 
               DataSource      =   "data1"
               Height          =   285
               Left            =   8475
               Locked          =   -1  'True
               TabIndex        =   63
               Top             =   810
               Width           =   930
            End
            Begin VB.TextBox Text29 
               DataField       =   "cantitatex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6825
               TabIndex        =   62
               Top             =   810
               Width           =   855
            End
            Begin VB.TextBox Text28 
               DataField       =   "kghora"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6810
               TabIndex        =   70
               Top             =   1425
               Width           =   630
            End
            Begin VB.TextBox Text27 
               DataField       =   "extrusora"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1125
               TabIndex        =   69
               Top             =   1425
               Width           =   630
            End
            Begin VB.TextBox Text26 
               DataField       =   "aditiuex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1125
               TabIndex        =   68
               Top             =   1125
               Width           =   630
            End
            Begin VB.TextBox Text25 
               DataField       =   "materialex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1125
               TabIndex        =   67
               Top             =   825
               Width           =   630
            End
            Begin VB.TextBox Text24 
               DataField       =   "colorex"
               DataSource      =   "data1"
               Height          =   285
               Left            =   1140
               TabIndex        =   66
               Top             =   525
               Width           =   630
            End
            Begin VB.TextBox Text22 
               DataSource      =   "data1"
               Height          =   285
               Left            =   8475
               Locked          =   -1  'True
               TabIndex        =   61
               Top             =   510
               Width           =   930
            End
            Begin VB.TextBox Text21 
               DataField       =   "espessor"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6825
               TabIndex        =   60
               Top             =   510
               Width           =   855
            End
            Begin VB.TextBox Text20 
               DataField       =   "solapa"
               DataSource      =   "data1"
               Height          =   285
               Left            =   5205
               TabIndex        =   57
               Top             =   225
               Width           =   795
            End
            Begin VB.TextBox Text19 
               DataField       =   "plegatesq"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3675
               TabIndex        =   56
               Top             =   225
               Width           =   855
            End
            Begin VB.TextBox Text18 
               DataField       =   "ampleesq"
               DataSource      =   "data1"
               Height          =   285
               Left            =   2685
               TabIndex        =   54
               Top             =   225
               Width           =   780
            End
            Begin VB.TextBox Text23 
               DataField       =   "mesuraesp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8070
               TabIndex        =   74
               Top             =   525
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label Label1 
               Caption         =   "Simult."
               DataSource      =   "data1"
               Height          =   255
               Index           =   33
               Left            =   7785
               TabIndex        =   98
               Top             =   1185
               Width           =   570
            End
            Begin VB.Label Label1 
               Caption         =   "Ref.Esq.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   32
               Left            =   7800
               TabIndex        =   97
               Top             =   300
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "Ref.Dret:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   31
               Left            =   6075
               TabIndex        =   96
               Top             =   300
               Width           =   750
            End
            Begin VB.Label Label1 
               Caption         =   "Arxiu:"
               DataSource      =   "clients"
               Height          =   255
               Index           =   30
               Left            =   7710
               TabIndex        =   95
               Top             =   1485
               Width           =   525
            End
            Begin VB.Label Label1 
               Caption         =   "Obs.  Extrussora"
               DataSource      =   "data1"
               Height          =   480
               Index           =   29
               Left            =   75
               TabIndex        =   93
               Top             =   1800
               Width           =   840
            End
            Begin VB.Label Label1 
               Caption         =   "Obs. Ext. del Client"
               DataSource      =   "data1"
               Height          =   480
               Index           =   28
               Left            =   75
               TabIndex        =   92
               Top             =   2400
               Width           =   1065
            End
            Begin VB.Label Label1 
               Caption         =   "Pesx1000"
               DataSource      =   "data1"
               Height          =   255
               Index           =   27
               Left            =   6075
               TabIndex        =   88
               Top             =   1185
               Width           =   840
            End
            Begin VB.Label Label1 
               Caption         =   "Mesura:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   26
               Left            =   7800
               TabIndex        =   86
               Top             =   885
               Width           =   690
            End
            Begin VB.Label Label1 
               Caption         =   "Quantitat:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   25
               Left            =   6075
               TabIndex        =   85
               Top             =   885
               Width           =   840
            End
            Begin VB.Label Label1 
               Caption         =   "Kgr./Hora."
               DataSource      =   "data1"
               Height          =   255
               Index           =   24
               Left            =   6060
               TabIndex        =   84
               Top             =   1500
               Width           =   915
            End
            Begin VB.Label nomextrussora 
               BackStyle       =   0  'Transparent
               Caption         =   "nomextrussora"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   0
               Left            =   1800
               TabIndex        =   83
               Top             =   1500
               Width           =   4185
            End
            Begin VB.Label Label1 
               Caption         =   "Extrussora:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   23
               Left            =   75
               TabIndex        =   82
               Top             =   1500
               Width           =   915
            End
            Begin VB.Label nomadditiu 
               BackStyle       =   0  'Transparent
               Caption         =   "Additiu:"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   23
               Left            =   1800
               TabIndex        =   81
               Top             =   1200
               Width           =   4185
            End
            Begin VB.Label nommaterial 
               BackStyle       =   0  'Transparent
               Caption         =   "Material:"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   23
               Left            =   1800
               TabIndex        =   80
               Top             =   900
               Width           =   4185
            End
            Begin VB.Label nomcolor 
               BackStyle       =   0  'Transparent
               Caption         =   "Color:"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   23
               Left            =   1800
               TabIndex        =   79
               Top             =   600
               Width           =   4185
            End
            Begin VB.Label Label1 
               Caption         =   "Additiu:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   22
               Left            =   75
               TabIndex        =   78
               Top             =   1200
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "Material:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   21
               Left            =   75
               TabIndex        =   77
               Top             =   900
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "Color:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   20
               Left            =   75
               TabIndex        =   76
               Top             =   600
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "Mesura:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   19
               Left            =   7800
               TabIndex        =   75
               Top             =   585
               Width           =   690
            End
            Begin VB.Label Label1 
               Caption         =   "Espessor:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   18
               Left            =   6075
               TabIndex        =   73
               Top             =   585
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "Solapa:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   16
               Left            =   4620
               TabIndex        =   55
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "/"
               DataSource      =   "data1"
               Height          =   255
               Index           =   14
               Left            =   3525
               TabIndex        =   53
               Top             =   300
               Width           =   165
            End
            Begin VB.Label Label1 
               Caption         =   "Ample/Pleg:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   13
               Left            =   1800
               TabIndex        =   51
               Top             =   300
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "Tubo o Lam:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   12
               Left            =   75
               TabIndex        =   50
               Top             =   300
               Width           =   915
            End
         End
         Begin VB.Frame cap 
            Height          =   3465
            Left            =   90
            TabIndex        =   10
            Top             =   135
            Width           =   9915
            Begin MSMask.MaskEdBox controlmask 
               Height          =   225
               Left            =   5865
               TabIndex        =   337
               Top             =   180
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   397
               _Version        =   327680
               PromptChar      =   "_"
            End
            Begin VB.TextBox Text45 
               DataField       =   "comandaclient"
               DataSource      =   "data1"
               Height          =   285
               Left            =   8865
               TabIndex        =   15
               Top             =   585
               Width           =   960
            End
            Begin VB.TextBox Text44 
               DataField       =   "refclient"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6375
               TabIndex        =   14
               Top             =   540
               Width           =   1530
            End
            Begin VB.Timer Timer2 
               Interval        =   100
               Left            =   570
               Top             =   180
            End
            Begin VB.Timer Timer1 
               Interval        =   1000
               Left            =   30
               Top             =   150
            End
            Begin VB.TextBox Text43 
               DataField       =   "marques"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6375
               TabIndex        =   18
               Top             =   840
               Width           =   1530
            End
            Begin VB.TextBox Text41 
               DataField       =   "numpressupost"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3675
               TabIndex        =   17
               Top             =   825
               Width           =   1530
            End
            Begin VB.TextBox Text15 
               DataField       =   "estatcomanda"
               DataSource      =   "data1"
               Height          =   285
               Left            =   9240
               TabIndex        =   23
               Top             =   1140
               Width           =   555
            End
            Begin VB.TextBox Text14 
               DataField       =   "obspedido2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   990
               TabIndex        =   29
               Top             =   2340
               Width           =   7500
            End
            Begin VB.TextBox Text13 
               DataField       =   "obspedido1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   990
               TabIndex        =   28
               Top             =   2100
               Width           =   7500
            End
            Begin VB.TextBox Text12 
               DataField       =   "obspedgen2"
               DataSource      =   "data1"
               Height          =   285
               Left            =   990
               TabIndex        =   31
               Top             =   2940
               Width           =   7500
            End
            Begin VB.TextBox Text32 
               DataField       =   "obspedgen1"
               DataSource      =   "data1"
               Height          =   285
               Left            =   990
               TabIndex        =   30
               Top             =   2700
               Width           =   7500
            End
            Begin VB.TextBox Text11 
               DataField       =   "tipoentrega"
               DataSource      =   "data1"
               Height          =   285
               Left            =   990
               TabIndex        =   27
               Top             =   1725
               Width           =   555
            End
            Begin VB.TextBox Text10 
               DataField       =   "pvpcliche"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6375
               TabIndex        =   26
               Top             =   1425
               Width           =   1530
            End
            Begin VB.TextBox Text9 
               DataField       =   "datapvprevisat"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3675
               TabIndex        =   25
               Top             =   1425
               Width           =   1530
            End
            Begin VB.TextBox Text8 
               DataField       =   "pvprevisat"
               DataSource      =   "data1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   975
               TabIndex        =   24
               Top             =   1425
               Width           =   1530
            End
            Begin VB.TextBox Text7 
               DataField       =   "mesurapvp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3300
               TabIndex        =   21
               Top             =   1125
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.TextBox Text6 
               DataField       =   "pvp"
               DataSource      =   "data1"
               Height          =   285
               Left            =   975
               TabIndex        =   19
               Top             =   1125
               Width           =   1530
            End
            Begin VB.TextBox Text5 
               DataField       =   "dataentrega"
               DataSource      =   "data1"
               Height          =   285
               Left            =   6375
               TabIndex        =   22
               Top             =   1140
               Width           =   1530
            End
            Begin VB.TextBox Text4 
               DataField       =   "datacomanda"
               DataSource      =   "data1"
               Height          =   285
               Left            =   975
               TabIndex        =   16
               Top             =   825
               Width           =   1530
            End
            Begin VB.TextBox Text3 
               DataField       =   "producte"
               DataSource      =   "data1"
               Height          =   285
               Left            =   975
               TabIndex        =   13
               Top             =   525
               Width           =   855
            End
            Begin VB.TextBox Text2 
               DataField       =   "client"
               DataSource      =   "data1"
               Height          =   285
               Left            =   3075
               TabIndex        =   12
               Top             =   225
               Width           =   1005
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H0080FFFF&
               DataField       =   "comanda"
               DataSource      =   "data1"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H008080FF&
               Height          =   285
               Left            =   975
               TabIndex        =   11
               Top             =   225
               Width           =   1455
            End
            Begin VB.TextBox Text16 
               Height          =   285
               Left            =   3675
               TabIndex        =   20
               Top             =   1125
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "NºCom.Cl.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   37
               Left            =   8040
               TabIndex        =   102
               Top             =   660
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "Ref. Client:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   36
               Left            =   5325
               TabIndex        =   101
               Top             =   615
               Width           =   990
            End
            Begin VB.Label Label1 
               Caption         =   "Marques Pres."
               DataSource      =   "data1"
               Height          =   255
               Index           =   35
               Left            =   5325
               TabIndex        =   100
               Top             =   915
               Width           =   1140
            End
            Begin VB.Label Label1 
               Caption         =   "Núm. Pressup."
               DataSource      =   "data1"
               Height          =   255
               Index           =   34
               Left            =   2625
               TabIndex        =   99
               Top             =   900
               Width           =   1140
            End
            Begin VB.Label Label1 
               Caption         =   "Estat Comanda:"
               DataSource      =   "data1"
               Height          =   330
               Index           =   11
               Left            =   8040
               TabIndex        =   48
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Obs.  de la Comanda"
               DataSource      =   "data1"
               Height          =   480
               Index           =   15
               Left            =   150
               TabIndex        =   47
               Top             =   2100
               Width           =   840
            End
            Begin VB.Label Label1 
               Caption         =   "Obs. del Client"
               DataSource      =   "data1"
               Height          =   480
               Index           =   17
               Left            =   150
               TabIndex        =   46
               Top             =   2700
               Width           =   765
            End
            Begin VB.Label Label3 
               Caption         =   "Descripcio del tipus d'entrega"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   315
               Left            =   1650
               TabIndex        =   45
               Top             =   1800
               Width           =   5940
            End
            Begin VB.Label Label1 
               Caption         =   "T. Entrega:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   10
               Left            =   150
               TabIndex        =   44
               Top             =   1800
               Width           =   1065
            End
            Begin VB.Label Label1 
               Caption         =   "Preu del Clixé:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   9
               Left            =   5325
               TabIndex        =   43
               Top             =   1500
               Width           =   1065
            End
            Begin VB.Label Label1 
               Caption         =   "Data Revisió:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   8
               Left            =   2625
               TabIndex        =   42
               Top             =   1500
               Width           =   1065
            End
            Begin VB.Label Label1 
               Caption         =   "PVP Rent.:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   7
               Left            =   150
               TabIndex        =   41
               Top             =   1500
               Width           =   930
            End
            Begin VB.Label Label1 
               Caption         =   "Mesura:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   6
               Left            =   2625
               TabIndex        =   40
               Top             =   1200
               Width           =   690
            End
            Begin VB.Label Label1 
               Caption         =   "PVP €:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   4
               Left            =   150
               TabIndex        =   39
               Top             =   1200
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "Data Entrega:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   3
               Left            =   5325
               TabIndex        =   38
               Top             =   1215
               Width           =   1140
            End
            Begin VB.Label Label1 
               Caption         =   "Data Com:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   2
               Left            =   150
               TabIndex        =   37
               Top             =   900
               Width           =   765
            End
            Begin VB.Label nomproducte 
               Caption         =   "Descripcio del producte"
               DataSource      =   "data1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   240
               Left            =   1950
               TabIndex        =   36
               Top             =   600
               Width           =   3285
            End
            Begin VB.Label Label1 
               Caption         =   "Producte:"
               DataSource      =   "data1"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   1
               Left            =   150
               TabIndex        =   35
               Top             =   600
               Width           =   765
            End
            Begin VB.Label nomclient 
               Caption         =   "Nom del client"
               DataSource      =   "data1"
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
               Height          =   315
               Left            =   4200
               TabIndex        =   34
               Top             =   225
               Width           =   5190
            End
            Begin VB.Label Label1 
               Caption         =   "Client"
               DataSource      =   "data1"
               Height          =   180
               Index           =   0
               Left            =   2475
               TabIndex        =   33
               Top             =   300
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "Comanda:"
               DataSource      =   "data1"
               Height          =   255
               Index           =   5
               Left            =   150
               TabIndex        =   32
               Top             =   300
               Width           =   765
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   285
         Left            =   2775
         TabIndex        =   336
         Top             =   345
         Width           =   840
      End
      Begin VB.CommandButton Command8 
         Height          =   450
         Left            =   8880
         Picture         =   "comandes.frx":364A
         Style           =   1  'Graphical
         TabIndex        =   265
         TabStop         =   0   'False
         ToolTipText     =   "Duplicar Comanda"
         Top             =   210
         Width           =   450
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   195
         Left            =   8280
         TabIndex        =   200
         Top             =   345
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton modificar 
         Height          =   450
         Left            =   555
         Picture         =   "comandes.frx":3BCC
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Modificar Registres"
         Top             =   225
         Width           =   450
      End
      Begin VB.CommandButton gravar 
         Height          =   450
         Left            =   9345
         Picture         =   "comandes.frx":3F1A
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Guardar Registres"
         Top             =   210
         Width           =   450
      End
      Begin VB.CommandButton eliminar 
         Height          =   450
         Left            =   1485
         Picture         =   "comandes.frx":425C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Eliminacio Registres"
         Top             =   225
         Width           =   450
      End
      Begin VB.CommandButton alta 
         Height          =   450
         Left            =   75
         Picture         =   "comandes.frx":456E
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Alta  Registres"
         Top             =   225
         Width           =   450
      End
      Begin VB.Data data1 
         BOFAction       =   1  'BOF
         Caption         =   "                     Comandes"
         Connect         =   "Access"
         DatabaseName    =   "y:\comandes\comandes.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         EOFAction       =   1  'EOF
         Exclusive       =   0   'False
         Height          =   390
         Left            =   3975
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "comandes"
         Top             =   225
         Width           =   3840
      End
      Begin VB.CommandButton sortir 
         Height          =   450
         Left            =   9915
         Picture         =   "comandes.frx":49A0
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Sortir a Menú"
         Top             =   210
         Width           =   450
      End
      Begin VB.CommandButton consultar 
         Height          =   450
         Left            =   1020
         Picture         =   "comandes.frx":4EA2
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Consulta Registres"
         Top             =   225
         Width           =   450
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
         TabIndex        =   7
         Top             =   300
         Width           =   105
      End
   End
End
Attribute VB_Name = "formcomandes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub camp1_Change()

End Sub

Private Sub adhesiu_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 113 Then triaralgu "Triar Adhesiu", "adhesius", vadhesiu, adhesiu
  possar_noms_adhesius False
End Sub
Sub possar_noms_adhesius(Optional lookup As Boolean)
Set rsttmp = dbtmp.OpenRecordset("select * from adhesius where codi=" + atrim(cadbl(vadhesiu)))
  If Not rsttmp.EOF Then
    enduridor(0) = atrim(rsttmp!enduridor)
    adhesiu = atrim(rsttmp!resina)
    grcm1(0) = cadbl(rsttmp!grmcm2resina)
    grcm2(0) = cadbl(rsttmp!grmcm2enduridor)
    ºC1(0) = cadbl(rsttmp!grausresina)
    ºC2(0) = cadbl(rsttmp!grausenduridor)
    If Not lookup Then
      pes1 = cadbl(rsttmp![%resina])
      pes2 = cadbl(rsttmp![%enduridor])
      grmt2 = cadbl(rsttmp!aportcola)
    End If
    
  End If
End Sub

Private Sub alta_Click()
alta_registre
End Sub

Private Sub form1_AccessKeyPress(tecla As String)
  MsgBox tecla
End Sub

Private Sub cimpressio_Click()
 Text64.Text = Mid(cimpressio.Text, 1, 1)
End Sub

Private Sub cimpressio_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub Combo1_LostFocus()
 If cadbl(Combo1.Text) > 4 Or cadbl(Combo1.Text) < 0 Then Combo1.Text = "0"
End Sub

Private Sub Combo10_LostFocus()
If Combo10 <> "B" And Combo10 <> "L" And Combo10 <> "F" And Combo10 <> "BB" Then
     Combo10 = ""
End If
End Sub

Private Sub Combo11_LostFocus()
If Combo11 <> "S" And Combo11 <> "N" Then
     Combo11 = "N"
End If
End Sub

Private Sub Combo14_Change()
  If Chr$(KeyAscii) <> "N" And Chr$(KeyAscii) <> "C" And Chr$(KeyAscii) <> "1" And Chr$(KeyAscii) <> "2" Then
     KeyAscii = Asc("N")
   Else: Combo14.Text = ""
  End If

End Sub

Private Sub Combo15_LostFocus()
 If cadbl(Combo15.Text) > 4 Or cadbl(Combo15.Text) < 0 Then Combo15.Text = "0"
End Sub

Private Sub Combo2_Change()
 If cadbl(Combo2.Text) > 4 Or cadbl(Combo2.Text) < 0 Then Combo2.Text = "0"
End Sub

Private Sub Combo4_LostFocus()
If Combo5 <> "E" And Combo5 <> "I" Then
     Combo5 = ""
End If
End Sub

Private Sub Combo5_LostFocus()
If Combo5 <> "S" And Combo5 <> "N" Then
     Combo5 = "N"
End If
End Sub

Private Sub Combo6_LostFocus()
If Combo6 <> "S" And Combo6 <> "N" Then
     Combo6 = "N"
End If
End Sub

Private Sub Combo7_LostFocus()
If Combo7 <> "S" And Combo7 <> "N" Then
     Combo7 = "N"
End If
End Sub

Private Sub Combo8_LostFocus()
If Combo8 <> "T" And Combo8 <> "L" And Combo8 <> "ST" Then
     Combo8 = ""
End If
End Sub

Private Sub Combo9_LostFocus()
If Combo9 <> "T" And Combo9 <> "L" And Combo9 <> "ST" Then
     Combo9 = ""
End If

End Sub

Private Sub Command6_Click()
MsgBox formscrooll.Values.VertValue

End Sub

Private Sub Command8_Click()
  Dim rsttmpdup As Recordset
  ratoli "espera"
  Set rsttmpdup = dbtmp.OpenRecordset("select * from comandes where comanda=" + atrim(cadbl(data1.Recordset!comanda)))
  If Not rsttmpdup.EOF Then
     alta_registre
     For i = 1 To data1.Recordset.Fields.Count
      data1.Recordset.Fields(i - 1) = Null
     Next i
     For i = 1 To rsttmpdup.Fields.Count
       data1.Recordset.Fields(i - 1) = rsttmpdup.Fields(i - 1)
     Next i
     gravar_registre
  End If
   ratoli "normal"
End Sub

Private Sub Command9_Click()
PrintForm
End Sub

Private Sub ctipusimp_Click()
  Text65.Text = Mid(ctipusimp.Text, 1, 1)
End Sub

Private Sub ctipusimp_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Data1_Reposition()
 
  carregar_lookups
  situar_seccions
  data1.Caption = "         Comandes " + atrim(data1.Recordset.AbsolutePosition + 1) + "/" + atrim(data1.Recordset.RecordCount)
  'On Error Resume Next
End Sub
Sub situar_seccions()
  Dim sec(9, 2)
  Dim ultimapos As Double
  Dim marge As Double
  marge = 100
  ultimapos = formscrooll.Top
  ultimapos = cap.Height + cap.Top
  ext.Visible = False
  imp1.Visible = False
  lam1.Visible = False
  sol.Visible = False
  reb.Visible = False
  For i = 1 To 10
    Select Case Mid(ruta, i, 1)
      Case "E"
         ext.Visible
         ext.Top = ultimapos + marge
         ultimapos = ultimapos + marge + ext.Height
      Case "I"
         imp1.Visible
         imp1.Top = ultimapos + marge
         ultimapos = ultimapos + marge + imp1.Height
      Case "L"
         lam1.Visible
         lam1.Top = ultimapos + marge
         ultimapos = ultimapos + marge + lam1.Height
      Case "R"
         reb.Visible
         reb.Top = ultimapos + marge
         ultimapos = ultimapos + marge + reb.Height
      Case "S"
         sol.Visible
         sol.Top = ultimapos + marge
         ultimapos = ultimapos + marge + sol.Height
    End Select
  Next i
End Sub
Sub triarclient()
  Load formseleccio
  formseleccio.data1.DatabaseName = data1.DatabaseName
  formseleccio.data1.RecordSource = "select * from clients"
  'formseleccio.DBGrid1.
  
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text2.Text = atrim(cadbl(formseleccio.data1.Recordset!codi))
   nomclient.Caption = atrim(formseleccio.data1.Recordset!nom)
  End If
  Unload formseleccio
  
End Sub
Sub triarmesura()
  Load formseleccio
  formseleccio.Caption = "Triar Unitat Mesura"
  formseleccio.data1.DatabaseName = data1.DatabaseName
  formseleccio.data1.RecordSource = "select * from mesures"
  formseleccio.refrescar
  formseleccio.DBGrid1.Columns(0).Visible = False
  formseleccio.DBGrid1.Columns(1).Width = 1200
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text7.Text = atrim(cadbl(formseleccio.data1.Recordset!codi))
   Text16.Text = atrim(formseleccio.data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub

Sub triarmesuraquantitat()
  Load formseleccio
  formseleccio.Caption = "Triar Unitat Mesura"
  formseleccio.data1.DatabaseName = data1.DatabaseName
  formseleccio.data1.RecordSource = "select * from mesureslineals"
  formseleccio.refrescar
  formseleccio.DBGrid1.Columns(0).Visible = False
  formseleccio.DBGrid1.Columns(1).Width = 1200
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text31.Text = atrim(cadbl(formseleccio.data1.Recordset!codi))
   Text30.Text = atrim(formseleccio.data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub


Sub triarmesuraespesor()
  Load formseleccio
  formseleccio.Caption = "Triar Unitat Mesura"
  formseleccio.data1.DatabaseName = data1.DatabaseName
  formseleccio.data1.RecordSource = "select * from mesureslineals"
  formseleccio.refrescar
  formseleccio.DBGrid1.Columns(0).Visible = False
  formseleccio.DBGrid1.Columns(1).Width = 1200
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text23.Text = atrim(cadbl(formseleccio.data1.Recordset!codi))
   Text22.Text = atrim(formseleccio.data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub



Sub triarextrussora()
  Load formseleccio
  formseleccio.Caption = "Triar Màquina Extrussora"
  formseleccio.data1.DatabaseName = data1.DatabaseName
  formseleccio.data1.RecordSource = "select * from maquines where maquina='E'"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text27.Text = atrim(formseleccio.data1.Recordset!codi)
   nomextrussora(0).Caption = atrim(formseleccio.data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub

Sub triaralgu(titol As String, taula As String, control1 As Control, control2 As Control, Optional camp As String, Optional anularcolsel As Byte)
  If atrim(camp) = "" Then camp = "descripcio"
  Load formseleccio
  If cadbl(anularcolsel) > 0 Then formseleccio.Tag = "1"
  formseleccio.Caption = titol
  formseleccio.data1.DatabaseName = data1.DatabaseName
  formseleccio.data1.RecordSource = IIf(Len(taula) < 10, "select * from " + taula, taula)
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
'   On Error Resume Next
   control1 = atrim(formseleccio.data1.Recordset!codi)
   control2 = atrim(formseleccio.data1.Recordset.Fields(camp))
  End If
  Unload formseleccio
  
End Sub

Sub triarproducte()
  Load formseleccio
  formseleccio.data1.DatabaseName = data1.DatabaseName
  formseleccio.data1.RecordSource = "select * from productes"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text3.Text = atrim(formseleccio.data1.Recordset!codi)
   nomproducte.Caption = atrim(formseleccio.data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub


Sub triartipusentrega()
  Load formseleccio
  formseleccio.data1.DatabaseName = data1.DatabaseName
  formseleccio.data1.RecordSource = "select * from tipusentregues"
  formseleccio.refrescar
  formseleccio.Show 1
  If seleccioret = 1 Then
   Text11.Text = atrim(formseleccio.data1.Recordset!codi)
   Label3.Caption = atrim(formseleccio.data1.Recordset!descripcio)
  End If
  Unload formseleccio
  
End Sub




Private Sub consultar_Click()
  buscant = True
  alta_registre
  deixartotblanc
  
End Sub

Private Sub data1_Validate(Action As Integer, Save As Integer)
   If data1.Recordset.EditMode > 0 And (Action = 2 Or Action = 3) Then
      If MsgBox("No has guardat canvis abans de passar a nou registre. Vols guardar ara?", vbCritical + vbYesNo, "Atenció") = vbNo Then
        Save = False
        cancelar_registre
      End If
   End If
End Sub

Private Sub eliminar_Click()
 On Error GoTo err
  If MsgBox("Segur que vols Eliminar?", vbYesNo + vbCritical, "Atenció") = 6 Then
    data1.Recordset.Delete
    data1.Recordset.MoveNext
    If data1.Recordset.EOF Then data1.Recordset.MovePrevious
  End If
 Exit Sub
err:
  MsgBox "No s'ha pogut eliminar possiblement perque tingui registres relacionats. O bé no hi ha res per eliminar."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 65 Then alta_registre: KeyCode = 0
If KeyCode = 69 Then buscar_registre
If KeyCode = 27 Then cancelar_registre
If KeyCode = 112 Then gravar_registre
If KeyCode = 13 Then SendKeys "{TAB}": KeyCode = 0
'34 pag avall    33 pag amunt
'esquerra 37  amunt 38 dreta 39 avall 40
'control    shift=2

If Shift = 2 And data1.Recordset.EditMode = 0 Then
   If KeyCode = 38 Then
     If data1.Recordset.AbsolutePosition > 0 Then data1.Recordset.MovePrevious
   End If
   If KeyCode = 40 Then If data1.Recordset.AbsolutePosition < data1.Recordset.RecordCount - 1 Then data1.Recordset.MoveNext
   If KeyCode = 37 Then If Not data1.Recordset.BOF Then data1.Recordset.MoveFirst
   If KeyCode = 39 Then If Not data1.Recordset.EOF Then data1.Recordset.MoveLast
End If
If KeyCode = 34 Then
   llocform = llocform + 1
   If llocform > topeform Then llocform = topeform
   formscrooll.SetValues formscrooll.Values.HorzValue, taulapos(llocform)
   DoEvents
End If

If KeyCode = 33 Then
   
   If llocform <> 0 Then llocform = llocform - 1
    formscrooll.SetValues formscrooll.Values.HorzValue, taulapos(llocform)
   DoEvents
End If


End Sub



Sub buscar_registre()
consultar_Click
End Sub
Sub alta_registre()
 If areadatos.Enabled = False Then
      areadatos.Enabled = True
      DoEvents
      data1.Recordset.AddNew
      possarvalordcamps 255
      DoEvents
      Text1.Enabled = True
      'busco el mes gran i el poso a codi +1
      If Not buscant Then
        Set rsttmp = dbtmp.OpenRecordset("select max(comanda) as [grancodi] from comandes")
        If Not rsttmp.EOF Then
          Text1 = atrim(cadbl(rsttmp!grancodi) + 1)
         Else: Text1 = "1"
        End If
      End If
      Text1.SetFocus
 End If
End Sub
Sub gravar_registre()
 If areadatos.Enabled And Not buscant Then
    Text1.Enabled = False
    sortir.SetFocus
    DoEvents
    If Screen.ActiveControl.Name = "sortir" Then
      data1.Recordset.Update
      areadatos.Enabled = False
      data1.Recordset.Bookmark = data1.Recordset.LastModified
    End If
 End If
 If buscant Then finalitzarbusqueda
 'formscrooll.SetValues formscrooll.Values.HorzMin, formscrooll.Values.VertMin
   formscrooll.SetValues formscrooll.Values.HorzValue, taulapos(llocform)
End Sub
Sub cancelar_registre()
  If data1.Recordset.EditMode > 0 Then
   data1.Recordset.CancelUpdate
   areadatos.Enabled = False
   Text1.Enabled = False
   buscant = False
   DoEvents
   DoEvents
   If Not data1.Recordset.EOF Then
       data1.Recordset.MoveNext: data1.Recordset.MovePrevious
     Else: If Not data1.Recordset.BOF Then data1.Recordset.MovePrevious: data1.Recordset.MoveNext
   End If
   
   'carregar_lookups
   possarvalordcamps
    Else: Unload Me
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = Asc("'") Then KeyAscii = Asc("´")
  If KeyAscii > 50 Then KeyAscii = Asc(UCase(Chr$(KeyAscii)))
 
End Sub

Private Sub Form_Load()
centerscreen Me
taulapos = Array(-32768, -29288, -26311, -22785, -18618, -15916)
data1.DatabaseName = cami
Set dbtmp = OpenDatabase(data1.DatabaseName)
data1.RecordSource = "comandes"
data1.Refresh
data1.Recordset.MoveLast
data1.Recordset.MoveFirst

possarvalordcamps
 Set ultimcontrol = Screen.ActiveControl
llocform = 0
End Sub

Private Sub gravar_Click()
gravar_registre
End Sub

Private Sub grmt2_LostFocus()
possarconsums
End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub modificar_Click()
   areadatos.Enabled = True
   DoEvents
   data1.Recordset.Edit
   Text2.SetFocus
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub pes1_LostFocus()
possarconsums
End Sub

Private Sub pes2_LostFocus()
possarconsums
End Sub

Private Sub sortir_Click()
 Unload Me
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_LostFocus()
 If Not buscant And data1.Recordset.EditMode > 0 Then
   Set rsttmp = dbtmp.OpenRecordset("select client from comandes where comanda=" + atrim(cadbl(Text1.Text)))
   If rsttmp.RecordCount > 0 Then MsgBox "Aquest codi ja existeix haurieu de canviar-lo": If areadatos.Enabled Then Text1.Enabled = True: Text1.SetFocus
 End If
End Sub


Private Sub Text120_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Soldadora", "select * from maquines where maquina='S'", Text120, nomsoldadora(1)
End Sub

Private Sub Text127_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
     triaralgu "Triar Mesura Espessor", "mesureslineals", Text128, Text127, , 1
  End If
End Sub

Private Sub Text100_LostFocus()
possar_desc_lot Text100.Text, desclot1(1)
End Sub

Private Sub Text101_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then
     triaralgu "Triar Tub Base", "tubbase", Text101, Text101, "cm_int", 1
  End If
End Sub

Private Sub Text129_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
     triaralgu "Triar Soldadura", "tipussoldadura", Text129, Text130, , 1
  End If
End Sub

Private Sub Text133_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Cinta", "select * from accessoris where Tipus_TNC='C'", Text133, cinta(0), , 1
End Sub

Private Sub Text134_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Ansa", "select * from accessoris where Tipus_TNC='N'", Text134, ansa(0), , 1
End Sub

Private Sub Text135_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Troquel", "select * from accessoris where Tipus_TNC='T'", Text135, truquel(0), , 1
End Sub

Private Sub Text64_Change()
If Text64 = "R" Then cimpressio.Text = "Repetida"
If Text64 = "N" Then cimpressio.Text = "Nova"
If Text64 = "M" Then cimpressio.Text = "Modificada"
If atrim(Text64) = "" Then cimpressio.Text = ""

End Sub

Private Sub Text65_Change()
If Text65 = "T" Then ctipusimp.Text = "Transparencia"
If Text65 = "N" Then ctipusimp.Text = "Normal"
If atrim(Text65) = "" Then ctipusimp.Text = ""

End Sub

Private Sub Text68_KeyPress(KeyAscii As Integer)
  If Chr$(KeyAscii) <> "N" And Chr$(KeyAscii) <> "C" And Chr$(KeyAscii) <> "1" And Chr$(KeyAscii) <> "2" Then
     KeyAscii = Asc("N")
   Else: Text68.Text = ""
  End If
End Sub

Private Sub Text68_LostFocus()
  If Len(Text68.Text) = 0 Then Text68.Text = "N"
End Sub

Private Sub Text73_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Impressora", "select * from maquines where maquina='I'", Text73, nomimpressora(1)
End Sub

Private Sub Text80_LostFocus()
  possar_desc_lot Text80.Text, desclot1(0)
End Sub
Sub possar_desc_lot(numlot As String, desclotx As Control)
  Dim desctmp As String
  Dim rsttmp2 As Recordset
  desctmp = ""
  desclotx = desctmp
  If cadbl(numlot) < 1 Then Exit Sub
  Set rsttmp = dbtmp.OpenRecordset("select producte,colorex,espessor,mesuraesp from comandes where comanda=" + atrim(cadbl(numlot)))
  If Not rsttmp.EOF Then
     Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from productes where codi='" + atrim(rsttmp!producte) + "'")
     If Not rsttmp2.EOF Then desctmp = rsttmp2!descripcio + " - "
     Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from colorants where codi=" + atrim(cadbl(rsttmp!colorex)))
     If Not rsttmp2.EOF Then desctmp = desctmp + rsttmp2!descripcio + "  "
     Set rsttmp2 = dbtmp.OpenRecordset("select descripcio from mesures where codi=" + atrim(cadbl(rsttmp!mesuraesp)))
     If Not rsttmp2.EOF Then desctmp = desctmp + rsttmp2!descripcio
  End If
  desclotx = desctmp
  Set rsttmp2 = Nothing
  Set rsttmp = Nothing
End Sub

Private Sub Text81_LostFocus()
  possar_desc_lot Text81.Text, desclot2(1)
End Sub

Private Sub Text82_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then triaralgu "Triar Laminadora", "select * from maquines where maquina='L'", Text82, nomlaminadora(0)
End Sub

Private Sub Text91_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 113 Then
     triaralgu "Triar Camisa", "camises", Text91, Text91, "cm", 1
  End If
End Sub

Private Sub Text99_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then triaralgu "Triar Rebobinadora", "select * from maquines where maquina='R'", Text99, nomrebobinadora(1)
End Sub

Private Sub Timer1_Timer()
  estattaula.Caption = textestattaula(data1.EditMode)
  If estattaula.ForeColor <> QBColor(0) Then
     estattaula.ForeColor = QBColor(0)
    Else: estattaula.ForeColor = QBColor(14)
  End If

   
End Sub


Sub recorregutregistres()
 Dim objecte As Object
 queryorder = ""
 querywhere = ""
 'On Error Resume Next
 For Each objecte In Me
    If TypeOf objecte Is TextBox Then
      If objecte.DataField <> "" Then ' Si Texto es igual "Hola".
        If objecte.Text <> "" Then evaluarcontingut objecte.DataField, objecte.Text, data1.Recordset.Fields(objecte.DataField).Type
     End If
    End If
Next

End Sub


Function evaluarcontingut(camp As String, valor As String, tipusdato As Byte) As String
  Dim rest As String
  rest = ""
  evaluarcontingut = ""
  If triarordre(camp, valor) Then Exit Function
  If tipusdato = 10 Then
   If InStr(1, valor, "*") Or InStr(1, valor, "?") Then
      rest = " like '" + valor + "'"
     Else
       If InStr(1, valor, ">") Or InStr(1, valor, "<") Or InStr(1, valor, "=") Then
           rest = "='" + valor + "'"
        Else: rest = "=" + "'" + valor + "'"
       End If
   End If
  End If
  If tipusdato <> 10 Then
    If InStr(1, valor, ">") Or InStr(1, valor, "<") Or InStr(1, valor, "=") Then
           rest = atrim((valor))
        Else: rest = "=" + atrim(cadbl(valor))
    End If
  End If
  rest = camp + rest
  evaluarcontingut = rest
  
  If querywhere = "" Then
     querywhere = rest
    Else
     querywhere = querywhere + " and " + rest + " "
  End If
End Function

Function triarordre(camp As String, valorord As String) As Boolean
  Dim ord As String
  triarordre = False
  If InStr(1, valorord, "<<") Then ord = camp + " " + " ASC"
  If InStr(1, valorord, ">>") Then ord = camp + " " + " DESC"
  If ord <> "" Then
      triarordre = True
    Else: Exit Function
  End If
  If queryorder = "" Then
     queryorder = ord
   Else: queryorder = queryorder + ", " + ord
  End If
  
End Function
Sub finalitzarbusqueda()
 recorregutregistres
 If data1.Recordset.EditMode > 0 Then data1.Recordset.CancelUpdate
 buscant = False
 Text1.Enabled = True
 areadatos.Enabled = False
 If queryorder <> "" Then queryorder = " Order By " + queryorder
 If querywhere <> "" Then querywhere = " Where " + querywhere
 data1.RecordSource = "select * from comandes " + querywhere + queryorder
 data1.Refresh
  If Not data1.Recordset.EOF Then data1.Recordset.MoveLast
  If Not data1.Recordset.BOF Then data1.Recordset.MoveFirst
 possarvalordcamps
End Sub

Sub deixartotblanc()
 For Each objecte In Me
    If TypeOf objecte Is TextBox Then
      If objecte.DataField <> "" Then ' Si Texto es igual "Hola".
        objecte.Text = ""
     End If
    End If
Next

End Sub

Sub carregar_lookups()

 lookupde "clients", Text2, nomclient, "nom"
  'LOOKUP DE producte
  Set rsttmp = dbtmp.OpenRecordset("select descripcio from productes where codi='" + atrim((Text3.Text)) + "'")
  If Not rsttmp.EOF Then
     nomproducte.Caption = atrim(rsttmp!descripcio)
    Else: nomproducte.Caption = ""
  End If
  'lookup de tipussoldadura
  Set rsttmp = dbtmp.OpenRecordset("select descripcio from tipussoldadura where codi='" + atrim((Text129.Text)) + "'")
  If Not rsttmp.EOF Then
     Text130 = atrim(rsttmp!descripcio)
    Else: Text130 = ""
  End If
  
  
lookupde "tipusentregues", Text11, Label3
lookupde "mesures", Text7, Text16
lookupde "mesureslineals", Text23, Text22
lookupde "mesureslineals", Text31, Text30
lookupde "colorants", Text24, nomcolor(23)
lookupde "materials", Text25, nommaterial(23)
lookupde "aditius", Text26, nomadditiu(23)
lookupde "select descripcio from maquines where maquina='E' and codi=" + atrim(cadbl((Text27.Text))), , nomextrussora(0)
lookupde "select descripcio from maquines where maquina='I' and codi=" + atrim(cadbl((Text73.Text))), , nomimpressora(1)
lookupde "select descripcio from maquines where maquina='L' and codi=" + atrim(cadbl((Text82.Text))), , nomlaminadora(0)
lookupde "select descripcio from maquines where maquina='S' and codi=" + atrim(cadbl((Text120.Text))), , nomsoldadora(0)
lookupde "select descripcio from maquines where maquina='R' and codi=" + atrim(cadbl((Text99.Text))), , nomrebobinadora(1)
lookupde "accessoris", Text133, cinta(0)
lookupde "accessoris", Text134, ansa(0)
lookupde "accessoris", Text135, truquel(0)

possar_desc_lot Text80.Text, desclot1(0)
possar_desc_lot Text81.Text, desclot2(1)
possar_desc_lot Text100.Text, desclot1(1)
lookupde "mesureslineals", Text128, Text127
possar_noms_adhesius True
possarconsums
Set rsttmp = Nothing
End Sub
Sub lookupde(taula As String, Optional control1 As Control, Optional control2 As Control, Optional camp As String)
If camp = "" Then camp = "descripcio"
If Len(taula) < 20 Then
    Set rsttmp = dbtmp.OpenRecordset("select " + camp + " from " + taula + " where codi=" + atrim(cadbl(control1.Text)))
   Else: Set rsttmp = dbtmp.OpenRecordset(taula)
End If
If Not rsttmp.EOF Then
     control2 = atrim(rsttmp.Fields(0))
    Else: control2 = ""
End If

End Sub

Sub possarvalordcamps(Optional tamany As Byte)
Dim t As Double
If cadbl(tamany) = 0 Then t = tamany
On Error Resume Next
 For Each objecte In Me
    If TypeOf objecte Is TextBox Then
      If objecte.DataField <> "" Then
         If cadbl(tamany) = 0 Then t = data1.Recordset.Fields(objecte.DataField).Size
         objecte.MaxLength = t
      End If
    End If
Next

End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triartipusentrega
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarmesura
End Sub

Private Sub Text16_LostFocus()
carregar_lookups
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarclient
End Sub

Private Sub Text2_LostFocus()
carregar_lookups
End Sub

Private Sub Text22_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarmesuraespesor
End Sub

Private Sub Text22_LostFocus()
  carregar_lookups
End Sub

Private Sub Text24_Change()
  If KeyCode = 113 Then triarextrussora
End Sub

Private Sub Text24_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Colorant", "colorants", Text24, nomcolor(23)
End Sub
Private Sub Text25_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Material", "materials", Text25, nommaterial(23)
End Sub

Private Sub Text26_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triaralgu "Triar Aditiu", "aditius", Text26, nomadditiu(23)
End Sub

Private Sub Text27_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarextrussora
End Sub

Private Sub Text27_LostFocus()
carregar_lookups
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarproducte
End Sub

Private Sub Text3_LostFocus()
carregar_lookups
End Sub

Private Sub Text30_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then triarmesuraquantitat
End Sub

Private Sub Text30_LostFocus()
carregar_lookups
End Sub

Private Sub Timer2_Timer()
'fa canviar el color del control que te el focus
On Error Resume Next
   If TypeOf Screen.ActiveControl Is TextBox Then
     If Screen.ActiveControl.BackColor = -2147483643 Then
         Screen.ActiveControl.BackColor = QBColor(11) 'possar aqui el color
         If TypeOf ultimcontrol Is TextBox Then
          ultimcontrol.BackColor = -2147483643
         End If
          Set ultimcontrol = Screen.ActiveControl
     End If
      Else: On Error GoTo err
         If TypeOf ultimcontrol Is TextBox Then
          ultimcontrol.BackColor = -2147483643
         End If
         Set ultimcontrol = Screen.ActiveControl
   End If
  Exit Sub
err:

End Sub

Sub possarconsums()
Dim valorscol
Dim val1, val2, val3, dens1, dens2 As Double
On Error Resume Next
val1 = 0: val2 = 0: val3 = 0: dens1 = 0
val1 = (cadbl(Text89.Text) / 100) * 1000
val2 = cadbl(grmt2) / 1000
val3 = cadbl(pes1.Text) / (cadbl(pes1.Text) + cadbl(pes2.Text))
dens1 = (val1 * val2 * val3) / cadbl(grcm1(0))

val1 = 0: val2 = 0: val3 = 0: dens2 = 0
val1 = (cadbl(Text89.Text) / 100) * 1000
val2 = cadbl(grmt2) / 1000
val3 = cadbl(pes2.Text) / (cadbl(pes1.Text) + cadbl(pes2.Text))
dens2 = (val1 * val2 * val3) / cadbl(grcm2(0))

On Error GoTo 0
valorscol = Array(1000, 2000, 3000, 4000, 5000, 7500, 10000, 15000, 20000, 30000, 40000, 50000, 75000, 100000, 150000, 200000)
'prepara l'amplada de les reixes i possa els titols
For i = 0 To 15
 If reixaconsums.Tag = "1" Then
  reixaconsums.ColWidth(i) = 620
  If i < 5 Then reixaconsums.ColWidth(i) = 520
 End If
 reixaconsums.Col = i
 reixaconsums.Row = 0
 If reixaconsums.Text <> valorscol(i) Then reixaconsums.Text = valorscol(i)
 reixaconsums.Row = 1
 If reixaconsums.Text <> Format(dens1 * (valorscol(i) / 1000), "##,##0.00") Then
   reixaconsums.Text = Format(dens1 * (valorscol(i) / 1000), "##,##0.00")
 End If
 reixaconsums.Row = 2
 If reixaconsums.Text <> Format(dens2 * (valorscol(i) / 1000), "##,##0.00") Then
   reixaconsums.Text = Format(dens2 * (valorscol(i) / 1000), "##,##0.00")
 End If
Next i
If reixaconsums.Tag = "1" Then
 reixaconsums.ColWidth(15) = 650
 reixaconsums.Width = (590 * 16) + 100
End If
reixaconsums.Tag = ""
End Sub

