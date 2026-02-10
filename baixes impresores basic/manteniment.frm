VERSION 5.00
Begin VB.Form manteniment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manteniments de rasquetes"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   Icon            =   "manteniment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Canvis de Rasqueta"
      Height          =   5040
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   7500
      Begin VB.Frame Frame2 
         Caption         =   "Recircular"
         Height          =   4335
         Left            =   660
         TabIndex        =   33
         Top             =   195
         Width           =   930
         Begin VB.CommandButton brecircular 
            BackColor       =   &H00F1B75F&
            Caption         =   "23/12/25"
            Height          =   345
            Index           =   7
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   3855
            Width           =   810
         End
         Begin VB.CommandButton brecircular 
            BackColor       =   &H00F1B75F&
            Caption         =   "23/12/25"
            Height          =   345
            Index           =   6
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   3330
            Width           =   810
         End
         Begin VB.CommandButton brecircular 
            BackColor       =   &H00F1B75F&
            Caption         =   "23/12/25"
            Height          =   345
            Index           =   5
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   2820
            Width           =   810
         End
         Begin VB.CommandButton brecircular 
            BackColor       =   &H00F1B75F&
            Caption         =   "23/12/25"
            Height          =   345
            Index           =   4
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   2295
            Width           =   810
         End
         Begin VB.CommandButton brecircular 
            BackColor       =   &H00F1B75F&
            Caption         =   "23/12/25"
            Height          =   345
            Index           =   3
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1785
            Width           =   810
         End
         Begin VB.CommandButton brecircular 
            BackColor       =   &H00F1B75F&
            Caption         =   "23/12/25"
            Height          =   345
            Index           =   2
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1260
            Width           =   810
         End
         Begin VB.CommandButton brecircular 
            BackColor       =   &H00F1B75F&
            Caption         =   "23/12/25"
            Height          =   345
            Index           =   1
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   750
            Width           =   810
         End
         Begin VB.CommandButton brecircular 
            BackColor       =   &H00F1B75F&
            Caption         =   "23/12/25"
            Height          =   345
            Index           =   0
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   225
            Width           =   810
         End
      End
      Begin VB.CommandButton bcanvibandeja 
         BackColor       =   &H00C78DFA&
         Caption         =   "Canvi Bandeja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   7
         Left            =   5655
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4005
         Width           =   855
      End
      Begin VB.CommandButton bcanvibandeja 
         BackColor       =   &H00C78DFA&
         Caption         =   "Canvi Bandeja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   6
         Left            =   5655
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton bcanvibandeja 
         BackColor       =   &H00C78DFA&
         Caption         =   "Canvi Bandeja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   5
         Left            =   5655
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2955
         Width           =   855
      End
      Begin VB.CommandButton bcanvibandeja 
         BackColor       =   &H00C78DFA&
         Caption         =   "Canvi Bandeja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   4
         Left            =   5655
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2445
         Width           =   855
      End
      Begin VB.CommandButton bcanvibandeja 
         BackColor       =   &H00C78DFA&
         Caption         =   "Canvi Bandeja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   3
         Left            =   5655
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton bcanvibandeja 
         BackColor       =   &H00C78DFA&
         Caption         =   "Canvi Bandeja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   5655
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1395
         Width           =   855
      End
      Begin VB.CommandButton bcanvibandeja 
         BackColor       =   &H00C78DFA&
         Caption         =   "Canvi Bandeja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   5655
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   885
         Width           =   855
      End
      Begin VB.CommandButton bcanvibandeja 
         BackColor       =   &H00C78DFA&
         Caption         =   "Canvi Bandeja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   5655
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton breset 
         BackColor       =   &H0000FF00&
         Caption         =   "Reset"
         Height          =   360
         Index           =   7
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4065
         Width           =   885
      End
      Begin VB.CommandButton breset 
         BackColor       =   &H0000FF00&
         Caption         =   "Reset"
         Height          =   360
         Index           =   6
         Left            =   1605
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3525
         Width           =   885
      End
      Begin VB.CommandButton breset 
         BackColor       =   &H0000FF00&
         Caption         =   "Reset"
         Height          =   360
         Index           =   5
         Left            =   1605
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3015
         Width           =   885
      End
      Begin VB.CommandButton breset 
         BackColor       =   &H0000FF00&
         Caption         =   "Reset"
         Height          =   360
         Index           =   4
         Left            =   1605
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2490
         Width           =   885
      End
      Begin VB.CommandButton breset 
         BackColor       =   &H0000FF00&
         Caption         =   "Reset"
         Height          =   360
         Index           =   3
         Left            =   1605
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1995
         Width           =   885
      End
      Begin VB.CommandButton breset 
         BackColor       =   &H0000FF00&
         Caption         =   "Reset"
         Height          =   360
         Index           =   2
         Left            =   1605
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1470
         Width           =   885
      End
      Begin VB.CommandButton breset 
         BackColor       =   &H0000FF00&
         Caption         =   "Reset"
         Height          =   360
         Index           =   1
         Left            =   1605
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   930
         Width           =   885
      End
      Begin VB.CommandButton breset 
         BackColor       =   &H0000FF00&
         Caption         =   "Reset"
         Height          =   360
         Index           =   0
         Left            =   1605
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   405
         Width           =   885
      End
      Begin VB.Label etmetres 
         BackStyle       =   0  'Transparent
         Caption         =   "---------------------------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2295
         TabIndex        =   24
         Top             =   4110
         Width           =   3195
      End
      Begin VB.Label etmetres 
         BackStyle       =   0  'Transparent
         Caption         =   "---------------------------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2295
         TabIndex        =   23
         Top             =   3585
         Width           =   3195
      End
      Begin VB.Label etmetres 
         BackStyle       =   0  'Transparent
         Caption         =   "---------------------------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2295
         TabIndex        =   22
         Top             =   3060
         Width           =   3195
      End
      Begin VB.Label etmetres 
         BackStyle       =   0  'Transparent
         Caption         =   "---------------------------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2295
         TabIndex        =   21
         Top             =   2535
         Width           =   3195
      End
      Begin VB.Label etmetres 
         BackStyle       =   0  'Transparent
         Caption         =   "---------------------------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2295
         TabIndex        =   20
         Top             =   2010
         Width           =   3195
      End
      Begin VB.Label etmetres 
         BackStyle       =   0  'Transparent
         Caption         =   "---------------------------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2295
         TabIndex        =   19
         Top             =   1500
         Width           =   3195
      End
      Begin VB.Label etmetres 
         BackStyle       =   0  'Transparent
         Caption         =   "---------------------------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2295
         TabIndex        =   18
         Top             =   975
         Width           =   3195
      End
      Begin VB.Label etmetres 
         BackStyle       =   0  'Transparent
         Caption         =   "---------------------------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2295
         TabIndex        =   17
         Top             =   450
         Width           =   3195
      End
      Begin VB.Label Label1 
         Caption         =   "Rasq-8:"
         Height          =   285
         Index           =   7
         Left            =   90
         TabIndex        =   8
         Top             =   4125
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Rasq-7:"
         Height          =   285
         Index           =   6
         Left            =   90
         TabIndex        =   7
         Top             =   3600
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Rasq-6:"
         Height          =   285
         Index           =   5
         Left            =   90
         TabIndex        =   6
         Top             =   3075
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Rasq-5:"
         Height          =   285
         Index           =   4
         Left            =   90
         TabIndex        =   5
         Top             =   2550
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Rasq-4:"
         Height          =   285
         Index           =   3
         Left            =   90
         TabIndex        =   4
         Top             =   2025
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Rasq-3:"
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   3
         Top             =   1500
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Rasq-2:"
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   2
         Top             =   975
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Rasq-1:"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   450
         Width           =   570
      End
   End
End
Attribute VB_Name = "manteniment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bcanvibandeja_Click(Index As Integer)
   Dim rst As Recordset
   If MsgBox("Has canviat la Bandeja de la rasqueta " + atrim(Index + 1) + Chr(13) + "ES CORRECTE?", vbDefaultButton2 + vbExclamation + vbYesNo, "Canvi bandeja") = vbNo Then Exit Sub
   
    Set rst = dbtmpb.OpenRecordset("select * from impresores_canvisrasquetes")
    rst.AddNew
    rst!Data = Now
    rst!nummaquina = nummaq
    rst!numrasqueta = Index + 1
    rst!metres = 0
    rst!numoperari = numop
    rst!nomoperari = form1.nomoperari
    rst!rasquetaobandeja = "B"
    rst.Update
    If Not form1.impresores.Recordset.EOF Then
       If form1.impresores.Recordset.EditMode = 0 Then form1.impresores.Recordset.Edit
       form1.impresores.Recordset!observacio = "[B-" + atrim(Index + 1) + "] " + atrim(form1.impresores.Recordset!observacio)
       form1.impresores.Recordset.Update
    End If
   carregar_dades
   Set rst = Nothing
End Sub
Sub resetrecirculacio()
    escriure_ini "recirculacioimpressores", "recirculacio" + atrim(nummaq), Now, rutadelfitxer(cami) + "valorsprograma.ini"
    etdatarecirculacio = "Data últim reset: " + Format(llegir_ini("recirculacioimpressores", "recirculacio" + atrim(nummaq), rutadelfitxer(cami) + "valorsprograma.ini"), "dd/mm/yyyy")
    
End Sub

Private Sub brecircular_Click(Index As Integer)
   If MsgBox("Vols fer la recirculació de la Rasqueta " + atrim(Index + 1) + "?", vbExclamation + vbDefaultButton2 + vbYesNo, "RECIRCULACIÓ") = vbYes Then
       dbtmpb.Execute "update impresores_rasquetes set recirculacio=now where nummaquina=" + atrim(nummaq) + " and numrasqueta=" + atrim(Index + 1)
       wait 1
       carregar_dades
   End If
   
End Sub

Private Sub breset_Click(Index As Integer)
   Dim vmetres As Double
   Dim rst As Recordset
   If Index = 8 Then resetrecirculacio: Exit Sub
   If MsgBox("Es posarà el comptador de metres a zero per la rasqueta " + atrim(Index + 1) + Chr(13) + "ES CORRECTE?", vbDefaultButton2 + vbExclamation + vbYesNo, "RESET de rasqueta") = vbNo Then Exit Sub
   dbtmpb.Execute "update impresores_rasquetes set metres=0 where nummaquina=" + atrim(nummaq) + " and numrasqueta=" + atrim(Index + 1)
   vmetres = cadbl(etmetres(Index).tag)
   If vmetres > 0 Then
       Set rst = dbtmpb.OpenRecordset("select * from impresores_canvisrasquetes")
       rst.AddNew
       rst!Data = Now
       rst!nummaquina = nummaq
       rst!numrasqueta = Index + 1
       rst!metres = vmetres
       rst!numoperari = numop
       rst!nomoperari = form1.nomoperari
       rst!rasquetaobandeja = "R"
       rst.Update
       If Not form1.impresores.Recordset.EOF Then
          If form1.impresores.Recordset.EditMode = 0 Then form1.impresores.Recordset.Edit
          form1.impresores.Recordset!observacio = "[R-" + atrim(Index + 1) + "] " + atrim(form1.impresores.Recordset!observacio)
          form1.impresores.Recordset.Update
       End If
   End If
   carregar_dades
   Set rst = Nothing
End Sub

Private Sub Form_Load()
   carregar_dades
End Sub
Sub carregar_dades()
   Dim rst As Recordset
   Set rst = dbtmpb.OpenRecordset("select * from impresores_rasquetes where nummaquina=" + atrim(nummaq))
   While Not rst.EOF
     etmetres(cadbl(rst!numrasqueta) - 1) = String(7 - Len(atrim(rst!metres)), " ") + atrim(rst!metres) + " Metres."
     etmetres(cadbl(rst!numrasqueta) - 1).tag = atrim(rst!metres)
     brecircular(cadbl(rst!numrasqueta) - 1).caption = IIf(atrim(rst!recirculacio) <> "", Format(atrim(rst!recirculacio), "dd/mm/yy"), "")
     rst.MoveNext
   Wend
   Set rst = Nothing
   etdatarecirculacio = "Data últim reset: " + atrim(Format(llegir_ini("recirculacioimpressores", "recirculacio" + atrim(nummaq), rutadelfitxer(cami) + "valorsprograma.ini"), "dd/mm/yyyy"))
End Sub

