VERSION 5.00
Begin VB.Form etbobina 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiqueta Bobina"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox obsetiq 
      Height          =   330
      Left            =   915
      TabIndex        =   12
      Top             =   7335
      Width           =   4665
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Etiquetes Bobina"
      Height          =   1485
      Left            =   3255
      TabIndex        =   8
      Top             =   2580
      Width           =   2325
      Begin VB.ComboBox etinteriorbob 
         Height          =   315
         ItemData        =   "etiquetabobina.frx":0000
         Left            =   90
         List            =   "etiquetabobina.frx":000A
         TabIndex        =   10
         Top             =   675
         Width           =   2130
      End
      Begin VB.CheckBox etmostra 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Et. mostra x enviar al client"
         Height          =   300
         Left            =   60
         TabIndex        =   9
         Top             =   1095
         Width           =   2205
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Interior Bobina"
         Height          =   270
         Left            =   495
         TabIndex        =   11
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Guardar i Tancar"
      Height          =   450
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   75
      Width           =   1755
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Codi de Barres"
      Height          =   1650
      Left            =   3240
      TabIndex        =   2
      Top             =   735
      Width           =   2385
      Begin VB.ComboBox combocampcodibarres 
         Height          =   315
         ItemData        =   "etiquetabobina.frx":0026
         Left            =   135
         List            =   "etiquetabobina.frx":0033
         TabIndex        =   4
         Top             =   495
         Width           =   2040
      End
      Begin VB.ComboBox combotipuscodibarres 
         Height          =   315
         ItemData        =   "etiquetabobina.frx":005D
         Left            =   150
         List            =   "etiquetabobina.frx":006A
         TabIndex        =   3
         Top             =   1215
         Width           =   2040
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipus codi de barres"
         Height          =   270
         Left            =   165
         TabIndex        =   6
         Top             =   960
         Width           =   1920
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Camp codi de barres"
         Height          =   270
         Left            =   165
         TabIndex        =   5
         Top             =   225
         Width           =   1920
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Camps que NO vull a la Et."
      Height          =   7200
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.CheckBox campsetiqueta 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Obs.Etiq.:"
      Height          =   285
      Left            =   165
      TabIndex        =   13
      Top             =   7365
      Width           =   825
   End
End
Attribute VB_Name = "etbobina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Set rsttmp = dbtmp.OpenRecordset("select * from clients_etbobina where id_envio=" + atrim(cadbl(formclients.envios.Recordset!ID)))
  If rsttmp.EOF Then MsgBox "No s'han gravat els canvis": Exit Sub
  rsttmp.Edit
  For i = 0 To campsetiqueta.Count - 1
     rsttmp.Fields(campsetiqueta(i).Caption) = IIf(campsetiqueta(i).Value = 1, True, False)
  Next i
  rsttmp!campcodibarres = combocampcodibarres.Text
  rsttmp!tipuscodibarres = combotipuscodibarres.Text
  rsttmp!etinteriorbob = etinteriorbob.Text
  rsttmp!etmostra = IIf(etmostra.Value = 1, True, False)
  rsttmp!obsetiq = obsetiq
  rsttmp.Update
  Unload etbobina
End Sub

Private Sub Form_Activate()
Dim i As Byte
  Dim c As String
  i = 1
  If campsetiqueta.Count = 0 Then Unload etbobina: Exit Sub
  While campsetiqueta.Count > 1
     If i <> 0 Then
      Unload campsetiqueta(i)
     End If
     i = i + 1
  Wend
  Set rsttmp = dbtmp.OpenRecordset("select * from clients_etbobina where id_envio=" + atrim(cadbl(formclients.envios.Recordset!ID)))
  If rsttmp.EOF Then
     rsttmp.AddNew
     rsttmp!id_envio = atrim(cadbl(formclients.envios.Recordset!ID))
     rsttmp.Update
     rsttmp.MoveFirst
  End If
  i = 2
  j = 0
  Me.Visible = True
  If Not rsttmp.EOF Then
     While i < rsttmp.Fields.Count
        c = rsttmp.Fields(i).Name
        If LCase(c) <> "obsetiq" And LCase(c) <> "campcodibarres" And LCase(c) <> "etinteriorbob" And LCase(c) <> "tipuscodibarres" And LCase(c) <> "etmostra" Then
         campsetiqueta(j).Caption = UCase(c)
         campsetiqueta(j).Value = IIf(rsttmp.Fields(i).Value, 1, 0)
        End If
        If i < rsttmp.Fields.Count Then
           c = rsttmp.Fields(i).Name
           If LCase(c) <> "campcodibarres" And LCase(c) <> "etinteriorbob" And LCase(c) <> "tipuscodibarres" And LCase(c) <> "etmostra" Then
            j = j + 1
            Load campsetiqueta(j): campsetiqueta(j).Top = campsetiqueta(j - 1).Top + campsetiqueta(j - 1).Height + 10
            campsetiqueta(j).Left = campsetiqueta(j - 1).Left
            campsetiqueta(j).Visible = True
           End If
              Else: If campsetiqueta(j).Caption <> c Then Unload campsetiqueta(j)
        End If
        i = i + 1
     Wend
     combocampcodibarres.Text = atrim(rsttmp!campcodibarres)
     combotipuscodibarres.Text = atrim(rsttmp!tipuscodibarres)
     etinteriorbob.Text = atrim(rsttmp!etinteriorbob)
     etmostra.Value = IIf(rsttmp.Fields("etmostra"), 1, 0)
     obsetiq = atrim(rsttmp!obsetiq)
  End If
End Sub

