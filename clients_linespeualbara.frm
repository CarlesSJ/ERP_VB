VERSION 5.00
Begin VB.Form formliniespeualbara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Linies al peu de l'albarà"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7095
   Icon            =   "clients_linespeualbara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton gravar 
      Height          =   450
      Left            =   5985
      Picture         =   "clients_linespeualbara.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Guardar Registres"
      Top             =   2055
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Height          =   1965
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   6900
      Begin VB.TextBox clinies 
         Height          =   315
         Index           =   3
         Left            =   135
         MaxLength       =   80
         TabIndex        =   4
         Top             =   1515
         Width           =   6585
      End
      Begin VB.TextBox clinies 
         Height          =   315
         Index           =   2
         Left            =   135
         MaxLength       =   80
         TabIndex        =   3
         Top             =   1105
         Width           =   6585
      End
      Begin VB.TextBox clinies 
         Height          =   315
         Index           =   1
         Left            =   135
         MaxLength       =   80
         TabIndex        =   2
         Top             =   695
         Width           =   6585
      End
      Begin VB.TextBox clinies 
         Height          =   315
         Index           =   0
         Left            =   135
         MaxLength       =   80
         TabIndex        =   1
         Tag             =   "1"
         Top             =   285
         Width           =   6585
      End
   End
End
Attribute VB_Name = "formliniespeualbara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   If clinies(0).Tag = "1" Then carregar_linies: clinies(0).Tag = ""
End Sub

Sub carregar_linies()
   Dim rst As Recordset
   Set rst = dbtmp.OpenRecordset("select * from clients_notespeu where id_direnvio=" + atrim(cadbl(formliniespeualbara.Tag)) + " order by ordre")
   While Not rst.EOF
      clinies(rst!ordre).Text = rst!descripcio
      rst.MoveNext
   Wend
   Set rst = Nothing
End Sub
Sub guardar_linies()
   Dim rst As Recordset
   Dim i As Byte
   Dim vordre As Byte
   dbtmp.Execute "delete * from clients_notespeu where id_direnvio=" + atrim(cadbl(formliniespeualbara.Tag))
   Set rst = dbtmp.OpenRecordset("select * from clients_notespeu where id_direnvio=" + atrim(cadbl(formliniespeualbara.Tag)) + " order by ordre")
   vordre = 0
   For i = 0 To 3
      If atrim(clinies(i).Text) <> "" Then
        rst.AddNew
        rst!id_direnvio = cadbl(formliniespeualbara.Tag)
        rst!ordre = vordre
        rst!descripcio = atrim(clinies(i).Text)
        rst.Update
        vordre = vordre + 1
      End If
   Next i
   Set rst = Nothing
End Sub

Private Sub gravar_Click()
  guardar_linies
  Unload formliniespeualbara
End Sub
