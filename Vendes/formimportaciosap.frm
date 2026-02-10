VERSION 5.00
Begin VB.Form formimportaciosap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importació SAP"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   330
      Top             =   660
   End
   Begin VB.ListBox llista 
      Height          =   3765
      Left            =   75
      TabIndex        =   0
      Top             =   315
      Width           =   9930
   End
   Begin VB.Label etnomfitxer 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   45
      Width           =   9885
   End
End
Attribute VB_Name = "formimportaciosap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vnomfitxerlog As String

Private Sub Form_Load()
   etnomfitxer = "Carregant..."
 
End Sub

Private Sub Timer1_Timer()
  Dim vlinia As String
  If vnomfitxerlog = "" Then
      vnomfitxerlog = buscarnomfitxerlog("\\servidorsap\seidor_COMUNICADOR\LOG\Inplacsa")
      DoEvents
      If vnomfitxerlog = "" Then Exit Sub
  End If
  etnomfitxer = vnomfitxerlog
  Open vnomfitxerlog For Input As #1
  llista.Clear
  While Not EOF(1)
       Line Input #1, vlinia
       If vlinia <> "" Then llista.AddItem vlinia
  Wend
  Close 1
End Sub
Function buscarnomfitxerlog(vdir As String) As String
 Dim v As String
  Dim vlinia As String
  Dim vmesactual As String
  v = Dir(vdir + "\Log_" + Format(Now, "yyyymmdd") + "*.txt")
  While v <> ""
    If cadbl(Mid(v, Len(v) - 9, 6)) > vgran Then
      vgran = cadbl(Mid(v, Len(v) - 9, 6))
      vmesactual = v
    End If
    'v = Dir(vdir + "\Log_" + Format(Now, "yyyymmdd") + "*.txt")
    v = Dir
  Wend
  buscarnomfitxerlog = vdir + "\" + vmesactual
  
End Function
