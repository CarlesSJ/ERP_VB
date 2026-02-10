VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form veurereport 
   Caption         =   "Llistat"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9690
   Icon            =   "veurereport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   9690
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer 
      Height          =   2370
      Left            =   285
      TabIndex        =   0
      Top             =   255
      Width           =   8970
      lastProp        =   500
      _cx             =   15822
      _cy             =   4180
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "veurereport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CRViewer_Clicked(ByVal x As Long, ByVal y As Long, EventInfo As Variant, UseDefault As Boolean)
  On Error GoTo fi
  If UCase(arguments(1)) = "DESBOBINADORS" Then veurereport.Hide
fi:
End Sub

Private Sub Form_Resize()
   On Error Resume Next
    CRViewer.Top = 0
    CRViewer.Left = 0
    CRViewer.Height = ScaleHeight
    CRViewer.width = ScaleWidth
End Sub

Private Sub mllistats_Click()

End Sub
