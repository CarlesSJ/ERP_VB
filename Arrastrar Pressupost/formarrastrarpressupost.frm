VERSION 5.00
Begin VB.Form Formarrastrarpressupost 
   BackColor       =   &H0080FF80&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1740
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1035
   ScaleWidth      =   1740
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   150
      Top             =   105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Arrastra PDF del pressupost acceptat"
      DragIcon        =   "formarrastrarpressupost.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   90
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   90
      Width           =   1530
   End
End
Attribute VB_Name = "Formarrastrarpressupost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
' -- Api SetForegroundWindow Para traer la ventana al frente
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Sub MakeNormal(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub


Private Sub Image1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
 
End Sub

Private Sub Form_Load()
   MakeTopMost Formarrastrarpressupost.hwnd
End Sub

Private Sub Label1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Data.GetFormat(vbCFFiles) Then
   If existeix(Data.Files(1)) Then
    Formarrastrarpressupost.Visible = False
    importarelfitxer Data.Files(1)
    Unload Formarrastrarpressupost
   End If
  End If
End Sub

Sub crearrutaiborrartemporals(vruta As String)
   On Error Resume Next
     If Not existeix("c:\temp") Then MkDir "c:\temp"
      MkDir vruta
      Kill vruta
End Sub


Private Sub Label1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Static horaentrada As Date
 If Data.GetFormat(vbCFFiles) Then
    If existeix(Data.Files(1)) Then GoTo fi
 End If
 If DateDiff("s", horaentrada, Now) > 5 Then
     horaentrada = Now
     If Not existeix("c:\temp\ImportarPressupost") Then
        crearrutaiborrartemporals "c:\temp\ImportarPressupost"
     End If
     Formarrastrarpressupost.Hide
     Shell "c:\windows\explorer.exe c:\temp\ImportarPressupost", vbNormalFocus
 End If
fi:
End Sub

Private Sub Timer1_Timer()
   mirarsihihaunfitxertemporal
End Sub
Sub mirarsihihaunfitxertemporal()
   Dim vfitxer As String
   vfitxer = Dir("c:\temp\ImportarPressupost\*.pdf")
   While vfitxer <> ""
     If vfitxer <> "." And vfitxer <> ".." Then
         importarelfitxer "c:\temp\ImportarPressupost\" + vfitxer
         Unload Formarrastrarpressupost
         GoTo fi
     End If
     vfitxer = Dir
   Wend
fi:
   
End Sub
Sub importarelfitxer(vfitxer As String, Optional vnumc As Double)
   Dim vcodiclient As Double
   Dim vnomclient As String
   Dim vnompressupost As String
   Dim ruta_documentacio_pressupostos As String
   
   triar_client_imp vcodiclient, vnomclient
   If vcodiclient = 0 Then GoTo fi
   vnompressupost = "                   "
   While Len(vnompressupost) > 12
     vnompressupost = InputBox("Entra el numero de pressupost assignat a aquest document." + Chr(10) + atrim(vcodiclient) + " - " + atrim(UCase(vnomclient)) + Chr(10) + "SI NO POSSES RES ES POSARÀ LA DATA D'AVUI COM A NUMERO DE PRESSUPOST.", "Numero de pressupost.", atrim(vnompressupost))
     If Len(vnompressupost) > 12 Then MsgBox "El nom del pressupost no pot superar els 12 caracters", vbCritical, "Error"
   Wend
   If vnompressupost = "" Then vnompressupost = Format(Now, "yyyymmddhhmmss")
   vnompressupost = treuresimbols(vnompressupost)
   ruta_documentacio_pressupostos = llegir_ini("ruta", "ruta_documentacio_pressupostos", rutadelfitxer(cami) + "valorsprograma.ini")
   If Not existeix(ruta_documentacio_pressupostos + "\" + atrim(vcodiclient)) Then
       MkDir ruta_documentacio_pressupostos + "\" + atrim(vcodiclient)
   End If
   If vnumc = Null Then
       mirarsihihacomandespendents vcodiclient, vnompressupost
        Else:
          vnompressupost = vnompressupost + "_" + atrim(vnumc)
          dbtmp.Execute "update comandes set numpressupost='" + vnompressupost + "' where comanda=" + atrim(vnumcomanda)
   End If
   Copiar_Fitxer vfitxer, ruta_documentacio_pressupostos + "\" + atrim(vcodiclient) + "\" + vnompressupost + ".pdf"
eliminarfitxer:
   Kill vfitxer
fi:
End Sub
Sub mirarsihihacomandespendents(vcodiclient As Double, vnompressupost As String)
   Dim rst As Recordset
   Dim vconsultasql As String
   Dim vnumcomanda As Double
   
   vconsultasql = "select comanda,refclient as [Referencia],marcailinia as [Marca i Linia], cantitatex as Quantitat from comandes where client=" + atrim(vcodiclient) + " and (numpressupost='' or numpressupost=null) and proximaseccio<>'T' and producte<>'PC' and producte<>'PC2' and producte<>'PCP' "
   Set rst = dbtmp.OpenRecordset(vconsultasql)
   If Not rst.EOF Then
      vnumcomanda = 0
      triar_comandapervincular vconsultasql, vnumcomanda
      If vnumcomanda > 0 Then
           vincularcomandaambpressupost vnumcomanda, vnompressupost
            Else: GoTo fi
      End If
      'Set rst = dbtmp.OpenRecordset(vconsultasql)
      'If Not rst.EOF Then If MsgBox("Hi ha mes comandes pendents de vincular, vols vincular mes comandes amb aquest pressupost?", vbInformation + vbYesNo + vbDefaultButton2, "Vincular pressupost amb comanda") = vbNo Then GoTo fi
   End If
fi:
   Set rst = Nothing
End Sub
Sub vincularcomandaambpressupost(vnumcomanda As Double, vnompressupost As String)
   If MsgBox("Segur que vols vincular la comanda " + atrim(vnumcomanda) + " amb el pressupost " + atrim(vnompressupost) + "?", vbYesNo + vbDefaultButton2 + vbExclamation, "Acceptar vinculació") = vbYes Then
'      vnompressupost = vnompressupost + "_" + vnumcomanda
      dbtmp.Execute "update comandes set numpressupost='" + vnompressupost + "_" + vnumcomanda + "' where comanda=" + atrim(vnumcomanda)
   End If
End Sub
Sub triar_comandapervincular(vconsulta As String, vnumcomanda As Double)
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = vconsulta
  formseleccio.refrescar
  formseleccio.DBGrid2.Columns(0).Width = 1500
  formseleccio.DBGrid2.Columns(1).Width = 1500
  formseleccio.DBGrid2.Columns(2).Width = 1500
  'formseleccio.DBGrid2.Columns(3).Width = 2000
  formseleccio.Left = (Screen.Width / 2) - (formseleccio.Width / 2)
  formseleccio.Top = (Screen.Height / 2) - (formseleccio.Height / 2)
  formseleccio.Show 1
  formseleccio.Left = (Screen.Width / 2) - (formseleccio.Width / 2)
  formseleccio.Top = (Screen.Height / 2) - (formseleccio.Height / 2)
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           vnumcomanda = cadbl(formseleccio.DBGrid2.Columns("comanda"))
        End If
   End If
    If seleccioret = 9 Then
       vcomanda = 0
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub


Sub triar_client_imp(vcodiclient As Double, vnomclient As String)
  Load formseleccio
  formseleccio.Command3.Tag = "filtre"
  formseleccio.Data1.DatabaseName = cami
  formseleccio.Data1.RecordSource = "select codi,nom  from clients"
  formseleccio.refrescar
  'formseleccio.DBGrid2.Columns(0).Visible = False
  formseleccio.DBGrid2.Columns(1).Width = 4200
  formseleccio.Left = (Screen.Width / 2) - (formseleccio.Width / 2)
  formseleccio.Top = (Screen.Height / 2) - (formseleccio.Height / 2)
  formseleccio.Show 1
  formseleccio.Left = (Screen.Width / 2) - (formseleccio.Width / 2)
  formseleccio.Top = (Screen.Height / 2) - (formseleccio.Height / 2)
   If seleccioret = 1 Then
        If Not formseleccio.Data1.Recordset.EOF Then
           vnomclient = formseleccio.DBGrid2.Columns("nom")
           vcodiclient = cadbl(formseleccio.DBGrid2.Columns("codi"))
        End If
   End If
    If seleccioret = 9 Then
        vnomclient = ""
        vcodiclient = ""
   End If
   formseleccio.Data1.RecordSource = ""
   formseleccio.Data1.Refresh
   Unload formseleccio
End Sub
