VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Compactar bases de dades"
   ClientHeight    =   795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6060
   Icon            =   "compactar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timercomençaracompactar 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5325
      Top             =   30
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1245
      Top             =   465
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Compactant..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2385
      TabIndex        =   1
      Top             =   0
      Width           =   2280
   End
   Begin VB.Label lfitxer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   525
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess _
    As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle _
    As Long, ByVal dwMilliseconds As Long) As Long
    
Dim vlog As String
Dim arguments As Variant
Function ObtenerLíneaComando(Optional MaxArgs)
    'Declara las variables.
    Dim C, LíneaComando, LonLínComando, ArgIn, i, NúmArgs
    'Ver si MaxArgs está.
    If IsMissing(MaxArgs) Then MaxArgs = 10
    'Crea una matriz del tamaño correcto.
    ReDim ArgArray(MaxArgs)
    NúmArgs = 0: ArgIn = False
    'Obtiene los argumentos de la línea de comandos.
    LíneaComando = Command()
    LonLínComando = Len(LíneaComando)
    'Recorre la línea de comando carácter a carácter
    'a la vez.

For i = 1 To LonLínComando
        C = Mid(LíneaComando, i, 1)
        'Comprueba espacio o tabulación.
        If (C <> " " And C <> vbTab) Then
            'Ningún espacio o tabulación.
            'Comprueba si está en el argumento.
            If Not ArgIn Then
            'Empieza el nuevo argumento.
            'Comprueba para más argumentos.
                If NúmArgs = MaxArgs Then Exit For
                    NúmArgs = NúmArgs + 1
                    ArgIn = True
                End If
            'Agrega el carácter al argumento actual.

ArgArray(NúmArgs) = ArgArray(NúmArgs) + C
        Else
            'Encontró un espacio o tabulador.
            'Establece ArgIn a False.
            ArgIn = False
        End If
    Next i
    'Redimensiona la matriz lo suficiente para contener los argumentos.
    'ReDim Preserve ArgArray(NúmArgs)
    'Devuelve la matriz en nombre de la función.
    ObtenerLíneaComando = ArgArray()
End Function

Private Sub Form_Activate()
'escriurelog
  Timercomençaracompactar.Enabled = True
End Sub
Sub comença_a_compactar()
  Timercomençaracompactar.Enabled = False
  Timercomençaracompactar.Interval = 0
  vblog = vbNewLine + "****  Comença a compactar -" + Str(Now) + "***********" + vbNewLine
If Not existeix(App.Path + "\jetcomp.exe") Then MsgBox "No hi ha el programa JetComp.exe a la ruta del programa de compactació.", vbCritical, "Error": End
If arguments(1) = "" Then MsgBox " Parametres: nombasededades.mdb o ruta dels mdbs i els farà tots. " + Chr(10) + "Ex: c:\dades\basededades.mdb        Ex: c:\dades  (així farà tots els mdbs de la carpeta)": GoTo fi
   If InStr(1, arguments(1), ".mdb") Then
      vlog = vlog + "  -Compactant... " + Trim(arguments(1)) + vbNewLine
      compactarbasededades Trim(arguments(1))
      vlog = vlog + "---------------------- " + vbNewLine + vbNewLine + vbNewLine
      GoTo fi
   End If
   ruta_relativa_docs = arguments(1)
   resp = Dir(ruta_relativa_docs + "\*.mdb")
   While resp <> ""
      If InStr(1, resp, ".mdb_") = 0 Then
        vlog = vlog + "-Compactant... " + Trim(resp) + vbNewLine
        compactarbasededades ruta_relativa_docs + "\" + resp
         vlog = vlog + "---------------------- " + vbNewLine + vbNewLine + vbNewLine
      End If
      resp = Dir
   Wend
fi:

If vlog <> "" Then
    If existeix("c:\temp\logcompactar.txt") Then
      Kill "c:\temp\logcompactar.txt"
    End If
    vlog = vlog + vbNewLine + "***  Fi de la compactació " + Trim(Now) + "   ***********"
    Open "c:\temp\logcompactar.txt" For Output As #1
    Print #1, vlog
    Close #1
End If
Timer1.Enabled = True
End Sub
Function existeix(nomfitxer As String) As Boolean
Dim a As Integer
On Error GoTo err:
 a = GetAttr(nomfitxer)
 existeix = True
 Exit Function
err:
 existeix = False
End Function
Private Sub Form_Load()
    Dim resp As String
    Dim ruta_relativa_docs As String
    arguments = ObtenerLíneaComando

   
End Sub
Function copiarfitxerpercompactar(fitxermdb As String) As Boolean
  On Error GoTo fi
  copiarfitxerpercompactar = True
  FileCopy fitxermdb, "c:\temp\fitxerpercompactar.mdb"
  Exit Function
fi:
  copiarfitxerpercompactar = False
End Function
Sub compactarbasededades(fitxermdb As String)
  On Error Resume Next
  If InStr(1, UCase(fitxermdb), "_OLD") > 0 Then Exit Sub
  Dim vliniacomandaments As String
  MkDir "c:\temp"
  Kill "c:\temp\fitxercompactant.mdb"
  Kill fitxermdb + "_old"
  Kill "c:\temp\fitxerpercompactar.mdb"
  On Error GoTo fi
  lfitxer = fitxermdb
  Label2.Caption = "Copiant..."
  DoEvents
  vlog = vlog + "Intentant copiar el fitxer " + fitxermdb + vbNewLine
  If Not copiarfitxerpercompactar(fitxermdb) Then GoTo fi
  vlog = vlog + "Fitxer copiat." + vbNewLine
  
  'DBEngine.CompactDatabase "c:\temp\fitxerpercompactar.mdb", "C:\temp\fitxercompactant.mdb"
  vliniacomandaments = """" + App.Path + "\jetcomp.exe""" + " -src:""" + "c:\temp\fitxerpercompactar.mdb" + """ -dest:""" + "C:\temp\fitxercompactant.mdb" + """  -v3"
  'MsgBox vliniacomandaments
  vlog = vlog + "Compactant... " + fitxermdb + vbNewLine
  Label2.Caption = "Compactant...": DoEvents
  ShellAndWait vliniacomandaments, vbNormalFocus
  If Not existeix("c:\temp\fitxercompactant.mdb") Then GoTo fi
  vlog = vlog + "Compactat. Canviant el nom a " + fitxermdb + "_old" + vbNewLine
  Name fitxermdb As fitxermdb + "_old"
  Label2.Caption = "Guardant els canvis...": DoEvents
  vlog = vlog + "Fitxer " + fitxermdb + " compactat." + vbNewLine
  FileCopy "c:\temp\fitxercompactant.mdb", fitxermdb
 ' escriurelog fitxermdb
 Exit Sub
fi:
 vlog = vlog + "Error... " + err.Description + vbNewLine
End Sub
Public Sub ShellAndWait(ByVal program_name As String, _
                         Optional ByVal window_style As VbAppWinStyle = vbNormalFocus, _
                         Optional ByVal max_wait_seconds As Long = 0)
Dim lngProcessId As Long
Dim lngProcessHandle As Long
Dim datStartTime As Date
Const WAIT_TIMEOUT = &H102
Const SYNCHRONIZE As Long = &H100000
Const INFINITE As Long = &HFFFFFFFF

    ' Start the program.
    On Error GoTo ShellError
    lngProcessId = Shell(program_name, window_style)
    On Error GoTo 0
    
    DoEvents

    ' Wait for the program to finish.
    ' Get the process handle.
    lngProcessHandle = OpenProcess(SYNCHRONIZE, 0, lngProcessId)
    If lngProcessHandle <> 0 Then
        datStartTime = Now
        Do
          If WaitForSingleObject(lngProcessHandle, 250) <> WAIT_TIMEOUT Then
            Exit Do
          End If
          DoEvents
          If max_wait_seconds > 0 Then
            If DateDiff("s", datStartTime, Now) > max_wait_seconds Then Exit Do
          End If
        Loop
        CloseHandle lngProcessHandle
    End If
    Exit Sub
    
ShellError:
End Sub

Sub escriurelog(vfitxer As String)
   Dim wShell As Object 'New wshShell
   Set wShell = CreateObject("WScript.Shell")
    MsgBox2 = wShell.PopUp("Que vols fer?", 3, "Prova", Buttons)
    Set wShell = Nothing
End Sub
Private Sub Label1_Click()

End Sub

Private Sub Timer1_Timer()
  End
End Sub

Private Sub Timercomençaracompactar_Timer()
   comença_a_compactar
End Sub
