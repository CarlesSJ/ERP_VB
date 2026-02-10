VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form formcapturarmicrometre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capturar micrometre."
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   3630
   ControlBox      =   0   'False
   Icon            =   "formcapturarmicrometre.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   195
      Top             =   3660
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   105
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   2745
      Picture         =   "formcapturarmicrometre.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Acceptar les micres del visor."
      Top             =   1485
      Width           =   795
   End
   Begin VB.CommandButton bmodificarmicres 
      Height          =   645
      Left            =   2760
      Picture         =   "formcapturarmicrometre.frx":0860
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Escriure les micres manualment"
      Top             =   720
      Width           =   795
   End
   Begin VB.TextBox vmicrometre 
      Alignment       =   2  'Center
      BackColor       =   &H00A6A58E&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   615
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0.0"
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label ettolerancia 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005C31DD&
      Height          =   390
      Left            =   150
      TabIndex        =   5
      Top             =   330
      Width           =   2520
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   1110
      Picture         =   "formcapturarmicrometre.frx":0DEA
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   630
   End
   Begin VB.Label etbobina 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005C31DD&
      Height          =   390
      Left            =   150
      TabIndex        =   4
      Top             =   15
      Width           =   2520
   End
   Begin VB.Label etstatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   135
      TabIndex        =   3
      Top             =   5250
      Width           =   3435
   End
   Begin VB.Image Image1 
      Height          =   4350
      Left            =   300
      Picture         =   "formcapturarmicrometre.frx":14D4
      Top             =   840
      Width           =   3000
   End
   Begin VB.Menu mopcions 
      Caption         =   "Opcions"
      Begin VB.Menu mconfigmicrometre 
         Caption         =   "Configuració micròmetre."
         Begin VB.Menu mportdeconnexio 
            Caption         =   "Escullir port de connexió"
         End
         Begin VB.Menu parametresdecom 
            Caption         =   "Parametres de comunicació"
         End
      End
   End
End
Attribute VB_Name = "formcapturarmicrometre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bmodificarmicres_Click()
   Dim v As String
   
   v = InputBox("Escriu les micres manualment." + vbNewLine + "S'ENVIARÀ UN MISSATGE A OFICINES SI ES FA MANUALMENT.", "MICRES MANUALMENT")
   If cadbl(v) > 0 Then
        vmicrometre = atrim(cadbl(v))
        'enviaremailgeneric "rebobinadoresinplacsa@gmail.com", "S'ha entrat les micres de la bobina " + etbobina + " [" + form1.comanda + "] manualment.", atrim(Now) + vbNewLine + atrim(numop) + "-" + atrim(form1.nomoperari) + vbNewLine + "Comanda: " + atrim(form1.comanda) + vbNewLine + "Motiu:" + vbNewLine + atrim("El micrometre no funciona correctament.")
        enviaremailgeneric "micresentradesmanualmentamaquina", "S'ha entrat les micres de la bobina " + etbobina + " [" + form1.comanda + "] manualment.", atrim(Now) + vbNewLine + atrim(numop) + "-" + atrim(form1.nomoperari) + vbNewLine + "Comanda: " + atrim(form1.comanda) + vbNewLine + "Motiu:" + vbNewLine + atrim("El micrometre no funciona correctament.")
   End If
End Sub

Private Sub Command1_Click()
   If cadbl(vmicrometre) = 0 Then MsgBox "No pots surtir sense posar un valor de micres.", vbCritical, "Error": Exit Sub
   formcapturarmicrometre.Hide
End Sub

Private Sub Form_Load()
  obrirportseriebascula
End Sub
Sub obrirportseriebascula(Optional vstringport As String)
  'Dim vstringport As String
 ' On Error GoTo errordeport
 Dim vnumport As Double
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    If Not MSComm1.PortOpen Then
      vnumport = cadbl(llegir_ini("Micrometre", "portconnexio", "comandes.ini"))
      If vnumport = 0 Then vnumport = 1
      MSComm1.CommPort = vnumport
     ' 9600 baudios, sin paridad, 7 bits de datos y 1 bit de parada.
      If vstringport = "" Then vstringport = llegir_ini("Micrometre", "connexio", "comandes.ini")
      If vstringport = "{[}]" Then vstringport = "4800,E,7,2"
      MSComm1.Settings = vstringport
      MSComm1.InputLen = 0
      MSComm1.DTREnable = True
      MSComm1.RTSEnable = False
      On Error GoTo errordeport
      MSComm1.PortOpen = True
      'MSComm1.Output = "<OUT1>" + Chr$(13)
    End If
    Exit Sub
errordeport:
   etstatus = "Error connexió micrometre."
    MSComm1.Tag = "error"
    
End Sub

Function contrasenyacorrecte() As Boolean
  Dim vcontrasenya As String
  Dim vdemanarcontrasenya As String
  vcontrasenya = "inplacsa"
  vdemanarcontrasenya = UCase(InputBoxEx("Escriu la contrasenya per canviar aquests paràmetres.", "Contrasenya", , , , , , SPassword))
  If UCase(vcontrasenya) = vdemanarcontrasenya Then contrasenyacorrecte = True
End Function



Private Sub mportdeconnexio_Click()
  Dim vport As String
   If Not contrasenyacorrecte Then MsgBox "Aquesta contrasenya no es vàlida.", vbCritical, "Error": Exit Sub
   vport = llegir_ini("Micrometre", "portconnexio", "comandes.ini")
   If vport = "{[}]" Or cadbl(vport) = 0 Then
         escriure_ini "Micrometre", "portconnexio", "4", "comandes.ini"
         vport = "4"
   End If
   vport = InputBox("Escriu els parametres del PORT de connexió amb el micrometre." + vbNewLine + " Ex: 4", "Parametres PORT", vstringport)
   If cadbl(vport) > 0 Then escriure_ini "Micrometre", "portconnexio", vport, "comandes.ini"
End Sub

Private Sub parametresdecom_Click()
   Dim vstringport As String
   If Not contrasenyacorrecte Then MsgBox "Aquesta contrasenya no es vàlida.", vbCritical, "Error": Exit Sub
   vstringport = llegir_ini("Micrometre", "connexio", "comandes.ini")
   If vstringport = "{[}]" Or vstringport = "" Then
         escriure_ini "Micrometre", "connexio", "4800,E,7,2", "comandes.ini"
         vstringport = "4800,E,7,2"
          
   End If
   vstringport = InputBox("Escriu els parametres de connexió amb el micrometre." + vbNewLine + " Ex: 4800,E,7,2", "Parametres connexió", vstringport)
   If atrim(vstringport) <> "" Then escriure_ini "Micrometre", "connexio", vstringport, "comandes.ini"
End Sub

Function llegirmicrometre() As Double
Dim buffer As String
Dim t As String
 On Error GoTo nopossarpes
 i = 0
 If MSComm1.InBufferCount > 8 Then
    buffer = MSComm1.Input
    If InStr(1, buffer, "-") Then buffer = ""
    If buffer <> "" Then
        etstatus = buffer
        If InStr(1, Trim(1 / 2), ",") > 0 Then buffer = substituir(buffer, ".", ",")
        vmicrometre = buffer * 1000
    End If
 End If
 Exit Function
nopossarpes:
   llegirmicrometre = 0
End Function

Private Sub Timer1_Timer()
   llegirmicrometre
   Image2.Visible = Not Image2.Visible
   ettolerancia.Visible = Not Image2.Visible
   If vmicrometre > 0 And Image2.Tag = "" Then
       Image2.Tag = "1": Image2.Left = Command1.Left + 30: Image2.Top = Command1.Top + Command1.Height
   End If
End Sub
