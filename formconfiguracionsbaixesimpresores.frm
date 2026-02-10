VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form formconfiguracions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuracions escaners i balances"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4320
   Icon            =   "formconfiguracionsbaixesimpresores.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm MSComm1 
      Left            =   465
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Acceptar canvis"
      Height          =   390
      Left            =   2520
      TabIndex        =   12
      Top             =   2790
      Width           =   1485
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EEE4D7&
      Caption         =   "Configuració balança   -   Retorn de tintes"
      Height          =   705
      Left            =   165
      TabIndex        =   8
      Top             =   960
      Width           =   3855
      Begin VB.ComboBox combocombalança 
         Height          =   315
         ItemData        =   "formconfiguracionsbaixesimpresores.frx":058A
         Left            =   1395
         List            =   "formconfiguracionsbaixesimpresores.frx":05CA
         TabIndex        =   10
         Top             =   285
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Provar"
         Height          =   285
         Left            =   2595
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label3 
         BackColor       =   &H00EEE4D7&
         Caption         =   "Nº de port serie:"
         Height          =   240
         Left            =   135
         TabIndex        =   11
         Top             =   315
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F3B378&
      Caption         =   "Configuració escaner   -   Tintes consumides"
      Height          =   705
      Left            =   165
      TabIndex        =   4
      Top             =   1980
      Width           =   3855
      Begin VB.ComboBox combocomconsumides 
         Height          =   315
         ItemData        =   "formconfiguracionsbaixesimpresores.frx":0651
         Left            =   1395
         List            =   "formconfiguracionsbaixesimpresores.frx":0691
         TabIndex        =   6
         Top             =   270
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Provar"
         Height          =   285
         Left            =   2595
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F3B378&
         Caption         =   "Nº de port serie:"
         Height          =   240
         Left            =   135
         TabIndex        =   7
         Top             =   315
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEE4D7&
      Caption         =   "Configuració escaner   -   Retorn de tintes"
      Height          =   705
      Left            =   165
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Provar"
         Height          =   285
         Left            =   2595
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   270
         Width           =   945
      End
      Begin VB.ComboBox combocomretorn 
         Height          =   315
         ItemData        =   "formconfiguracionsbaixesimpresores.frx":0718
         Left            =   1395
         List            =   "formconfiguracionsbaixesimpresores.frx":0758
         TabIndex        =   1
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EEE4D7&
         Caption         =   "Nº de port serie:"
         Height          =   240
         Left            =   135
         TabIndex        =   2
         Top             =   315
         Width           =   1260
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Totes les lectures es faran amb els parametres: ""9600,n,8,1"""
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   90
      TabIndex        =   13
      Top             =   2745
      Width           =   2745
   End
End
Attribute VB_Name = "formconfiguracions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  comprovar_connexio combocomretorn
End Sub
Sub comprovar_connexio(vvalordelcombo As String)
Dim vinici As Date
  Dim v As String
  MSComm1.Tag = ""
  obrirportseriebascula Mid(vvalordelcombo + "    ", 4, 1)
  If MSComm1.Tag = "" Then
     MsgBox "El port s'ha obert correctament despres de fer Acceptar tindràs 10 segons per fer una lectura des del dispositiu per comprovar si llegeix correctament.", vbInformation + vbOKOnly + vbDefaultButton1, "Atenció"
     vinici = Now
     While DateDiff("s", vinici, Now) < 11
        v = pesbascula
        If v <> "" Then MsgBox "El valor rebut es: " + v, vbInformation, "Lectura": GoTo fi
     Wend
     If v = "" Then MsgBox "No s'ha rebut cap lectura del dispositiu", vbCritical, "Error"
fi:
     MSComm1.PortOpen = False
  End If
  formconfiguracions.SetFocus
End Sub

Sub obrirportseriebascula(vnumport As Byte)
  On Error GoTo errordeport
    If Not MSComm1.PortOpen Then
      MSComm1.CommPort = vnumport
     ' 9600 baudios, sin paridad, 7 bits de datos y 1 bit de parada.
      MSComm1.Settings = "9600,n,8,1"
     ' If nummaq = 1 Then MSComm1.Settings = "2400,n,8,1"
     ' Indicar al control que lea todo el búfer al usar Input.
      MSComm1.InputLen = 0
     
      MSComm1.RTSEnable = True 'Por si necesitas habilitar el RTS
     
     'Abrir Puertos
     
      MSComm1.PortOpen = True
    End If
    Exit Sub
errordeport:
    MsgBox "No s'ha pogut connectar amb el dispositiu a la porta Com" + atrim(vnumport), vbCritical, "Error"
    MSComm1.Tag = "error"
End Sub
Function pesbascula() As String
Static buffer As String
Static nobascula As Boolean
 On Error GoTo nopossarpes
 i = 0
 
 buffer = buffer & MSComm1.Input
 If Len(buffer) > 1 Then
   If InStr(1, buffer, Chr$(13)) > 0 Then buffer = Mid(buffer, InStr(1, buffer, "+") + 1, InStr(1, buffer, Chr$(13)))
   pesbascula = buffer
 '  MSComm1.Output = "9" + Chr(7)
   buffer = ""
 End If
 Exit Function
nopossarpes:
   pesbascula = ""
End Function

Private Sub Command2_Click()
comprovar_connexio combocomconsumides
End Sub

Private Sub Command3_Click()
comprovar_connexio combocombalança
End Sub

Private Sub Command4_Click()
   escriure_parametres_configuracions
   Unload formconfiguracions
End Sub

Private Sub Form_Load()
   llegir_parametres_configuracions
End Sub
Sub llegir_parametres_configuracions()

   combocomretorn = llegir_ini("Baixes", "EscanerRetornCom", "comandes.ini")
   combocomconsumides = llegir_ini("Baixes", "EscanerConsumidesCom", "comandes.ini")
   combocombalança = llegir_ini("Baixes", "EscanerBalançaCom", "comandes.ini")
   
End Sub
Sub escriure_parametres_configuracions()
    escriure_ini "Baixes", "EscanerRetornCom", combocomretorn, "comandes.ini"
    escriure_ini "Baixes", "EscanerConsumidesCom", combocomconsumides, "comandes.ini"
    escriure_ini "Baixes", "EscanerBalançaCom", combocombalança, "comandes.ini"
End Sub
