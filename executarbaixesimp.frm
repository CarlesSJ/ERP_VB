VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080FF80&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opcions "
      Height          =   420
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Opcions varies de configuració d'escaners i balances"
      Top             =   45
      Width           =   870
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00008000&
      Caption         =   "Baixes Impresores (Versió anterior)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1035
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1590
      Width           =   2505
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000080&
      Caption         =   "Parar PC"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1470
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2385
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Baixes Impresores"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   1035
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Width           =   2520
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   435
      Width           =   4305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim X As String
 X = "\\serverprodu\dades\progcomandes\aplicacio\baixesimpresoramaquina1.exe"
 If existeix(X) Then
    r = Shell(X, vbNormalFocus)
     Else: MsgBox "No trobo " + X, vbInformation, "Atenció"
 End If
End Sub

Private Sub Command2_Click()

r = Shell("c:\windows\system32\shutdown -s -t 0", vbNormalFocus)
End Sub

Private Sub Command3_Click()
 
  
End Sub

Private Sub Command4_Click()
 Dim X As String
 X = "\\serverprodu\dades\progcomandes\aplicacio\baixesimpresoramaquina2.exe"
 If Not existeix(X) Then
    MsgBox "No he trobat la versió anterior. Hauras de continuar amb l'actual.", vbCritical, "Error de versió"
   Else: r = Shell(X, vbNormalFocus)
 End If
End Sub

Private Sub Command5_Click()
   If UCase(InputBox("Entra la contrasenya per canviar aquests paràmetres." + Chr(10) + "ATENCIÓ SI ES CANVIA ALGUN PARAMETRE LES LECTURES D'ESCANER I BALANÇA PODEN NO FUNCIONAR.", "Parametres de configuració")) = "INPLACSA" Then
      obrir_finestra_de_configuracions
   End If
End Sub
Sub obrir_finestra_de_configuracions()
   formconfiguracions.Show 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = 2 Then
    If InputBoxEx("Entra la contrasenya de configuració, la de sempre però de 4", "Programador", , , , , , SPassword) = "9909" Then
     If MsgBox("Prem si per activar el bloqueig del Ctrl+Alt+Supr i no per desactivar-lo", vbYesNo, "Atenció") = vbYes Then
       Shell "c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\desactivarctrl.reg"
        Else: Shell "c:\windows\regedit.exe /s \\serverprodu\dades\progcomandes\aplicacio\activarctrl.reg"
     End If
    End If
  End If
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
