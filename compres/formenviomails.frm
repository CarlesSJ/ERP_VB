VERSION 5.00
Begin VB.Form formenviomails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dades envio mail al proveïdor"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   6780
   Icon            =   "formenviomails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton enviar 
      Height          =   465
      Left            =   5655
      Picture         =   "formenviomails.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Enviar per mail amb pdf."
      Top             =   30
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cos del Missatge"
      Height          =   2700
      Left            =   75
      TabIndex        =   1
      Top             =   1320
      Width           =   6690
      Begin VB.TextBox cosdelmissatge 
         Height          =   2415
         Left            =   75
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   195
         Width           =   6525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dades Capçalera"
      Height          =   1335
      Left            =   75
      TabIndex        =   0
      Top             =   -15
      Width           =   5550
      Begin VB.TextBox asumpte 
         Height          =   285
         Left            =   1305
         TabIndex        =   4
         Top             =   645
         Width           =   4140
      End
      Begin VB.TextBox destinatari 
         Height          =   285
         Left            =   1305
         TabIndex        =   2
         Top             =   285
         Width           =   4155
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   405
         Picture         =   "formenviomails.frx":0B14
         Top             =   975
         Width           =   240
      End
      Begin VB.Label nomfitxeradjunt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   660
         TabIndex        =   6
         Top             =   1020
         Width           =   4980
      End
      Begin VB.Label Label2 
         Caption         =   "Asumpte"
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Destinatari"
         Height          =   255
         Left            =   165
         TabIndex        =   3
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Menu mopcions 
      Caption         =   "Opcions"
      Begin VB.Menu mprametres 
         Caption         =   "Parametres enviament"
         Begin VB.Menu musuari 
            Caption         =   "Usuari"
         End
         Begin VB.Menu mcontrasenya 
            Caption         =   "Contrasenya"
         End
      End
   End
End
Attribute VB_Name = "formenviomails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command7_Click()

End Sub

Private Sub destinatari_LostFocus()
  If Not comprovar_emails(destinatari) Then MsgBox "La direcció o direccions de correu no son en el format correcte.", vbCritical, "Error": destinatari.SetFocus
End Sub
Function comprovar_emails(ByVal vdest As String) As Boolean
   Dim v As String
   comprovar_emails = True
   vdest = vdest + ";"
   While vdest <> ""
     v = Mid(vdest, 1, InStr(1, vdest, ";") - 1)
     If IsEmail(v) Then
         vdest = Mid(vdest, InStr(1, vdest, ";") + 1)
         
           Else: GoTo errnomail
     End If
   Wend
Exit Function
errnomail:
  comprovar_emails = False
End Function
Function comptarcaracters(v As String, c As String) As Byte
   Dim i As Byte
   For i = 1 To Len(v)
      If Mid(v, i, 1) = c Then comptarcaracters = comptarcaracters + 1
   Next i
End Function
Private Function IsEmail(ByVal strEmail As String) As Boolean
Dim strTemp As String
If comptarcaracters(strEmail, "@") <> 1 Then Exit Function
    If Not InStr(strEmail, "@") > 0 Then
        IsEmail = False
    Else
        If Not InStr(strEmail, ".") > 0 Then
            IsEmail = False
        Else
            If Not Len(Left(strEmail, InStr(strEmail, "@") - 1)) >= 3 Then
                IsEmail = False
            Else
                strTemp = Mid(strEmail, InStr(strEmail, "@") + 1, Len(strEmail))
                If Not Len(Left(strTemp, InStr(strTemp, ".") - 1)) >= 3 Then
                    IsEmail = False
                Else
                    If Not Len(Right(strTemp, Len(strTemp) - InStr(strTemp, "."))) >= 2 Then
                        IsEmail = False
                    Else
                        IsEmail = True
                    End If
                End If
            End If
        End If
    End If
   If IsEmail Then IsEmail = verificarsimbolsextranysemail(strEmail)
   
End Function
Function verificarsimbolsextranysemail(vemail As String) As Boolean
  Dim oVBRegE As Object
  vemail = Trim(vemail)
  Set oVBRegE = CreateObject("VBScript.RegExp")
  oVBRegE.IgnoreCase = True
  oVBRegE.Pattern = "^[A-Za-z0-9](([_.-]?[a-zA-Z0-9]+)*)@([A-Za-z0-9]+)(([.-]?[a-zA-Z0-9]+)*).([A-Za-z]{2,})$"
 ' oVBRegE.Pattern = "^((?:[A-Z0-9_%+-]+\.?)+)@((?:[A-Z0-9-]+\.)+[A-Z]{2,4})$"
  verificarsimbolsextranysemail = IIf(oVBRegE.Test(vemail), True, False)
End Function
Private Sub enviar_Click()
  If Not comprovar_emails(destinatari) Then MsgBox "La direcció o direccions de correu no son en el format correcte.", vbCritical, "Error": destinatari.SetFocus: Exit Sub
  comandescompra.Command7.tag = "enviar"
  Me.Hide
End Sub

Private Sub Form_Load()
  comprovar_usuari_i_contrasenya
End Sub
Sub comprovar_usuari_i_contrasenya()
   If llegir_ini("Enviomails", "usuari", "comandes.ini") = "{[}]" Then MsgBox "Pensa a entrar l'usuari i la contrasenya en el menu opcions d'aquesta finestra", vbCritical, "Atenció": Exit Sub
   If llegir_ini("Enviomails", "contrasenya", "comandes.ini") = "{[}]" Then MsgBox "Pensa a entrar l'usuari i la contrasenya en el menu opcions d'aquesta finestra", vbCritical, "Atenció"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  enviar.tag = ""
End Sub

Private Sub mcontrasenya_Click()
Dim usr As String
   usr = InputBoxEx("Entra la contrasenya d'enviament del correu:" + Chr(10) + "(Respecteu majúscules i minúscules)", "Contrasenya", , , , , , SPassword)
   If usr <> "" Then
      escriure_ini "Enviomails", "contrasenya", usr, "comandes.ini"
      MsgBox "Contrasenya canviada.", vbInformation, "D'acord"
   End If
End Sub

Private Sub musuari_Click()
   Dim usr As String
   usr = InputBox("Entra l'usuari d'enviament del correu:" + Chr(10) + "Ex: usuari@inplacsa.com", "Usuari")
   If usr <> "" Then
      escriure_ini "Enviomails", "usuari", usr, "comandes.ini"
      MsgBox "Usuari canviat.", vbInformation, "D'acord"
   End If
   
End Sub
