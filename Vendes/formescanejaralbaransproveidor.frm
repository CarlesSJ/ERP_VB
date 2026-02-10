VERSION 5.00
Begin VB.Form formescanejaralbaransproveidor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albarans de proveidor"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6285
   Icon            =   "formescanejaralbaransproveidor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2925
      Left            =   105
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   15
      Width           =   6090
      Begin VB.CheckBox checktotselsproveidors 
         Caption         =   "Tots els proveidors"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   120
         Width           =   2145
      End
      Begin VB.Timer timerdragdrop 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   5370
         Top             =   975
      End
      Begin VB.Timer Timer1 
         Interval        =   400
         Left            =   3720
         Top             =   330
      End
      Begin VB.CheckBox bcmr 
         Caption         =   "CMR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4335
         TabIndex        =   4
         Top             =   270
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label vcarpeta 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   195
         TabIndex        =   2
         Top             =   2640
         Width           =   5730
      End
      Begin VB.Label ettipusescaneig 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Albarans del Proveïdor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   1
         Top             =   2145
         Width           =   5715
      End
      Begin VB.Image espera2 
         Height          =   900
         Left            =   2565
         OLEDropMode     =   1  'Manual
         Picture         =   "formescanejaralbaransproveidor.frx":058A
         Top             =   900
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image espera1 
         Height          =   900
         Left            =   2565
         OLEDropMode     =   1  'Manual
         Picture         =   "formescanejaralbaransproveidor.frx":0994
         Top             =   900
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image Image1 
         Height          =   1725
         Left            =   2130
         OLEDropMode     =   1  'Manual
         Picture         =   "formescanejaralbaransproveidor.frx":0D9D
         Top             =   360
         Width           =   1950
      End
      Begin VB.Image Image2 
         Height          =   825
         Left            =   60
         Picture         =   "formescanejaralbaransproveidor.frx":3062
         Stretch         =   -1  'True
         Top             =   990
         Visible         =   0   'False
         Width           =   1365
      End
   End
End
Attribute VB_Name = "formescanejaralbaransproveidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub espera1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    control_arrastrarideixar Data
End Sub

Private Sub espera2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   control_arrastrarideixar Data
End Sub

Private Sub Form_Activate()
  Static vjaheentrat As Boolean
  If vcarpeta = "" Then vcarpeta = "\\ord_copies\Proveidors Albarans Scanner\EscanejarDesdeExpedicions"
  If Not vjaheentrat Then
   vjaheentrat = True
  ' vcarpeta = "C:\temp\escaner"
   If Not existeix(vcarpeta) Then
       MkDir vcarpeta
       If Not existeix(vcarpeta) Then
         MsgBox "Error no es troba la carpeta on es guarden les escanejades." + vbNewLine + vcarpeta + vbNewLine + "AVISA AL RESPONSABLE.", vbCritical, "ERROR"
         Unload formescanejaralbaransproveidor
       End If
   End If
   eliminar_fitxersdelacarpetaescaner
  End If
End Sub

Sub eliminar_fitxersdelacarpetaescaner()
 On Error Resume Next
 Kill vcarpeta + "\*.*"
End Sub

Private Sub Form_Click()
 Dim rst As Recordset
  Dim rst2 As Recordset
  Dim rst3 As Recordset
  Dim v As String
  Dim vlot As String
  Dim vprov As String
  enviaremailgenericambadjunt "miquel.inplacsa@gmail.com", "CMR escanejat Nº: " + atrim(vnumcmr), "Escaneig d'aquest CMR i verificat que els albarans també estan inclosos, REVISA QUE EL CMR ADJUNT SIGUI EL CORRECTE.", "c:\temp\a.pdf"
  Exit Sub
  Set rst = dbtmp.OpenRecordset("select * from registre_escanejades_expedicions where tipus='CQ'")
  While Not rst.EOF
     v = rst!nomfitxer
     vlot = atrim(Mid(v, 4, InStr(1, v, "[") - 4))
     vprov = Mid(v, InStr(1, v, "[") + 1, InStr(1, v, "]") - InStr(1, v, "[") - 1)
     Set rst2 = dbtmp.OpenRecordset("select * from proveidors_comercial where codiproduccio=" + atrim(cadbl(vprov)))
     If Not rst2.EOF Then
         Set rst3 = dbtmp.OpenRecordset("select * from albaransbip where numlotproveidor='" + atrim(vlot) + "' and codiproveidorcomercial=" + atrim(rst2!codicomptable))
         If rst3.EOF Then Stop
         dbtmp.Execute "update albaransbip set lotescanejat=true where numlotproveidor='" + atrim(vlot) + "' and codiproveidorcomercial=" + atrim(rst2!codicomptable)
           Else: Stop
     End If
     rst.MoveNext
  Wend
   
  


'  Dim vnomproveidor As String
'  Dim vcodiproveidor As Double
'  triar_proveidor vnomproveidor, vcodiproveidor
'  escullir_albaradelproveidor vcodiproveidor

End Sub

Private Sub Form_Load()
   Set dbtmp = OpenDatabase(rutadelfitxer(cami) + "comandes.mdb")
End Sub

Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    control_arrastrarideixar Data
End Sub
Sub executareldragdrop(vfitxer As String)
     If Not existeix(vcarpeta) Then MkDir vcarpeta
     eliminar_fitxersdelacarpetaescaner
     Copiar_Fitxer vfitxer, vcarpeta
     comprovar_carpetaescaner
End Sub

Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   control_arrastrarideixar Data
End Sub
Sub control_arrastrarideixar(Data As DataObject)
 If Data.GetFormat(15) Then
     timerdragdrop.Tag = Data.Files(1)
     timerdragdrop.Enabled = True
 End If
End Sub

Private Sub Timer1_Timer()
  If espera1.Visible Then
       espera2.Visible = True
       espera1.Visible = False
         Else
           espera2.Visible = False
           espera1.Visible = True
  End If
  
  comprovar_carpetaescaner
End Sub
Sub comprovar_carpetaescaner()
   Dim vdir As String
   vdir = Dir(vcarpeta + "\*.*")
   If vdir <> "" Then
      Timer1.Enabled = False
      anomenar_i_desar_fitxer vcarpeta + "\" + vdir
      If existeix(vcarpeta + "\" + vdir) And vcarpeta <> "" Then Kill vcarpeta + "\" + vdir
      Unload formescanejaralbaransproveidor
'      Timer1.Enabled = True
   End If
End Sub
Sub anomenar_i_desar_fitxer(vfitxer As String)
    If formescanejaralbaransproveidor.Tag = "albarans" Then desar_albaransproveidors vfitxer
    If formescanejaralbaransproveidor.Tag = "certificats" Then desar_certificatsLOTS vfitxer
    If formescanejaralbaransproveidor.Tag = "albaransSAP" Then
          If bcmr.Value <> 1 Then
                desar_albaransSAP vfitxer
                 Else: desar_CMRS vfitxer
          End If
    End If
End Sub
Sub desar_CMRS(vfitxer As String)
   Dim vnomfitxerfinal As String
   Dim vnumalb As Double
   Dim vnumcmr As String
inici:
   vnomfitxerfinal = escull_numerodeCMR
   vnumcmr = vnomfitxerfinal
   'vnomfitxerfinal = InputBox("Escaneja el codi de barres de l'albarà o escriu-lo." + vbNewLine + "ASSEGUREU-VOS QUE SIGUI CORRECTE AQUEST NUMERO.", "ATENCIÓ")
   If vnomfitxerfinal <> "" Then
      vnomfitxerfinal = substituir(vnomfitxerfinal, ".", "")
      vnomfitxerfinal = "CMR_" + vnomfitxerfinal
      If Not comprovarsiexisteixeixentotselsalbaransSAPescanejats(cadbl(vnumcmr)) Then
            MsgBox "No he trobat tos els ALBARANS DEL SAP escanejats? " + vbNewLine + vbNewLine + "   PRIMER S'HAN D'ESCANEJAR ELS ALBARANS SAP RELACIONATS.", vbExclamation, "A T E N C I Ó"
            Kill vfitxer
            Exit Sub
      End If
      Else: GoTo fi
   End If
   vnomfitxerfinal = vnomfitxerfinal + ".pdf"
   vnomfitxerfinal = treuresimbolsnovalidsnomfitxer(vnomfitxerfinal)
   If existeix(vfitxer) Then
       FileCopy vfitxer, rutadelfitxer(cami) + "Cache_escanejarexpedicions\AlbaransSAP\" + vnomfitxerfinal
       If SiNoexisteix(rutadelfitxer(cami) + "Cache_escanejarexpedicions\AlbaransSAP\" + vnomfitxerfinal) Then
            MsgBox "Hi ha hagut un error al copiar el fitxer al Servidor... " + vbNewLine + "Torna-ho a provar.", vbCritical, "Error"
            GoTo fi
              Else: passarCMRaescanejat cadbl(vnumcmr), rutadelfitxer(cami) + "Cache_escanejarexpedicions\AlbaransSAP\" + vnomfitxerfinal
       End If
   End If
fi:

End Sub
Function SiNoexisteix(vnomfitxer As String) As Boolean
   SiNoexisteix = True
   If existeix(vnomfitxer) Then SiNoexisteix = False: GoTo fi
   wait 1
   If existeix(vnomfitxer) Then SiNoexisteix = False: GoTo fi
   wait 1
   If existeix(vnomfitxer) Then SiNoexisteix = False: GoTo fi
fi:
End Function
Sub passarCMRaescanejat(vnumcmr As Double, vfitxer As String)
  Dim dbvendestmp As Database
  Set dbvendestmp = OpenDatabase(rutadelfitxer(cami) + "Vendes.mdb")
  If cadbl(vnumcmr) = 0 Then Exit Sub
  dbvendestmp.Execute "update transportistes_avisos set escanejat=true where numeroavis='" + atrim(vnumcmr) + "'"
  enviaremailgenericambadjunt "expedicions@inplacsa.com", "CMR escanejat Nº: " + atrim(vnumcmr), "Escaneig d'aquest CMR i verificat que els albarans també estan inclosos, REVISA QUE EL CMR ADJUNT SIGUI EL CORRECTE.", vfitxer
  Set dbvendestmp = Nothing
End Sub
Function comprovarsiexisteixeixentotselsalbaransSAPescanejats(vnumcmr As Double) As Boolean
   Dim rst As Recordset
   Dim vnumalbSAP As String
   Dim dbvendestmp As Database
   Set dbvendestmp = OpenDatabase(rutadelfitxer(cami) + "Vendes.mdb")
   Set rst = dbvendestmp.OpenRecordset("SELECT Transportistes_avisos.numeroavis, capcaleraalbara.numalbara, capcaleraalbara.numalbaraSAP FROM Transportistes_avisos LEFT JOIN capcaleraalbara ON Transportistes_avisos.numalbara = capcaleraalbara.numalbara where numeroavis='" + atrim(vnumcmr) + "'")
   If Not rst.EOF Then comprovarsiexisteixeixentotselsalbaransSAPescanejats = True
   While Not rst.EOF
     vnumalbSAP = atrim(rst!numalbaraSAP)
     If vnumalbSAP <> "" Then
       If Not existeix("\\ord_copies\AlbaransSAPClients\" + vnumalbSAP + ".pdf") Then
         If Not existeix(rutadelfitxer(cami) + "Cache_escanejarexpedicions\AlbaransSAP\" + vnumalbSAP + ".pdf") Then
          comprovarsiexisteixeixentotselsalbaransSAPescanejats = False
         End If
       End If
     End If
     rst.MoveNext
   Wend
   Set rst = Nothing
   Set dbvendestmp = Nothing
End Function
Function escull_numerodeCMR() As String
Dim rst As Recordset
   Load formseleccio
'   formseleccio.sortirs.tag = "filtre"
   formseleccio.data1.DatabaseName = rutadelfitxer(cami) + "vendes.mdb"
   formseleccio.data1.RecordSource = "SELECT distinct numeroavis as Numero_CMR FROM transportistes_avisos where year(datarecullida)>=2025 and not escanejat and matriculacamio<>''"
  
  formseleccio.refrescar
  formseleccio.Width = 4000
  formseleccio.DBGrid2.Columns(0).Width = 2200
  'formseleccio.DBGrid2.Columns(1).visible = False
  
  If formseleccio.data1.Recordset.EOF Then MsgBox "No hi ha cap CMR pendent d'escanejar.": GoTo fi
  formseleccio.Show 1
  If seleccioret = 1 Then escull_numerodeCMR = formseleccio.data1.Recordset!Numero_CMR
  If seleccioret = 9 Then escull_numerodeCMR = " "
  
fi:
  Unload formseleccio
      
End Function
Function comprovarsiexisteixaquestalbaraalSAP(vnumalbSAP As String, vnumalb As Double) As Boolean
   Dim rst As Recordset
   If cadbl(vnumalbSAP) = 0 Then vnumalb = 0: GoTo fi
   Set rst = formvendes.datacapcalera.Database.OpenRecordset("select * from capcaleraalbara where numalbaraSAP=" + atrim(cadbl(vnumalbSAP)))
   If Not rst.EOF Then vnumalb = cadbl(rst!numalbara): comprovarsiexisteixaquestalbaraalSAP = True
fi:
   Set rst = Nothing
End Function
Sub desar_albaransSAP(vfitxer As String)
   Dim vnomfitxerfinal As String
   Dim vnumalb As Double
inici:
   vnomfitxerfinal = InputBox("Escaneja el codi de barres de l'albarà o escriu-lo." + vbNewLine + "ASSEGUREU-VOS QUE SIGUI CORRECTE AQUEST NUMERO.", "ATENCIÓ")
   If vnomfitxerfinal <> "" Then
      vnomfitxerfinal = substituir(vnomfitxerfinal, ".", "")
      If Not comprovarsiexisteixaquestalbaraalSAP(vnomfitxerfinal, vnumalb) Then
            If MsgBox("No he trobat aquest ALBARÀ AL SAP? " + vnomfitxerfinal + vbNewLine + "ES CORRECTE?", vbExclamation + vbDefaultButton2 + vbYesNo, "A T E N C I Ó") = vbNo Then GoTo inici
        End If
       Else: GoTo fi
   End If
   vnomfitxerfinal = vnomfitxerfinal + ".pdf"
   vnomfitxerfinal = treuresimbolsnovalidsnomfitxer(vnomfitxerfinal)
   FileCopy vfitxer, rutadelfitxer(cami) + "Cache_escanejarexpedicions\AlbaransSAP\" + vnomfitxerfinal
   If SiNoexisteix(rutadelfitxer(cami) + "Cache_escanejarexpedicions\AlbaransSAP\" + vnomfitxerfinal) Then
          MsgBox "Hi ha hagut un error al copiar el fitxer al Servidor... " + vbNewLine + "Torna-ho a provar.", vbCritical, "Error"
          GoTo fi
   End If
                     
   If vnumalb > 0 Then passar_albara_a_enviat vnumalb, Now
fi:
End Sub
Sub passar_albara_a_enviat(vnumalb As Double, vdataenviament As Date)
    dbbaixes.Execute "update bobinesent set dataentrega=#" + atrim(Format(vdataenviament, "mm/dd/yy")) + "# where numalbara=" + atrim(cadbl(vnumalb))
    formvendes.datacapcalera.Database.Execute "update linies_expedicions set enviat=true where albara=" + atrim(cadbl(vnumalb))
End Sub

Sub desar_certificatsLOTS(vfitxer As String)
  desar_albaransproveidors vfitxer, True
End Sub
Function escullir_lotproveidor(vcodiproveidor As Double) As String
   Dim rst As Recordset
   Load formseleccio
   formseleccio.sortirs.Tag = "filtre"
   formseleccio.data1.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
   formseleccio.data1.RecordSource = "SELECT numlotproveidor FROM albarans_i_lots_proveidor where codiproveidorcomercial=" + atrim(vcodiproveidor)
  
  formseleccio.refrescar
  formseleccio.Width = 6000
  formseleccio.DBGrid2.Columns(0).Width = 1200
  'formseleccio.DBGrid2.Columns(1).visible = False
  
  If formseleccio.data1.Recordset.EOF Then GoTo fi
  formseleccio.Show 1
  If seleccioret = 1 Then escullir_lotproveidor = formseleccio.data1.Recordset!numlotproveidor
  If seleccioret = 9 Then escullir_lotproveidor = " "
  
fi:
  Unload formseleccio
   
End Function

Function totselscodiscomptables(vcodi As Double) As String
  Dim rst As Recordset
  totselscodiscomptables = ""
  If Len(atrim(vcodi)) < 5 Then
     Set rst = dbtmp.OpenRecordset("select codiproduccio from proveidors_comercial where codiproduccio=" + atrim(vcodi))
      Else: Set rst = dbtmp.OpenRecordset("select codiproduccio from proveidors_comercial where codicomptable='" + atrim(vcodi) + "'")
  End If
  If Not rst.EOF Then
     Set rst = dbtmp.OpenRecordset("select codicomptable from proveidors_comercial where codiproduccio=" + atrim(rst!codiproduccio))
     While Not rst.EOF
         totselscodiscomptables = totselscodiscomptables + IIf(totselscodiscomptables = "", "", ",") + atrim(rst!codicomptable)
         rst.MoveNext
     Wend
  End If
  If totselscodiscomptables = "" Then totselscodiscomptables = atrim(vcodi)
  Set rst = Nothing
End Function
Function escullir_albaradelproveidor(vcodiproveidor As Double) As String
   Dim rst As Recordset
   Load formseleccio
   formseleccio.sortirs.Tag = "filtre"
   formseleccio.data1.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
   formseleccio.data1.RecordSource = "SELECT codiproveidorcomercial,num_alb FROM albarans_proveidor where codiproveidorcomercial in (" + totselscodiscomptables(vcodiproveidor) + ")"
  
  formseleccio.refrescar
  formseleccio.Width = 5000
  formseleccio.DBGrid2.Columns(1).Width = 3000
  formseleccio.DBGrid2.Columns(0).Visible = False
  If formseleccio.data1.Recordset.EOF Then GoTo fi
  formseleccio.Show 1
  If seleccioret = 1 Then
        escullir_albaradelproveidor = formseleccio.data1.Recordset!num_alb
        vcodiproveidor = formseleccio.data1.Recordset!codiproveidorcomercial
  End If
  If seleccioret = 0 Then escullir_albaradelproveidor = " "
  
fi:
  Unload formseleccio
   
End Function
Sub desar_albaransproveidors(vfitxer As String, Optional escertificat As Boolean)
   Dim vnomproveidor As String
   Dim vcodiproveidor As Double
   Dim vnomfitxerfinal As String
   Dim vnumeros(100) As String
   Dim i As Byte
   Dim v As String
   
   i = 0
   triar_proveidor vnomproveidor, vcodiproveidor
   If vnomproveidor = "" Then Exit Sub
inici:
   If escertificat Then
       vnomfitxerfinal = " "
       While vnomfitxerfinal <> ""
        vnomfitxerfinal = escullir_lotproveidor(vcodiproveidor)
        If vnomfitxerfinal = " " Or vnomfitxerfinal = "0" Then
            vnomfitxerfinal = InputBox("Entra el " + atrim(i + 1) + "º numero del LOT del proveïdor." + vbNewLine + "ASSEGUREU-VOS QUE SIGUI CORRECTE AQUEST NUMERO.", "ATENCIÓ")
            If vnomfitxerfinal <> "" Then
                If Not comprovarsiexisteixaquestlotoalbdeproveidor(vnomfitxerfinal, vcodiproveidor, False) Then
                 If MsgBox("No he trobat aquest numero de LOT entrat als palets segur que aquest es el número del LOT DEL PROVEÏDOR? " + vnomfitxerfinal + "DEIXA EN BLANC O CANCELA PER DEIXAR DE DEMANAR NUMEROS", vbExclamation + vbDefaultButton2 + vbYesNo, "A T E N C I Ó") = vbNo Then GoTo inici
                End If
            End If
        End If
        vnumeros(i) = vnomfitxerfinal
        If vnomfitxerfinal <> " " Then i = i + 1
       Wend
          Else
           vnomfitxerfinal = " "
           
           While vnomfitxerfinal <> ""
              vnomfitxerfinal = escullir_albaradelproveidor(vcodiproveidor)
              If vnomfitxerfinal = " " Or vnomfitxerfinal = "0" Then vnomfitxerfinal = InputBox("Entra el " + atrim(i + 1) + "º numero d'ALBARÀ del proveïdor." + vbNewLine + "ASSEGUREU-VOS QUE SIGUI CORRECTE AQUEST NUMERO." + vbNewLine + "DEIXA EN BLANC O CANCELA PER DEIXAR DE DEMANAR NUMEROS", "ATENCIÓ")
              If vnomfitxerfinal <> "" Then
               If Not comprovarsiexisteixaquestlotoalbdeproveidor(vnomfitxerfinal, vcodiproveidor, True) Then
                 If MsgBox("No he trobat aquest numero d'ALBARÀ entrat als palets segur que aquest es el número d'ALBARÀ? " + vnomfitxerfinal, vbExclamation + vbDefaultButton2 + vbYesNo, "A T E N C I Ó") = vbNo Then GoTo inici
               End If
              End If
              vnumeros(i) = vnomfitxerfinal
              i = i + 1
           Wend
   End If
   
   i = 0
   If vnomproveidor <> "" Then
         While vnumeros(i) <> ""
              vnomfitxerfinal = vnumeros(i)
              v = vnomfitxerfinal
              If atrim(vnomfitxerfinal) <> "" Then
                vnomfitxerfinal = IIf(escertificat, "CQ_", "") + vnomfitxerfinal + " [" + atrim(vcodiproveidor) + "]-" + atrim(vnomproveidor) + ".pdf"
                vnomfitxerfinal = treuresimbolsnovalidsnomfitxer(vnomfitxerfinal)
                If escertificat Then
                     FileCopy vfitxer, rutadelfitxer(cami) + "Cache_escanejarexpedicions\CQ\" + vnomfitxerfinal
                     If SiNoexisteix(rutadelfitxer(cami) + "Cache_escanejarexpedicions\CQ\" + vnomfitxerfinal) Then
                           MsgBox "Hi ha hagut un error al copiar el fitxer al Servidor... " + vbNewLine + "Torna-ho a provar.", vbCritical, "Error"
                           GoTo fi
                     End If
                     dbtmp.Execute "update albaransbip set lotescanejat=true where numlotproveidor='" + atrim(v) + "' and codiproveidorcomercial=" + atrim(vcodiproveidor)
                      Else
                        FileCopy vfitxer, rutadelfitxer(cami) + "Cache_escanejarexpedicions\AlbaransProveidor\" + vnomfitxerfinal
                        If SiNoexisteix(rutadelfitxer(cami) + "Cache_escanejarexpedicions\AlbaransProveidor\" + vnomfitxerfinal) Then
                           MsgBox "Hi ha hagut un error al copiar el fitxer al Servidor... " + vbNewLine + "Torna-ho a provar.", vbCritical, "Error"
                           GoTo fi
                        End If
                        dbtmp.Execute "update albaransbip set albaraescanejat=true where numalbaraprov='" + atrim(v) + "' and codiproveidorcomercial=" + atrim(vcodiproveidor)
                        RevisaCQdetotselslots "numalbaraprov='" + atrim(v) + "' and codiproveidorcomercial=" + atrim(vcodiproveidor)
                        If UCase(App.EXEName) = "VENDES" Then
                               enviar_email_arribadadematerial atrim(v), atrim(vcodiproveidor)
                              Else
                                 If MsgBox("Vols avisar a COMPRES que ja ha arribat el material a magatzem?", vbExclamation + vbDefaultButton2 + vbYesNo, "ATENCIÓ") = vbYes Then
                                     enviar_email_arribadadematerial atrim(v), atrim(vcodiproveidor)
                                 End If
                        End If
                End If
              End If
              i = i + 1
         Wend
         If i > 0 Then MsgBox "FET"
   End If
   
fi:
End Sub
Sub enviar_email_arribadadematerial(valbprov As String, vcodi As String)
  Dim vsqlquery As String
  Dim vcos As String
  Dim vdata As String
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim rstcompres As Recordset
  
  vdata = Format(Now, "dd/mm/yy")
  vsqlquery = "numalbaraprov='" + atrim(valbprov) + "' and codiproveidorcomercial=" + atrim(vcodi)
  Set rst = dbtmp.OpenRecordset("select * from albaransbip where " + vsqlquery)
  If Not rst.EOF Then
      Set rstcompres = dbtmp.OpenRecordset("select tipusmaterialcomprat from liniescompra where idliniacompra=" + atrim(cadbl(rst!idliniacompra)))
      If rstcompres.EOF Then GoTo fi
      'If rstcompres!tipusmaterialcomprat <> "M" Then GoTo fi
  End If
  If Not rst.EOF Then vcos = vbNewLine + vbNewLine + "Data arribada: " + atrim(vdata) + "->  Proveidor: " + atrim(rst!nomproveidorcomercial) + vbNewLine + vbNewLine + "Detall: " + vbNewLine
  If rst.EOF Then
     Set rst2 = dbtmp.OpenRecordset("select * from proveidors_comercial where codicomptable='" + vcodi + "'")
     If Not rst2.EOF Then
       vcos = vbNewLine + vbNewLine + "Data arribada: " + atrim(vdata) + "-> Proveidor: " + atrim(rst2!nom)
         Else: vcos = vbNewLine + vbNewLine + "Data arribada: " + atrim(vdata) + "->proveidor: " + vcodi
     End If
  End If
  vcos = vcos + vbNewLine + "Ha arribat l'albarà de proveidor Nº: " + valbprov + vbNewLine + "Detall: " + vbNewLine
  While Not rst.EOF
    vcos = vcos + "Comanda: " + justificar(rst!numcomanda, 10, "E") + " -> " + justificar(rst!article, 10, "D") + "  " + atrim(rst!descripcio) + vbNewLine
    rst.MoveNext
  Wend
  enviaremailgeneric "expedicions@inplacsa.com; calidad@inplacsa.com; compres@inplacsa.com", Format(Now, "dd/mm/yy") + " - Arribada de material al magatzem. ", vcos
fi:
  Set rst = Nothing
  Set rst2 = Nothing
  Set rstcompres = Nothing
End Sub
Sub RevisaCQdetotselslots(vquerysql As String)
   Dim rst As Recordset
   Dim rst2 As Recordset
   
   Set rst = dbtmp.OpenRecordset("select * from albaransbip where " + vquerysql)
   While Not rst.EOF
      Set rst2 = dbtmp.OpenRecordset("select * from albaransbip where " + vquerysql + " and numlotproveidor='" + atrim(rst!numlotproveidor) + "' and lotescanejat=true")
      If Not rst2.EOF Then
          rst.Edit
          rst!lotescanejat = True
          rst.Update
      End If
      rst.MoveNext
   Wend
   
   Set rst = Nothing
   Set rst2 = Nothing
End Sub
Function comprovarsiexisteixaquestlotoalbdeproveidor(vnumlot As String, vcodiproveidor As Double, esalbara As Boolean) As Boolean
  Dim rst As Recordset
  If Not esalbara Then
     Set rst = dbtmp.OpenRecordset("select * from albaransbip where numlotproveidor='" + atrim(vnumlot) + "' and codiproveidorcomercial=" + atrim(vcodiproveidor))
      Else: Set rst = dbtmp.OpenRecordset("select * from albaransbip where numalbaraprov='" + atrim(vnumlot) + "' and codiproveidorcomercial=" + atrim(vcodiproveidor))
  End If
  If rst.EOF Then
       comprovarsiexisteixaquestlotoalbdeproveidor = False
        Else: comprovarsiexisteixaquestlotoalbdeproveidor = True
  End If
  Set rst = Nothing
End Function
Sub triar_proveidor(vnomproveidor As String, vcodiproveidor As Double)
  Set dbcomandes = OpenDatabase(rutadelfitxer(cami) + "comandes.mdb")
  Load formseleccio
  formseleccio.sortirs.Tag = "filtre"
  formseleccio.data1.DatabaseName = rutadelfitxer(cami) + "compres.mdb"
  If checktotselsproveidors.Value = 1 Then
         formseleccio.data1.RecordSource = "SELECT DISTINCT codiproduccio AS codi, nom AS nom_proveidor, IIf([proveidors_comercial].[codicomptable]<>'',[proveidors_comercial].[codicomptable],[proveidors_comercial].[codiproduccio]) AS codi_comptable from proveidors_comercial ORDER BY proveidors_comercial.nom"
       Else: formseleccio.data1.RecordSource = "SELECT * from proveidors_ambcomprespendents"
  End If

  formseleccio.refrescar
  formseleccio.Width = 9000
  formseleccio.DBGrid2.Columns(0).Width = 1000
  formseleccio.DBGrid2.Columns(1).Width = 4000
  formseleccio.DBGrid2.Columns(2).Width = 2000
  formseleccio.DBGrid2.Columns(2).Visible = True
  
  formseleccio.Show 1
  If seleccioret = 1 Then
   vcodiproveidor = cadbl(formseleccio.data1.Recordset!codi_comptable)
   vnomproveidor = atrim(formseleccio.data1.Recordset!nom_proveidor)
   Unload formseleccio
  End If
  Unload formseleccio
End Sub

Private Sub timerdragdrop_Timer()
   timerdragdrop.Enabled = False
   'MsgBox vcarpeta
   executareldragdrop timerdragdrop.Tag
   timerdragdrop.Tag = ""
   Unload formescanejaralbaransproveidor
End Sub
