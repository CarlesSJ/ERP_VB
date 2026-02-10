VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Llistatproduccio 
   Caption         =   "Llistat de Produccions"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Checkexcelnou 
      Caption         =   "Excel nou"
      Height          =   195
      Left            =   5145
      TabIndex        =   24
      Top             =   3150
      Width           =   1080
   End
   Begin VB.CheckBox CheckExcel 
      Caption         =   "a Excel"
      Height          =   195
      Left            =   4170
      TabIndex        =   23
      Top             =   3150
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acceptar"
      Height          =   810
      Left            =   4170
      Picture         =   "llistatproduccio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3360
      Width           =   1035
   End
   Begin Crystal.CrystalReport llistat 
      Left            =   6135
      Top             =   2130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton acceptar 
      Caption         =   "Acceptar"
      Height          =   810
      Left            =   5280
      Picture         =   "llistatproduccio.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mides"
      Height          =   1035
      Left            =   555
      TabIndex        =   10
      Top             =   3180
      Width           =   3525
      Begin VB.TextBox gruix 
         Height          =   285
         Left            =   2235
         TabIndex        =   15
         Top             =   495
         Width           =   795
      End
      Begin VB.TextBox llarg 
         Height          =   285
         Left            =   1290
         TabIndex        =   13
         Top             =   495
         Width           =   795
      End
      Begin VB.TextBox ample 
         Height          =   285
         Left            =   285
         TabIndex        =   11
         Top             =   525
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "Gruix"
         Height          =   225
         Left            =   2430
         TabIndex        =   16
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Llarg"
         Height          =   225
         Left            =   1470
         TabIndex        =   14
         Top             =   255
         Width           =   420
      End
      Begin VB.Label Label5 
         Caption         =   "Ample"
         Height          =   225
         Left            =   345
         TabIndex        =   12
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dades de sel.lecció"
      Height          =   2280
      Left            =   540
      TabIndex        =   5
      Top             =   840
      Width           =   5655
      Begin VB.ComboBox combooperaris 
         Height          =   315
         ItemData        =   "llistatproduccio.frx":074C
         Left            =   120
         List            =   "llistatproduccio.frx":074E
         TabIndex        =   22
         Text            =   "Tots"
         Top             =   870
         Width           =   2250
      End
      Begin VB.ListBox maquina 
         Height          =   2010
         ItemData        =   "llistatproduccio.frx":0750
         Left            =   2385
         List            =   "llistatproduccio.frx":0752
         MultiSelect     =   1  'Simple
         TabIndex        =   9
         Top             =   180
         Width           =   3150
      End
      Begin VB.ComboBox seccio 
         Height          =   315
         ItemData        =   "llistatproduccio.frx":0754
         Left            =   750
         List            =   "llistatproduccio.frx":0767
         TabIndex        =   8
         Text            =   "I"
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label8 
         Caption         =   "Operari:"
         Height          =   210
         Left            =   255
         TabIndex        =   21
         Top             =   645
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "Maquina:"
         Height          =   270
         Left            =   1680
         TabIndex        =   7
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Secció:"
         Height          =   210
         Left            =   105
         TabIndex        =   6
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dates"
      Height          =   615
      Left            =   510
      TabIndex        =   0
      Top             =   210
      Width           =   5655
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Height          =   405
         Left            =   5190
         Picture         =   "llistatproduccio.frx":077A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Afegir un mes a la data"
         Top             =   165
         Width           =   390
      End
      Begin MSMask.MaskEdBox inici 
         Height          =   330
         Left            =   720
         TabIndex        =   3
         Top             =   195
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         _Version        =   327681
         Format          =   "dd/mm/yy hh:nn"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox fi 
         Height          =   330
         Left            =   3030
         TabIndex        =   4
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         _Version        =   327681
         Format          =   "dd/mm/yy hh:nn"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Fi:"
         Height          =   240
         Left            =   2805
         TabIndex        =   2
         Top             =   285
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "Inici:"
         Height          =   240
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.Label maquines 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   150
      TabIndex        =   19
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "Llistatproduccio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub acceptar_Click()
  If Not existeix("c:\ordprog.ini") Then
        Me.Tag = "1"
        llistat_produccions
    Else:
      llistat_produccions True
      Formreixallistatproduccions.Show 1
  End If
End Sub
Sub crear_taula_temporal(dbbaixes As Database, vnomtaula As String)
  On Error GoTo fi
  dbbaixes.Execute ("select * into tmp" + vnomtaula + " in '" + rutadelfitxer(cami) + "comandes.mdb" + "' from " + vnomtaula + " where 1=0")
fi:
  
  On Error GoTo 0
End Sub
Function buscaravaries(vinici As Date, vfi As Date, voperari As Double, vsqlmaquina As String, db2 As Database) As String
  Dim rst As Recordset
  Dim vrasquetes As String
  Dim vbandejes As String
  'Set rst = db2.OpenRecordset("select sum(totalhores) as Th,first(tipificacioavaria) as Tipus from tmpimpressores where  TIPUS='V' AND (datainici>=#" + Format(vinici, "mm/dd/yy hh:nn") + "# and datafi<=#" + Format(vfi, "mm/dd/yy hh:nn") + "#)" + IIf(voperari > 0, " and numoperari=" + atrim(voperari), "") + " and numeromaquina " + vsqlmaquina + " group by tipificacioavaria")
  Set rst = db2.OpenRecordset("select sum(totalhores) as Th,first(tipificacioavaria) as Tipus from tmpimpressores where  TIPUS='V' " + IIf(voperari > 0, " and operari=" + atrim(voperari), "") + " and numeromaquina " + vsqlmaquina + " group by tipificacioavaria")
  While Not rst.EOF
    buscaravaries = buscaravaries + " " + IIf(atrim(rst!tipus) = "", "N/T", atrim(rst!tipus)) + "->" + atrim(rst!Th) + "h"
    rst.MoveNext
  Wend
  If buscaravaries <> "" Then buscaravaries = "Avaries: " + atrim(buscaravaries)
End Function
Function calculrasquetesibandejes(vinici As Date, vfi As Date, voperari As Double, vsqlmaquina As String) As String
  Dim rst As Recordset
  Dim vrasquetes As String
  Dim vbandejes As String
  
  Set rst = dbbaixes.OpenRecordset("select count(*) as Q from impresores_canvisrasquetes where rasquetaobandeja='R' and (data>=#" + Format(vinici, "mm/dd/yy hh:nn") + "# and data<=#" + Format(vfi, "mm/dd/yy hh:nn") + "#)" + IIf(voperari > 0, " and numoperari=" + atrim(voperari), "") + " and nummaquina " + vsqlmaquina)
  vrasquetes = rst!q
  Set rst = dbbaixes.OpenRecordset("select count(*) as Q from impresores_canvisrasquetes where rasquetaobandeja='B' and (data>=#" + Format(vinici, "mm/dd/yy hh:nn") + "# and data<=#" + Format(vfi, "mm/dd/yy hh:nn") + "#)" + IIf(voperari > 0, " and numoperari=" + atrim(voperari), "") + " and nummaquina " + vsqlmaquina)
  vbandejes = rst!q
  calculrasquetesibandejes = "Neteja: " + atrim(vrasquetes) + " Rasquetes i " + atrim(vbandejes) + " Bandejes"
End Function
Sub borrar_fitxers_temporals(vruta As String)
   On Error Resume Next
   MkDir vruta
   Kill vruta + "\*.*"
   On Error GoTo 0
End Sub
Sub llistat_produccions(Optional vnomesgenerardades As Boolean)
  Dim vample As String
  Dim vllarg As String
  Dim vgruix As String
  Dim voperaris As String
  Dim vmaquina As String
  Dim taulatemp As String
  Dim db As Database
  Dim db2 As Database
  'Dim dbbaixes As Database
  Dim camps As String
  Dim rst As Recordset
  Dim rst2 As Recordset
  Dim rstbaixes As Recordset
  Dim com As Double
  Dim titol As String
  Dim vtotalrasquetesibandejes As String
  Dim f As Double
  Dim vtotalavarias As String
  
  'taulatemp = Environ("TEMP") + "\temporal.mdb"
  borrar_fitxers_temporals "c:\temp\taules_llistat_produccions"
  taulatemp = "c:\temp\taules_llistat_produccions\temporal_" + Format(Now, "hhmmss") + ".mdb"
  ratoli "espera"
  Me.Caption = "Processant... "
  If Not existeix(taulatemp) Then DBEngine.CreateDatabase taulatemp, dbLangGeneral, DatabaseTypeEnum.dbVersion30
  vllarg = "longitudsol"
  If seccio <> "S" Then
     vample = "ampleesq": vgruix = "espessor"
   Else: vample = "amplesol": vgruix = "espessorsol"
  End If
  vample = eval(ample, vample)
  vllarg = eval(llarg, vllarg)
  vgruix = eval(gruix, vgruix)
  

  Set dbbaixes = OpenDatabase(llegir_ini("General", "camibaixes", fitxerini), , True, DAO.LockTypeEnum.dbPessimistic)
  Set db = OpenDatabase(rutadelfitxer(cami) + "comandes.mdb", , True, DAO.LockTypeEnum.dbPessimistic)

  On Error Resume Next
  Set db2 = OpenDatabase(taulatemp)
  db2.Execute ("drop table llistatprodu")
  db2.Execute ("drop table tmpcomandes")
  db2.Execute ("drop table llistatprodu")
  On Error GoTo 0
 
  For i = 0 To maquina.ListCount - 1
     If maquina.Selected(i) Then
       If vmaquina <> "" Then vmaquina = vmaquina + ","
       vmaquina = vmaquina + atrim(maquina.ItemData(i))
     End If
  Next i
  If combooperaris.Text <> "Tots" And combooperaris.Text <> "" Then voperaris = " and " + IIf(seccio = "M" Or seccio = "R" Or seccio = "S", "operari1", "operari") + "=" + atrim(cadbl(combooperaris.ItemData(combooperaris.ListIndex)))
  If vmaquina <> "" Then vmaquina = " in (" + vmaquina + ")"
 
 If vmaquina = "" Then MsgBox "No hi ha cap maquina escullida": ratoli "normal": Exit Sub
 
  'faig la maquina
  Me.Caption = "Processant...  (Filtrant les condicions)"
  DoEvents
  Select Case seccio
    Case "I"
       r = "impressora"
       camps = " 'I' as seccio,0 as operari,numeromaquina as Maq,client,comanda,ampleesq,0 as sim,plegatesq,longitudsol,espessor,mesuraesp,'         ' as descmesura,0.0 as hmaquina, 0.0 as hcanvi,0.0 as havaria,0.0 as hfuncionament,0.0 as totalbobines, 0.0 as totalkg, 0.0 as totalmtrs, 0.0 as mtrsmin, #01/01/1900# as dia, 0 as hora,'' as horaformat,numerotintes,0.0 as gramsm2,'' as nomclient "
       camps2 = ",0 as tmetresdolents, 0 as tmtrsajust, 0 as tmtrsllencats"
       campss = " tmpimpressores.numeromaquina ,tmpimpressores.operari,client,comandes.comanda,ampleesq,plegatesq,longitudsol,espessor,mesuraesp,numerotintes,0 as tmetresdolents, 0 as tmtrsajust"
       'dbbaixes.Execute ("select comanda,numeromaquina ,operari into tmpimpressores in '" + cami + "' from impressores where 1=0")
       Me.Caption = "Processant...  (Filtrant les condicions) creant temporal"
       crear_taula_temporal dbbaixes, "impressores"
       Me.Caption = "Processant...  (Filtrant les condicions) borrant temporal"
       db.Execute "delete * from tmpimpressores"
       Me.Caption = "Processant...  (Filtrant les condicions) insertant registres"
       
       dbbaixes.Execute ("insert into tmpimpressores in '" + rutadelfitxer(cami) + "comandes.mdb" + "' select comanda,numeromaquina ,operari from impressores where (tipus='A' or tipus='F') and numeromaquina " + vmaquina + voperaris + " and not isnull(datafi) and (cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn')) between #" & (Format(inici, "mm/dd/yy hh:nn")) & "# and #" & (Format(fi, "mm/dd/yy hh:nn")) & "#)  order by comanda")
       
'       Clipboard.Clear
'       Clipboard.SetText "insert into tmpimpressores in '" + rutadelfitxer(cami) + "comandes.mdb" + "' select comanda,numeromaquina ,operari from impressores where (tipus='A' or tipus='F') and numeromaquina " + vmaquina + voperaris + " and not isnull(datafi) and (cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn')) between #" & (Format(inici, "mm/dd/yy hh:nn")) & "# and #" & (Format(fi, "mm/dd/yy hh:nn")) & "#)  order by comanda"
       'dbbaixes.Execute ("select comanda,numeromaquina as impressora,operari into tmpimpressores in '" + cami + "' from impressores where numeromaquina " + vmaquina + voperaris + " and datafi<>null and (cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn')) between #" & (Format(inici, "mm/dd/yy hh:nn")) & "# and #" & (Format(fi, "mm/dd/yy hh:nn")) & "#)  order by comanda")
       Me.Caption = "Processant...  (Filtrant les condicions) select into"
      ' Clipboard.Clear
      ' Clipboard.SetText "select " + campss + " into tmpcomandes in'" + taulatemp + "' from tmpimpressores ,comandes where [comandes]![comanda] in (select distinct comanda from tmpimpressores) "
       
       'db.Execute "select " + campss + " into tmpcomandes in'" + taulatemp + "' from tmpimpressores ,comandes where [comandes]![comanda] in (select distinct comanda from tmpimpressores) ", dbSQLPassThrough
       db.Execute "select " + campss + " into tmpcomandes in'" + taulatemp + "' FROM comandes RIGHT JOIN tmpimpressores ON comandes.comanda = tmpimpressores.comanda  ", dbSQLPassThrough
       'Set rst = db.OpenRecordset("tmpimpressores")
       'While Not rst.EOF
       '   db2.Execute ("delete * from tmpcomandes where comanda="+atrim(cadbl(rst!comanda))")
       '   rst.MoveNext
       'Wend
       Me.Caption = "Processant...  (Filtrant les condicions) select tmpcomandes"
       Set rst = db2.OpenRecordset("select comanda from tmpcomandes order by comanda")
       titol = "Llistat d´impressores. Data Inici: " + inici + "    Data Fi: " + fi
       
       
    Case "R"
        r = "rebobinadora"
       camps = " 'R' as seccio,0 as operari,rebobinadora as Maq,client,comanda,amplereb as ampleesq,simulteneitatreb as sim,plegatesq,longitudsol,espessor,mesuraesp,'         ' as descmesura,0.0 as hmaquina, 0.0 as hcanvi,0.0 as havaria,0.0 as hfuncionament,0.0 as totalbobines, 0.0 as totalkg, 0.0 as totalmtrs, 0.0 as mtrsmin, #01/01/1900# as dia, 0 as hora,'' as horaformat,numerotintes,0.0 as gramsm2,'' as nomclient "
      ' MsgBox "select comanda into tmprebobinadores in '" + cami + "' from rebobinadores where numeromaquina " + vmaquina + " and (cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn')) between #" & (Format(inici, "mm/dd/yy hh:nn")) & "# and #" & (Format(fi, "mm/dd/yy hh:nn")) & "#) order by comanda"
      'dbbaixes.Execute ("select comanda into tmprebobinadores in '" + cami + "' from rebobinadores where 1=0")
      crear_taula_temporal dbbaixes, "rebobinadores"
       db.Execute "delete * from tmprebobinadores"
       dbbaixes.Execute ("insert into tmprebobinadores in '" + rutadelfitxer(cami) + "comandes.mdb" + "' select comanda from rebobinadores where (tipus='C' or tipus='F') and datafi<>null and numeromaquina " + vmaquina + voperaris + " and (cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn')) between #" & (Format(inici, "mm/dd/yy hh:nn")) & "# and #" & (Format(fi, "mm/dd/yy hh:nn")) & "#) order by comanda")
       'dbbaixes.Execute ("select comanda into tmprebobinadores in '" + cami + "' from rebobinadores where datafi<>null and numeromaquina " + vmaquina + voperaris + " and (cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn')) between #" & (Format(inici, "mm/dd/yy hh:nn")) & "# and #" & (Format(fi, "mm/dd/yy hh:nn")) & "#) order by comanda")
       
       'db.Execute "select comandes.* into tmpcomandes in'" + taulatemp + "' from tmprebobinadores ,comandes where [comandes]![comanda] in (select distinct comanda from tmprebobinadores)" ' [tmprebobinadores]![comanda]
       db.Execute "select comandes.* into tmpcomandes in'" + taulatemp + "' from comandes RIGHT JOIN tmprebobinadores ON comandes.comanda = tmprebobinadores.comanda"
       
       Set rst = db2.OpenRecordset("select comanda from tmpcomandes order by comanda")
       titol = "Llistat de Rebobinadores.  Data Inici: " + inici + "    Data Fi: " + fi
     Case "L"
        r = "laminadora"
       camps = " 'L' as seccio,0 as operari,laminadora as Maq,client,comanda,camisa as ampleesq,simulteneitatlam as sim,plegatesq,longitudsol,espessor,mesuraesp,'         ' as descmesura,0.0 as hmaquina, 0.0 as hcanvi,0.0 as havaria,0.0 as hfuncionament,0.0 as totalbobines, 0.0 as totalkg, 0.0 as totalmtrs, 0.0 as mtrsmin, #01/01/1900# as dia, 0 as hora,'' as horaformat,numerotintes,0.0 as gramsm2,'' as nomclient "
      ' dbbaixes.Execute ("select comanda into tmplaminadores in '" + cami + "' from laminadores where 1=0")
       crear_taula_temporal dbbaixes, "laminadores"
       db.Execute "delete * from tmplaminadores"
       dbbaixes.Execute ("insert into tmplaminadores in '" + rutadelfitxer(cami) + "comandes.mdb" + "' select comanda from laminadores where tipus='C' and numeromaquina " + vmaquina + voperaris + " and datafi<>null and (cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn')) between #" & (Format(inici, "mm/dd/yy hh:nn")) & "# and #" & (Format(fi, "mm/dd/yy hh:nn")) & "#) order by comanda")
       
       db.Execute "select comandes.* into tmpcomandes in'" + taulatemp + "' from comandes RIGHT JOIN tmplaminadores ON comandes.comanda = tmplaminadores.comanda"
       Set rst = db2.OpenRecordset("select comanda from tmpcomandes order by comanda")
       titol = "Llistat de Laminadores.  Data Inici: " + inici + "    Data Fi: " + fi
     Case "S"
        r = "soldadora"
        camps = " 'S' as seccio,0 as operari,soldadora as Maq,client,comanda,amplesol as ampleesq,simulteneitatsol as sim,ampleplegsol as plegatesq,longitudsol,espessor,mesuraesp,'         ' as descmesura,0.0 as hmaquina, 0.0 as hcanvi,0.0 as havaria,0.0 as hfuncionament,0.0 as totalbobines, 0.0 as totalkg, 0.0 as totalmtrs, 0.0 as mtrsmin, #01/01/1900# as dia, 0 as hora,'' as horaformat,numerotintes,0.0 as gramsm2,'' as nomclient "
        'MsgBox camps
        crear_taula_temporal dbbaixes, "soldadores"
       db.Execute "delete * from tmpsoldadores"
       'dbbaixes.Execute ("select comanda into tmpsoldadores in '" + cami + "' from soldadores where 1=0")
       dbbaixes.Execute ("insert into tmpsoldadores in '" + rutadelfitxer(cami) + "comandes.mdb" + "' select comanda from soldadores where tipus='F' and  numeromaquina " + vmaquina + voperaris + " and datafi<>null and (cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn')) between #" & (Format(inici, "mm/dd/yy hh:nn")) & "# and #" & (Format(fi, "mm/dd/yy hh:nn")) & "#) order by comanda")
       db.Execute "select comandes.* into tmpcomandes in'" + taulatemp + "' from comandes RIGHT JOIN tmpsoldadores ON comandes.comanda = tmpsoldadores.comanda"
       Set rst = db2.OpenRecordset("select comanda,impressora from tmpcomandes order by comanda")
       titol = "Llistat de Soldadores.  Data Inici: " + inici + "    Data Fi: " + fi
    Case "M"
       r = "muntadora"
       camps = " 'M' as seccio,0 as operari,0 as Maq,client,comanda,ampleesq,0 as sim,plegatesq,longitudsol,espessor,mesuraesp,'         ' as descmesura,0.0 as hmaquina, 0.0 as hcanvi,0.0 as havaria,0.0 as hfuncionament,0.0 as totalbobines, 0.0 as totalkg, 0.0 as totalmtrs, 0.0 as mtrsmin, #01/01/1900# as dia, 0 as hora,'' as horaformat,numerotintes,0.0 as gramsm2,'' as nomclient "
       campss = " tmpmuntadores.numeromaquina,client,comandes.comanda,ampleesq,plegatesq,longitudsol,espessor,mesuraesp,numerotintes"
       'dbbaixes.Execute ("select comanda,numeromaquina as muntadora into tmpmuntadores in '" + cami + "' from muntadores where 1=0")
       crear_taula_temporal dbbaixes, "muntadores"
       db.Execute "delete * from tmpmuntadores"
       dbbaixes.Execute ("insert into tmpmuntadores in '" + rutadelfitxer(cami) + "comandes.mdb" + "' select comanda,numeromaquina  from muntadores where datafi<>null " + voperaris + " and (cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn')) between #" & (Format(inici, "mm/dd/yy hh:nn")) & "# and #" & (Format(fi, "mm/dd/yy hh:nn")) & "#) order by comanda")
       db.Execute "select " + campss + " into tmpcomandes in'" + taulatemp + "' from comandes RIGHT JOIN tmpmuntadores ON comandes.comanda = tmpmuntadores.comanda " ' [tmpimpressores]![comanda]
       
       Set rst = db2.OpenRecordset("select comanda from tmpcomandes order by comanda")
       titol = "Llistat de muntadores. Data Inici: " + inici + "    Data Fi: " + fi
  End Select
  
  'faig la consulta
 
'que no es repeteixin
Me.Caption = "Processant...  (Eliminant duplicats)"
DoEvents
 com = 9999
   While Not rst.EOF
    If rst!comanda <> com Then
        com = rst!comanda
       Else: rst.Delete
    End If
    rst.MoveNext
   Wend
 
 
 Me.Caption = "Processant...  (Calculant totals)"
 'si es impressora faig els totals
 If r = "impressora" Then
  t = IIf(vllarg <> "", " and " + vllarg, "") + IIf(vample <> "", " and " + vample, "") + IIf(vgruix <> "", " and " + vgruix, "")
  db2.Execute ("select " + camps + camps2 + " into llistatprodu from tmpcomandes " + IIf(t <> "", " where ", "") + t)
 ' MsgBox "select " + camps + camps2 + " into llistatprodu from tmpcomandes " + IIf(t <> "", " where ", "") + t
  'db.Execute ("drop table tmpimpressores")
  db.Execute "delete * from tmpimpressores"
  'wait 2
  r = "select distinct comanda from llistatprodu"
  Me.Caption = "Processant...  (Insert into llistat)": DoEvents
  modificarcamps_tmpimpressores
  
  db2.Execute ("insert into tmpimpressores in '" + rutadelfitxer(cami) + "comandes.mdb" + "' select * from impressores in '" + llegir_ini("General", "camibaixes", fitxerini) + "' where comanda in (" + r + ")")
  Set rst = db2.OpenRecordset("select * from llistatprodu")
  While Not rst.EOF
    db.Execute ("delete * from tmpimpressores where datafi=null")
    Set rst2 = db.OpenRecordset("select min(numeromaquina) as imp,min(operari) as operaris, min(cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn'))) as dia,sum(totalbobines) as totalb,sum(totalkilos) as totalk ,sum(totalmetres) as totalm from tmpimpressores where comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst.Edit
   
    rst!maq = rst2!imp
    rst!operari = rst2!operaris
    rst!totalbobines = rst2!totalb
    rst!totalkg = rst2!totalk
    rst!gramsm2 = 0
    rst!totalmtrs = rst2!totalm
    If Not IsNull(rst2!dia) Then
     rst!dia = rst2!dia
     rst!hora = Format(rst2!dia, "hhnn")
     rst!horaformat = Format(rst!dia, "hh:nn")
    End If
    'busco el nom client
    Set rst2 = db.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(rst!client)), , dbReadOnly)
    If Not rst2.EOF Then rst!nomclient = atrim(rst2!nom)
    
    
    Set rst2 = db.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rst!mesuraesp)), , dbReadOnly)
    If Not rst2.EOF Then rst!descmesura = atrim(rst2!descripcio)
    
    Set rst2 = db.OpenRecordset("select sum(totalhores) as horesf from tmpimpressores where tipus='F' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hfuncionament = cadbl(rst2!horesf)
    
    Set rst2 = db.OpenRecordset("select sum(totalhores) as horesa from tmpimpressores where tipus='A' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!havaria = cadbl(rst2!horesa)
    Set rst2 = db.OpenRecordset("select sum(totalhores) as hmaquina from tmpimpressores where tipus='M' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hmaquina = cadbl(rst2!hmaquina)
    
     Set rst2 = db.OpenRecordset("select sum(totalhores) as horesc from tmpimpressores where tipus='C' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hcanvi = cadbl(rst2!horesc) + rst!havaria + rst!hmaquina
    rst!havaria = cadbl(rst!havaria) 'es hora d'ajust
    rst!hmaquina = rst!hmaquina
    Set rst2 = dbbaixes.OpenRecordset("select sum(metres) as Tmetres from parcials where comanda='" + atrim(rst!comanda) + "' and orcomassignacio='500'")
    rst!tmtrsajust = cadbl(rst2!tmetres)
    Set rst2 = dbbaixes.OpenRecordset("select sum(metres) as Tmetres from parcials where comanda='" + atrim(rst!comanda) + "' and orcomassignacio='500' and instr(1,[observacions],'#llençar')>0")
    rst!tmtrsllencats = cadbl(rst2!tmetres)
    
    'Set rst2 = db.OpenRecordset("select sum(metresprova) as tmetresprova from tmpimpressores where paletprova<>11111 and tipus='A' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    'rst!tmtrsajust = cadbl(rst2!tmetresprova)
    
    'Set rst2 = db.OpenRecordset("select sum(metresprova2) as tmetresprova2 from tmpimpressores where paletprova2<>11111 and tipus='A' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    'rst!tmtrsajust = rst!tmtrsajust + cadbl(rst2!tmetresprova2)
    'Set rst2 = db.OpenRecordset("select sum(metresprova) as tmetresprova from tmpimpressores where paletprova=11111 and tipus='A' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    'rst!tmtrsllencats = cadbl(rst2!tmetresprova)
    
    'Set rst2 = db.OpenRecordset("select sum(metresprova2) as tmetresprova2 from tmpimpressores where paletprova2=11111 and tipus='A' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    'rst!tmtrsllencats = rst!tmtrsllencats + cadbl(rst2!tmetresprova2)
    
    Set rst2 = dbbaixes.OpenRecordset("select tmetresdolents from impressorestot where comanda=" + atrim(rst!comanda), , dbReadOnly)
    If Not rst2.EOF Then rst!tmetresdolents = rst2!tmetresdolents
    
    
    If cadbl(rst!hfuncionament) < 1000 Then
      f = cadbl(rst!hfuncionament)
    Else: f = 0
    End If
    '(Int(v) * 60) + (((v - Int(v)) * 100) * 60 / 100)
    f = (Int(f) * 60) + (((f - Int(f)) * 100) * 60 / 100)
    rst!mtrsmin = 0
    If f > 0 Then rst!mtrsmin = cadbl(rst!totalmtrs) / f
    
    rst.Update
    rst.MoveNext
  Wend
 End If
 
 'si es rebobinadora faig els totals
 If r = "rebobinadora" Then
 t = IIf(vllarg <> "", " and " + vllarg, "") + IIf(vample <> "", " and " + vample, "") + IIf(vgruix <> "", " and " + vgruix, "")
 If t <> "" Then t = " where " + t
  db2.Execute ("select " + camps + camps2 + " into llistatprodu from tmpcomandes " + t)
  'db.Execute ("drop table tmprebobinadores")
  db.Execute "delete * from tmprebobinadores"
  db2.Execute ("insert into tmprebobinadores in '" + rutadelfitxer(cami) + "comandes.mdb" + "' select * from rebobinadores in '" + llegir_ini("General", "camibaixes", fitxerini) + "' where comanda in (select distinct comanda from llistatprodu)")
  Set rst = db2.OpenRecordset("select * from llistatprodu")
  While Not rst.EOF
    db.Execute ("delete * from tmprebobinadores where datafi=null")
    Set rst2 = db.OpenRecordset("select min(numeromaquina) as imp,min(operari1) as operari, min(cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn'))) as dia,sum(totalbobines) as totalb,sum(totalkilos) as totalk ,sum(totalmetres) as totalm from tmprebobinadores where comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst.Edit
    rst!maq = rst2!imp
    rst!operari = rst2!operari
    rst!totalbobines = rst2!totalb
    rst!totalkg = rst2!totalk
    rst!gramsm2 = 0
    rst!totalmtrs = rst2!totalm
    rst!dia = rst2!dia
    rst!hora = Format(rst2!dia, "hhnn")
    rst!horaformat = Format(rst!dia, "hh:nn")
    
    'busco el nom client
    Set rst2 = db.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(rst!client)), , dbReadOnly)
    If Not rst2.EOF Then rst!nomclient = atrim(rst2!nom)
    
    
    Set rst2 = db.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rst!mesuraesp)), , dbReadOnly)
    If Not rst2.EOF Then rst!descmesura = atrim(rst2!descripcio)
    
    Set rst2 = db.OpenRecordset("select sum(totalhores) as horesf from tmprebobinadores where tipus='F' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hfuncionament = cadbl(rst2!horesf)
    
    Set rst2 = db.OpenRecordset("select sum(totalhores) as horesa from tmprebobinadores where tipus='A' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!havaria = cadbl(rst2!horesa)
    Set rst2 = db.OpenRecordset("select sum(totalhores) as hmaquina from tmprebobinadores where tipus='M' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hmaquina = cadbl(rst2!hmaquina)
    
     Set rst2 = db.OpenRecordset("select sum(totalhores) as horesc from tmprebobinadores where tipus='C' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hcanvi = cadbl(rst2!horesc) + rst!havaria + rst!hmaquina
    rst!havaria = 0
    rst!hmaquina = 0
    
    
    
    If cadbl(rst!hfuncionament) < 1000 Then
      f = cadbl(rst!hfuncionament)
    Else: f = 0
    End If
    '(Int(v) * 60) + (((v - Int(v)) * 100) * 60 / 100)
    f = (Int(f) * 60) + (((f - Int(f)) * 100) * 60 / 100)
    rst!mtrsmin = 0
    If f > 0 Then rst!mtrsmin = cadbl(rst!totalmtrs) / f
    
    rst.Update
    rst.MoveNext
  Wend
 End If
 
 'si es laminadora faig els totals
 If r = "laminadora" Then
   t = IIf(vllarg <> "", " and " + vllarg, "") + IIf(vample <> "", " and " + vample, "") + IIf(vgruix <> "", " and " + vgruix, "")
   If t <> "" Then t = " where " + t
  db2.Execute ("select " + camps + camps2 + " into llistatprodu from tmpcomandes " + t)
  'db.Execute ("drop table tmplaminadores")
  db.Execute "delete * from tmplaminadores"
  db2.Execute ("insert into tmplaminadores in '" + rutadelfitxer(cami) + "comandes.mdb" + "' select * from laminadores in '" + llegir_ini("General", "camibaixes", fitxerini) + "' where comanda in (select distinct comanda from llistatprodu)")
  Set rst = db2.OpenRecordset("select * from llistatprodu")
  While Not rst.EOF
    db.Execute ("delete * from tmplaminadores where datafi=null")
    Set rst2 = db.OpenRecordset("select min(numeromaquina) as imp,min(operari) as operaris, min(cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn'))) as dia,sum(totalbobines) as totalb,sum(totalmetres) as totalm from tmplaminadores where comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst.Edit
    rst!maq = rst2!imp
    rst!operari = rst2!operaris
    rst!totalbobines = rst2!totalb
    rst!totalkg = 0
    rst!gramsm2 = 0
    rst!totalmtrs = rst2!totalm
    rst!dia = rst2!dia
    rst!hora = Format(rst2!dia, "hhnn")
    rst!horaformat = Format(rst!dia, "hh:nn")
    
    'busco el nom client
    Set rst2 = db.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(rst!client)), , dbReadOnly)
    If Not rst2.EOF Then rst!nomclient = atrim(rst2!nom)
    
    
    Set rst2 = dbbaixes.OpenRecordset("select grmmtr2 from laminadorestot where comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    If Not rst2.EOF Then
         rst!gramsm2 = cadbl(rst2!grmmtr2)
          Else: rst!gramsm2 = 0
    End If
    Set rst2 = db.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rst!mesuraesp)), , dbReadOnly)
    If Not rst2.EOF Then rst!descmesura = atrim(rst2!descripcio)
    
    Set rst2 = db.OpenRecordset("select sum(totalhores) as horesf from tmplaminadores where tipus='F' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hfuncionament = cadbl(rst2!horesf)
    
    Set rst2 = db.OpenRecordset("select sum(totalhores) as horesa from tmplaminadores where tipus='A' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!havaria = cadbl(rst2!horesa)
    Set rst2 = db.OpenRecordset("select sum(totalhores) as hmaquina from tmplaminadores where tipus='M' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hmaquina = cadbl(rst2!hmaquina)
    
     Set rst2 = db.OpenRecordset("select sum(totalhores) as horesc from tmplaminadores where tipus='C' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hcanvi = cadbl(rst2!horesc) + rst!havaria + rst!hmaquina
    rst!havaria = 0
    rst!hmaquina = 0
    
    
    
    If cadbl(rst!hfuncionament) < 1000 Then
      f = cadbl(rst!hfuncionament)
    Else: f = 0
    End If
    '(Int(v) * 60) + (((v - Int(v)) * 100) * 60 / 100)
    f = (Int(f) * 60) + (((f - Int(f)) * 100) * 60 / 100)
    rst!mtrsmin = 0
    If f > 0 Then rst!mtrsmin = cadbl(rst!totalmtrs) / f
    
    rst.Update
    rst.MoveNext
  Wend
 End If
 
 'si es sodladora faig els totals
 If r = "soldadora" Then
   t = IIf(vllarg <> "", " and " + vllarg, "") + IIf(vample <> "", " and " + vample, "") + IIf(vgruix <> "", " and " + vgruix, "")
   If t <> "" Then t = " where " + t
   'MsgBox camps
  db2.Execute ("select " + camps + camps2 + " into llistatprodu from tmpcomandes " + t)
  'db.Execute ("drop table tmpsoldadores")
  db.Execute "delete * from tmpsoldadores"
  db2.Execute ("insert into tmpsoldadores in '" + rutadelfitxer(cami) + "comandes.mdb" + "' select * from soldadores in '" + llegir_ini("General", "camibaixes", fitxerini) + "' where comanda in (select distinct comanda from llistatprodu)")
  Set rst = db2.OpenRecordset("select * from llistatprodu")
  While Not rst.EOF
    db.Execute ("delete * from tmpsoldadores where datafi=null")
    Set rst2 = db.OpenRecordset("select min(numeromaquina) as imp,min(operari1) as operari, min(cvdate(format(datafi,'dd/mm/yy')+' '+format(horafi,'hh:nn'))) as dia,sum(totalsacs) as totalb ,sum(totalunitats) as totalm from tmpsoldadores where comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst.Edit
    rst!maq = rst2!imp
    rst!operari = rst!operari
    rst!totalbobines = rst2!totalb
    rst!totalkg = 0
    rst!gramsm2 = 0
    rst!totalmtrs = rst2!totalm
    rst!dia = rst2!dia
    rst!hora = Format(rst2!dia, "hhnn")
    rst!horaformat = Format(rst!dia, "hh:nn")
    
    'busco el nom client
    Set rst2 = db.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(rst!client)), , dbReadOnly)
    If Not rst2.EOF Then rst!nomclient = atrim(rst2!nom)
    
    
    Set rst2 = db.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rst!mesuraesp)), , dbReadOnly)
    If Not rst2.EOF Then rst!descmesura = atrim(rst2!descripcio)
    ' APCF
    Set rst2 = db.OpenRecordset("select sum(totalhores) as horesf from tmpsoldadores where tipus='F' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hfuncionament = cadbl(rst2!horesf)
    
    Set rst2 = db.OpenRecordset("select sum(totalhores) as horesa from tmpsoldadores where tipus='A' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!havaria = cadbl(rst2!horesa)
    Set rst2 = db.OpenRecordset("select sum(totalhores) as hmaquina from tmpsoldadores where tipus='P' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hmaquina = cadbl(rst2!hmaquina)
    
     Set rst2 = db.OpenRecordset("select sum(totalhores) as horesc from tmpsoldadores where tipus='C' and comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    rst!hcanvi = cadbl(rst2!horesc) + rst!havaria + rst!hmaquina
    rst!havaria = 0
    rst!hmaquina = 0
    
    
    
    If cadbl(rst!hfuncionament) < 1000 Then
      f = cadbl(rst!hfuncionament)
    Else: f = 0
    End If
    '(Int(v) * 60) + (((v - Int(v)) * 100) * 60 / 100)
    'f = (Int(f) * 60) + (((f - Int(f)) * 100) * 60 / 100)
    rst!mtrsmin = 0
    If f > 0 Then rst!mtrsmin = cadbl(rst!totalmtrs) / f
    
    rst.Update
    rst.MoveNext
  Wend
 End If
 
 
 
 'si es muntadora faig els totals
 If r = "muntadora" Then
   t = IIf(vllarg <> "", " and " + vllarg, "") + IIf(vample <> "", " and " + vample, "") + IIf(vgruix <> "", " and " + vgruix, "")
   If t <> "" Then t = " where " + t
  db2.Execute ("select " + camps + camps2 + " into llistatprodu from tmpcomandes " + t)
  'db.Execute ("drop table tmpmuntadores")
  db.Execute "delete * from tmpmuntadores"
  db2.Execute ("insert into tmpmuntadores in '" + rutadelfitxer(cami) + "comandes.mdb" + "' select * from muntadores in '" + llegir_ini("General", "camibaixes", fitxerini) + "' where comanda in (select distinct comanda from llistatprodu)")
  Set rst = db2.OpenRecordset("select * from llistatprodu")
  While Not rst.EOF
    db.Execute ("delete * from tmpmuntadores where datafi=null")
    'Set rst2 = db.OpenRecordset("select min(numeromaquina) as imp, min(cvdate(format(datainici,'dd/mm/yy')+' '+format(horainici,'hh:nn'))) as dia,sum(totalsacs) as totalb ,sum(polimers) as totalm from tmpmuntadores where comanda=" + atrim(cadbl(rst!comanda)))
    Set rst2 = dbbaixes.OpenRecordset("SELECT muntadoratot.*, operari1,cvdate(format(datafi,'dd/mm/yy') +' '+format(muntadores.horafi,'hh:nn')) as dia FROM muntadores RIGHT JOIN muntadoratot ON muntadores.comanda = muntadoratot.comanda where muntadoratot.comanda=" + atrim(cadbl(rst!comanda)), , dbReadOnly)
    If Not rst2.EOF Then
     If atrim(rst2!dia) = "" Then MsgBox "La comanda " + atrim(rst!comanda) + " li falta alguna data o hora": GoTo cont
     rst.Edit
     rst!maq = 0
     
     rst!operari = rst2!operari1
     'rst!totalbobines = rst2!totalb
     rst!totalkg = 0
     rst!gramsm2 = 0
     rst!totalmtrs = cadbl(rst2!totalpolimers)
     rst!dia = rst2!dia
     rst!hora = Format(rst2!dia, "hhnn")
     rst!horaformat = Format(rst!dia, "hh:nn")
     rst!hfuncionament = cadbl(rst2!totalhores)
     If cadbl(rst2!totalhores) > 0 Then rst!mtrsmin = (cadbl(rst2!totalpolimers) * 60) / (cadbl(rst2!totalhores) * 60)
    End If
    'busco el nom client
    Set rst2 = db.OpenRecordset("select nom from clients where codi=" + atrim(cadbl(rst!client)), , dbReadOnly)
    If Not rst2.EOF Then rst!nomclient = atrim(rst2!nom)
    
    
    Set rst2 = db.OpenRecordset("select descripcio from mesureslineals where codi=" + atrim(cadbl(rst!mesuraesp)), , dbReadOnly)
    If Not rst2.EOF Then rst!descmesura = atrim(rst2!descripcio)
    ' APCF
    
    rst!havaria = 0
    rst!hmaquina = 0
    
    rst.Update
cont:
    rst.MoveNext
  Wend
 End If
 

  wait (4)
  If seccio = "I" Then
       vtotalrasquetesibandejes = calculrasquetesibandejes(inici, fi, cadbl(combooperaris.ItemData(combooperaris.ListIndex)), vmaquina)
       vtotalavarias = buscaravaries(inici, fi, cadbl(combooperaris.ItemData(combooperaris.ListIndex)), vmaquina, dbtmp)
 End If
 'tirar el llistat
 If vnomesgenerardades Then GoTo fi
 GoTo reportvell
 
 Dim oapp As CRAXDDRT.Application
  Dim oreport As CRAXDDRT.report
  
  Set oapp = New CRAXDDRT.Application
  Set oreport = oapp.OpenReport(llegir_ini("General", "rutallistats", fitxerini) + "temporal produccio.rpt", 1)
  oreport.Database.Tables.Item(1).Location = taulatemp
  'oreport.RecordSelectionFormula = "{aniloxosinformacio.id} in (SELECT First({aniloxos_informacio.id]) From {aniloxos_informacio} GROUP BY {aniloxos_informacio.matricula}"
  oreport.FormulaFields.GetItemByName("titol").Text = "'" + titol + "'"
  oreport.FormulaFields.GetItemByName("maquines").Text = "'" + treure_apostruf(maquines) + "Op: " + combooperaris.Text + " '"
  oreport.FormulaFields.GetItemByName("totalrasquetesibandejes").Text = "'" + vrasquetesibandejes + " '"
  'oreport.EnableParameterPrompting = False
  oreport.ExportOptions.DiskFileName = "c:\temp\prova.pdf"
  oreport.ExportOptions.DestinationType = crEDTDiskFile
  oreport.ExportOptions.FormatType = crEFTPortableDocFormat

   oreport.Export True
   GoTo fi
   
   Load veurereport
   veurereport.CRViewer.ReportSource = oreport
   veurereport.CRViewer.DisplayGroupTree = False
   veurereport.CRViewer.ViewReport
   veurereport.WindowState = 2
   veurereport.Show
 
 
  GoTo fi
reportvell:
 
  llistat.ReportFileName = llegir_ini("General", "rutallistats", fitxerini) + "temporal produccio.rpt"
  llistat.Destination = crptToWindow
  

  If Me.Tag = "1" Then
     triar_impressora
     llistat.Destination = crptToPrinter
  End If
  For i = 1 To 10
   llistat.SectionFormat(i) = ""
   llistat.Formulas(i) = ""
  Next i
  llistat.DataFiles(0) = taulatemp
  llistat.DiscardSavedData = True
  llistat.Formulas(0) = "titol='" + titol + "'"
  llistat.Formulas(1) = "maquines='" + treure_apostruf(maquines) + "Op: " + combooperaris.Text + " '"
  llistat.Formulas(2) = "totalrasquetesibandejes='" + vtotalrasquetesibandejes + " '"
  llistat.Formulas(3) = "totalavarias='" + Mid(treure_apostruf(vtotalavarias), 1, 250) + " '"
  
  
  
  'If previprint <> 1 Then If cadbl(llegir_ini("General", "programador", fitxerini)) = 0 Then llistat.Destination = crptToPrinter
  If CheckExcel.Value <> 1 And Checkexcelnou <> 1 Then llistat.WindowState = crptMaximized: llistat.Action = 1
  If CheckExcel.Value = 1 Then
         'exportar el llistat a CSV
         exportar_llistat_CSV titol, treure_apostruf(maquines) + "Op: " + combooperaris.Text, vtotalrasquetesibandejes, treure_apostruf(vtotalavarias), rst
  End If
  If Checkexcelnou.Value = 1 Then
         exportar_noullistat_csv titol, treure_apostruf(maquines) + "Op: " + combooperaris.Text, vtotalrasquetesibandejes, treure_apostruf(vtotalavarias), rst
  End If
 
fi:
 's´acava el llistat
 Me.Caption = "Listat de Produccions"
 ratoli "normal"
 Set rst2 = Nothing
 Set rst = Nothing
 Set db = Nothing
 Set db2 = Nothing
 'SET DBBAIXES = NOTHING
End Sub
Sub exportar_noullistat_csv(vtitol As String, vmaquines As String, vrasquetes As String, vavaries As String, rst As Recordset)
  Dim vnomfitxer As String
  Dim vlinia As String
  Dim vcont As Double
  Dim vhcanvi As Double
  Dim vhc As Double
  Dim vmtrsllençats As Double
  Dim vmll As Double
  If rst.EOF And rst.BOF Then MsgBox "No hi ha dades en aquesta consulta.", vbExclamation, "Atenció": GoTo fi
  Set dbtmpb = OpenDatabase(rutadelfitxer(cami) + "baixes.mdb", , True)
  vnomfitxer = "c:\temp\llistat_produccions_CSV.csv"
  Open vnomfitxer For Output As #1
  rst.MoveFirst
  vlinia = "Seccio;Nº Maq;Nº Op;Lot;Client;Nº Treball;Texte_Imp;Ample;Espessor;Cilindre;Tintes;Concepte;Tipificació;Data H.Inici;Data H.Fi;Temps descans;Temps paro;Temps Avaria;Temps Prep.Maq;Temps Ajust;Temps Func;Mt/min teoric;Mt/min Real;Mt Prova;Mt Merma;Mt Dolents;Mt Func;Mt Dif Real Teor.;Mt Dif Fulla Real;Nº Bobs;Kg Mat.;Nº Anil;Nº Rasq;Nº Rent;Kg Tinta;Observacions"
  Print #1, vtitol
  Print #1, vmaquines
  Print #1, vrasquetes
  Print #1, vavaries
  Print #1, ""
  Print #1, ""
  Print #1, vlinia
  While Not rst.EOF
    vcont = vcont + 1
    vhc = 0: vmll = 0
    vlinia = generar_linia_Nou_csv(rst, vhc, vmll)
    Print #1, "FER EL TOTAL DE LA COMANDA " + atrim(rst!comanda) + "   " + String(100, "-") 'linia de total de comanda vlinia
    vhcanvi = vhcanvi + vhc
    vmtrsllençats = vmtrsllençats + vmll
    rst.MoveNext
  Wend
  Print #1, ";;;TComandes:;" + atrim(vcont) + ";;;;;;;;Mig.H.Canvi:;" + atrim(Redondejar(vhcanvi / vcont, 2)) + ";;;;Mig.MtrsLlençar:;" + atrim(Redondejar(vmtrsllençats / vcont, 0))
  Close 1
  If existeix(vnomfitxer) Then obrir_document vnomfitxer
fi:

End Sub
Sub exportar_llistat_CSV(vtitol As String, vmaquines As String, vrasquetes As String, vavaries As String, rst As Recordset)
  Dim vnomfitxer As String
  Dim vlinia As String
  Dim vcont As Double
  Dim vhcanvi As Double
  Dim vhc As Double
  Dim vmtrsllençats As Double
  Dim vmll As Double
  If rst.EOF And rst.BOF Then MsgBox "No hi ha dades en aquesta consulta.", vbExclamation, "Atenció": GoTo fi
  vnomfitxer = "c:\temp\llistat_produccions_CSV.csv"
  Open vnomfitxer For Output As #1
  rst.MoveFirst
  vlinia = "Dia/Hora;Maquina;Operari;Client;Lot;Ample;Simulteineitat;Plegat;Longitud;Espessor;Mesura espessor;Tintes;G/M;Hores Canvi;Hores Avaria;Hores Funcionament;Bobines/Sacs;Total Kg;Metres llençats;Metres ajust;Metres dolents;Total metres;Metres minut"
  Print #1, vtitol
  Print #1, vmaquines
  Print #1, vrasquetes
  Print #1, vavaries
  Print #1, ""
  Print #1, ""
  Print #1, vlinia
  While Not rst.EOF
    vcont = vcont + 1
    vhc = 0: vmll = 0
    vlinia = generar_linia_csv(rst, vhc, vmll)
    Print #1, vlinia
    vhcanvi = vhcanvi + vhc
    vmtrsllençats = vmtrsllençats + vmll
    rst.MoveNext
  Wend
  Print #1, ";;;TComandes:;" + atrim(vcont) + ";;;;;;;;Mig.H.Canvi:;" + atrim(Redondejar(vhcanvi / vcont, 2)) + ";;;;Mig.MtrsLlençar:;" + atrim(Redondejar(vmtrsllençats / vcont, 0))
  Close 1
  If existeix(vnomfitxer) Then obrir_document vnomfitxer
fi:
End Sub
Function generar_linia_Nou_csv(rst As Recordset, vhcanvi As Double, vmtrsllençats As Double) As String
Dim v As String
  'Dim vhcanvi As String
  Dim vhavaria As String
  Dim vmtrsmin As Double
  Dim vfactorhoresminuts As Double
  Dim rstb As Recordset
  Dim rste As Recordset
  Dim vt As String
  
  Set rstb = dbtmpb.OpenRecordset("select * from impressores where comanda=" + atrim(rst!comanda) + " order by datainici,horainici")
  If rstb.EOF Then Exit Function
  Set rste = dbtmpb.OpenRecordset("SELECT comandes.comanda, Modificacions.id_treball, clients.nom, Modificacions.desarroll, Clixes.marca, Clixes.linia, comandes.cilindres, Modificacions.tinters FROM clients RIGHT JOIN (comandes LEFT JOIN (Clixes RIGHT JOIN Modificacions ON Clixes.id_treball = Modificacions.id_treball) ON (comandes.numordremodificacio = Modificacions.ordre) AND (comandes.numtreball = Modificacions.id_treball)) ON clients.codi = comandes.client WHERE (((comandes.comanda)=" + atrim(rstb!comanda) + "));")

  vfactorhoresminuts = IIf(rst!seccio = "S", 1, 60)
  vmtrsmin = IIf(rst!hfuncionament > 0, IIf(rst!seccio = "M", rst!mtrsmin, rst!totalmtrs / IIf((rst!hfuncionament * vfactorhoresminuts) <> 0, (rst!hfuncionament * vfactorhoresminuts), 1)), 0)
  vhavaria = 0
  If cadbl(rst!hmaquina) > 0 Then
       vhcanvi = rst!hmaquina
       vhavaria = atrim(rst!havaria)
      Else
         vhcanvi = atrim(Redondejar(rst!hcanvi, 1))
  End If
  v = "I;" + atrim(rst!maq) + ";" + atrim(rst!operari) + ";" + atrim(rst!comanda) + ";" + Mid(atrim(rste!nom) + " ", 1, 12) + ";" + atrim(rste!id_treball) + ";" + Mid(atrim(rste!linia) + " ", 1, 12) + ";"
  v = v + atrim(rst!ampleesq) + ";"
  v = v + atrim(rst!espessor) + ";"
  v = v + atrim(rste!cilindres) + ";"
  v = v + atrim(rste!tinters) + ";"
  While Not rstb.EOF
     vt = rstb!tipus + ";;" + atrim(rstb!datainici) + " " + atrim(Format(rstb!horainici, "hh:nn")) + ";" + atrim(rstb!datafi) + " " + atrim(Format(rstb!horafi, "hh:nn")) + ";;;"
     vt = vt + atrim(IIf(rstb!tipus = "V", cadbl(rstb!totalhores), 0)) + ";"
     vt = vt + atrim(IIf(rstb!tipus = "M", cadbl(rstb!totalhores), 0)) + ";"
     vt = vt + atrim(IIf(rstb!tipus = "A", cadbl(rstb!totalhores), 0)) + ";"
     vt = vt + atrim(IIf(rstb!tipus = "F", cadbl(rstb!totalhores), 0)) + ";"
     vt = vt + atrim(rstb!mtrsminut) + ";" + atrim(Redondejar(vmtrsmin, 0)) + ";" + atrim(rstb!mtrsprova) + ";;" + atrim(rst!tmetresdolents) + ";" + atrim(rstb!totalmetres)
     'fins aqui total metres funcionament
     Print #1, v + vt
     rstb.MoveNext
  Wend
  
  vmtrsllençats = rst!tmtrsllencats + rst!tmtrsajust '+ rst!tmetresdolents
  generar_linia_Nou_csv = v
End Function
Function generar_linia_csv(rst As Recordset, vhcanvi As Double, vmtrsllençats As Double) As String
  Dim v As String
  'Dim vhcanvi As String
  Dim vhavaria As String
  Dim vmtrsmin As Double
  Dim vfactorhoresminuts As Double
  vfactorhoresminuts = IIf(rst!seccio = "S", 1, 60)
  vmtrsmin = IIf(rst!hfuncionament > 0, IIf(rst!seccio = "M", rst!mtrsmin, rst!totalmtrs / IIf((rst!hfuncionament * vfactorhoresminuts) <> 0, (rst!hfuncionament * vfactorhoresminuts), 1)), 0)
  vhavaria = 0
  If cadbl(rst!hmaquina) > 0 Then
       vhcanvi = rst!hmaquina
       vhavaria = atrim(rst!havaria)
      Else
         vhcanvi = atrim(Redondejar(rst!hcanvi, 1))
  End If
  
  v = Format(rst!dia, "dd/mm/yyyy hh:nn") + ";"
  v = v + atrim(rst!maq) + ";"
  v = v + atrim(rst!operari) + ";"
  v = v + atrim(rst!nomclient) + ";"
  v = v + atrim(rst!comanda) + ";"
  v = v + atrim(rst!ampleesq) + ";"
  v = v + atrim(rst!sim) + ";"
  v = v + atrim(rst!plegatesq) + ";"
  v = v + atrim(rst!longitudsol) + ";"
  v = v + atrim(rst!espessor) + ";"
  v = v + atrim(rst!descmesura) + ";"
  v = v + atrim(rst!numerotintes) + ";"
  v = v + atrim(rst!gramsm2) + ";"
  v = v + atrim(vhcanvi) + ";"
  v = v + atrim(vhavaria) + ";"
  v = v + atrim(rst!hfuncionament) + ";"
  v = v + atrim(rst!totalbobines) + ";"
  v = v + atrim(rst!totalkg) + ";"
  If rst!seccio = "I" Then
    v = v + atrim(rst!tmtrsllencats) + ";"
    v = v + atrim(rst!tmtrsajust) + ";"
    v = v + atrim(rst!tmetresdolents) + ";"
    vmtrsllençats = rst!tmtrsllencats + rst!tmtrsajust '+ rst!tmetresdolents
      Else: v = v + "0;0;0;"
  End If
  v = v + atrim(rst!totalmtrs) + ";"
  v = v + atrim(Redondejar(vmtrsmin, 0))
  generar_linia_csv = v
End Function
Sub modificarcamps_tmpimpressores()
  On Error Resume Next
  db.Execute "alter table tmpimpressores add column operari2 double"
  On Error GoTo 0
End Sub
 Sub triar_impressora()
' Seleccionar la impresora a usar (23/Ene/00)

' La detección de errores es por si no hay impresora instalada
On Error Resume Next

With CommonDialog1
.DialogTitle = "Seleccionar impresora"
.flags = cdlPDPrintSetup
.ShowPrinter
End With

err = 0
End Sub
Function eval(am As String, vam As String) As String
   If atrim(am) = "" Then eval = "": Exit Function
   If InStr(1, am, "-") Then
        eval = vam + ">=" + Mid(am, 1, InStr(1, am, "-") - 1)
        eval = eval + " and " + vam + "<=" + Mid(am, InStr(1, am, "-") + 1, Len(am))
        eval = "(" + eval + ")"
   End If
   If InStr(1, am, "<") = 0 And InStr(1, am, ">") = 0 And InStr(1, am, "=") = 0 Then
        eval = vam + "=" + am
         Else: eval = vam + am
   End If
   
End Function

Private Sub combooperari_Change()

End Sub

Private Sub combooperaris_DropDown()
   DoEvents
   carregaroperaris
End Sub
Sub carregaroperaris()
   Dim rst As Recordset
   combooperaris.Clear
   Set rst = dbtmp.OpenRecordset("select * from operaris where maquina='" + atrim(seccio.Text) + "' and actiu=1")
   While Not rst.EOF
     combooperaris.AddItem atrim(rst!codi) + " - " + atrim(rst!descripcio)
     combooperaris.ItemData(combooperaris.NewIndex) = rst!codi
     rst.MoveNext
   Wend
   combooperaris.AddItem "Tots"
   combooperaris.ListIndex = combooperaris.NewIndex
   Set rst = Nothing
End Sub
Private Sub Command1_Click()
  Me.Tag = "0"
  llistat_produccions
End Sub

Private Sub Command2_Click()
  Dim ini As Date
  If IsDate(inici) Then
   ini = Format(DateAdd("m", 1, inici), "dd/mm/yy")
   inici = Format(ini, "dd/mm/yy")
   ini = DateAdd("m", 1, ini)
   fi = Format(DateSerial(Year(ini), Month(ini), 0), "dd/mm/yy") + " 06:00"
   
  End If
End Sub

Private Sub Command3_Click()
End Sub

Private Sub Form_Activate()
seccio_LostFocus
End Sub

Private Sub Form_Load()
 Dim ara As Date
 inici = "01/" + Format(Now, "mm/yy") + " 06:00"
 ara = DateAdd("m", 1, inici)
 fi = Format(DateSerial(Year(ara), Month(ara), 0), "dd/mm/yy") + " 06:00"
 carregaroperaris
End Sub

Private Sub maquina_Click()
   maquines = "Maquina: "
   For i = 0 To maquina.ListCount - 1
     If maquina.Selected(i) Then
       
       maquines = maquines + atrim(maquina.ItemData(i)) + ":" + atrim(maquina.List(i)) + "   "
     End If
  Next i
End Sub

Private Sub seccio_Click()
  seccio_LostFocus
  carregaroperaris
  If seccio = "I" Then Checkexcelnou.Value = 0: Checkexcelnou.Enabled = True Else Checkexcelnou.Value = 0: Checkexcelnou.Enabled = False
End Sub

Private Sub seccio_LostFocus()
 Dim db As Database
 Dim rst As Recordset
 Set db = OpenDatabase(rutadelfitxer(cami) + "comandes.mdb")
 Set rst = db.OpenRecordset("select * from maquines where isnull(donadadebaixa) and maquina='" + atrim(seccio.Text) + "'")
  ' where donadadebaixa=null and
 maquina.Clear
 carregaroperaris
 'combooperaris.Clear
 'combooperaris.Text = "Tots"
 While Not rst.EOF
   'If rst!maquina = atrim(seccio.Text) And IsNull(rst!donadadebaixa) Then
    If Mid(rst!descripcio + " ", 1, 1) <> "#" Then
      maquina.AddItem atrim(rst!descripcio)
      maquina.ItemData(maquina.NewIndex) = rst!codi
    End If
    'End If
    rst.MoveNext
 Wend
 'For i = 0 To maquina.ListCount - 1
 '   maquina.Selected(i) = True
 'Next i
End Sub

