VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 
Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Load()
 Dim comanda As String
 Dim dbstocks As Database
 Dim dbcomanda As Database
 Dim rststocks As Recordset
 Dim rsttmp As Recordset
 Dim rsttmp2 As Recordset
 Dim dbbaixes As Database
 Dim rstbaixes As Recordset
 Dim camistocks As String
 Dim camibaixes As String
 
 Dim mtrs As Double
 Dim id As Double
 Dim ruta As String
 Dim estat As String
 comanda = Command
  '  comanda = "122088"
 If Not IsNumeric(comanda) Then End
 r = llegir_ini("General", "ruta_stocksmdb", "comandes.ini")
 If r = "{[}]" Then
    escriure_ini "General", "ruta_stocksmdb", "\\Ser2\documentos\Stock Reclamaciones\Estoc inplacsa.mdb", "comandes.ini"
    r = llegir_ini("General", "ruta_stocksmdb", "comandes.ini")
 End If
 camistocks = r
 camibaixes = llegir_ini("General", "camibaixes", "comandes.ini")
 Set dbcomanda = OpenDatabase(llegir_ini("General", "cami", "comandes.ini"))
    
 Set dbstocks = OpenDatabase(camistocks, , True)
 Set dbbaixes = OpenDatabase(camibaixes)
 Set rststocks = dbstocks.OpenRecordset("select idpalet,mts from bobines where numcom='" + comanda + "'")
 
 'borro la seccio d'extrussora de la comanda en concret
 Set rsttmp = dbbaixes.OpenRecordset("select id from extrussores where comanda=" + aTrim(cadbl(comanda)))
 While Not rsttmp.EOF
     dbbaixes.Execute "delete * from bobinesext where controlid=" + aTrim(rsttmp!id)
     rsttmp.MoveNext
 Wend
 dbbaixes.Execute "delete * from extrussores where comanda=" + aTrim(cadbl(comanda))
' dbcomanda.Execute "update comandes set proximaseccio='E' where comanda=" + aTrim(cadbl(comanda))
 
 'fins aqui borra seccio baixa
 
 'faig l'alta de la seccio d'extrussores a baixes
 dbbaixes.Execute "insert into extrussores (comanda,tipus) values (" + comanda + ",'F')"
     
  
  Set rsttmp = dbbaixes.OpenRecordset("select id from extrussores where comanda=" + aTrim(cadbl(comanda)))
  id = cadbl(rsttmp!id)
  Set rsttmp = dbstocks.OpenRecordset("select idprod,ample,plegat,solapa from palets where idpalet=" + aTrim(cadbl(rststocks!idpalet)))
  cont = 1
  While Not rststocks.EOF
     Set rsttmp2 = dbstocks.OpenRecordset("select Grmt2 from productes where IdProd=" + aTrim(cadbl(rsttmp!IdProd)))
     r = ((rsttmp!Ample + (rsttmp!Solapa / 2)) / 100) * (rsttmp2!grmt2 / 1000) * (rststocks!mts)
     r = Format(r, "#,##0.00")
     dbbaixes.Execute "insert into bobinesext (controlid,metres,kilos,numerodebobina) values (" + aTrim(id) + "," + aTrim(cadbl(rststocks!mts)) + "," + passardecomaapunt(cadbl(r)) + "," + Trim(cont) + ")"
     cont = cont + 1
     rststocks.MoveNext
  Wend
  
 'fins aqui l'alta de seccio
 'si hi ha alguna bobina passo l'estat de la comanda a la proxima seccio
 If cont > 1 Then
   'passo l'estat de comanda a la proxima
   Set rsttmp = dbcomanda.OpenRecordset("select producte,proximaseccio from comandes where comanda=" + aTrim(comanda))
   If Not rsttmp.EOF Then
     estat = aTrim(rsttmp!proximaseccio)
     If estat = "" Then estat = "E"
   End If
   Set rsttmp = dbcomanda.OpenRecordset("select ruta from productes where codi='" + rsttmp!producte + "'")
   ruta = rsttmp!ruta
   If estat = "E" Then
     seccio = Mid(ruta, 2, 1)
     If seccio = "" Then seccio = "V"
     dbcomanda.Execute "update comandes set seccioactual='E' where comanda=" + aTrim(comanda)
     dbcomanda.Execute "update comandes set proximaseccio='" + seccio + "' where comanda=" + aTrim(comanda)
   End If
 End If
 Set dbbaixes = Nothing
 Set dbcomanda = Nothing
 Set rsttmp = Nothing
 Set rststocks = Nothing
 Set dbcomanda = Nothing
 Set dbstocks = Nothing
 End
End Sub
Function passardecomaapunt(valo As String) As String
   While InStr(1, valo, ",")
      valo = Mid(valo, 1, InStr(1, valo, ",") - 1) + "." + Mid(valo, InStr(1, valo, ",") + 1, Len(valo))
   Wend
   passardecomaapunt = valo
End Function
