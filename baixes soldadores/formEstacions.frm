VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form formEstacions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estacions Soldadora"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bmacules 
      BackColor       =   &H00FF80FF&
      Caption         =   "Màcules"
      Height          =   420
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   75
      Width           =   1665
   End
   Begin VB.CommandButton bSoldadors 
      BackColor       =   &H00FF80FF&
      Caption         =   "Soldadors"
      Height          =   420
      Left            =   1875
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   75
      Width           =   1665
   End
   Begin VB.CommandButton bAccessoris 
      BackColor       =   &H00FF80FF&
      Caption         =   "Accessoris"
      Height          =   420
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   1665
   End
   Begin MSFlexGridLib.MSFlexGrid reixa 
      Height          =   3270
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   5768
      _Version        =   393216
      BackColorFixed  =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "formEstacions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vRefInplacsa As String
Dim vComandaRef As String
Dim dbbaixesAnnex As Database


Private Sub bAccessoris_Click()
   carregar_reixa "Accessoris"
End Sub
Sub carregar_reixa(vtipus As String)
  Dim rst As Recordset
  Dim vultimcos As Byte
  Dim vnumcol As Byte
  Dim vNomCos As Variant
  Dim vDadesCarregades As Boolean
  vNomCos = Array("Guillotina", "Alim.Frontal", "Càmera", "PA", "PB", "PC", _
                "Refilet", "Màcula Frontal", "Sold. Refred. A", "Sold. Refred. B", _
                "Soldador A", "Soldador B", "Soldador C", "Soldador D", "Soldador E", "Ultrasó", "Alim.Central", "Màcula Central", "Sold.Long. Fondo", "Sold.Long. Zip", "Sold.Long. Z Fondo", "Sold.Long. Z Zip", _
                "Alim.Z", "Màcula final", "Alim. Final", "Triange", "Desbob.")
  
  
  If vtipus = "Accessoris" Then vtipus = "4,5,6"
  configurar_reixa
  Set rst = dbbaixesAnnex.OpenRecordset("select * from Estacions_tot where refinplacsa='Plantilla' and comanda=-1 and Num_cos in (" + vtipus + ") order by Num_cos,num_opcio")
Possar_Dades:
  vultimcos = 0
  vnumcol = 0 'columna inicial -1
  While Not rst.EOF
    If vultimcos <> rst!num_cos Then
          vnumcol = vnumcol + 1
          If vnumcol >= reixa.Cols Then reixa.Cols = vnumcol + 1
          reixa.col = vnumcol
          vultimcos = rst!num_cos
        'posso la casella ultima en groc que es observacio
          reixa.row = reixa.Rows - 1
          reixa.CellBackColor = QBColor(14)
    End If
    reixa.row = 1: reixa.CellAlignment = 3
    reixa.row = 0: reixa.CellAlignment = 3
    reixa.TextMatrix(0, vnumcol) = atrim(rst!num_cos)
    reixa.TextMatrix(1, vnumcol) = vNomCos(rst!num_cos - 1)
    reixa.TextMatrix(rst!num_opcio, vnumcol) = atrim(rst!valor_opcio)
    reixa.row = rst!num_opcio
    reixa.CellAlignment = 3
    If rst!valor_opcio = "-" Then reixa.CellBackColor = QBColor(14)
    rst.MoveNext
    
  Wend
  If Not vDadesCarregades Then
        Set rst = dbbaixesAnnex.OpenRecordset("select * from Estacions_tot where refinplacsa='" + vRefInplacsa + "' and comanda=" + atrim(vComandaRef) + " and Num_cos in (" + vtipus + ") order by Num_cos,num_opcio")
        If Not rst.EOF Then vDadesCarregades = True: GoTo Possar_Dades
  End If
  Set rst = Nothing
End Sub
Sub configurar_reixa()
  
                
  reixa.Clear
  reixa.Rows = 10
  reixa.Cols = 2
  reixa.col = 0
  reixa.ColWidth(0) = 1500
  reixa.TextMatrix(0, 0) = "Cos Núm.:"
  reixa.TextMatrix(1, 0) = "Descripció:"
  reixa.TextMatrix(2, 0) = "On/Off:"
  reixa.TextMatrix(3, 0) = "Codi Esq.:"
  reixa.TextMatrix(4, 0) = "Codi Dret:"
  reixa.TextMatrix(5, 0) = "Presió Esq:"
  reixa.TextMatrix(6, 0) = "Presió Dret:"
  reixa.TextMatrix(7, 0) = "Temp dalt:"
  reixa.TextMatrix(8, 0) = "Temp baix:"
  reixa.TextMatrix(9, 0) = "Observacions:"
End Sub
Private Sub Form_Load()
   vRefInplacsa = "01I1234PROVA"
   vComandaRef = 123456
   
   Set dbbaixesAnnex = OpenDatabase(rutadelfitxer(cami) + "baixes_annex.mdb")
   carregar_reixa "Accessoris"
End Sub
