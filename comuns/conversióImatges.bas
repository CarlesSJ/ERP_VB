Attribute VB_Name = "conversioimatges"
Public Sub ConvertirFormats(ByVal sP_Fichero_Origen_01 As String, _
                           Optional ByVal sFichero_Destino_01 As String, Optional vqualitat As Double)
    '
    '-------------------------------------------------§§§----'
    ' PROCEDIMIENTO PARA CONVERTIR UNA IMAGEN
    '
    '-------------------------------------------------§§§----'
    '
    Dim objImageMagik_01 As Object ' DEFINIMOS UN OBJECTO
    '
    ' CHEQUEO Y MAS CHEQUEOS… INTENTADO HACER LAS COSAS A PRUEBA DE BOMBAS….
    If cadbl(vqualitat) = 0 Then vqualitat = 150
    If (Dir(sP_Fichero_Origen_01) <> "") Then
        If Not (sFichero_Destino_01 <> "") Then
            sFichero_Destino_01 = sP_Fichero_Origen_01
        End If
        '
        Set objImageMagik_01 = CreateObject("ImageMagickObject.MagickImage.1") ' CREAMOS UN OBJECTO, SERA NUESTRO OBJETO IMAGEMAGICK
        'If vqualitat = 0 Then
       '      objImageMagik_01.Convert sP_Fichero_Origen_01, sFichero_Destino_01  ' LLAMAMOS A .CONVERT
       '     Else
                On Error GoTo errors
                objImageMagik_01.Convert "-density", vqualitat, sP_Fichero_Origen_01, sFichero_Destino_01
                'objImageMagik_01.Convert "-resize", "25%", "-density", 300, sP_Fichero_Origen_01, sFichero_Destino_01
                If InStr(1, UCase(sP_Fichero_Origen_01), ".JPG") > 0 And InStr(1, UCase(sFichero_Destino_01), ".JPG") > 0 Then
                    objImageMagik_01.Convert "-quality", 8, sP_Fichero_Origen_01, sFichero_Destino_01
                End If
       ' End If
                                                                                                           ' VEIS QUE SENCILLO ES ROTAR UN FICHERO
        '
        Set objImageMagik_01 = Nothing ' ELIMINAMOS O VACIAMOS EL OBJETO IMAGEMAGICK
        '
    Else
        MsgBox "ERROR_CONVERTIN_PDF A JPG" & vbCrLf & _
               vbCrLf & _
               vbCrLf & _
               "NO s'HA trobat el fitxer[" & sP_Fichero_Origen_01 & "]", vbCritical
        '
    End If
    '
    Exit Sub
errors:
    MsgBox "Hi ha hagut un error al fer la miniatura del fitxer " & sP_Fichero_Origen_01 + vbNewLine + " REVISA QUE EL PDF ESTIGUI CORRECTE."
End Sub
Public Sub retallarimatgeFitxer(ByVal sP_Fichero_Origen_01 As String, _
                           Optional ByVal sFichero_Destino_01 As String, Optional vtanxcent As Byte)
    '
    '-------------------------------------------------§§§----'
    ' PROCEDIMIENTO treure les parts sobrants dels costats de la imatge AUTOCROP
    '
    '-------------------------------------------------§§§----'
    '
    Dim objImageMagik_01 As Object ' DEFINIMOS UN OBJECTO
    '
    ' CHEQUEO Y MAS CHEQUEOS… INTENTADO HACER LAS COSAS A PRUEBA DE BOMBAS….
    If Not existeix(sP_Fichero_Origen_01) Then Exit Sub
    If FileLen(sP_Fichero_Origen_01) < 10 Then Exit Sub
    If cadbl(vtanxcent) = 0 Then vtanxcent = 80
    If (Dir(sP_Fichero_Origen_01) <> "") Then
        If Not (sFichero_Destino_01 <> "") Then
            sFichero_Destino_01 = sP_Fichero_Origen_01
        End If
        '
        
        Set objImageMagik_01 = CreateObject("ImageMagickObject.MagickImage.1") ' CREAMOS UN OBJECTO, SERA NUESTRO OBJETO IMAGEMAGICK
        '
       ' objImageMagik_01.Convert sP_Fichero_Origen_01, sFichero_Destino_01  ' LLAMAMOS A .CONVERT
       'convert image -fuzz 1% -trim +repage result
       
       objImageMagik_01.Convert "-fuzz", atrim(vtanxcent) + "%", "-trim", sP_Fichero_Origen_01, sFichero_Destino_01
'        objImageMagik_01.convert "-density", 150, sP_Fichero_Origen_01, sFichero_Destino_01
        'objImageMagik_01.Convert "-resize", "25%", "-density", 300, sP_Fichero_Origen_01, sFichero_Destino_01
        'convert -density 150 -quality 95
                                                                                                           ' VEIS QUE SENCILLO ES ROTAR UN FICHERO
        '
        Set objImageMagik_01 = Nothing ' ELIMINAMOS O VACIAMOS EL OBJETO IMAGEMAGICK
        '
    Else
        MsgBox "ERROR_CONVERTIN_PDF A JPG" & vbCrLf & _
               vbCrLf & _
               vbCrLf & _
               "NO s'HA trobat el fitxer[" & sP_Fichero_Origen_01 & "]", vbCritical
        '
    End If
    '
End Sub





Public Sub RotarImatge(ByVal sP_Fichero_Origen_01 As String, _
                           Optional ByVal sFichero_Destino_01 As String, _
                           Optional ByVal fP_Angulo_01 As Double = 90)
    '
    '-------------------------------------------------§§§----'
    ' PROCEDIMIENTO PARA ROTAR UNA IMAGEN
    '
    ' POR DEFECTO LA IMAGEN ES ROTADA 90, SEGUN LAS AGUJAS DEL RELOJ
    '-------------------------------------------------§§§----'
    '
    Dim objImageMagik_01 As Object ' DEFINIMOS UN OBJECTO
    '
    ' CHEQUEO Y MAS CHEQUEOS… INTENTADO HACER LAS COSAS A PRUEBA DE BOMBAS….
    If Not existeix(sP_Fichero_Origen_01) Then Exit Sub
    If (Dir(sP_Fichero_Origen_01) <> "") Then
        If Not (sFichero_Destino_01 <> "") Then
            sFichero_Destino_01 = sP_Fichero_Origen_01
        End If
        '
        Set objImageMagik_01 = CreateObject("ImageMagickObject.MagickImage.1") ' CREAMOS UN OBJECTO, SERA NUESTRO OBJETO IMAGEMAGICK
        '
        objImageMagik_01.Convert "-rotate", CStr(fP_Angulo_01), sP_Fichero_Origen_01, sFichero_Destino_01 '" ' LLAMAMOS A .CONVERT
        'objImageMagik_01.Convert "-fx", "p{w-i-1,j}", sFichero_Destino_01, sFichero_Destino_01
        'objImageMagik_01.Convert "-flop", sFichero_Destino_01, sFichero_Destino_01
                                                                                                           ' VEIS QUE SENCILLO ES ROTAR UN FICHERO
        '
        Set objImageMagik_01 = Nothing ' ELIMINAMOS O VACIAMOS EL OBJETO IMAGEMAGICK
        '
    Else
        MsgBox "ERROR_Rotar_Imagen_01_01" & vbCrLf & _
               vbCrLf & _
               "ERROR AL EJECUTAR COMANDO DE ROTACION DE IMAGEN" & vbCrLf & _
               vbCrLf & _
               "NO SE HA ENCONTRADO EL FICHERO DE DESTINO [" & sP_Fichero_Origen_01 & "]", vbCritical
        '
    End If
    '
End Sub

Public Sub TallarImatge(ByVal sP_Fichero_Origen_01 As String, _
                           Optional ByVal sFichero_Destino_01 As String, _
                           Optional ByVal vX As Double = 10, Optional ByVal vY As Double = 10)
    '
    '-------------------------------------------------§§§----'
    ' PROCEDIMIENTO PARA cortar UNA IMAGEN
    '
    
    '-------------------------------------------------§§§----'
    '
    Dim objImageMagik_01 As Object ' DEFINIMOS UN OBJECTO
    '
    ' CHEQUEO Y MAS CHEQUEOS… INTENTADO HACER LAS COSAS A PRUEBA DE BOMBAS….
    If Not existeix(sP_Fichero_Origen_01) Then Exit Sub
    If (Dir(sP_Fichero_Origen_01) <> "") Then
        If Not (sFichero_Destino_01 <> "") Then
            sFichero_Destino_01 = sP_Fichero_Origen_01
        End If
        '
        Set objImageMagik_01 = CreateObject("ImageMagickObject.MagickImage.1") ' CREAMOS UN OBJECTO, SERA NUESTRO OBJETO IMAGEMAGICK
        '
        objImageMagik_01.Convert "-crop", "+" + atrim(vX) + "+" + atrim(vY), sP_Fichero_Origen_01, sFichero_Destino_01 '" ' LLAMAMOS A .CONVERT
        objImageMagik_01.Convert "+repage", sP_Fichero_Origen_01, sFichero_Destino_01  '" ' LLAMAMOS A .CONVERT
        
        'objImageMagik_01.Convert "-fx", "p{w-i-1,j}", sFichero_Destino_01, sFichero_Destino_01
        'objImageMagik_01.Convert "-flop", sFichero_Destino_01, sFichero_Destino_01
                                                                                                           ' VEIS QUE SENCILLO ES ROTAR UN FICHERO
        '
        Set objImageMagik_01 = Nothing ' ELIMINAMOS O VACIAMOS EL OBJETO IMAGEMAGICK
        '
    Else
        MsgBox "ERROR_Rotar_Imagen_01_01" & vbCrLf & _
               vbCrLf & _
               "ERROR AL EJECUTAR COMANDO DE ROTACION DE IMAGEN" & vbCrLf & _
               vbCrLf & _
               "NO SE HA ENCONTRADO EL FICHERO DE DESTINO [" & sP_Fichero_Origen_01 & "]", vbCritical
        '
    End If
    '
End Sub


Public Sub InvertirHVImatge(ByVal sP_Fichero_Origen_01 As String, _
                           Optional ByVal sFichero_Destino_01 As String, Optional vVertical As Boolean)
    '
    '-------------------------------------------------§§§----'
    ' PROCEDIMIENTO PARA ROTAR UNA IMAGEN
    '
    ' POR DEFECTO LA IMAGEN ES ROTADA 90, SEGUN LAS AGUJAS DEL RELOJ
    '-------------------------------------------------§§§----'
    '
    Dim objImageMagik_01 As Object ' DEFINIMOS UN OBJECTO
    '
    ' CHEQUEO Y MAS CHEQUEOS… INTENTADO HACER LAS COSAS A PRUEBA DE BOMBAS….
    If Not existeix(sP_Fichero_Origen_01) Then Exit Sub
    If (Dir(sP_Fichero_Origen_01) <> "") Then
        If Not (sFichero_Destino_01 <> "") Then
            sFichero_Destino_01 = sP_Fichero_Origen_01
        End If
        '
        Set objImageMagik_01 = CreateObject("ImageMagickObject.MagickImage.1") ' CREAMOS UN OBJECTO, SERA NUESTRO OBJETO IMAGEMAGICK
        '
        'objImageMagik_01.Convert "-rotate", CStr(fP_Angulo_01), sP_Fichero_Origen_01, sFichero_Destino_01 '" ' LLAMAMOS A .CONVERT
        'objImageMagik_01.Convert "-fx", "p{w-i-1,j}", sFichero_Destino_01, sFichero_Destino_01
        If vVertical Then
               objImageMagik_01.Convert "-flip", sP_Fichero_Origen_01, sFichero_Destino_01
           Else
             objImageMagik_01.Convert "-flop", sP_Fichero_Origen_01, sFichero_Destino_01
        End If
                                                                                                           ' VEIS QUE SENCILLO ES ROTAR UN FICHERO
        '
        Set objImageMagik_01 = Nothing ' ELIMINAMOS O VACIAMOS EL OBJETO IMAGEMAGICK
        '
    Else
        MsgBox "ERROR_Rotar_Imagen_01_01" & vbCrLf & _
               vbCrLf & _
               "ERROR AL EJECUTAR COMANDO DE ROTACION DE IMAGEN" & vbCrLf & _
               vbCrLf & _
               "NO SE HA ENCONTRADO EL FICHERO DE DESTINO [" & sP_Fichero_Origen_01 & "]", vbCritical
        '
    End If
    '
End Sub


