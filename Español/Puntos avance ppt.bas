Attribute VB_Name = "Puntos_avance"
' Declaración de variables globales
Public colorAvanzado As Long
Public colorPendiente As Long
Public bordeAvanzado As Long
Public bordePendiente As Long
Public grosorBordeCirculos As Single
Public radius As Single
Public spacing As Single
Public puntoAltura As Single

Sub InicializarVariables()
    ' Inicialización de las variables globales
    colorAvanzado = RGB(0, 0, 0) ' Color para las diapositivas avanzadas
    bordeAvanzado = RGB(256, 256, 256) ' Color del borde para las diapositivas avanzadas

    colorPendiente = RGB(256, 256, 256) ' Color para las diapositivas pendientes
    bordePendiente = RGB(0, 0, 0) ' Color del borde para las diapositivas pendientes

    grosorBordeCirculos = 0.025 ' Grosor del borde en mm para todos los círculos

    radius = 5 ' Radio del círculo
    spacing = 10 ' Espacio entre los puntos
    puntoAltura = 50 - (5.5 * 2.835) ' Altura a la que se presentan los puntos
End Sub

Sub DibujarPuntos()
    Dim sld As Slide
    Dim totalSlides As Integer
    Dim currentSlide As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim totalWidth As Single
    Dim i As Integer

    ' Asegurarse de que las variables estén inicializadas
    Call InicializarVariables

    totalSlides = ActivePresentation.Slides.Count

    ' Calcular el ancho total de los puntos y espacios
    totalWidth = (totalSlides * (radius * 2)) + ((totalSlides - 1) * spacing)

    ' Calcular la posición inicial en X para centrar los puntos
    xPos = (ActivePresentation.PageSetup.SlideWidth - totalWidth) / 2
    yPos = ActivePresentation.PageSetup.SlideHeight - puntoAltura ' Posición en Y ajustable

    ' Borrar puntos anteriores
    Call BorrarPuntos_TodasLasDiapositivas

        ' Dibujar puntos en todas las diapositivas
    For Each sld In ActivePresentation.Slides
        currentSlide = sld.SlideIndex
        
        ' Dibujar puntos para las diapositivas pasadas y la actual
        For i = 1 To currentSlide
            With sld.Shapes.AddShape(msoShapeOval, xPos, yPos, radius * 2, radius * 2)
                .Fill.ForeColor.RGB = colorAvanzado
                .Line.ForeColor.RGB = bordeAvanzado
                .Line.Weight = grosorBordeCirculos / 0.0352778 ' Convertir mm a puntos
                .Name = "ProgressDot"
            End With
            xPos = xPos + (radius * 2) + spacing ' Incremento en X para el siguiente punto
        Next i

        ' Dibujar puntos para las diapositivas restantes
        For i = currentSlide + 1 To totalSlides
            With sld.Shapes.AddShape(msoShapeOval, xPos, yPos, radius * 2, radius * 2)
                .Fill.ForeColor.RGB = colorPendiente
                .Line.ForeColor.RGB = bordePendiente
                .Line.Weight = grosorBordeCirculos / 0.0352778 ' Convertir mm a puntos Jose Chirif
                .Name = "ProgressDot"
            End With
            xPos = xPos + (radius * 2) + spacing ' Incremento en X para el siguiente punto
        Next i

        ' Restablecer la posición inicial en X para la siguiente diapositiva
        xPos = (ActivePresentation.PageSetup.SlideWidth - totalWidth) / 2
    Next sld
    

    
    
End Sub

Sub BorrarPuntos_TodasLasDiapositivas()
    Dim sld As Slide
    Dim shp As Shape
    ' Limpiar puntos anteriores generados por esta macro en todas las diapositivas
    For Each sld In ActivePresentation.Slides
        For i = sld.Shapes.Count To 1 Step -1
            Set shp = sld.Shapes(i)
            If shp.Name = "ProgressDot" Then
                shp.Delete
            End If
        Next i
    Next sld
End Sub

Sub BorrarPuntos_EstaDiapositiva()
    Dim sld As Slide
    Dim shp As Shape
    ' Limpiar puntos anteriores generados por esta macro en la diapositiva actual
    Set sld = Application.ActiveWindow.View.Slide
    For i = sld.Shapes.Count To 1 Step -1
        Set shp = sld.Shapes(i)
        If shp.Name = "ProgressDot" Then
            shp.Delete
        End If
    Next i
End Sub

Sub EliminarPrimerPuntoYCentrar()
    Dim sld As Slide
    Dim shp As Shape
    Dim totalSlides As Integer
    Dim xPos As Single
    Dim totalWidth As Single
    Dim pointsCount As Integer
    Dim i As Integer
    Dim remainingDots As Integer

    ' Asegurarse de que las variables estén inicializadas
    Call InicializarVariables

    totalSlides = ActivePresentation.Slides.Count

    ' Iterar sobre cada diapositiva para eliminar el primer punto
    For Each sld In ActivePresentation.Slides
        pointsCount = 0
        
        ' Contar los puntos "ProgressDot"
        For i = 1 To sld.Shapes.Count
            If sld.Shapes(i).Name = "ProgressDot" Then
                pointsCount = pointsCount + 1
            End If
        Next i
        
        ' Eliminar el primer punto encontrado
        If pointsCount > 0 Then
            For i = 1 To sld.Shapes.Count
                If sld.Shapes(i).Name = "ProgressDot" Then
                    sld.Shapes(i).Delete ' Eliminar el primer punto
                    Exit For
                End If
            Next i
        End If
        
        ' Calcular el número de puntos restantes
        remainingDots = pointsCount - 1
        
        ' Recalcular la posición inicial en X para centrar los puntos restantes
        totalWidth = (remainingDots * radius * 2) + ((remainingDots - 1) * spacing)
        xPos = (ActivePresentation.PageSetup.SlideWidth - totalWidth) / 2
        
        ' Reajustar la posición de los puntos restantes
        pointsCount = 0
        For i = 1 To sld.Shapes.Count
            Set shp = sld.Shapes(i)
            If shp.Name = "ProgressDot" Then
                shp.Left = xPos + (pointsCount * (radius * 2 + spacing))
                pointsCount = pointsCount + 1
            End If
        Next i
    Next sld
End Sub


Sub EliminarUltimoPuntoYCentrar()
    Dim sld As Slide
    Dim shp As Shape
    Dim totalSlides As Integer
    Dim xPos As Single
    Dim totalWidth As Single
    Dim pointsCount As Integer
    Dim i As Integer

    ' Asegurarse de que las variables estén inicializadas
    Call InicializarVariables

    totalSlides = ActivePresentation.Slides.Count

    ' Iterar sobre cada diapositiva para eliminar el último punto
    For Each sld In ActivePresentation.Slides
        pointsCount = 0
        
        ' Contar los puntos "ProgressDot"
        For i = 1 To sld.Shapes.Count
            If sld.Shapes(i).Name = "ProgressDot" Then
                pointsCount = pointsCount + 1
            End If
        Next i
        
        ' Eliminar el último punto encontrado
        If pointsCount > 0 Then
            sld.Shapes(sld.Shapes.Count).Delete
        End If
        
        ' Recalcular la posición inicial en X para centrar los puntos restantes
        totalWidth = ((pointsCount - 1) * (radius * 2)) + ((pointsCount - 2) * spacing)
        xPos = (ActivePresentation.PageSetup.SlideWidth - totalWidth) / 2
        
        ' Reajustar la posición de los puntos restantes
        pointsCount = 0
        For i = 1 To sld.Shapes.Count
            Set shp = sld.Shapes(i)
            If shp.Name = "ProgressDot" Then
                shp.Left = xPos + (pointsCount * (radius * 2 + spacing))
                pointsCount = pointsCount + 1
            End If
        Next i
    Next sld
End Sub



Sub EliminarPuntosPrimeraDiapositiva()
    Dim sld As Slide
    Dim shp As Shape
    Dim i As Integer

    ' Obtener la primera diapositiva
    Set sld = ActivePresentation.Slides(1)

    ' Recorrer las formas de la diapositiva y eliminar los puntos "ProgressDot"
    For i = sld.Shapes.Count To 1 Step -1
        If sld.Shapes(i).Name = "ProgressDot" Then
            sld.Shapes(i).Delete
        End If
    Next i
End Sub

Sub EliminarPuntosUltimaDiapositiva()
    Dim sld As Slide
    Dim shp As Shape
    Dim i As Integer
    Dim totalSlides As Integer

    ' Obtener el número total de diapositivas
    totalSlides = ActivePresentation.Slides.Count

    ' Obtener la última diapositiva
    Set sld = ActivePresentation.Slides(totalSlides)

    ' Recorrer las formas de la diapositiva y eliminar los puntos "ProgressDot"
    For i = sld.Shapes.Count To 1 Step -1
        If sld.Shapes(i).Name = "ProgressDot" Then
            sld.Shapes(i).Delete
        End If
    Next i
End Sub

