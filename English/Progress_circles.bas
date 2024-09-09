Attribute VB_Name = "Progress_circles"
 ' Global variables declaration
Public ProgressCircleFillColor As Long
Public RemainingSlidesCircleFillColor As Long
Public ProgressCircleBorderColor As Long
Public RemainingSlidesCircleBorderColor As Long
Public CircleBorderWidth As Single
Public radius As Single
Public spacing As Single
Public CircleHeight As Single

Sub InitializeVariables()
    ' Global variables initialization
    ProgressCircleFillColor = RGB(0, 0, 0) ' Color of the circles fill that simulate advanced slides
    ProgressCircleBorderColor = RGB(256, 256, 256) ' Color of the circles border that simulate advanced slides

    RemainingSlidesCircleFillColor = RGB(256, 256, 256) ' Color of the circles fill that remaining slides
    RemainingSlidesCircleBorderColor = RGB(0, 0, 0) ' Color of the circles border that remaining slides

    CircleBorderWidth = 0.025 ' Border size in mm for all circles

    radius = 5 ' Radius of all circles
    spacing = 10 ' Spacing between circles
    CircleHeight = 50 - (5.5 * 2.835) ' Height at which the circles are displayed
End Sub

Sub DrawCircles()
    Dim sld As Slide
    Dim totalSlides As Integer
    Dim currentSlide As Integer
    Dim xPos As Single
    Dim yPos As Single
    Dim totalWidth As Single
    Dim i As Integer

    ' Ensure that variables are initialized
    Call InitializeVariables

    totalSlides = ActivePresentation.Slides.Count

    '  Calculate total width of circles and spaces
    totalWidth = (totalSlides * (radius * 2)) + ((totalSlides - 1) * spacing)

    ' Calculate the initial position in X to center the points.
    xPos = (ActivePresentation.PageSetup.SlideWidth - totalWidth) / 2
    yPos = ActivePresentation.PageSetup.SlideHeight - CircleHeight ' Adjustable Y-position

    ' Delete progress circles drawed before
    Call DeleteCircles_AllSlides

        ' Draw progress circles on all slides
    For Each sld In ActivePresentation.Slides
        currentSlide = sld.SlideIndex
        
        ' Draw advanced progress circles for the previous and the current slide
        For i = 1 To currentSlide
            With sld.Shapes.AddShape(msoShapeOval, xPos, yPos, radius * 2, radius * 2)
                .Fill.ForeColor.RGB = ProgressCircleFillColor
                .Line.ForeColor.RGB = ProgressCircleBorderColor
                .Line.Weight = CircleBorderWidth / 0.0352778 ' Converte circles in mm
                .Name = "ProgressDot"
            End With
            xPos = xPos + (radius * 2) + spacing ' Increase in X for the next circle
        Next i

        ' Drawing next progress circles for the remaining slides
        For i = currentSlide + 1 To totalSlides
            With sld.Shapes.AddShape(msoShapeOval, xPos, yPos, radius * 2, radius * 2)
                .Fill.ForeColor.RGB = RemainingSlidesCircleFillColor
                .Line.ForeColor.RGB = RemainingSlidesCircleBorderColor
                .Line.Weight = CircleBorderWidth / 0.0352778 ' Convert circles border to mm     Jose Chirif
                .Name = "ProgressDot"
            End With
            xPos = xPos + (radius * 2) + spacing ' Increase in X for the next circle
        Next i

        ' Restore the initial position in X for the next slide
        xPos = (ActivePresentation.PageSetup.SlideWidth - totalWidth) / 2
    Next sld
    

    
    
End Sub

Sub DeleteCircles_AllSlides()
    Dim sld As Slide
    Dim shp As Shape
    ' Delete circles that have been done with the DrawCircles macro
    For Each sld In ActivePresentation.Slides
        For i = sld.Shapes.Count To 1 Step -1
            Set shp = sld.Shapes(i)
            If shp.Name = "ProgressDot" Then
                shp.Delete
            End If
        Next i
    Next sld
End Sub

Sub DeleteCircles_CurrentSlide()
    Dim sld As Slide
    Dim shp As Shape
    ' Delete circles that have been done with the DrawCircles macro only in the current slide
    Set sld = Application.ActiveWindow.View.Slide
    For i = sld.Shapes.Count To 1 Step -1
        Set shp = sld.Shapes(i)
        If shp.Name = "ProgressDot" Then
            shp.Delete
        End If
    Next i
End Sub

Sub DeleteFirstCircleAndCenter()
    Dim sld As Slide
    Dim shp As Shape
    Dim totalSlides As Integer
    Dim xPos As Single
    Dim totalWidth As Single
    Dim pointsCount As Integer
    Dim i As Integer
    Dim remainingDots As Integer

    '  Ensure that variables are initialized
    Call InitializeVariables

    totalSlides = ActivePresentation.Slides.Count

    ' Iterate over each slide to remove the first circle
    For Each sld In ActivePresentation.Slides
        pointsCount = 0
        
        ' Count the progress circles
        For i = 1 To sld.Shapes.Count
            If sld.Shapes(i).Name = "ProgressDot" Then
                pointsCount = pointsCount + 1
            End If
        Next i
        
        ' Delete the first progress circle
        If pointsCount > 0 Then
            For i = 1 To sld.Shapes.Count
                If sld.Shapes(i).Name = "ProgressDot" Then
                    sld.Shapes(i).Delete ' Eliminar el primer punto
                    Exit For
                End If
            Next i
        End If
        
        ' Count the progress circles
        remainingDots = pointsCount - 1
        
        ' Recalculate the initial position in X to center the remaining points.
        totalWidth = (remainingDots * radius * 2) + ((remainingDots - 1) * spacing)
        xPos = (ActivePresentation.PageSetup.SlideWidth - totalWidth) / 2
        
        ' Readjusting the position of the progress circles
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


Sub DeleteLastCircleAndCenter()
    Dim sld As Slide
    Dim shp As Shape
    Dim totalSlides As Integer
    Dim xPos As Single
    Dim totalWidth As Single
    Dim pointsCount As Integer
    Dim i As Integer

    '  Ensure that variables are initialized
    Call InitializeVariables

    totalSlides = ActivePresentation.Slides.Count

    ' Iterate over each slide to remove the last circle
    For Each sld In ActivePresentation.Slides
        pointsCount = 0
        
        ' Count the progress circles
        For i = 1 To sld.Shapes.Count
            If sld.Shapes(i).Name = "ProgressDot" Then
                pointsCount = pointsCount + 1
            End If
        Next i
        
        ' Delete the last progress circle
        If pointsCount > 0 Then
            sld.Shapes(sld.Shapes.Count).Delete
        End If
        
        ' Recalculate the initial position in X to center the remaining points.
        totalWidth = ((pointsCount - 1) * (radius * 2)) + ((pointsCount - 2) * spacing)
        xPos = (ActivePresentation.PageSetup.SlideWidth - totalWidth) / 2
        
        ' Readjusting the position of the progress circles
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



Sub DeleteAllCirclesInFirstSlide()
    Dim sld As Slide
    Dim shp As Shape
    Dim i As Integer

    ' Get teh first slide
    Set sld = ActivePresentation.Slides(1)

    ' Delete progress circles
    For i = sld.Shapes.Count To 1 Step -1
        If sld.Shapes(i).Name = "ProgressDot" Then
            sld.Shapes(i).Delete
        End If
    Next i
End Sub

Sub DeleteAllCirclesInLastSlide()
    Dim sld As Slide
    Dim shp As Shape
    Dim i As Integer
    Dim totalSlides As Integer

    ' Get the number of slides
    totalSlides = ActivePresentation.Slides.Count

    '  Get the last slide
    Set sld = ActivePresentation.Slides(totalSlides)

    ' Delete progress circles
    For i = sld.Shapes.Count To 1 Step -1
        If sld.Shapes(i).Name = "ProgressDot" Then
            sld.Shapes(i).Delete
        End If
    Next i
End Sub

