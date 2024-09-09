# English
# Progress points ppt
This VBA Macros add progress circles to your slides (powerpoint). Similar to a progress bar for ppt presentations.

## Execute
To draw the progress points on all slides, run the macro DrawCircles.
The configuration of the points (as circles) is done with the variables, and all macros are explained below in this README.

## Variables
The variables are declared in the first macro (InitializeVariables).
- ProgressCircleFillColor: Color of the circles fill that simulate advanced slides.
- ProgressCircleBorderColor: Color of the circles border that simulate advanced slides.
- RemainingSlidesCircleFillColor: Color of the circles fill that remaining slides.
- RemainingSlidesCircleBorderColor: Color of the circles border that remaining slides.
- CircleBorderWidth: Border size in mm for all circles.
- radius = Radius of all circles.
- spacing = Spacing between circles.
- CircleHeight = Height at which the circles are displayed.

## Macros
- InitializeVariables: Declares the variables (circle size, border, height, spacing and colors).

- DrawCircles: Draws circles with ProgressCircleFillColor and ProgressCircleBorderColor according to the slide you are on (e.g. slide 2, it will draw 2 circles with these characteristics). And then draw as many circles with RemainingSlidesCircleFillColor and RemainingSlidesCircleBorderColor as slides are next to the current. 
If there were already progress circles previously, it will delete them before drawing the new ones will draw.

- DeleteCircles_AllSlides: Deletes all the progress circles of all slides.
- DeleteCircles_CurrentSlide: Deletes all the progress circles of the slide you are currently on.

- DeleteFirstCircleAndCenter: Removes the first progress circle from all slides and re-centers the remaining circles.
- DeleteLastCircleAndCenter: Removes the last progress circle from all slides and re-centers the remaining circles.

- DeleteAllCirclesInFirstSlide: Deletes all the progress pcircles of the first slide.
- DeleteAllCirclesInLastSlide: Deletes all the progress circles of the last slide.

### Activate developer tab in power point
https://youtu.be/VbR-YYA2yRk?si=ZsaZesv_8MILX2dD&t=13


### Personal recommendations
- Personally, I do not recommend using these progress points for an exhibition or class (it can distract the audience and we have the view of the moderator who gives us this data). However, it can be optimal for presentations that will not be exposed, giving the final client the correct notion of progress.
- Create in your template a border or background to reserve the space of the circles and run the macros you consider once the presentation is finished.
- Delete the advance points of the first slide (cover) with the macro DeleteAllCirclesInFirstSlide.
- In case you end with a “thank you” slide without feedback, delete the progress points of that slide (macro DeleteAllCirclesInLastSlide) and delete the last progress point of all slides (macro DeleteLastCircleAndCenter). 

# Author
- [@Jose Chirif](https://github.com/JoseChirif)

## 🚀 About me
I'm an Industrial Engineer specialized in process optimization, business intelligence and data science.
[Porfolio - Network - Contact](https://linktr.ee/jchirif)


----------
# Español
# Puntos avance ppt
Macros VBA para añadir puntos de avance de diapositivas. Similar a una barra de progreso para presentaciones ppt.

## Ejecutar
Para Dibujar los puntos de avance en todas las diapositivas, ejecutar la macro DibujarPuntos.
La configuración de los puntos se realiza con las variables (Siguiente sub-título), y las demás macros se explican más abajo en este mismo README.

## Variables
Las variables están declaradas en el primer macro (InicializarVariables).
- colorAvanzado: Color del relleno de los circulos que simulan las diapositivas avanzadas.
- bordeAvanzado: Color del borde de los circulos que simulan las diapositivas avanzadas.
- colorPendiente: Color del relleno de los circulos que simulan las diapositivas faltantes.
- bordePendiente: Color del borde de los circulos que simulan las diapositivas faltantes.
- grosorBordeCirculos: Grosor del borde en mm para todos los círculos.
- radius = Radio de todos los círculos.
- spacing = Espacio entre los círculos.
- puntoAltura = Altura a la que se presentan los círculos.

## Macros
- InicializarVariables: Declara las variables (tamaño circulos, borde, altura, espaciado y colores).

- DibujarPuntos: Dibuja circulos con  colorAvanzado y bordeAvanzado según la diapositiva en la estes (ej. diapositiva 2, dibujará 2 circulos con estas carácteristicas). Y seguido dibuja tantos circulos con colorPendiente y bordePendiente como diapositivas faltan. 
Si ya habían circulos de avance previamente, los eliminirá antes de dibujar los nuevos.

- BorrarPuntos_TodasLasDiapositivas: Borra los circulos de avance de todas las diapositivas.
- BorrarPuntos_EstaDiapositiva:  Borra los circulos de avance de la diapositiva donde estas actualmente.

- EliminarPrimerPuntoYCentrar: Elimina el primer circulo de avance de todas las diapositivas y vuelve a centrar los puntos restantes.
- EliminarUltimoPuntoYCentrar: Elimina el último circulo de avance de todas las diapositivas y vuelve a centrar los puntos restantes.

- EliminarPuntosPrimeraDiapositiva: Elimina todos los puntos de avance de la primera diapositiva.
- EliminarPuntosUltimaDiapositiva: Elimina todos los puntos de avance de la última diapositiva.


### Activar macros en power point
https://youtu.be/hiWvuBARspc?si=jaV932PODH8bWpSV&t=13


### Recomendaciones personales
- Personalmente, no recomiendo utilizar estos puntos para una exposición o clase (puede distraer la audiencia y tenemos la vista del moderador que nos da estos datos). Sin embargo, puede ser óptimo para presentaciones que no se expondran, dando al cliente final la noción de avance correcta.
- Crear en tu plantilla un borde o fondo para reservar el espacio de los circulos y ejecutar las macros que consideres una vez terminada la presentación.
- Eliminar los puntos de avance de la primera diapositiva (caratula) con el macro EliminarPuntosPrimeraDiapositiva.
- En caso finalice con una diapositiva de "gracias" sin feedback, eliminar los puntos de progreso de esa diapositiva (macro EliminarPuntosUltimaDiapositiva) y eliminar el último punto de progreso de todas las diapositivas (macro EliminarUltimoPuntoYCentrar). 

# Autor
- [@Jose Chirif](https://github.com/JoseChirif)

## 🚀 Acerca de mi
Ingeniero Industrial especializado en optimización de procesos, business intelligence y ciencia de datos.
[Portafolio - Redes - Contacto](https://linktr.ee/josechirif)

