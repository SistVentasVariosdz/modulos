Attribute VB_Name = "ECNLIB03_WINEVE_TIMER_MDL"
Option Explicit
'================================================
' ADVERTENCIA:  NO presione el botón Fin al
'   depurar este proyecto. En el modo de
'   ruptura, NO haga cambios que restablezcan
'   el proyecto.
'
' Este módulo es peligroso, ya que usa la API
'   SetTimer y el operador AddressOf para
'   establecer un cronómetro de sólo código.
'   Una vez establecido un cronómetro de este
'   tipo, el sistema continuará llamando a la
'   función TimerProc INCLUSO DESPUÉS DE VOLVER
'   AL TIEMPO DE DISEÑO.
'
' Como TimerProc no está disponible en tiempo
'   de diseño, el sistema provocará un FALLO DE
'   PROGRAMA en Visual Basic.
'
' Al depurar este módulo, debe asegurarse de que
'   todos los cronómetros del sistema estén
'   detenidos (con KillTimer) antes de volver al
'   tiempo de diseño. Puede conseguirlo si llama
'   a SCRUB desde la ventana Inmediato.
'
' Los cronómetros de llamada de retorno son
'   inherentemente peligrosos. Es mucho más seguro
'   usar controles Timer para la mayoría del
'   proceso de desarrollo, y reemplazarlos por
'   cronómetros de llamada de retorno muy al
'   final.
'==================================================

' Cantidad para incrementar el tamaño de la matriz
'   maxti cuando son necesarios más cronómetros.

Const MAXTIMERINCREMEMT = 5

Private Type XTIMERINFO   ' xti húngaro
    xt As ECNLIB03_WINEVE_TIMER
    id As Long
    blnReentered As Boolean
End Type

Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerProc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

' maxti es una matriz de objetos XTimer activos. La razón
' -----   para usar una matriz de tipo definido por el
'   usuario en lugar de un objeto Collection es lograr la
'   vinculación temprana al producir el evento Tick del
'   objeto XTimer.
Private maxti() As XTIMERINFO
'
' mintMaxTimers nos indica el tamaño de la matriz maxti en
' -------------   cualquier momento.
Private mintMaxTimers As Integer

' La función BeginTimer es invocada por un objeto XTimer
' ---------------------   cuando se establece un valor
'   distinto de cero en su propiedad Interval.
'
' La función hace las llamadas de API necesarias para
'   configurar un cronómetro. Si se crea un cronómetro
'   con éxito, la función incluye una referencia al objeto
'   XTimer en la matriz maxti. Esta referencia se empleará
'   para llamar al método que produce el evento Tick del
'   XTimer.
'
Public Function BeginTimer(ByVal xt As ECNLIB03_WINEVE_TIMER, _
                           ByVal Interval As Long)
    Dim lngTimerID As Long
    Dim intTimerNumber As Integer
    
    lngTimerID = SetTimer(0, 0, Interval, AddressOf TimerProc)
    ' Hay éxito cuando SetTimer devuelve un valor distinto de cero.
    '   Si no podemos obtener un cronómetro, hay un error.
    If lngTimerID = 0 Then Err.Raise vbObjectError + 31013, , "No hay cronómetros disponibles"
    
    ' El bucle siguiente localiza el primer lugar disponible
    '   en la matriz maxti. Si se supera el límite superior,
    '   se produce un error y se amplía la matriz. (Si compila
    '   esta DLL con Código nativo, NO desactive la comprobación
    '   de límites).
    For intTimerNumber = 1 To mintMaxTimers
        If maxti(intTimerNumber).id = 0 Then Exit For
    Next
    '
    ' Si no se encuentra espacio libre, se
    '   aumenta el tamaño de la matriz.
    If intTimerNumber > mintMaxTimers Then
        mintMaxTimers = mintMaxTimers + MAXTIMERINCREMEMT
        ReDim Preserve maxti(1 To mintMaxTimers)
    End If
    '
    ' Guarda una referencia para utilizarla al
    '  producir el evento Tick del objeto XTimer.
    Set maxti(intTimerNumber).xt = xt
    '
    ' Guarda el Id. de cronómetro que devuelve la API
    '   SetTimer, y devuelve el valor al objeto XTimer.
    maxti(intTimerNumber).id = lngTimerID
    maxti(intTimerNumber).blnReentered = False
    BeginTimer = lngTimerID
End Function

' TimerProc es el procedimiento de cronómetro al que
' ---------   se llama siempre que se dispare uno de los cronómetros.
'
' IMPORTANTE -- Este procedimiento debe estar en un módulo
'   estándar, por lo que todos los objetos cronómetro deben
'   compartirlo. Esto supone que el procedimiento debe
'   identificar qué cronómetro se ha disparado. Para ello
'   busca en la matriz maxti el ID del cronómetro (idEvent).
'
' Si esta declaración Sub es incorrecta, se producirán FALLOS
'   de programa. Este es uno de los peligros de utilizar API
'   que requieren funciones de llamada de retorno.
'
Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal lngSysTime As Long)
    Dim intCt As Integer

    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id = idEvent Then
            ' No provoca el evento si aún se está
            '   procesando una instancia anterior
            '   del mismo.
            If maxti(intCt).blnReentered Then Exit Sub
            ' El indicador blnReentered bloquea otras
            '   instancias del evento hasta que termina
            '   la instancia actual.
            maxti(intCt).blnReentered = True
            On Error Resume Next
            ' Provoca el evento Tick del objeto XTimer
            '   correspondiente.
            maxti(intCt).xt.RaiseTick
            If Err.Number <> 0 Then
                ' Si se produce un error, el XTimer ha
                '   conseguido de algún modo terminar sin
                '   disparar primero su cronómetro. Cierra
                '   el cronómetro huérfano, para evitar
                '   errores de protección general posteriores.
                KillTimer 0, idEvent
                maxti(intCt).id = 0
                '
                ' Libera la referencia al
                '   objeto XTimer.
                Set maxti(intCt).xt = Nothing
            End If
            '
            ' Permite que este evento entre de
            '   nuevo en TimerProc.
            maxti(intCt).blnReentered = False
            Exit Sub
        End If
    Next
    ' La línea siguiente provocará un fallo en caso de
    '   que se haya liberado de algún modo un XTimer
    '   sin haber cerrado el cronómetro del sistema
    '   Windows.
    '
    ' La ejecución también puede llegar a este punto a
    '   causa de un error conocido del NT 3.51, que
    '   hace posible recibir un evento de cronómetro
    '   adicional DESPUÉS de haber ejecutado el API
    '   KillTimer.
    KillTimer 0, idEvent
End Sub

' El procedimiento EndTimer es invocado por el XTimer
' -------------------------   siempre que se establece
'   en la propiedad Enabled, y cada vez que sea necesario
'   un nuevo intervalo de cronómetro. No hay forma de
'   restablecer un cronómetro del sistema, por lo que la
'   única manera de modificar el intervalo es cerrar el
'   cronómetro existente y llamar a BeginTimer para
'   iniciar uno nuevo.
'
Public Sub EndTimer(ByVal xt As ECNLIB03_WINEVE_TIMER)
    Dim lngTimerID As Long
    Dim intCt As Integer
    
    ' Pregunta al XTimer su TimerID, para poder buscar en
    '   la matriz la XTIMERINFO correcta. (Podría buscar la
    '   propia referencia al XTimer, usando el operador Is
    '   para comparar xt con maxti(intCt).xt, pero no sería
    '   tan rápido.
    lngTimerID = xt.TimerID
    '
    ' Si TimerID es cero, se ha llamado a EndTimer
    '   por error.
    If lngTimerID = 0 Then Exit Sub
    '
    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id = lngTimerID Then
            ' Cierra el cronómetro del sistema.
            KillTimer 0, lngTimerID
            '
            ' Libera la referencia al
            '   objeto XTimer.
            Set maxti(intCt).xt = Nothing
            '
            ' Borra el ID para dejar espacio
            '   un nuevo cronómetro activo.
            maxti(intCt).id = 0
            Exit Sub
        End If
    Next
End Sub

' El procedimiento Scrub es una válvula de seguridad sólo
' ----------------------   para la depuración: si tiene que
'   terminar con End este proyecto mientras hay objetos
'   XTimer activos, llame a Scrub en el panel Inmediato.
'   Con ello llamará a KillTimer para todos los cronómetros
'   del sistema, de forma que el entorno de desarrollo pueda
'   volver al modo de diseño con seguridad.
'
Public Sub Scrub()
    Dim intCt As Integer
    ' Cierra los cronómetros del sistema que queden activos.
    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id <> 0 Then KillTimer 0, maxti(intCt).id
    Next
End Sub

