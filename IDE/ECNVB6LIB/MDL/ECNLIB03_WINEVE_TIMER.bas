Attribute VB_Name = "ECNLIB03_WINEVE_TIMER_MDL"
Option Explicit
'================================================
' ADVERTENCIA:  NO presione el bot�n Fin al
'   depurar este proyecto. En el modo de
'   ruptura, NO haga cambios que restablezcan
'   el proyecto.
'
' Este m�dulo es peligroso, ya que usa la API
'   SetTimer y el operador AddressOf para
'   establecer un cron�metro de s�lo c�digo.
'   Una vez establecido un cron�metro de este
'   tipo, el sistema continuar� llamando a la
'   funci�n TimerProc INCLUSO DESPU�S DE VOLVER
'   AL TIEMPO DE DISE�O.
'
' Como TimerProc no est� disponible en tiempo
'   de dise�o, el sistema provocar� un FALLO DE
'   PROGRAMA en Visual Basic.
'
' Al depurar este m�dulo, debe asegurarse de que
'   todos los cron�metros del sistema est�n
'   detenidos (con KillTimer) antes de volver al
'   tiempo de dise�o. Puede conseguirlo si llama
'   a SCRUB desde la ventana Inmediato.
'
' Los cron�metros de llamada de retorno son
'   inherentemente peligrosos. Es mucho m�s seguro
'   usar controles Timer para la mayor�a del
'   proceso de desarrollo, y reemplazarlos por
'   cron�metros de llamada de retorno muy al
'   final.
'==================================================

' Cantidad para incrementar el tama�o de la matriz
'   maxti cuando son necesarios m�s cron�metros.

Const MAXTIMERINCREMEMT = 5

Private Type XTIMERINFO   ' xti h�ngaro
    xt As ECNLIB03_WINEVE_TIMER
    id As Long
    blnReentered As Boolean
End Type

Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerProc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

' maxti es una matriz de objetos XTimer activos. La raz�n
' -----   para usar una matriz de tipo definido por el
'   usuario en lugar de un objeto Collection es lograr la
'   vinculaci�n temprana al producir el evento Tick del
'   objeto XTimer.
Private maxti() As XTIMERINFO
'
' mintMaxTimers nos indica el tama�o de la matriz maxti en
' -------------   cualquier momento.
Private mintMaxTimers As Integer

' La funci�n BeginTimer es invocada por un objeto XTimer
' ---------------------   cuando se establece un valor
'   distinto de cero en su propiedad Interval.
'
' La funci�n hace las llamadas de API necesarias para
'   configurar un cron�metro. Si se crea un cron�metro
'   con �xito, la funci�n incluye una referencia al objeto
'   XTimer en la matriz maxti. Esta referencia se emplear�
'   para llamar al m�todo que produce el evento Tick del
'   XTimer.
'
Public Function BeginTimer(ByVal xt As ECNLIB03_WINEVE_TIMER, _
                           ByVal Interval As Long)
    Dim lngTimerID As Long
    Dim intTimerNumber As Integer
    
    lngTimerID = SetTimer(0, 0, Interval, AddressOf TimerProc)
    ' Hay �xito cuando SetTimer devuelve un valor distinto de cero.
    '   Si no podemos obtener un cron�metro, hay un error.
    If lngTimerID = 0 Then Err.Raise vbObjectError + 31013, , "No hay cron�metros disponibles"
    
    ' El bucle siguiente localiza el primer lugar disponible
    '   en la matriz maxti. Si se supera el l�mite superior,
    '   se produce un error y se ampl�a la matriz. (Si compila
    '   esta DLL con C�digo nativo, NO desactive la comprobaci�n
    '   de l�mites).
    For intTimerNumber = 1 To mintMaxTimers
        If maxti(intTimerNumber).id = 0 Then Exit For
    Next
    '
    ' Si no se encuentra espacio libre, se
    '   aumenta el tama�o de la matriz.
    If intTimerNumber > mintMaxTimers Then
        mintMaxTimers = mintMaxTimers + MAXTIMERINCREMEMT
        ReDim Preserve maxti(1 To mintMaxTimers)
    End If
    '
    ' Guarda una referencia para utilizarla al
    '  producir el evento Tick del objeto XTimer.
    Set maxti(intTimerNumber).xt = xt
    '
    ' Guarda el Id. de cron�metro que devuelve la API
    '   SetTimer, y devuelve el valor al objeto XTimer.
    maxti(intTimerNumber).id = lngTimerID
    maxti(intTimerNumber).blnReentered = False
    BeginTimer = lngTimerID
End Function

' TimerProc es el procedimiento de cron�metro al que
' ---------   se llama siempre que se dispare uno de los cron�metros.
'
' IMPORTANTE -- Este procedimiento debe estar en un m�dulo
'   est�ndar, por lo que todos los objetos cron�metro deben
'   compartirlo. Esto supone que el procedimiento debe
'   identificar qu� cron�metro se ha disparado. Para ello
'   busca en la matriz maxti el ID del cron�metro (idEvent).
'
' Si esta declaraci�n Sub es incorrecta, se producir�n FALLOS
'   de programa. Este es uno de los peligros de utilizar API
'   que requieren funciones de llamada de retorno.
'
Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal lngSysTime As Long)
    Dim intCt As Integer

    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id = idEvent Then
            ' No provoca el evento si a�n se est�
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
                '   conseguido de alg�n modo terminar sin
                '   disparar primero su cron�metro. Cierra
                '   el cron�metro hu�rfano, para evitar
                '   errores de protecci�n general posteriores.
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
    ' La l�nea siguiente provocar� un fallo en caso de
    '   que se haya liberado de alg�n modo un XTimer
    '   sin haber cerrado el cron�metro del sistema
    '   Windows.
    '
    ' La ejecuci�n tambi�n puede llegar a este punto a
    '   causa de un error conocido del NT 3.51, que
    '   hace posible recibir un evento de cron�metro
    '   adicional DESPU�S de haber ejecutado el API
    '   KillTimer.
    KillTimer 0, idEvent
End Sub

' El procedimiento EndTimer es invocado por el XTimer
' -------------------------   siempre que se establece
'   en la propiedad Enabled, y cada vez que sea necesario
'   un nuevo intervalo de cron�metro. No hay forma de
'   restablecer un cron�metro del sistema, por lo que la
'   �nica manera de modificar el intervalo es cerrar el
'   cron�metro existente y llamar a BeginTimer para
'   iniciar uno nuevo.
'
Public Sub EndTimer(ByVal xt As ECNLIB03_WINEVE_TIMER)
    Dim lngTimerID As Long
    Dim intCt As Integer
    
    ' Pregunta al XTimer su TimerID, para poder buscar en
    '   la matriz la XTIMERINFO correcta. (Podr�a buscar la
    '   propia referencia al XTimer, usando el operador Is
    '   para comparar xt con maxti(intCt).xt, pero no ser�a
    '   tan r�pido.
    lngTimerID = xt.TimerID
    '
    ' Si TimerID es cero, se ha llamado a EndTimer
    '   por error.
    If lngTimerID = 0 Then Exit Sub
    '
    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id = lngTimerID Then
            ' Cierra el cron�metro del sistema.
            KillTimer 0, lngTimerID
            '
            ' Libera la referencia al
            '   objeto XTimer.
            Set maxti(intCt).xt = Nothing
            '
            ' Borra el ID para dejar espacio
            '   un nuevo cron�metro activo.
            maxti(intCt).id = 0
            Exit Sub
        End If
    Next
End Sub

' El procedimiento Scrub es una v�lvula de seguridad s�lo
' ----------------------   para la depuraci�n: si tiene que
'   terminar con End este proyecto mientras hay objetos
'   XTimer activos, llame a Scrub en el panel Inmediato.
'   Con ello llamar� a KillTimer para todos los cron�metros
'   del sistema, de forma que el entorno de desarrollo pueda
'   volver al modo de dise�o con seguridad.
'
Public Sub Scrub()
    Dim intCt As Integer
    ' Cierra los cron�metros del sistema que queden activos.
    For intCt = 1 To mintMaxTimers
        If maxti(intCt).id <> 0 Then KillTimer 0, maxti(intCt).id
    Next
End Sub

