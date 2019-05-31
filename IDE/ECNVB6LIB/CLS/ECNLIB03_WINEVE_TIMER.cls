VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ECNLIB03_WINEVE_TIMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'==============================================
' ADVERTENCIA:  NO presione el botón Fin al
'   depurar este proyecto. Vea la explicación
'   al comienzo del módulo XTimerSupport
'   (XTimerS.bas).
'==============================================

' Almacenamiento privado para las
' propiedades de XTimer:
Private mlngTimerID As Long
Private mlngInterval As Long
Private mblnEnabled As Boolean

' El único evento de XTimer es Tick. Este evento no tiene
'   argumentos (eliminando los argumentos se acelera
'   ligeramente el evento), pero no hay razón por la que
'   no pueda indicarlos, si lo desea.
Event Tick()

' La propiedad TimerID es necesaria para el procedimiento
' --------------------   EndTimer, con el fin de localizar
'   rápidamente el cronómetro en la matriz de cronómetros
'   activos del módulo de soporte.
'
' No hay razón para que el cliente use esta propiedad, de
'   modo que se declara con Friend en lugar de Public.
'
Friend Property Get TimerID() As Long
    TimerID = mlngTimerID
End Property

' La propiedad Enabled activa y desactiva el cronómetro. Esto
' --------------------      se hace terminando el cronómetro del
'   sistema, ya que no hay forma de suspenderlo. Si existe, está
'   funcionando.
'
Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property
'
Public Property Let Enabled(ByVal NewValue As Boolean)
    ' Si no hay ningún cambio en el estado
    '   de la propiedad, sale. Así se evita
    '   iniciar un segundo cronómetro del
    '   sistema cuando ya está funcionando
    '   uno, etcétera.
    If NewValue = mblnEnabled Then Exit Property
    '
    ' Guarda el nuevo valor de la propiedad.
    mblnEnabled = NewValue
    '
    ' Si Interval es cero, el cronómetro
    '   ya se ha detenido. No lo inicia.
    If mlngInterval = 0 Then Exit Property
    '
    ' Activa o desactiva el cronómetro.
    If mblnEnabled Then
        Debug.Assert mlngTimerID = 0
        mlngTimerID = BeginTimer(Me, mlngInterval)
    Else
        ' Lo siguiente es necesario, ya que
        '   un XTimer puede cerrar su cronómetro
        '   del sistema de dos formas:
        '   Enabled = False, o Interval = 0.
        If mlngTimerID <> 0 Then
            Call EndTimer(Me)
            mlngTimerID = 0
        End If
    End If
End Property

' La propiedad Interval debe hacer algo más que
' ---------------------   establecer el intervalo del
'   cronómetro. Si está activado el XTimer, y se cambia
'   Interval de cero a un valor distinto, debe iniciarse
'   un cronómetro del sistema. Del mismo modo, si
'   Interval se pone a cero, el cronómetro del sistema
'   debe detenerse.
'
' El procedimiento Property Let también termina un
'   cronómetro del sistema e inicia otro siempre que
'   cambia el intervalo. Ello se debe a que no hay forma
'   de cambiar el intervalo de un cronómetro del sistema.
'
Public Property Get Interval() As Long
    Interval = mlngInterval
End Property
'
Public Property Let Interval(ByVal NewInterval As Long)
    ' Si el nuevo valor de Interval es igual al antiguo,
    '   no hay por qué hacer nada.
    If NewInterval = mlngInterval Then Exit Property
    '
    ' Guarda el nuevo valor.
    mlngInterval = NewInterval
    '
    ' Si el XTimer está activo, mlngTimerID es
    '   distinto de cero. En este caso, el anterior
    '   cronómetro del sistema debe terminar antes
    '   de iniciar uno nuevo.
    If mlngTimerID <> 0 Then
        Call EndTimer(Me)
        mlngTimerID = 0
    End If
    '
    ' Si el nuevo intervalo es cero, el XTimer queda
    '   inactivo, independientemente del valor actual
    '   de Enabled. Si el nuevo intervalo no es cero,
    '   y la propiedad Enabled es True, se inicia un
    '   nuevo cronómetro del sistema, y su ID se
    '   almacena en mlngTimerID.
    If (NewInterval <> 0) And mblnEnabled Then
        mlngTimerID = BeginTimer(Me, NewInterval)
    End If
End Property

' El método RaiseTick es invocado por el módulo de
' -------------------   soporte cuando se produce el
'   evento de cronómetro de sistema para el cronómetro
'   correspondiente al XTimer.
'
' Detalle de implementación: puede que espere declarar
'   este método como Friend en lugar de Public, ya que
'   el cliente no lo necesita para llamar a RaiseTick. Pero,
'   es crítico que RaiseTick se declare como Public,
'   ya que puede que se libere XTimer mientras se está
'   controlando el evento Tick. Un objeto no terminará
'   mientras uno de sus métodos Public esté en
'   la pila, pero PUEDE terminar mientras uno de sus
'   métodos Friend esté en la pila. Si el objeto termina
'   antes de que vuelva el método Friend (puede pasar
'   si el cliente ejecuta mucho código en
'   el evento Tick de XTimer), ocurrirá un error de protección.
'   (Tenga en cuentar que no es un escenario habitual que
'   depende de un evento externo; no ocurre con el uso normal
'   de funciones Friend.)
'
Friend Sub RaiseTick()
    RaiseEvent Tick
End Sub

Private Sub Class_Terminate()
    ' Cuando el cliente libera su última referencia
    '   a un objeto XTimer, éste desaparece, pero
    '   sólo si la propiedad Enabled del XTimer es
    '   False, o si su propiedad Interval es True.
    '
    ' El motivo es que mientras funcione el cronómetro
    '   del sistema asociado al XTimer, el módulo
    '   XTimerSupport tiene que tener una referencia
    '   al XTimer para poder producir su evento Tick.
    '   Así, si el cliente no desactiva los objetos
    '   XTimer antes de liberarlos, se perderán
    '   cronómetros del sistema.
    '
    ' Estos cronómetros del sistema perdidos no se
    '   recuperarán hasta que se cierre el componente
    '   XTimers; es decir, hasta que el cliente que
    '   usa la DLL se cierre. La DLL NO se descargará
    '   cuando se liberen todos los objetos XTimer,
    '   ya que las referencias a objetos públicos (en
    '   este caso, las que mantiene XTimerSupport)
    '   impiden que se descargue una DLL.
    '
    ' De modo que ¿por qué preocuparse de cerrar el
    '   cronómetro del sistema en el evento Terminate?
    '   Porque cuando la DLL se cierra, todas las
    '   referencias al objeto XTimer se cerrarán, y
    '   el XTimer recibirá su evento Terminate. El
    '   cronómetro del sistema debe destruirse en este
    '   momento.
    On Error Resume Next
    If mlngTimerID <> 0 Then KillTimer 0, mlngTimerID
    '
    ' Lo siguiente es lo que debería hacer XTimer si
    '   pudiese ser liberado de algún modo antes del
    '   cierre de la DLL.
    'If mlngTimerID <> 0 Then Call EndTimer(Me)
End Sub


