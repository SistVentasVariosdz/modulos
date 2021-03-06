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
' ADVERTENCIA:  NO presione el bot�n Fin al
'   depurar este proyecto. Vea la explicaci�n
'   al comienzo del m�dulo XTimerSupport
'   (XTimerS.bas).
'==============================================

' Almacenamiento privado para las
' propiedades de XTimer:
Private mlngTimerID As Long
Private mlngInterval As Long
Private mblnEnabled As Boolean

' El �nico evento de XTimer es Tick. Este evento no tiene
'   argumentos (eliminando los argumentos se acelera
'   ligeramente el evento), pero no hay raz�n por la que
'   no pueda indicarlos, si lo desea.
Event Tick()

' La propiedad TimerID es necesaria para el procedimiento
' --------------------   EndTimer, con el fin de localizar
'   r�pidamente el cron�metro en la matriz de cron�metros
'   activos del m�dulo de soporte.
'
' No hay raz�n para que el cliente use esta propiedad, de
'   modo que se declara con Friend en lugar de Public.
'
Friend Property Get TimerID() As Long
    TimerID = mlngTimerID
End Property

' La propiedad Enabled activa y desactiva el cron�metro. Esto
' --------------------      se hace terminando el cron�metro del
'   sistema, ya que no hay forma de suspenderlo. Si existe, est�
'   funcionando.
'
Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property
'
Public Property Let Enabled(ByVal NewValue As Boolean)
    ' Si no hay ning�n cambio en el estado
    '   de la propiedad, sale. As� se evita
    '   iniciar un segundo cron�metro del
    '   sistema cuando ya est� funcionando
    '   uno, etc�tera.
    If NewValue = mblnEnabled Then Exit Property
    '
    ' Guarda el nuevo valor de la propiedad.
    mblnEnabled = NewValue
    '
    ' Si Interval es cero, el cron�metro
    '   ya se ha detenido. No lo inicia.
    If mlngInterval = 0 Then Exit Property
    '
    ' Activa o desactiva el cron�metro.
    If mblnEnabled Then
        Debug.Assert mlngTimerID = 0
        mlngTimerID = BeginTimer(Me, mlngInterval)
    Else
        ' Lo siguiente es necesario, ya que
        '   un XTimer puede cerrar su cron�metro
        '   del sistema de dos formas:
        '   Enabled = False, o Interval = 0.
        If mlngTimerID <> 0 Then
            Call EndTimer(Me)
            mlngTimerID = 0
        End If
    End If
End Property

' La propiedad Interval debe hacer algo m�s que
' ---------------------   establecer el intervalo del
'   cron�metro. Si est� activado el XTimer, y se cambia
'   Interval de cero a un valor distinto, debe iniciarse
'   un cron�metro del sistema. Del mismo modo, si
'   Interval se pone a cero, el cron�metro del sistema
'   debe detenerse.
'
' El procedimiento Property Let tambi�n termina un
'   cron�metro del sistema e inicia otro siempre que
'   cambia el intervalo. Ello se debe a que no hay forma
'   de cambiar el intervalo de un cron�metro del sistema.
'
Public Property Get Interval() As Long
    Interval = mlngInterval
End Property
'
Public Property Let Interval(ByVal NewInterval As Long)
    ' Si el nuevo valor de Interval es igual al antiguo,
    '   no hay por qu� hacer nada.
    If NewInterval = mlngInterval Then Exit Property
    '
    ' Guarda el nuevo valor.
    mlngInterval = NewInterval
    '
    ' Si el XTimer est� activo, mlngTimerID es
    '   distinto de cero. En este caso, el anterior
    '   cron�metro del sistema debe terminar antes
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
    '   nuevo cron�metro del sistema, y su ID se
    '   almacena en mlngTimerID.
    If (NewInterval <> 0) And mblnEnabled Then
        mlngTimerID = BeginTimer(Me, NewInterval)
    End If
End Property

' El m�todo RaiseTick es invocado por el m�dulo de
' -------------------   soporte cuando se produce el
'   evento de cron�metro de sistema para el cron�metro
'   correspondiente al XTimer.
'
' Detalle de implementaci�n: puede que espere declarar
'   este m�todo como Friend en lugar de Public, ya que
'   el cliente no lo necesita para llamar a RaiseTick. Pero,
'   es cr�tico que RaiseTick se declare como Public,
'   ya que puede que se libere XTimer mientras se est�
'   controlando el evento Tick. Un objeto no terminar�
'   mientras uno de sus m�todos Public est� en
'   la pila, pero PUEDE terminar mientras uno de sus
'   m�todos Friend est� en la pila. Si el objeto termina
'   antes de que vuelva el m�todo Friend (puede pasar
'   si el cliente ejecuta mucho c�digo en
'   el evento Tick de XTimer), ocurrir� un error de protecci�n.
'   (Tenga en cuentar que no es un escenario habitual que
'   depende de un evento externo; no ocurre con el uso normal
'   de funciones Friend.)
'
Friend Sub RaiseTick()
    RaiseEvent Tick
End Sub

Private Sub Class_Terminate()
    ' Cuando el cliente libera su �ltima referencia
    '   a un objeto XTimer, �ste desaparece, pero
    '   s�lo si la propiedad Enabled del XTimer es
    '   False, o si su propiedad Interval es True.
    '
    ' El motivo es que mientras funcione el cron�metro
    '   del sistema asociado al XTimer, el m�dulo
    '   XTimerSupport tiene que tener una referencia
    '   al XTimer para poder producir su evento Tick.
    '   As�, si el cliente no desactiva los objetos
    '   XTimer antes de liberarlos, se perder�n
    '   cron�metros del sistema.
    '
    ' Estos cron�metros del sistema perdidos no se
    '   recuperar�n hasta que se cierre el componente
    '   XTimers; es decir, hasta que el cliente que
    '   usa la DLL se cierre. La DLL NO se descargar�
    '   cuando se liberen todos los objetos XTimer,
    '   ya que las referencias a objetos p�blicos (en
    '   este caso, las que mantiene XTimerSupport)
    '   impiden que se descargue una DLL.
    '
    ' De modo que �por qu� preocuparse de cerrar el
    '   cron�metro del sistema en el evento Terminate?
    '   Porque cuando la DLL se cierra, todas las
    '   referencias al objeto XTimer se cerrar�n, y
    '   el XTimer recibir� su evento Terminate. El
    '   cron�metro del sistema debe destruirse en este
    '   momento.
    On Error Resume Next
    If mlngTimerID <> 0 Then KillTimer 0, mlngTimerID
    '
    ' Lo siguiente es lo que deber�a hacer XTimer si
    '   pudiese ser liberado de alg�n modo antes del
    '   cierre de la DLL.
    'If mlngTimerID <> 0 Then Call EndTimer(Me)
End Sub


