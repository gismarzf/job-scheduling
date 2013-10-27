VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFabricaDeRecorridos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pRecorridos As New Collection
Private pRecorridoCritico As cRecorrido
Private pModelo As cModeloDisyuntivo
Private pSumaMaxima As Integer

'********************************************************************************
'PRINCIPAL
'********************************************************************************

'********************************************************************************
'RECORRIDOS CRITICOS
'********************************************************************************
'*
'*
'*
Public Function dameRecorridosQueContienenOperacion(o As cOperacion) As Collection
    Dim rec As cRecorrido
    Dim recQueContienen As New Collection
    
    For Each rec In recorridos
        If coleccionContieneObjeto(o, rec.Recorrido) Then recQueContienen.Add rec
    Next
    
    Set dameRecorridosQueContienenOperacion = recQueContienen
End Function

Public Function dameRecorridosQueContienenArco(a As cArco) As Collection
    Dim rec As cRecorrido
    Dim recQueContienen As New Collection
    
    For Each rec In recorridos
        If coleccionContieneObjeto(a, rec.Arcos) Then recQueContienen.Add rec
    Next
    
    Set dameRecorridosQueContienenArco = recQueContienen
End Function

Public Function obtenerRecorridoCriticoDesdeMemoria(o As cOperacion) As cRecorrido

    Dim recorridosEncontrados As Collection, recorridosCortados As New Collection
    Set recorridosEncontrados = dameRecorridosQueContienenOperacion(o)
    
    Dim colaInf As cRecorrido
    Dim rec As cRecorrido
    '*
    '*
    '* se podria mejorar, busca todos los recorridos aunk so
    For Each rec In recorridosEncontrados
        Set colaInf = rec.dameColaInferior(o)
        If Not coleccionContieneRecorrido(recorridosCortados, colaInf) Then
            recorridosCortados.Add colaInf
        End If
    Next
          
    Set obtenerRecorridoCriticoDesdeMemoria = elegirRecorridoCritico(recorridosCortados)
    
End Function

Public Function coleccionContieneRecorrido(col As Collection, r As cRecorrido) As Boolean
    Dim rtemp As cRecorrido

    If col.count = 0 Then
        coleccionContieneRecorrido = False
    Else
        For Each rtemp In col
            If rtemp.mismoRecorrido(r) Then coleccionContieneRecorrido = True
            Exit For
        Next
    End If
        
End Function

Public Function obtenerRecorridosCritico() As Collection
    
    Dim rtemp As New Collection
    
    Dim o As New cOperacion
    For Each o In Modelo.Operaciones
        ' agregar la ultima operacion en cada trabajo como punto final para el Recorrido critico
        If o.OrdenEnElTrabajo = Modelo.cantOperacionesPorTrabajo Then
        '********************************************************************************
            Set rtemp = sumarColecciones(rtemp, obtenerTodosRecorridosHasta(o))
            ' **** ESTO ES SUPER LENTO ****
        '********************************************************************************
        End If
    Next
        
    Set obtenerRecorridosCritico = rtemp

End Function

'*
'*
'*
Public Function elegirRecorridoCritico(r As Collection) As cRecorrido
    Dim rcrit As cRecorrido
    Dim rtemp As cRecorrido
    Dim max As Integer
    max = 0
    
    For Each rtemp In r
        rtemp.Suma = rtemp.sumarPesos
        If rtemp.Suma > max Then
            Set rcrit = rtemp
            max = rtemp.Suma
        End If
    Next
    
    Set elegirRecorridoCritico = rcrit
    
End Function
'*
'*
'*
Public Function obtenerRecorridoCriticoHasta(o As cOperacion)
   Set obtenerRecorridoCriticoHasta = elegirRecorridoCritico(obtenerTodosRecorridosHasta(o))
End Function
'*
'*
'*
'********************************************************************************
'RECORRIDOS
'********************************************************************************
'*
'*
'*
' actualiza todos los recorridos en el modelo
Public Sub MALactualizarRecorridos(arco As cArco)
    ' superiores pasan por donde termina el arco, inferiores al reves
    ' de los superiores necesito las colas inferiores
    Dim colaInf As cRecorrido, colaSup As cRecorrido, rec As cRecorrido
    Dim recorridosParaRemover As Collection
    Set recorridosParaRemover = dameRecorridosQueContienenArco(arco)
    
    For Each rec In recorridosParaRemover
        removerRecorrido rec
    Next

    Dim recorridosSuperiores As Collection, recorridosInferiores As Collection
    Set recorridosSuperiores = dameRecorridosQueContienenOperacion(arco.InicioOperacion)
    Set recorridosInferiores = dameRecorridosQueContienenOperacion(arco.FinalOperacion)
    
    ' obtengo las colas inferiores y superiores
    Dim colasInferiores As New Collection, colasSuperiores As New Collection

    For Each rec In recorridosSuperiores
        Set colaInf = rec.dameColaInferior(arco.InicioOperacion)
        If Not coleccionContieneRecorrido(colasInferiores, colaInf) Then
            colasInferiores.Add colaInf
        End If
    Next
    
    For Each rec In recorridosInferiores
        Set colaSup = rec.dameColaSuperior(arco.FinalOperacion)
        If Not coleccionContieneRecorrido(colasSuperiores, colaSup) Then
            colasSuperiores.Add colaSup
        End If
    Next
    

    ' ahora tengo que conectar las colas inferiores y superiores de una forma que
    ' genero nuevos recorridos
    
    Dim nuevasColas As Collection
    Dim rec2 As cRecorrido, recnuevo As New cRecorrido
    For Each rec In colasInferiores
        For Each rec2 In colasSuperiores
            Set recnuevo = pegaDosSubRecorridos(rec, rec2)
            recnuevo.Suma = recnuevo.sumarPesos
            recorridos.Add recnuevo
        Next
    Next
    
    
End Sub

Public Sub removerRecorrido(r As cRecorrido)

    Dim i As Integer
    For i = 1 To recorridos.count
        If recorridos(i) Is r Then
            recorridos.Remove i
            Exit For ' tengo que salir ya que cambia el indice maximo
        End If
    Next i

End Sub
'*
'*
'*
' ojo que esto va del ultimo al primero
Public Function obtenerTodosRecorridosHasta(o As cOperacion)

    Dim recorridosIncompletos As New Collection
    Dim subrecorridos As New Collection
    Dim recorridosCompletos As New Collection
    Set recorridosCompletos = New Collection
    
    Dim r As cRecorrido
    Set r = New cRecorrido
    
    ' agregar el punto de inicio
    r.Recorrido.Add o
    recorridosIncompletos.Add r
    
    '********************************************************************************
    ' SUPER LENTO
    '********************************************************************************
    While hayMasPosiblesRecorridos(recorridosIncompletos)
    
        Set subrecorridos = New Collection
        
        ' tomo todos los recorridosIncompletos que tengo en la coleccion, voy a la primer operacion de la lista
        ' y agrego las operaciones que pueden estar antes de esa operacion
        For Each r In recorridosIncompletos

                ' veo si puedo agregar mas operaciones al Recorrido este, si NO, no lo agrego a mi lista
                ' ya que esta completo, lo mando a otra lista (para recorridosIncompletos completos)
                
                If hayMasSubRecorridos(r) Then
                    Set subrecorridos = sumarColecciones(subrecorridos, obtenerSubRecorridos(r))
                Else
                    recorridosCompletos.Add obtenerSubRecorridos(r).item(1)

                End If
        Next
        
        Set recorridosIncompletos = subrecorridos

    Wend
    
    ' esto para poder obtener agregar los ultimos recorridosIncompletos
    Set recorridosCompletos = sumarColecciones(recorridosCompletos, recorridosIncompletos)
    
    Set obtenerTodosRecorridosHasta = recorridosCompletos
    
End Function
'*
'*
'*
Private Function hayMasPosiblesRecorridos(recs As Collection) As Boolean

    Dim r As cRecorrido
    For Each r In recs
        If r.primerOperacionDelRecorrido.posiblesPrecedores.count > 0 Then
            hayMasPosiblesRecorridos = True
            Exit Function ' salimos al primero, no hace falter iterar todos
        End If
    Next
    
End Function
'*
'*
'*
'********************************************************************************
'SUBRECORRIDOS
'********************************************************************************
'*
'*
'*
' voy a la primera operacion en el camino, y agrego las operaciones que estan antes de esta
' devuelvo los nuevos caminos que tienen una operacion mas, si llegue al inicio y no hay operacion
' anterior, devuelvo el camino con cual llamamos a la funcion que va ser uno solo
Private Function obtenerSubRecorridos(r As cRecorrido) As Collection
    Dim col As New Collection
    Dim r2 As cRecorrido
    Dim o As cOperacion
    
    For Each o In r.primerOperacionDelRecorrido.posiblesPrecedores
        Set r2 = r.copiarRecorrido
        r2.agregarOperacionAntesDeTodos o
        col.Add r2
    Next
    
    ' si no habia mas subRecorridos, agrego el Recorrido original para no perderlo
    If col.count = 0 Then col.Add r
    
    Set obtenerSubRecorridos = col
End Function
'*
'*
'*
Public Function pegaDosSubRecorridos(primerRecorrido As cRecorrido, segundoRecorrido As cRecorrido) _
                                                                            As cRecorrido
    Dim recorridoPegado As New cRecorrido
    
    ' agrego los arcos del primer recorrido
    Dim i As Integer
    For i = 1 To primerRecorrido.Arcos.count
        recorridoPegado.Arcos.Add primerRecorrido.Arcos(i)
    Next i
    
    For i = 1 To primerRecorrido.Recorrido.count
        recorridoPegado.Recorrido.Add primerRecorrido.Recorrido(i)
    Next i
    
    ' necesito el arco entre el final del primer recorrido y el inicio del segundo
    Dim a As cArco
    i = primerRecorrido.Recorrido.count
    Set a = Modelo.dameArcoEntreOperaciones(primerRecorrido.Recorrido(i), segundoRecorrido.Recorrido(1))
    recorridoPegado.Arcos.Add a
                                        
    
    ' agrego los arcos del segundo recorrido
    For i = 1 To segundoRecorrido.Arcos.count
        recorridoPegado.Arcos.Add segundoRecorrido.Arcos(i)
    Next i
        
    For i = 1 To segundoRecorrido.Recorrido.count
        recorridoPegado.Recorrido.Add segundoRecorrido.Recorrido(i)
    Next i

    
    ' tengo que agregar el final de la ultima operacion
    'recorridoPegado.Recorrido.Add segundoRecorrido.Arcos(segundoRecorrido.Arcos.count).FinalOperacion
    
    Set pegaDosSubRecorridos = recorridoPegado

End Function
'*
'*
'*
Private Function hayMasSubRecorridos(c As cRecorrido) As Boolean
    
    If c.primerOperacionDelRecorrido.posiblesPrecedores.count > 0 Then
        hayMasSubRecorridos = True
    End If
    
End Function
'*
'*
'*
'********************************************************************************
'AUXILIAR
'********************************************************************************
'*
'*
'*
Public Function recorridosQueContienenArco(a As cArco) As Collection
    Dim recs As New Collection
    Dim r As cRecorrido

    For Each r In recorridos
        If coleccionContieneObjeto(a, r.Arcos) Then recs.Add r
    Next

    Set recorridosQueContienenArco = recs
End Function
'*
'*
'*
Public Function calcularSumaTotal() As Integer
    Dim sum As Integer
    Dim r As cRecorrido
    For Each r In recorridos
        sum = sum + r.Suma
    Next
    
    calcularSumaTotal = sum
End Function
Public Function calcularSumaMaxima()

    ' itero por todos Recorridos y  calculo la suma y devuelvo la suma maxima
    Dim c As cRecorrido
    Dim maxSuma As Integer
    maxSuma = 0
    For Each c In recorridos
        If c.Suma > maxSuma Then
            Set calcularSumaMaxima = c
            SumaMaxima = c.Suma
        End If
    Next
    
End Function
'*
'*
'*
Public Sub mostrarRecorridos(r As cRecorrido)
    Dim a As cArco
    
        For Each a In r.Arcos
            a.Conector.marcarLinea
        Next
    
End Sub
'*
'*
'*
'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************
'*
'*
'*
Property Get RecorridoCritico() As cRecorrido
    Set RecorridoCritico = pRecorridoCritico
End Property
Property Set RecorridoCritico(rc As cRecorrido)
    Set pRecorridoCritico = rc
End Property
Property Get SumaMaxima() As Integer
    SumaMaxima = pSumaMaxima
End Property
Property Let SumaMaxima(i As Integer)
    pSumaMaxima = i
End Property
Property Get Modelo() As cModeloDisyuntivo
    Set Modelo = pModelo
End Property
Property Set Modelo(m As cModeloDisyuntivo)
    Set pModelo = m
End Property
Property Get recorridos() As Collection
    Set recorridos = pRecorridos
End Property
Property Set recorridos(c As Collection)
    Set pRecorridos = c
End Property
