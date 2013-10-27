Attribute VB_Name = "mTestSolucion"
Option Explicit
Option Base 1


' NOT tested
' va del primero al ultimo, y no del ultimo al primero!!!
Private Function hayMasSubRecorridos(c() As Integer) As Boolean
    hayMasSubRecorridos = True
    If damePosiblesPrecedores(c())(1) = 0 Then hayMasSubRecorridos = False
End Function

' NOT tested
Private Function damePosiblesPrecedores(c() As Integer) As Integer()
    Dim pp(cantOperaciones) As Integer
    
    Dim indicePrimeraOperacionDelRecorrido As Integer
    indicePrimeraOperacionDelRecorrido = c(UBound(c, 1))
       
    Dim i As Integer
    For i = 1 To cantOperaciones
        If operacionPrecedores(indicePrimeraOperacionDelRecorrido, 1) = 0 Then
            Exit For
        End If
        
        
End Function
''*
''*
''*
'Private Function hayMasSubRecorridos(c As cRecorrido) As Boolean
'
'    If c.primerOperacionDelRecorrido.posiblesPrecedores.count > 0 Then
'        hayMasSubRecorridos = True
'    End If
'
'End Function

'' voy a la primera operacion en el camino, y agrego las operaciones que estan antes de esta
'' devuelvo los nuevos caminos que tienen una operacion mas, si llegue al inicio y no hay operacion
'' anterior, devuelvo el camino con cual llamamos a la funcion que va ser uno solo
'Private Function obtenerSubRecorridos(r As cRecorrido) As Collection
'    Dim col As New Collection
'    Dim r2 As cRecorrido
'    Dim o As cOperacion
'
'    For Each o In r.primerOperacionDelRecorrido.posiblesPrecedores
'        Set r2 = r.copiarRecorrido
'        r2.agregarOperacionAntesDeTodos o
'        col.Add r2
'    Next
'
'    ' si no habia mas subRecorridos, agrego el Recorrido original para no perderlo
'    If col.count = 0 Then col.Add r
'
'    Set obtenerSubRecorridos = col
'End Function
''*
''*
''*

'Public Function testSolucion(bvector As Collection) As Integer
'
'    Dim valant As New Collection
'
'    Dim i As Integer
'    For i = 1 To ArcosDisyuntivos.count
'        ' guardar los valores anteriores
'        valant.Add ArcosDisyuntivos(i).Direccion
'        ArcosDisyuntivos(i).Direccion = bvector(i)
'    Next i
'
'    testSolucion = _
'        FabricaDeRecorridos.elegirRecorridoCritico(FabricaDeRecorridos.obtenerRecorridosCritico).sumarPesos
'
'    For i = 1 To ArcosDisyuntivos.count
'        ' volver a los valores anteriores
'        ArcosDisyuntivos(i).Direccion = valant(i)
'    Next i
'
'End Function


'Public Function elegirRecorridoCritico(r As Collection) As cRecorrido
'    Dim rcrit As cRecorrido
'    Dim rtemp As cRecorrido
'    Dim max As Integer
'    max = 0
'
'    For Each rtemp In r
'        rtemp.Suma = rtemp.sumarPesos
'        If rtemp.Suma > max Then
'            Set rcrit = rtemp
'            max = rtemp.Suma
'        End If
'    Next
'
'    Set elegirRecorridoCritico = rcrit
'
'End Function
''*
''*
''*
'Public Function obtenerRecorridoCriticoHasta(o As cOperacion)
'   Set obtenerRecorridoCriticoHasta = elegirRecorridoCritico(obtenerTodosRecorridosHasta(o))
'End Function

''*
''*
''*
'' ojo que esto va del ultimo al primero
'Public Function obtenerTodosRecorridosHasta(o As cOperacion)
'
'    Dim recorridosIncompletos As New Collection
'    Dim subrecorridos As New Collection
'    Dim recorridosCompletos As New Collection
'    Set recorridosCompletos = New Collection
'
'    Dim r As cRecorrido
'    Set r = New cRecorrido
'
'    ' agregar el punto de inicio
'    r.Recorrido.Add o
'    recorridosIncompletos.Add r
'
'    '********************************************************************************
'    ' SUPER LENTO
'    '********************************************************************************
'    While hayMasPosiblesRecorridos(recorridosIncompletos)
'
'        Set subrecorridos = New Collection
'
'        ' tomo todos los recorridosIncompletos que tengo en la coleccion, voy a la primer operacion de la lista
'        ' y agrego las operaciones que pueden estar antes de esa operacion
'        For Each r In recorridosIncompletos
'
'                ' veo si puedo agregar mas operaciones al Recorrido este, si NO, no lo agrego a mi lista
'                ' ya que esta completo, lo mando a otra lista (para recorridosIncompletos completos)
'
'                If hayMasSubRecorridos(r) Then
'                    Set subrecorridos = sumarColecciones(subrecorridos, obtenerSubRecorridos(r))
'                Else
'                    recorridosCompletos.Add obtenerSubRecorridos(r).item(1)
'
'                End If
'        Next
'
'        Set recorridosIncompletos = subrecorridos
'
'    Wend
'
'    ' esto para poder obtener agregar los ultimos recorridosIncompletos
'    Set recorridosCompletos = sumarColecciones(recorridosCompletos, recorridosIncompletos)
'
'    Set obtenerTodosRecorridosHasta = recorridosCompletos
'
'End Function
''*
''*
''*
'Private Function hayMasPosiblesRecorridos(recs As Collection) As Boolean
'
'    Dim r As cRecorrido
'    For Each r In recs
'        If r.primerOperacionDelRecorrido.posiblesPrecedores.count > 0 Then
'            hayMasPosiblesRecorridos = True
'            Exit Function ' salimos al primero, no hace falter iterar todos
'        End If
'    Next
'
'End Function

