Attribute VB_Name = "mTestSolucion"
Option Explicit
Option Base 1

Public Const maxRecorridos As Integer = 256
Public recorridosCompletos() As Integer
Public recorridosIncompletos() As Integer

' ojo que
' en los vectores, el recorrido va del ultimo hasta el primero
Private Sub atest()
    Dim start As Long, finish As Long


    init
    ReDim c(cantOperaciones) As Integer
    Dim testrec As New cRecorrido
    Set testrec = _
        Modelo.FabricaDeRecorridos.RecorridoCritico.dameColaSuperior _
            (Modelo.FabricaDeRecorridos.RecorridoCritico.Recorrido(6))
    
    
    c = coleccionRecoridoEnVectorRecorrido(testrec.Recorrido)

    Dim testrec2 As New Collection
    Set testrec2 = Modelo.FabricaDeRecorridos.obtenerSubRecorridos(testrec)

    ReDim r2(cantOperaciones, cantOperaciones) As Integer
    r2 = obtenerSubRecorridos(c)

    ReDim v(cantOperaciones) As Integer


    ReDim recorridosCompletos(maxRecorridos, cantOperaciones) As Integer
    ReDim recorridosIncompletos(maxRecorridos, cantOperaciones) As Integer

    v = extraerVectorDeMatriz(r2, 1)

    Dim ctest As New Collection
    Dim vtest() As Long
    Dim r1() As Integer
    
    r2 = reemplazarPrimerVectorZero(r2, v)
    r1 = r2
    
    
    r2 = juntarMatriz(r2, r1)
    
'    start = GetTickCount()
'    Dim i As Long
'    For i = 1 To 1000000
'        ctest.Add i
'    Next i
'    Debug.Print GetTickCount() - start
'
'    start = GetTickCount()
'    For i = 1 To 1000000
'        ReDim vtest(i)
'        vtest(i) = i
'    Next i
'    Debug.Print GetTickCount() - start


    End
End Sub

Public Function obtenerTodosRecorridosHasta(o As Integer)
    
    
End Function

Public Function juntarMatriz(m1() As Integer, m2() As Integer) As Integer()
    
    ReDim mjuntos(maxRecorridos, cantOperaciones) As Integer
    ReDim v(cantOperaciones) As Integer
    Dim i As Integer
    For i = 1 To maxRecorridos
        If m2(i, 1) = 0 Then
            juntarMatriz = m1
            Exit Function
        End If
        v = extraerVectorDeMatriz(m2, i)
        
        m1 = reemplazarPrimerVectorZero(m1, v)
    Next i
    
End Function

' reemplaza el primer vector que empieza con 0 de una matriz
Public Function reemplazarPrimerVectorZero(mat() As Integer, v() As Integer) As Integer()
    
    Dim mtemp() As Integer
    mtemp = copiarMatriz(mat)
    
    Dim i As Integer, j As Integer
    For i = 1 To UBound(mat, 1)
        If mat(i, 1) = 0 Then
            For j = 1 To UBound(v, 1)
                mat(i, j) = v(j)
            Next j
            reemplazarPrimerVectorZero = mat
            Exit Function
        End If
    Next i
    
End Function

Public Function extraerVectorDeMatriz(r() As Integer, ind As Integer) As Integer()
    ReDim v(UBound(r, 2)) As Integer
    
    Dim i As Integer
    For i = 1 To UBound(r, 2)
        v(i) = r(ind, i)
    Next i
    
    extraerVectorDeMatriz = v

End Function

' falta
Public Function borrarVectorDeMatriz(m() As Integer, ind As Integer) As Integer()
    
End Function


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





Public Function coleccionRecoridoEnVectorRecorrido(c As Collection) As Integer()
    ReDim r(cantOperaciones) As Integer
    
    Dim i As Integer
    ' el vector va del ultimo al primero
    For i = c.count To 1 Step -1
        r(c.count + 1 - i) = c(i).indice
    Next
    
    coleccionRecoridoEnVectorRecorrido = r
End Function

Private Function hayMasSubRecorridos(c() As Integer) As Boolean
    hayMasSubRecorridos = True
    If damePosiblesPrecedores(c())(1) = 0 Then hayMasSubRecorridos = False
End Function

Public Function dameIndiceUltimoLugarNoZero(c() As Integer) As Integer
    Dim i As Integer
    For i = 1 To UBound(c, 1)
        If c(i) = 0 Then
            dameIndiceUltimoLugarNoZero = i - 1
            Exit Function
        End If
    Next i
End Function

Private Function damePosiblesPrecedores(c() As Integer) As Integer()
    ReDim pp(cantOperaciones) As Integer
    
    Dim indicePrimeraOperacionDelRecorrido As Integer
    indicePrimeraOperacionDelRecorrido = c(dameIndiceUltimoLugarNoZero(c))
       
    Dim i As Integer
    For i = 1 To cantOperaciones
        If operacionPrecedores(indicePrimeraOperacionDelRecorrido, i) = 0 Then
            Exit For
        End If

        pp = vectorReemplazoPrimerCero(operacionPrecedores(indicePrimeraOperacionDelRecorrido, i), pp)
        
    Next i
            
    damePosiblesPrecedores = pp
            
End Function

Private Function obtenerSubRecorridos(r() As Integer) As Integer()
    ReDim sr(cantOperaciones, cantOperaciones) As Integer
    
    Dim indicePrimerOperacionDelRecorrido As Integer
    indicePrimerOperacionDelRecorrido = r(dameIndiceUltimoLugarNoZero(r))
        
    Dim i As Integer, j As Integer
    For i = 1 To cantOperaciones
        If operacionPrecedores(indicePrimerOperacionDelRecorrido, i) = 0 Then _
            Exit For
        
        ' copiar la primera parte de los recorridos
        For j = 1 To cantOperaciones
            If r(j) = 0 Then Exit For
            sr(i, j) = r(j)
        Next j
        
        sr = _
            matrizReemplazoPrimerCero(operacionPrecedores( _
                indicePrimerOperacionDelRecorrido, i), sr, i)
    Next i
    
    ' si no hubo posibles subrecorridos, devuelvo el recorrido original
    If sr(1, 1) = 0 Then
        For j = 1 To cantOperaciones
            If r(j) = 0 Then Exit For
            sr(i, j) = r(j)
        Next j
    End If
    
    obtenerSubRecorridos = sr
    
End Function


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



