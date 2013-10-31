Attribute VB_Name = "mTestSolucion"
Option Explicit
Option Base 1

Public Const maxRecorridos As Integer = 256
Public recorridosCompletos() As Integer


Public Function coleccionRecoridoEnVectorRecorrido(c As Collection) As Integer()
    ReDim r(cantOperaciones) As Integer

    Dim i As Integer
    ' el vector va del ultimo al primero
    For i = c.count To 1 Step -1
        r(c.count + 1 - i) = c(i).indice
    Next

    coleccionRecoridoEnVectorRecorrido = r
End Function

Public Function contarHastaPrimerVectorZero(m() As Integer) As Integer
    Dim i As Integer
    For i = 1 To UBound(m, 1)
        If m(i, 1) = 0 Then
            contarHastaPrimerVectorZero = i - 1
            Exit Function
        End If
    Next i
    
    contarHastaPrimerVectorZero = UBound(m, 1)
End Function

Public Function contarHastaPrimerZero(v() As Integer) As Integer
    Dim i As Integer
    For i = 1 To UBound(v, 1)
        If v(i) = 0 Then
            contarHastaPrimerZero = i - 1
            Exit Function
        End If
    Next i
    
    contarHastaPrimerZero = UBound(v, 1)
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

Public Function eliminarVectorDeMatriz(m() As Integer, ind As Integer) As Integer()

    ReDim mtemp(UBound(m, 1), UBound(m, 2)) As Integer
    
    Dim i As Integer, j As Integer
    For i = 1 To UBound(m, 1)
        If i < ind Then
            For j = 1 To UBound(m, 2)
                mtemp(i, j) = m(i, j)
            Next j
        ElseIf i > ind Then
            For j = 1 To UBound(m, 2)
                mtemp(i - 1, j) = m(i, j)
            Next j
        End If
    Next i

    eliminarVectorDeMatriz = mtemp

End Function

Public Sub extraerRecorridosCompletos(r() As Integer)

    Dim i As Integer
    For i = 1 To maxRecorridos
        If r(i, 1) = 0 Then Exit Sub
        Dim v() As Integer
        v = extraerVectorDeMatriz(r, i)
        If Not hayMasSubRecorridos(v) Then
            r = eliminarVectorDeMatriz(r, i)
            reemplazarPrimerVectorZero recorridosCompletos, v
            i = i - 1 ' tengo que hacer esto, ya que la matriz se achico
        End If
    Next i
       
End Sub

Public Function extraerVectorDeMatriz(r() As Integer, ind As Integer) As Integer()
    ReDim v(cantOperaciones) As Integer

    Dim i As Integer
    For i = 1 To cantOperaciones
        v(i) = r(ind, i)
    Next i

    extraerVectorDeMatriz = v

End Function

Private Function hayMasRecorridos(r() As Integer) As Boolean
    
    Dim i As Integer
    For i = 1 To UBound(r, 1)
        Dim v() As Integer
        v = extraerVectorDeMatriz(r, i)
        If hayMasSubRecorridos(v) Then
            hayMasRecorridos = True
            Exit Function
        End If
    Next i
    
End Function

Private Function hayMasSubRecorridos(c() As Integer) As Boolean
    hayMasSubRecorridos = True
    If damePosiblesPrecedores(c())(1) = 0 Then hayMasSubRecorridos = False
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

        reemplazarPrimerVectorZero m1, v
    Next i

End Function

Public Function mCalcularFuncional() As Integer

    ReDim recorridosCompletos(maxRecorridos, cantOperaciones)

    obtenerTodosRecorridos ' SUPER slow
        
    Dim sum As Integer, maxsum As Integer
    maxsum = 0

    Dim j As Integer, i As Integer
    For i = 1 To maxRecorridos
        If recorridosCompletos(i, 1) = 0 Then Exit For
        
        sum = 0
        For j = 1 To cantOperaciones
            If recorridosCompletos(i, j) = 0 Then Exit For
            sum = sum + operacionPeso(recorridosCompletos(i, j))
        Next j
        
        If sum > maxsum Then
            maxsum = sum
        End If
    Next i
    
    mCalcularFuncional = maxsum
    
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

Public Function obtenerTodosRecorridos()
    ReDim recorridosIncompletos(maxRecorridos, cantOperaciones) As Integer

    'agregar las ultimas operaciones como ptos de partida1
    Dim i As Integer
    For i = 1 To cantTrabajos
        recorridosIncompletos(i, 1) = i * cantOpxTrabajo
    Next i

    ' como sacamos cada vez un camino completo de la matriz de recorridos incompletos
    ' cuando no haya mas (primer indice zero) es pq obtuvimos todos
    While recorridosIncompletos(1, 1) <> 0
        Dim subrecorridos() As Integer
        For i = 1 To maxRecorridos
            If recorridosIncompletos(i, 1) = 0 Then Exit For
            Dim recorrido() As Integer
            recorrido = extraerVectorDeMatriz(recorridosIncompletos, i)
            subrecorridos = juntarMatriz(subrecorridos, obtenerSubRecorridos(recorrido))
        Next i
    
        extraerRecorridosCompletos subrecorridos
        recorridosIncompletos = subrecorridos
    Wend
End Function

' reemplaza el primer vector que empieza con 0 de una matriz
Public Function reemplazarPrimerVectorZero(mat() As Integer, v() As Integer) As Integer()

    Dim mtemp() As Integer
    mtemp = copiarMatriz(mat)

    Dim i As Integer, j As Integer
    For i = 1 To maxRecorridos
        If mat(i, 1) = 0 Then
            For j = 1 To cantOperaciones
                mat(i, j) = v(j)
            Next j
            reemplazarPrimerVectorZero = mat
            Exit Function
        End If
    Next i

End Function

