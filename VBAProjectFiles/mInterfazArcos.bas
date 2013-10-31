Attribute VB_Name = "mInterfazArcos"
Option Explicit
Option Base 1

Public operacionSucedores() As Integer
Public operacionPrecedores() As Integer
Public operacionPeso() As Integer
Public arcosBit() As Boolean
Public cantOperaciones As Integer
Public cantTrabajos As Integer
Public cantOpxTrabajo As Integer

Private arcosCInicioOperacion() As Integer
Private arcosCFinalOperacion() As Integer
Private arcosDInicioOperacion() As Integer
Private arcosDFinalOperacion() As Integer

'********************************************************************************
Public Sub mInitInterfazArcos(m As cModeloDisyuntivo)
    cantOperaciones = m.Operaciones.count
    cantTrabajos = m.Trabajos.count
    cantOpxTrabajo = m.cantOperacionesPorTrabajo
    
    ReDim arcosBit(m.ArcosDisyuntivos.count)

    ' inicializar los vectores p arcos, max = cant. arcos disyuntivos
    initArcosConyuntivos m.ArcosConyuntivos, m.ArcosDisyuntivos.count
    initArcosDisyuntivos m.ArcosDisyuntivos, m.ArcosDisyuntivos.count
        
    ReDim operacionPeso(cantOperaciones)
    Dim i As Integer
    For i = 1 To cantOperaciones
        operacionPeso(i) = m.Operaciones(i).Duracion
    Next i

    cargarVecinos
    
End Sub

Public Sub mImplementarBitVector(c As Collection)
    ReDim arcosBit(c.count)

    
    Dim i As Integer
    For i = 1 To c.count
        arcosBit(i) = c(i)
    Next i

    cargarVecinos

End Sub

'********************************************************************************

Private Sub initArcosDisyuntivos(c As Collection, max As Integer)
    ReDim arcosDInicioOperacion(max)
    ReDim arcosDFinalOperacion(max)
    
    Dim i As Integer
    For i = 1 To c.count
        If c(i).Direccion = True Then
            arcosDInicioOperacion(i) = c(i).InicioOperacion.indice
            arcosDFinalOperacion(i) = c(i).FinalOperacion.indice
        ElseIf c(i).Direccion = False Then
            arcosDFinalOperacion(i) = c(i).InicioOperacion.indice
            arcosDInicioOperacion(i) = c(i).FinalOperacion.indice
        End If
    Next i
    
End Sub

Private Sub initArcosConyuntivos(c As Collection, max As Integer)
    ReDim arcosCInicioOperacion(max)
    ReDim arcosCFinalOperacion(max)
    
    Dim i As Integer
    For i = 1 To c.count
        If c(i).Direccion = True Then
            arcosCInicioOperacion(i) = c(i).InicioOperacion.indice
            arcosCFinalOperacion(i) = c(i).FinalOperacion.indice
        ElseIf c(i).Direccion = False Then
            arcosCFinalOperacion(i) = c(i).InicioOperacion.indice
            arcosCInicioOperacion(i) = c(i).FinalOperacion.indice
        End If
    Next i
    
End Sub

Private Sub cargarVecinos()

    ReDim operacionSucedores(cantOperaciones, cantOperaciones)
    ReDim operacionPrecedores(cantOperaciones, cantOperaciones)
    ReDim tempSucedores(cantOperaciones) As Integer
    ReDim tempPrecedores(cantOperaciones) As Integer

    Dim i As Integer, j As Integer
    For i = 1 To cantOperaciones
        ' dame un vector con todos los proximas operaciones de la operacion i
        tempSucedores = dameSucedores(i)
        For j = 1 To cantOperaciones
            If tempSucedores(j) = 0 Then
                Exit For
            Else
                operacionSucedores(i, j) = tempSucedores(j)
            End If
        Next j
    Next i
    
    For i = 1 To cantOperaciones
        ' dame un vector con todos los anteriores operaciones de la operacion i
        tempPrecedores = damePrecedores(i)
        For j = 1 To cantOperaciones
            If tempPrecedores(j) = 0 Then
                Exit For
            Else
                operacionPrecedores(i, j) = tempPrecedores(j)
            End If
        Next j
    Next i

End Sub

Private Function dameSucedores(o As Integer) As Integer()

    ReDim tempSucedores(cantOperaciones) As Integer
    
    Dim i As Integer
    For i = 1 To UBound(arcosDInicioOperacion, 1)
        
        ' los conyuntivos
        If arcosCInicioOperacion(i) = o Then
            tempSucedores = vectorReemplazoPrimerCero(arcosCFinalOperacion(i), _
                tempSucedores)
        End If

        ' los disyuntivos
        If arcosDInicioOperacion(i) = o And arcosBit(i) = True Then
            tempSucedores = vectorReemplazoPrimerCero(arcosDFinalOperacion(i), _
                tempSucedores)
        End If
    
        If arcosDFinalOperacion(i) = o And arcosBit(i) = False Then
            tempSucedores = vectorReemplazoPrimerCero(arcosDInicioOperacion(i), _
                tempSucedores)
        End If
    Next i


    dameSucedores = tempSucedores

End Function

Private Function damePrecedores(o As Integer) As Integer()

    ReDim tempPrecedores(cantOperaciones) As Integer
    
    Dim i As Integer
    For i = 1 To UBound(arcosDInicioOperacion, 1)
        
        ' los conyuntivos
        If arcosCFinalOperacion(i) = o Then
            tempPrecedores = vectorReemplazoPrimerCero(arcosCInicioOperacion(i), _
                tempPrecedores)
        End If

        ' los disyuntivos
        If arcosDFinalOperacion(i) = o And arcosBit(i) = True Then
            tempPrecedores = vectorReemplazoPrimerCero(arcosDInicioOperacion(i), _
                tempPrecedores)
        End If
    
        If arcosDInicioOperacion(i) = o And arcosBit(i) = False Then
            tempPrecedores = vectorReemplazoPrimerCero(arcosDFinalOperacion(i), _
                tempPrecedores)
        End If
    Next i


    damePrecedores = tempPrecedores

End Function
