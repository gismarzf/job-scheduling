Attribute VB_Name = "mCalculoDeErrores"
Option Explicit
Option Base 1


'********************************************************************************

Public Function mCalcularNumeroDeRelacionesCirculares(bv As Collection) As Integer
    
    mImplementarBitVector bv
    
    Dim count As Integer
    count = 0
    
    Dim i As Integer
    For i = 1 To cantOperaciones
        If relacionCircular(i) Then
            count = count + 1
        End If
    Next i
    
    mCalcularNumeroDeRelacionesCirculares = count
    
End Function

Public Function mHayAlgunError(bv As Collection) As Boolean
    
    mImplementarBitVector bv

    Dim i As Integer
    For i = 1 To cantOperaciones
        If relacionCircular(i) Then
            mHayAlgunError = True
            Exit Function
        End If
    Next i
    
End Function
Private Function relacionCircular(operacion As Integer) As Boolean

    ReDim col(1 To cantOperaciones) As Integer
    
    ' salimos si no hay ningun sucedor
    If operacionSucedores(operacion, 1) = 0 Then
        relacionCircular = False
        Exit Function
    End If
        
    Dim i As Integer, j As Integer
    For i = 1 To cantOperaciones
        If operacionSucedores(operacion, i) = 0 Then
            ' break cuando no hay mas sucedores
            Exit For
        Else
            col = vectorReemplazoPrimerCero(operacionSucedores(operacion, i), _
                col)
        End If
    Next i

    While hayMasOperacionesACualesPuedoLlegar(col)
        For i = 1 To UBound(col, 1)
            For j = 1 To cantOperaciones
            If col(i) = 0 Then Exit For ' break si no hay mas suc. en el camino
            
            If operacionSucedores(col(i), j) = 0 Then
                Exit For
            ElseIf operacionSucedores(col(i), j) = operacion Then
                relacionCircular = True
                Exit Function
            Else
                If Not vectorContieneInteger(operacionSucedores(col(i), j), _
                    col) Then
                    col = vectorReemplazoPrimerCero(operacionSucedores(col(i), _
                        j), col)
                End If
            End If
            Next j
        Next i
    Wend
    
End Function

Private Function hayMasOperacionesACualesPuedoLlegar(camino() As Integer) As _
    Boolean
    
    ' por cada operacion veo si las posibles sucedores ya estan en mi lista o no
    Dim i As Integer, j As Integer
    For i = 1 To UBound(camino, 1)
        For j = 1 To cantOperaciones
        If camino(i) = 0 Then Exit For
        If operacionSucedores(camino(i), j) = 0 Then
            Exit For
        Else
            If Not vectorContieneInteger(operacionSucedores(camino(i), j), _
                camino) Then
                hayMasOperacionesACualesPuedoLlegar = True
                Exit Function
            End If
        End If
        Next j
    Next i

End Function

