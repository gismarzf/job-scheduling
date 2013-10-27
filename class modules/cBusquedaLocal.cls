VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBusquedaLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pMetaheuristica As cMetaheuristica
Private pListaTabu As Collection
Private pMaxListaTabu As Integer
Private countIgual As Integer, countOpt As Integer

Public Sub elegirProximaSolucion()
    Dim nuevaSol As cSolucion
    Dim tempSol As cSolucion
    Dim indice As Integer
    
    Dim min As Integer
    For Each tempSol In Metaheuristica.Vecindario
            indice = tempSol.IndiceDeDiferencia ' el indice del arco que fue cambiado en cada vecindario
        If Not coleccionContieneInteger(indice, ListaTabu) Then
            min = tempSol.Funcional ' busco el minimo de todos los vecindarios que NO son tabu
            Exit For                ' (para comparar despues)
        End If
    Next

    For Each tempSol In Metaheuristica.Vecindario
        indice = tempSol.IndiceDeDiferencia
        ' si el funcional es menor y no esta tabu, es la nueva solucion
        ' si es menor que TODOS los anteriores y si esta tabu, tambien
        If tempSol.Funcional <= min And Not coleccionContieneInteger(indice, _
            ListaTabu) Or coleccionContieneInteger(indice, ListaTabu) And _
            tempSol.Funcional < Metaheuristica.MejorSolucion.Funcional Then
            min = tempSol.Funcional
            Set Metaheuristica.Solucion = tempSol

        End If
    Next
End Sub

'*
'*
'*
'********************************************************************************
'FUNCIONES
'********************************************************************************
'*
'*
'*
Public Sub generarVecindarioCritico()
    
    Dim indices As Collection
    Set indices = Metaheuristica.Modelo.obtenerIndicesCriticos
    
    Dim i As Integer
    For i = 1 To indices.count
        Dim nuevaSolucion As New cSolucion
        Set nuevaSolucion = Metaheuristica.Solucion.copiaDeEstaSolucion
        
        ' invertir el valor de la posicion i
        nuevaSolucion.invertirBit (indices(i))

        nuevaSolucion.Funcional = _
            Metaheuristica.Modelo.obtenerFuncionalDeBitVector(nuevaSolucion.BitVector)
        nuevaSolucion.IndiceDeDiferencia = indices(i)
        Metaheuristica.Vecindario.Add nuevaSolucion

    Next i
        
End Sub

Public Sub generarVecindariosTodos()
    Dim start As Long
    Dim a As cArco
    Dim i As Integer
    For i = 1 To Metaheuristica.Modelo.ArcosDisyuntivos.count
    
        Dim nuevaSolucion As New cSolucion
        Set nuevaSolucion = Metaheuristica.Solucion.copiaDeEstaSolucion
        
        ' invertir el valor de la posicion i
        nuevaSolucion.invertirBit i
        
        ' la nueva solucion debe ser valida
        If Not mHayAlgunError(nuevaSolucion.BitVector) Then
            
            Metaheuristica.Modelo.implementarBitVector nuevaSolucion.BitVector

            nuevaSolucion.Funcional = _
                Metaheuristica.Modelo.testSolucion(nuevaSolucion.BitVector)
            nuevaSolucion.IndiceDeDiferencia = i
            Metaheuristica.Vecindario.Add nuevaSolucion

        End If

    Next i
        
End Sub

Public Sub start(iteraciones As Integer)

    Dim funcionalAnterior As Integer, funcionalActual As Integer
    
    Dim i As Integer
    For i = 1 To iteraciones
        ' saco el primer indice de la lista tabu
        If ListaTabu.count > MaxListaTabu Then ListaTabu.Remove (1)
        
        funcionalAnterior = Metaheuristica.Solucion.Funcional
        
        ' tengo que borrar el vecindario anterior
        Set Metaheuristica.Vecindario = New Collection
        
        ' genero los vecindarios (todos)
        generarVecindariosTodos
        elegirProximaSolucion
        
        funcionalActual = Metaheuristica.Solucion.Funcional
        
        If Metaheuristica.Solucion.Funcional < _
            Metaheuristica.MejorSolucion.Funcional Then
            Set Metaheuristica.MejorSolucion = _
                Metaheuristica.Solucion.copiaDeEstaSolucion
        End If
        
        ListaTabu.Add Metaheuristica.Solucion.IndiceDeDiferencia
       
    Next i
    
End Sub

Property Set ListaTabu(c As Collection)
    Set pListaTabu = c
End Property

Property Set Metaheuristica(m As cMetaheuristica)
    Set pMetaheuristica = m
End Property

Property Let MaxListaTabu(i As Integer)
    pMaxListaTabu = i
End Property

Property Get ListaTabu() As Collection
    Set ListaTabu = pListaTabu
End Property

'*
'*
'*
'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************
'*
'*
'*
Property Get MaxListaTabu() As Integer
    MaxListaTabu = pMaxListaTabu
End Property

Property Get Metaheuristica() As cMetaheuristica
    Set Metaheuristica = pMetaheuristica
End Property

