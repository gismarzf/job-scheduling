VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pBitVector As New Collection
Private pFuncional As Integer
Private pErrores As Integer
Private pIndiceDeDiferencia As Integer
'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************

Property Get IndiceDeDiferencia() As Integer
    IndiceDeDiferencia = pIndiceDeDiferencia
End Property

Property Let IndiceDeDiferencia(i As Integer)
    pIndiceDeDiferencia = i
End Property
Property Get Errores() As Integer
    Errores = pErrores
End Property

Property Let Errores(i As Integer)
    pErrores = i
End Property

Property Get BitVector() As Collection
    Set BitVector = pBitVector
End Property

Property Set BitVector(c As Collection)
    Set pBitVector = c
End Property

Property Get Funcional() As Integer
    Funcional = pFuncional
End Property

Public Property Let Funcional(i As Integer)
    pFuncional = i
End Property

'********************************************************************************
'FUNCIONES
'********************************************************************************

Public Function invertirBit(posicion As Integer)
    ' tengo que remover el bit de la posicion
    ' despues tengo que agregar uno antes de la nueva posicion (ocupada por el prox. bit)
    ' si la posicion es la ultima tengo que agregar el bit nuevo a la lista
        
    If BitVector(posicion) = True Then
    
        If Not posicion = BitVector.count Then
            BitVector.Remove posicion
            BitVector.Add False, , posicion
        ElseIf posicion = BitVector.count Then
            BitVector.Remove posicion
            BitVector.Add True
        End If

    ElseIf BitVector(posicion) = False Then
        If Not posicion = BitVector.count Then
            BitVector.Remove posicion
            BitVector.Add True, , posicion
        ElseIf posicion = BitVector.count Then
            BitVector.Remove posicion
            BitVector.Add True
        End If
    End If

End Function

Public Function copiaDeEstaSolucion() As cSolucion
    Dim s As New cSolucion
    Set s = New cSolucion
    
    Dim b As Boolean
    Dim i As Integer
    For i = 1 To BitVector.count
        s.BitVector.Add BitVector(i)
    Next
    
    s.Funcional = Funcional
    s.Errores = Errores
    s.IndiceDeDiferencia = IndiceDeDiferencia
    
    Set copiaDeEstaSolucion = s
End Function

Public Function esIgual(sol2 As cSolucion) As Boolean
    esIgual = True
    
    Dim i As Integer
    For i = 1 To Me.BitVector.count
        If Me.BitVector(i) <> sol2.BitVector(i) Then esIgual = False
    Next
End Function

