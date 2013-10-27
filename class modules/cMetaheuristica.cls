VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMetaheuristica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pSolucion As New cSolucion, pMejorSolucion As New cSolucion
Private pVecindarioCritico As New Collection
Private pVecindario As New Collection
Private pModelo As cModeloDisyuntivo
'*
'*
'*
'********************************************************************************
'FUNCIONES
'********************************************************************************
'*
'*
'*
Public Sub implementarSolucion()
    Modelo.implementarBitVector Solucion.BitVector
End Sub

Public Sub implementarMejorSolucion()
    Modelo.implementarBitVector MejorSolucion.BitVector
    Modelo.actualizar
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
Property Get Solucion() As cSolucion
    Set Solucion = pSolucion
End Property

Property Set Solucion(s As cSolucion)
    Set pSolucion = s
End Property

Property Get MejorSolucion() As cSolucion
    Set MejorSolucion = pMejorSolucion
End Property

Property Set MejorSolucion(s As cSolucion)
    Set pMejorSolucion = s
End Property

Property Get VecindarioCritico() As Collection
    Set VecindarioCritico = pVecindarioCritico
End Property

Property Set VecindarioCritico(c As Collection)
    Set pVecindarioCritico = c
End Property

Property Get Vecindario() As Collection
    Set Vecindario = pVecindario
End Property

Property Set Vecindario(c As Collection)
    Set pVecindario = c
End Property

Property Get Modelo() As cModeloDisyuntivo
    Set Modelo = pModelo
End Property

Property Set Modelo(m As cModeloDisyuntivo)
    Set pModelo = m
End Property
