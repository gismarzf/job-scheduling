VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFabricaDeSolucionesIniciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pMetaheuristica As cMetaheuristica
Private pListaTabu As Collection
Private pVecindario As New Collection
Private pMaxListaTabu As Integer
'*
'*
'*
'********************************************************************************
'FUNCIONES
'********************************************************************************
'*
'*
'*
Public Function porTodos1s(longitud As Integer) As cSolucion
    Dim Solucion As New cSolucion
    Dim i As Integer
    For i = 1 To longitud
        Solucion.BitVector.Add True
    Next i
    Solucion.Funcional = Metaheuristica.Modelo.testSolucion(Solucion.BitVector)
    
    Set porTodos1s = Solucion
End Function


Private Function generarSolucionAlAzar(longitud As Integer) As cSolucion
    Dim s As New cSolucion
    Dim i As Integer
    For i = 1 To longitud
        If Rnd() > 0.5 Then
            s.BitVector.Add True
        Else: s.BitVector.Add False
        End If
    Next i

    Set generarSolucionAlAzar = s
End Function

Public Function porListaTabu(longitud As Integer) As cSolucion
    Dim Solucion As cSolucion
        
    Set Solucion = generarSolucionAlAzar(longitud)
    
    Solucion.Errores = mCalcularNumeroDeRelacionesCirculares(Solucion.BitVector)


    ' uso una busqueda local para llegar a 0 errores..
    ' es medio sucio, teoricamente podria quedarme en un local minimo y nunca llegar
    ' pero como hay muchas configuraciones que me den 0 errores, soy optimista
    
    Set listaTabu = New Collection
    Dim count As Integer
    count = 0
    
    While Solucion.Errores > 0
    
        Set Vecindario = New Collection
        
        If listaTabu.count > MaxListaTabu Then
            listaTabu.Remove 1
        End If
        
        Dim s As New cSolucion
        Set s = Solucion.copiaDeEstaSolucion
        
        Dim i As Integer
        i = Int((s.BitVector.count - 1 + 1) * Rnd + 1)
        
        While coleccionContieneInteger(i, listaTabu)
            i = Int((s.BitVector.count - 1 + 1) * Rnd + 1)
        Wend
        
        s.invertirBit i
        
        s.Errores = mCalcularNumeroDeRelacionesCirculares(s.BitVector)
        
        If s.Errores <= Solucion.Errores Then
            Set Solucion = s.copiaDeEstaSolucion
            listaTabu.Add i
        End If
        count = count + 1
        
    Wend
    
    Solucion.Funcional = Metaheuristica.Modelo.testSolucion(s.BitVector)
    
    Set porListaTabu = Solucion
    
End Function
'*
'*
'*
'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************
'*
'*
'*
Property Get Metaheuristica() As cMetaheuristica
    Set Metaheuristica = pMetaheuristica
End Property

Property Set Metaheuristica(m As cMetaheuristica)
    Set pMetaheuristica = m
End Property

Property Get Solucion() As cSolucion
    Set Solucion = pSolucion
End Property

Property Set Solucion(s As cSolucion)
    Set pSolucion = s
End Property

Property Get listaTabu() As Collection
    Set listaTabu = pListaTabu
End Property

Property Set listaTabu(c As Collection)
    Set pListaTabu = c
End Property

Property Get Vecindario() As Collection
    Set Vecindario = pVecindario
End Property

Property Set Vecindario(c As Collection)
    Set pVecindario = c
End Property

Property Get MaxListaTabu() As Integer
    MaxListaTabu = pMaxListaTabu
End Property

Property Let MaxListaTabu(i As Integer)
    pMaxListaTabu = i
End Property


