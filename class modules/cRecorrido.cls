VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRecorrido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pSuma As Integer
Private pRecorrido As Collection
Private pArcos As Collection

Private Sub Class_Initialize()
    Set Recorrido = New Collection
    Set Arcos = New Collection
End Sub

'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************

Property Get Arcos() As Collection
    Set Arcos = pArcos
End Property

Property Set Arcos(c As Collection)
    Set pArcos = c
End Property

Property Get Suma() As Integer
    Suma = pSuma
End Property

Property Let Suma(i As Integer)
    pSuma = i
End Property

Property Get Recorrido() As Collection
    Set Recorrido = pRecorrido
End Property

Property Set Recorrido(c As Collection)
    Set pRecorrido = c
End Property

'********************************************************************************
'FUNCIONES
'********************************************************************************

Public Sub agregarOperacionAntesDeTodos(o As cOperacion)
    
    ' agregar la operacion como primer operacion en la cola
    Recorrido.Add o, , 1
    
    ' agregamos el arco que esta entre la primera y segunda operacion, tambien asi
    agregarArcoAntesDeTodos o, Recorrido(2)
    
End Sub

Public Function primerOperacionDelRecorrido() As cOperacion
    Set primerOperacionDelRecorrido = Recorrido(1)
End Function

Public Sub agregarArcoAntesDeTodos(operacion1 As cOperacion, operacion2 As cOperacion)
    
    Dim a As cArco
    For Each a In operacion1.Arcos
        If (a.FinalOperacion Is operacion2) And (a.InicioOperacion Is operacion1) Then
            ' si la lista es vacia agrego al fin, sino agrego al ppo
            If Arcos.count = 0 Then
                Arcos.Add a
            Else
                Arcos.Add a, , 1
            End If
        End If
    Next
        
End Sub

Public Function copiarRecorrido() As cRecorrido
    Dim c As cRecorrido
    Set c = New cRecorrido
    
    c.Suma = Suma
    
    Dim o As cOperacion
    For Each o In Recorrido
        c.Recorrido.Add o
    Next
        
    Dim a As cArco
    For Each a In Arcos
        c.Arcos.Add a
    Next
        
    Set copiarRecorrido = c
    
End Function

Public Function sumarPesos() As Integer

     Dim temp As Integer
    
    Dim o As cOperacion
    For Each o In Recorrido
        temp = temp + o.Duracion
    Next
    
    sumarPesos = temp
End Function

' devuelve el recorrido hasta a.iniciooperacion, inclusive
Public Function dameColaInferior(o As cOperacion) As cRecorrido
    Dim rnew As New cRecorrido
    Dim a As cArco
        
    If dameIndiceDeOperacion(o) = 1 Then
        rnew.Recorrido.Add o
        Set dameColaInferior = rnew
    ElseIf dameIndiceDeOperacion(o) = Recorrido.count Then
        Set rnew = Me
    Else
        Set a = dameProximoArco(o)
        ' agregar todos los arcos, hasta que llego al arco donde corto
        Dim i As Integer
        For i = 1 To Arcos.count
            If Not a Is Arcos(i) Then
                rnew.Arcos.Add Arcos(i)
            Else
                Exit For
            End If
        Next i
        
        For i = 1 To rnew.Arcos.count
            rnew.Recorrido.Add Arcos(i).InicioOperacion
            
            ' el ultimo arco debe tener la ultima operacion
            If i = rnew.Arcos.count Then rnew.Recorrido.Add Arcos(i).FinalOperacion
        Next i
    End If
    
    Set dameColaInferior = rnew

End Function


' devuelve el recorrido desde el punto de corte hasta la ultima operacion, inclusive
' TAMBIEN devuelva el arco de corte!!
Public Function dameColaSuperior(o As cOperacion) As cRecorrido
    Dim rnew As New cRecorrido
    Dim i As Integer
    
    ' si el indice es la ultima operacion, devuelvo todo el recorrido original
    If dameIndiceDeOperacion(o) = Recorrido.count Then
        rnew.Recorrido.Add o
        Set dameColaSuperior = Me
        
    ElseIf dameIndiceDeOperacion(o) = 1 Then
        Set rnew = Me
    Else
    
        Dim a As cArco
        
        Set a = dameProximoArco(o)

        ' agregar todos los arcos desde el punto de corte hasta el ultimo
        For i = dameIndiceDelArco(a) To Arcos.count
            If Not a Is Arcos(i) Then
                rnew.Arcos.Add Arcos(i)
            Else
                rnew.Arcos.Add Arcos(i)
                rnew.Recorrido.Add Arcos(i).InicioOperacion
            End If
        Next i
            
        For i = 1 To rnew.Arcos.count
            rnew.Recorrido.Add rnew.Arcos(i).FinalOperacion
        Next i
        
    End If
    Set dameColaSuperior = rnew
End Function

Public Function dameIndiceDelArco(a As cArco) As Integer
        
    ' buscar el indice del arco dado
    Dim i As Integer
    For i = 1 To Arcos.count
        If Arcos(i) Is a Then dameIndiceDelArco = i
    Next i

End Function

Public Function dameIndiceDeOperacion(o As cOperacion) As Integer
    
    ' buscar el indice del arco dado
    Dim i As Integer
    For i = 1 To Recorrido.count
        If Recorrido(i) Is o Then dameIndiceDeOperacion = i
    Next i
    
End Function

Public Function mismoRecorrido(rec As cRecorrido) As Boolean
    Dim i As Integer
    mismoRecorrido = True
    
    If rec.Recorrido.count = Recorrido.count Then
        For i = 1 To Recorrido.count
            If rec.Recorrido(i).indice <> Recorrido(i).indice Then mismoRecorrido = False
        Next i
    Else
        mismoRecorrido = False
    End If
    
End Function

Public Function dameProximoArco(o As cOperacion) As cArco
    Dim a As cArco
    
    ' si la operacion es la ultima en la lista
    For Each a In Arcos
            If a.InicioOperacion Is o Then Set dameProximoArco = a
    Next
    
End Function

Public Function dameArcoAnterior(o As cOperacion) As cArco
    Dim a As cArco
    
    ' si es la primera operacion
    If dameIndiceDeOperacion(o) = 1 Then
        Set a = Arcos(1)
    Else
    
    For Each a In Arcos
        If a.FinalOperacion Is o Then Set dameProximoArco = a
    Next
End Function

Public Function marcar()
    Dim a As cArco
    For Each a In Arcos
        a.Conector.marcarLinea
    Next
End Function
