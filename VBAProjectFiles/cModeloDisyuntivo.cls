VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cModeloDisyuntivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public cantMaquinas As Integer
Public cantTrabajos As Integer
Public cantOperacionesPorTrabajo As Integer
Public cantOperacionesPorMaquina As Integer

Private pTrabajos As New Collection
Private pMaquinas As New Collection

Private pOperaciones As New Collection
Private pOperacionesConMismoTrabajo() As New cOperacion
Private pOperacionesEnMismaMaquina() As New cOperacion

Private pArcosConyuntivos As New Collection
Private pArcosDisyuntivos As New Collection

Private pPlanillaDatos As New Worksheet
Private pPlanillaGraficoDisyuntivo As New Worksheet
Private pPlanillaGraficoGantt As New Worksheet

Private pDiagrama As cDiagrama

Private pFabricaDeRecorridos As cFabricaDeRecorridos

'********************************************************************************
'FUNCIONES
'********************************************************************************
Public Function calcularRecorridos()
    ' me actualiza los recorridos de la fabrica
    Set FabricaDeRecorridos.recorridos = FabricaDeRecorridos.obtenerRecorridosCritico
    Set FabricaDeRecorridos.RecorridoCritico = _
                FabricaDeRecorridos.elegirRecorridoCritico(FabricaDeRecorridos.recorridos)
End Function

Public Function testSolucion(bvector As Collection) As Integer

    Dim valant As New Collection

    Dim i As Integer
    For i = 1 To ArcosDisyuntivos.count
        ' guardar los valores anteriores
        valant.Add ArcosDisyuntivos(i).Direccion
        ArcosDisyuntivos(i).Direccion = bvector(i)
    Next i
    
    testSolucion = _
        FabricaDeRecorridos.elegirRecorridoCritico(FabricaDeRecorridos.obtenerRecorridosCritico).sumarPesos
    
    For i = 1 To ArcosDisyuntivos.count
        ' volver a los valores anteriores
        ArcosDisyuntivos(i).Direccion = valant(i)
    Next i
    
End Function

Public Function obtenerIndicesCriticos() As Collection
    Dim indices As New Collection
 
    Dim i As Integer
    For i = 1 To ArcosDisyuntivos.count
        If coleccionContieneObjeto(ArcosDisyuntivos(i), FabricaDeRecorridios.RecorridoCritico.Arcos) Then indices.Add i
    Next
            
    Set obtenerIndicesCriticos = indices
            
End Function

Public Sub implementarBitVector(bvector As Collection)
    
    Dim i As Integer
    For i = 1 To ArcosDisyuntivos.count
        ' guardar los valores anteriores
        ArcosDisyuntivos(i).Direccion = bvector(i)
    Next i

End Sub

Public Function Errores(bvector As Collection) As Integer
  '  Errores = mCalcularNumeroDeRelacionesCirculares(bvector)
    Errores = relacionesCirculares
End Function


Public Function hayAlgunError(bvector As Collection) As Integer
    hayAlgunError = mHayAlgunError(bvector)
End Function


Public Function esArcosDisyuntivosCompatible(Solucion As Collection) As Boolean
    
    Dim tempArcos As New Collection
    Set tempArcos = ArcosDisyuntivos
    
    Set ArcosDisyuntivos = Solucion
    If Not existeAlgunaRelacionCircular Then esSolucionCompatible = True

    ' volver a poner los arcos iniciales
    Set ArcosDisyuntivos = tempArcos

End Function

Public Sub programarOperaciones()

    Dim o As cOperacion
    For Each o In Operaciones
            o.HorarioFin = FabricaDeRecorridos.obtenerRecorridoCriticoDesdeMemoria(o).Suma
            o.HorarioInicio = o.HorarioFin - o.Duracion
    Next
    
End Sub

Public Function printOrden(c As Collection)
    
    Dim i As Integer
    For i = 1 To c.count
       PlanillaDatos.Cells(30 + i, 2).Value = c(i).indice
       PlanillaDatos.Cells(30 + i, 3).Value = c(i).Maquina.Numero
       PlanillaDatos.Cells(30 + i, 4).Value = c(i).Trabajo.Numero
       PlanillaDatos.Cells(30 + i, 5).Value = c(i).Duracion
        
       PlanillaDatos.Cells(30 + i, 6).Value = PlanillaDatos.Cells(30 + i - 1, 6).Value + c(i).Duracion
    Next i
    
End Function

Public Function algunaOperacionSinHacer() As Boolean
    
    Dim o As cOperacion
    For Each o In Operaciones
        If o.Completado = False Then algunaOperacionSinHacer = True
    Next
    
End Function

Public Function relacionesCirculares() As Integer
    
    Dim i As Integer
    i = 0
    Dim o As cOperacion
    For Each o In Operaciones
        If relacionCircular(o) Then
            i = i + 1
        End If

    Next
    
    relacionesCirculares = i
    
End Function

Private Function relacionCircular(operacion As cOperacion) As Boolean
    If coleccionContieneObjeto(operacion, obtenerTodasOperacionesACualesPuedoLLegar(operacion)) Then _
        relacionCircular = True
End Function

Private Function obtenerTodasOperacionesACualesPuedoLLegar(operacion As cOperacion) As Collection
    
    Dim c As New Collection
    Dim o As New cOperacion, o2 As New cOperacion
    
    ' agrego los primeros operaciones siguientes
    For Each o In operacion.posiblesSucedores
        c.Add o
    Next
    
    While hayMasOperacionesACualesPuedoLlegar(c)
    
        ' para cada operacion en c
        For Each o In c
            ' busco los siguientes
            For Each o2 In o.posiblesSucedores
                ' y las agrego a c
                If Not coleccionContieneObjeto(o2, c) Then c.Add o2
            Next
        Next
        
    Wend
    
    Set obtenerTodasOperacionesACualesPuedoLLegar = c
        
    
End Function

Private Function hayMasOperacionesACualesPuedoLlegar(operacionesAnteriores As Collection) As Boolean
    ' por cada operacion veo si las posibles sucedores ya estan en mi lista o no
    Dim o As cOperacion, o2 As cOperacion
    For Each o In operacionesAnteriores
        For Each o2 In o.posiblesSucedores
            If Not coleccionContieneObjeto(o2, operacionesAnteriores) Then _
                                                    hayMasOperacionesACualesPuedoLlegar = True
        Next
    Next
End Function

Public Sub setDireccionesDeArcosDisyuntivosAlAzar()

    Dim a As cArco
    For Each a In ArcosDisyuntivos
        If Rnd() < 0.5 Then
            a.Direccion = True
        Else
            a.Direccion = False
        End If
    Next

End Sub

Public Sub setDireccionesDeArcosConyuntivos()

    Dim a As cArco
    For Each a In ArcosConyuntivos
        a.Direccion = True
    Next

End Sub

Private Sub mostrarTodosArcosDisyuntivos()
    
    Dim a As cArco
    For Each a In ArcosDisyuntivos
        a.show
    Next
    
End Sub

Private Sub soloMostrarEsteArcoDisyuntivo(nroMaquina As Integer)
    
    mostrarArcosDisyuntivos (nroMaquina)
    
    Dim i As Integer
    For i = 1 To cantMaquinas
        If i <> nroMaquina Then ocultarArcosDisyuntivos (i)
    Next i
    
End Sub

Private Sub mostrarArcosDisyuntivos(nroMaquina As Integer)

    Dim a As cArco
    For Each a In ArcosDisyuntivos
        If a.Maquina.Numero = nroMaquina Then a.show
    Next
    
End Sub

Private Sub ocultarArcosDisyuntivos(nroMaquina As Integer)
    
    Dim a As cArco
    For Each a In ArcosDisyuntivos
        If a.Maquina.Numero = nroMaquina Then a.hide
    Next

End Sub

Public Sub actualizar()
    calcularRecorridos
    programarOperaciones
    
    Diagrama.actualizarDiagramaDisyuntivo
    Diagrama.actualizarDiagramaGantt
    Diagrama.actualizarRecorridoCritico
    Diagrama.actualizarPlanilla
End Sub

Public Sub inicializar(PlanillaDatos As Worksheet, PlanillaGraficoDisyuntivo As Worksheet, _
                                                    PlanillaGraficoGantt As Worksheet)

    Set Diagrama = New cDiagrama
    
    Set Me.PlanillaDatos = PlanillaDatos
    Set Me.PlanillaGraficoDisyuntivo = PlanillaGraficoDisyuntivo
    Set Me.PlanillaGraficoGantt = PlanillaGraficoGantt
    
    Set Diagrama.PlanillaGraficoDisyuntivo = Me.PlanillaGraficoDisyuntivo
    Set Diagrama.PlanillaGraficoGantt = Me.PlanillaGraficoGantt
    
    cantTrabajos = PlanillaDatos.Range("TRABAJOS").Cells(1, 1).Value
    cantMaquinas = PlanillaDatos.Range("MAQUINAS").Cells(1, 1).Value
    cantOperacionesPorTrabajo = PlanillaDatos.Range("OPXTRABAJO").Cells(1, 1).Value
    cantOperacionesPorMaquina = PlanillaDatos.Range("OPXMAQUINA").Cells(1, 1).Value
    
    removerGraficos
    
    crearTrabajos
    crearMaquinas
    
    crearOperaciones
       
    crearVectorConIgualesTrabajos
    crearVectorConIgualesMaquinas
    crearColeccionDeArcosConyuntivos
    crearColeccionDeArcosDisyuntivos
    
    setDireccionesDeArcosConyuntivos
    
    Set Diagrama.Operaciones = Operaciones
    Set Diagrama.Arcos = sumarColecciones(ArcosConyuntivos, ArcosDisyuntivos)
    Set Diagrama.Modelo = Me
    
    Diagrama.crearDiagramaDisyuntivo
    Diagrama.crearDiagramaGantt

    Set FabricaDeRecorridos = New cFabricaDeRecorridos

    Set FabricaDeRecorridos.Modelo = Me
    
End Sub

Private Sub crearColeccionDeArcosDisyuntivos()

    ' crear coleccion con arcos disyuntivos (en misma maquina)
    Dim a As New cArco
    Dim i As Integer, j As Integer, k As Integer, indice As Integer
    indice = 0
    For i = 1 To cantMaquinas
        For j = 1 To cantOperacionesPorMaquina - 1
            For k = 1 To cantOperacionesPorMaquina - j
                Set a = New cArco
                
                indice = indice + 1
                
                Set a.InicioOperacion = OperacionesEnMismaMaquina(i, j)
                Set a.FinalOperacion = OperacionesEnMismaMaquina(i, j + k)
                Set a.Maquina = Maquinas(CStr(i))
                a.Tipo = "Disyuntivo"
                a.indice = indice
                
                'agregar este arco a los operaciones de inicio y fin
                a.InicioOperacion.Arcos.Add a
                a.FinalOperacion.Arcos.Add a
                
                ArcosDisyuntivos.Add a
            Next k
        Next j
    Next i
    
End Sub


Private Sub crearColeccionDeArcosConyuntivos()
    
    ' crear coleccion con arcos conyuntivos (en mismo trabajo)
    Dim a As New cArco
    Dim i As Integer, j As Integer, indice As Integer
    For i = 1 To cantTrabajos
        For j = 1 To cantOperacionesPorTrabajo - 1
            Set a = New cArco
                            
            indice = indice + 1
                
            ' es importante el orden, para las flechas del grafico despues (siempre se deben
            ' conectar del indice menor al mayor
            Set a.InicioOperacion = OperacionesConMismoTrabajo(i, j)
            Set a.FinalOperacion = OperacionesConMismoTrabajo(i, j + 1)
            Set a.Trabajo = Trabajos(CStr(i))
            a.Tipo = "Conyuntivo"
            a.indice = indice
            
            ' hay que darle la direccion al ppo
            a.Direccion = True
                
            'agregar este arco a los operaciones de inicio y fin
            a.InicioOperacion.Arcos.Add a
            a.FinalOperacion.Arcos.Add a
            
            ArcosConyuntivos.Add a
        Next j
    Next i
    
End Sub

Private Sub crearVectorConIgualesMaquinas()

    ' crear vector con operaciones en misma maquina
    ReDim Preserve pOperacionesEnMismaMaquina(0 To cantMaquinas, cantOperacionesPorMaquina)
    Dim i As Integer, j As Integer
    Dim o As cOperacion
    For i = 1 To cantMaquinas
        j = 1
        Set o = New cOperacion
        For Each o In Operaciones
            If o.Maquina.Numero = i Then
                Set OperacionesEnMismaMaquina(i, j) = o
                j = j + 1
            End If
        Next
    Next i
    
End Sub

Private Sub crearVectorConIgualesTrabajos()
    
    ' crear vector con operaciones de mismo trabajo
    ReDim Preserve pOperacionesConMismoTrabajo(cantTrabajos, cantOperacionesPorTrabajo)
    Dim i As Integer, j As Integer
    Dim o As cOperacion
    For i = 1 To cantTrabajos
        j = 1
        Set o = New cOperacion
        For Each o In Operaciones
            If o.Trabajo.Numero = i Then
                Set OperacionesConMismoTrabajo(i, j) = o
                j = j + 1
            End If
        Next
    Next i
    
End Sub

Private Sub crearOperaciones()

    Dim o As New cOperacion
    
    ' indice
    Dim k As Integer
    k = 1
    
    ' leer datos
    Dim i As Integer
    For i = 1 To cantTrabajos
        Dim j As Integer
        For j = 1 To cantOperacionesPorTrabajo
            Set o = New cOperacion
            Set o.Trabajo = Trabajos(CStr(PlanillaDatos.Range("DATOS").Cells(k, 2).Value))
            Set o.Maquina = Maquinas(CStr(PlanillaDatos.Range("DATOS").Cells(k, 3).Value))

            o.Duracion = PlanillaDatos.Range("DATOS").Cells(k, 4).Value
            o.OrdenEnElTrabajo = j
            o.indice = k
            Operaciones.Add o, CStr(k)
            k = k + 1
        Next j
    Next i
    
End Sub

' inicializar coleccion de maquinas
Private Sub crearMaquinas()
    
    Dim i As Integer
    For i = 1 To cantMaquinas
        Dim m As New cMaquina
        Set m = New cMaquina
        m.Numero = i
        m.Color = Round(((maxColor - minColor) / cantTrabajos) * i + minColor, 0)
        
        ' numero de maquina es el key (string) en la colleccion
        Maquinas.Add m, CStr(i)
    Next i
    
End Sub

' inicializar coleccion de trabajos
Private Sub crearTrabajos()
    Dim dividerColor As Integer
    Dim i As Integer
    For i = 1 To cantTrabajos
        Dim t As New cTrabajo
        Set t = New cTrabajo
        t.Numero = i ' hago esto abajo para no tener los mismos colores que para mis maquinas
        
        If cantTrabajos Mod cantMaquinas = 0 Then
            dividerColor = Round(cantTrabajos * colorDif)
        Else
            dividerColor = cantTrabajos
        End If
        
        t.Color = Round(((maxColor - minColor) / dividerColor) * i * colorDif + minColor, 0)
        ' numero de trabajo es el key (string) en la colleccion
        Trabajos.Add t, CStr(i)
    Next i
    
End Sub

Public Sub removerGraficos()
    
    Dim shp As Shape
    For Each shp In PlanillaGraficoDisyuntivo.Shapes
        If shp.Type = msoAutoShape Or shp.Type = msoTextBox Then shp.Delete
    Next shp
        
    For Each shp In PlanillaGraficoGantt.Shapes
        If shp.Type = msoAutoShape Or shp.Type = msoTextBox Then shp.Delete
    Next shp
    
End Sub

Public Function dameArcoEntreOperaciones(o1 As cOperacion, o2 As cOperacion) As cArco
    
    Dim a As cArco
    For Each a In sumarColecciones(ArcosDisyuntivos, ArcosConyuntivos)
        If (a.InicioOperacion Is o1 And a.FinalOperacion Is o2) Or _
                (a.FinalOperacion Is o1 And a.InicioOperacion Is o2) Then Set dameArcoEntreOperaciones = a
    Next
    
End Function



'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************

Property Get FabricaDeRecorridos() As cFabricaDeRecorridos
    Set FabricaDeRecorridos = pFabricaDeRecorridos
End Property

Property Set FabricaDeRecorridos(f As cFabricaDeRecorridos)
    Set pFabricaDeRecorridos = f
End Property

Property Get Diagrama() As cDiagrama
    Set Diagrama = pDiagrama
End Property

Property Set Diagrama(d As cDiagrama)
    Set pDiagrama = d
End Property

Property Get ArcosConyuntivos() As Collection
    Set ArcosConyuntivos = pArcosConyuntivos
End Property

Property Set ArcosConyuntivos(c As Collection)
    Set pArcosConyuntivos = c
End Property

Property Get ArcosDisyuntivos() As Collection
    Set ArcosDisyuntivos = pArcosDisyuntivos
End Property

Property Set ArcosDisyuntivos(c As Collection)
    Set pArcosDisyuntivos = c
End Property

Property Get OperacionesEnMismaMaquina(i As Integer, j As Integer) As cOperacion
    Set OperacionesEnMismaMaquina = pOperacionesEnMismaMaquina(i, j)
End Property

Property Set OperacionesEnMismaMaquina(i As Integer, j As Integer, o As cOperacion)
    Set pOperacionesEnMismaMaquina(i, j) = o
End Property

Property Get OperacionesConMismoTrabajo(i As Integer, j As Integer) As cOperacion
    Set OperacionesConMismoTrabajo = pOperacionesConMismoTrabajo(i, j)
End Property

Property Set OperacionesConMismoTrabajo(i As Integer, j As Integer, o As cOperacion)
    Set pOperacionesConMismoTrabajo(i, j) = o
End Property

Property Get Trabajos() As Collection
    Set Trabajos = pTrabajos
End Property

Property Set Trabajos(c As Collection)
    Set pTrabajos = c
End Property

Property Get Maquinas() As Collection
    Set Maquinas = pMaquinas
End Property

Property Set Maquinas(c As Collection)
    Set pMaquinas = c
End Property

Property Get Operaciones() As Collection
    Set Operaciones = pOperaciones
End Property

Property Set Operaciones(c As Collection)
    Set pOperaciones = c
End Property

Property Get PlanillaDatos() As Worksheet
    Set PlanillaDatos = pPlanillaDatos
End Property

Property Set PlanillaDatos(w As Worksheet)
    Set pPlanillaDatos = w
End Property

Property Get PlanillaGraficoDisyuntivo() As Worksheet
    Set PlanillaGraficoDisyuntivo = pPlanillaGraficoDisyuntivo
End Property

Property Set PlanillaGraficoDisyuntivo(w As Worksheet)
    Set pPlanillaGraficoDisyuntivo = w
End Property

Property Get PlanillaGraficoGantt() As Worksheet
    Set PlanillaGraficoGantt = pPlanillaGraficoGantt
End Property

Property Set PlanillaGraficoGantt(w As Worksheet)
    Set pPlanillaGraficoGantt = w
End Property

