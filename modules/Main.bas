Attribute VB_Name = "Main"
Option Explicit
Option Base 1

Public Sub TestAlfa()
    Dim start As Long, finish As Long
    
    Randomize
    
    performanceOn
    
    ' constantes
    separacionHorizontalCirculos = 125
    Factor_Gantt = 1
    zeroXRectangulo = 2
    
    ' crear & inicializar nuestro modelo
    Dim Modelo As New cModeloDisyuntivo
    Modelo.inicializar PlanillaDatos:=Worksheets("DATOS"), _
                    PlanillaGraficoDisyuntivo:=Worksheets("DISYUNTIVO"), _
                 PlanillaGraficoGantt:=Worksheets("GANTT")
    
    ' crear & inicializar la base metaheuristica, asignar el modelo
    Dim Metaheuristica As New cMetaheuristica
    Set Metaheuristica.Modelo = Modelo
    
    ' obtener una solucion inicial
    Dim SolucionesIniciales As New cFabricaDeSolucionesIniciales
    Set SolucionesIniciales.Metaheuristica = Metaheuristica
    SolucionesIniciales.MaxListaTabu = 5
    
    ' setear la solucion inicial a la metaheuristica
    Set Metaheuristica.Solucion = _
        SolucionesIniciales.porListaTabu(longitud:=Modelo.ArcosDisyuntivos.count)
        ' la longitud de la solucion = cantidad de arcos disyuntivos
    Set Metaheuristica.MejorSolucion = Metaheuristica.Solucion
    Metaheuristica.implementarSolucion ' solucion inicial

    Modelo.actualizar
    

    ' crear & inicializar la busqueda local
    Dim busquedaLocal As New cBusquedaLocal
    Set busquedaLocal.Metaheuristica = Metaheuristica
    busquedaLocal.MaxListaTabu = Worksheets("DATOS").Range("TABU").Cells(1, 1).Value
    Set busquedaLocal.Metaheuristica.Modelo = Modelo
    busquedaLocal.MaxListaTabu = 10
    Set busquedaLocal.ListaTabu = New Collection
    
    ' start busqueda local
    start = GetTickCount() ' benchmark
    busquedaLocal.start Worksheets("DATOS").Range("ITERACIONES").Cells(1, 1).Value
    finish = GetTickCount() ' benchmark
    
    
    ' mejor solucion y actualizar la planilla
    Metaheuristica.implementarMejorSolucion
    Modelo.actualizar

    ' benchmark
    Worksheets("DATOS").Range("TICK").Cells(1, 1).Value = finish - start
    
    performanceOff
    
End Sub

Public Sub Benchmark()

    Dim v As New Collection
    Dim i As Integer
    For i = 1 To Worksheets("DEMO").Range("BENCHMARK").Cells(1, 1).Value
        v.Add test
    Next i
                
                
    Dim fmedio As Integer, tmedio As Integer
    For i = 1 To v.count
        fmedio = fmedio + v(i)(1, 1)
        tmedio = tmedio + v(i)(1, 2)
    Next i
    
    Worksheets("DEMO").Range("FMEDIO").Cells(1, 1).Value = fmedio / v.count
    Worksheets("DEMO").Range("TMEDIO").Cells(1, 1).Value = tmedio / v.count
        
End Sub
