Attribute VB_Name = "Main"
Option Explicit
Option Base 1

Public Sub init()
    Dim start As Long, finish As Long
    
    start = GetTickCount()
    
    Randomize
    
    performanceOn
    
    ' constantes
    separacionHorizontalCirculos = 125
    Factor_Gantt = 1
    zeroXRectangulo = 2
    
    ' crear & inicializar nuestro modelo
    Modelo.inicializar PlanillaDatos:=Worksheets("DATOS"), _
                    PlanillaGraficoDisyuntivo:=Worksheets("DISYUNTIVO"), _
                 PlanillaGraficoGantt:=Worksheets("GANTT")
    
    ' crear & inicializar la base metaheuristica, asignar el modelo
    Set Metaheuristica.Modelo = Modelo
    
    mInitInterfazArcos Modelo
    
    ' obtener una solucion inicial
    Dim SolucionesIniciales As New cFabricaDeSolucionesIniciales
    Set SolucionesIniciales.Metaheuristica = Metaheuristica
    SolucionesIniciales.MaxListaTabu = 5
    Set Metaheuristica.Solucion = _
        SolucionesIniciales.porListaTabu(longitud:=Modelo.ArcosDisyuntivos.count)
        ' la longitud de la solucion = cantidad de arcos disyuntivos

    Set Metaheuristica.MejorSolucion = Metaheuristica.Solucion
    Metaheuristica.implementarSolucion ' solucion inicial

    Modelo.actualizar
    Modelo.Diagrama.crearLineasSeparacion

'********************************************************************************
    ' crear & inicializar la busqueda local
    Set busquedaLocal.Metaheuristica = Metaheuristica
    busquedaLocal.MaxListaTabu = Worksheets("DATOS").Range("TABU").Cells(1, 1).Value
    Set busquedaLocal.Metaheuristica.Modelo = Modelo
    Set busquedaLocal.ListaTabu = New Collection

    Debug.Print GetTickCount() - start
    
End Sub

Public Sub step()
    busquedaLocal.start 1
    Metaheuristica.implementarMejorSolucion
    Modelo.actualizar
End Sub
