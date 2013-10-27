Attribute VB_Name = "Demo"
Option Explicit
Option Base 1

Public MaxListaTabu As Integer, maxListaTabuInicial As Integer
Public start As Long, finish As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Modelo As New cModeloDisyuntivo
Private Metaheuristica As New cMetaheuristica
Private SolucionesIniciales As New cFabricaDeSolucionesIniciales
Private busquedaLocal As New cBusquedaLocal
Private stepArcosMaquina As Integer
Private pausa As Boolean
Private countHastaDiv As Integer

Public Sub getTimeElapsed()
        Worksheets("DEMO").Range("TICK").Cells(1, 1).Value = (GetTickCount() - _
            start) / 1000
End Sub

Public Sub setTimeElapsedZero()
        Worksheets("DEMO").Range("TICK").Cells(1, 1).Value = 0
End Sub

Public Sub init()
    Dim start As Long, finish As Long
    
    Randomize

    ' constantes
    separacionHorizontalCirculos = 100
    Factor_Gantt = 5
    zeroXRectangulo = 430

    ' crear & inicializar nuestro modelo

    Modelo.inicializar PlanillaDatos:=Worksheets("DEMO"), _
                       PlanillaGraficoDisyuntivo:=Worksheets("DEMO"), _
                       PlanillaGraficoGantt:=Worksheets("DEMO")

    ' inicializar el modulo para poder rapidamente hacer los calculos
    mInitInterfazArcos Modelo

    ' crear & inicializar la metaheuristica, asignar el modelo
    Set Metaheuristica.Modelo = Modelo
    
    ' obtener una solucion inicial
    Set SolucionesIniciales.Metaheuristica = Metaheuristica
    SolucionesIniciales.MaxListaTabu = 5
    
    ' o al azar (lista tabu) o predefinida
'********************************************************************************
    If ActiveSheet.Shapes("Check Box 11").ControlFormat.Value = 1 Then

        Set Metaheuristica.Solucion = _
        SolucionesIniciales.porListaTabu(longitud:=Modelo.ArcosDisyuntivos.count)
    Else
        Set Metaheuristica.Solucion = _
        SolucionesIniciales.porTodos1s(longitud:=Modelo.ArcosDisyuntivos.count)
    End If
'********************************************************************************
    
    ' la longitud de la solucion = cantidad de arcos disyuntivos
    Set Metaheuristica.MejorSolucion = Metaheuristica.Solucion
    Metaheuristica.implementarSolucion    ' solucion inicial
    Modelo.actualizar
    Modelo.Diagrama.crearLineasSeparacion ' las lineas deberian estar espaciados con rspecto a la 1er sol
    
    ' crear & inicializar la busqueda local
    Set busquedaLocal.Metaheuristica = Metaheuristica
    busquedaLocal.MaxListaTabu = Worksheets("DEMO").Range("TABU").Cells(1, 1).Value
    Set busquedaLocal.Metaheuristica.Modelo = Modelo
    Set busquedaLocal.ListaTabu = New Collection

    setTimeElapsedZero

End Sub

Sub startBusqueda()
    Worksheets("DEMO").Range("TICK").Cells(1, 1).Value = 0
    start = GetTickCount()
    pausa = False
    While Metaheuristica.MejorSolucion.Funcional > 60 And pausa = False
        DoEvents
        step
        
        If pausa = True Then
            Exit Sub
        End If
            
    Wend
    getTimeElapsed
    End
End Sub

Sub pausaBusqueda()
    pausa = True
    Metaheuristica.implementarMejorSolucion
    Modelo.actualizar
End Sub

Sub stopBusqueda()
    getTimeElapsed
    Metaheuristica.implementarMejorSolucion
    Modelo.actualizar
    End
End Sub


Public Sub step()
' mejor solucion y actualizar la planilla
    
    ' dsps de tantas repeticiones hacemos una solucion totalmente al azar
    countHastaDiv = countHastaDiv + 1
    If countHastaDiv > 30 Then
        Set Metaheuristica.Solucion = _
            SolucionesIniciales.porListaTabu(longitud:=Modelo.ArcosDisyuntivos.count)
        countHastaDiv = 0
    Else
        busquedaLocal.start 1
    End If
    
    Metaheuristica.implementarSolucion
    Modelo.actualizar
End Sub

Public Sub stepArcos()
    If stepArcosMaquina = 4 Then stepArcosMaquina = 0
    stepArcosMaquina = stepArcosMaquina + 1
    Modelo.Diagrama.mostrarArcosDeMaquina Modelo.Maquinas(stepArcosMaquina)
End Sub
