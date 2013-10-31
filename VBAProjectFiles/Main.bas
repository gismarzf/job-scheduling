Attribute VB_Name = "Main"
Option Explicit
Option Base 1

Private planillaDatos As Worksheet
Private planillaDisyuntivo As Worksheet
Private planillaGantt As Worksheet

Private Sub actualizarModeloConSolucion()
    mImplementarBitVector Metaheuristica.Solucion.BitVector
    Metaheuristica.implementarSolucion
    Modelo.actualizar
End Sub

Private Sub ajusteDeDimensionesParaPlanillaDemo()
    separacionHorizontalCirculos = 100
    Factor_Gantt = 5
    zeroXRectangulo = 430
    Set planillaDatos = Worksheets("DEMO")
    Set planillaDisyuntivo = Worksheets("DEMO")
    Set planillaGantt = Worksheets("DEMO")
    
    planillaDatos.Range("TICK").Cells(1, 1).Value = 0
    
End Sub

Private Sub ajusteDeDimensionesParaPlanillaDatos()
    separacionHorizontalCirculos = 125
    Factor_Gantt = 1
    zeroXRectangulo = 2
    Set planillaDatos = Worksheets("DATOS")
    Set planillaDisyuntivo = Worksheets("DISYUNTIVO")
    Set planillaGantt = Worksheets("GANTT")
    
    planillaDatos.Range("TICK").Cells(1, 1).Value = 0
End Sub

Public Sub initMain()
    Randomize
    
    ajusteDeDimensionesParaPlanillaDatos
    
    initModelo
    initMetaheuristica
    initListaTabu
    
    obtenerSolucionInicialAlAzar
End Sub

Public Sub initDemo()
    Randomize
    
    ajusteDeDimensionesParaPlanillaDemo
    
    initModelo
    initMetaheuristica
    initListaTabu
    
    obtenerSolucionInicialAlAzar
End Sub

Private Sub initListaTabu()
    Set listaTabu.Metaheuristica = Metaheuristica
    listaTabu.MaxListaTabu = planillaDatos.Range("TABU").Cells(1, 1).Value
    Set listaTabu.Metaheuristica.Modelo = Modelo
    
    Set listaTabu.listaTabu = New Collection ' porque AHORA???
End Sub

Private Sub initMetaheuristica()
    ' crear & inicializar la metaheuristica, asignar el modelo
    Set Metaheuristica.Modelo = Modelo
    Set Metaheuristica.MejorSolucion = Metaheuristica.Solucion

End Sub

Private Sub initModelo()
    ' crear & inicializar nuestro modelo
    Modelo.inicializar planillaDatos:=planillaDatos, _
                       PlanillaGraficoDisyuntivo:=planillaDisyuntivo, _
                       PlanillaGraficoGantt:=planillaGantt

    ' inicializar el modulo para poder rapidamente hacer los calculos
    mInitInterfazArcos Modelo
End Sub

Private Sub obtenerSolucionInicialAlAzar()

    Set SolucionesIniciales.Metaheuristica = Metaheuristica
    SolucionesIniciales.MaxListaTabu = 5
    
    Set Metaheuristica.Solucion = _
        SolucionesIniciales.porListaTabu(longitud:=Modelo.ArcosDisyuntivos.count)
    
    actualizarModeloConSolucion
    Modelo.Diagrama.crearLineasSeparacion
    ' las lineas tengo que hacer ahora, ya que el espaciado depende de la
    ' primera solucion

    Set Metaheuristica.MejorSolucion = Metaheuristica.Solucion

End Sub

Private Sub obtenerSolucionInicialTodos1s()
    
    ' obtener una solucion inicial
    Set SolucionesIniciales.Metaheuristica = Metaheuristica
    SolucionesIniciales.MaxListaTabu = 5
    
    Set Metaheuristica.Solucion = _
        SolucionesIniciales.porListaTabu(longitud:=Modelo.ArcosDisyuntivos.count)

    
    actualizarModeloConSolucion
    Modelo.Diagrama.crearLineasSeparacion
    ' las lineas tengo que hacer ahora, ya que el espaciado depende de la
    ' primera solucion
    
    
    Set Metaheuristica.MejorSolucion = Metaheuristica.Solucion

End Sub

Sub pausaBusqueda()
    pausa = True
    Metaheuristica.implementarMejorSolucion
    Modelo.actualizar
End Sub

Sub startBusqueda()

    planillaDatos.Range("TICK").Cells(1, 1).Value = 0
    
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

Public Sub step()
' mejor solucion y actualizar la planilla
    
    ' dsps de tantas repeticiones hacemos una solucion totalmente al azar
    countHastaDiv = countHastaDiv + 1
    If countHastaDiv > 30 Then
        Set Metaheuristica.Solucion = _
            SolucionesIniciales.porListaTabu(longitud:=Modelo.ArcosDisyuntivos.count)
        countHastaDiv = 0
    Else
        listaTabu.start 1
    End If
    
    Metaheuristica.implementarSolucion
    Modelo.actualizar
End Sub

Sub stopBusqueda()
    getTimeElapsed
    Metaheuristica.implementarMejorSolucion
    Modelo.actualizar
    End
End Sub

