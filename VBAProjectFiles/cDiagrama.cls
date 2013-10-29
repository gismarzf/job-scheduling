VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDiagrama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pCirculos As New Collection
Private pConectores As New Collection
Private pRectangulos As New Collection
Private pOperaciones As New Collection
Private pArcos As New Collection
Private pModelo As cModeloDisyuntivo
Private pcolLineasSeparacion As New Collection

Private pPlanillaGraficoDisyuntivo As New Worksheet
Private pPlanillaGraficoGantt As New Worksheet

Private Sub Class_Terminate()
    Set pcolLineasSeparacion = Nothing
End Sub


'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************


Public Property Get colLineasSeparacion() As Collection
    Set pcolLineasSeparacion = pcolLineasSeparacion
End Property

Public Property Set colLineasSeparacion(colpcolLineasSeparacion As Collection)
    Set pcolLineasSeparacion = colpcolLineasSeparacion
End Property

Property Get Modelo() As cModeloDisyuntivo
    Set Modelo = pModelo
End Property

Property Set Modelo(m As cModeloDisyuntivo)
    Set pModelo = m
End Property

Property Get Operaciones() As Collection
    Set Operaciones = pOperaciones
End Property

Property Set Operaciones(c As Collection)
    Set pOperaciones = c
End Property

Property Get Arcos() As Collection
    Set Arcos = pArcos
End Property

Property Set Arcos(c As Collection)
    Set pArcos = c
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

Property Get Rectangulos() As Collection
    Set Rectangulos = pRectangulos
End Property

Property Set Rectangulos(c As Collection)
    Set pRectangulos = c
End Property

Property Get Circulos() As Collection
    Set Circulos = pCirculos
End Property

Property Set Circulos(c As Collection)
    Set pCirculo = s
End Property

Property Get Conectores() As Collection
    Set Conectores = pConectores
End Property

Property Set Conectores(c As Collection)
    Set pConectores = c
End Property

'********************************************************************************
'FUNCIONES
'********************************************************************************

Public Sub mandarAlFrenteCirculos()
    Dim c As cDiagramaCirculo
    For Each c In Circulos
        c.miShape.mandarAlFrente
    Next
End Sub

Public Sub mandarAtrasCirculos()
    Dim c As cDiagramaCirculo
    For Each c In Circulos
        c.miShape.mandarAtras
    Next
End Sub

Public Sub mandarAlFrenteConectores(Tipo As String)
    Dim c As cDiagramaConector
    For Each c In Conectores
        If c.arco.Tipo = Tipo Then c.miShape.mandarAlFrente
    Next
End Sub

Public Sub mandarAtrasConectores(Tipo As String)
    Dim c As cDiagramaConector
    For Each c In Conectores
        If c.arco.Tipo = Tipo Then c.miShape.mandarAtras
    Next
    
End Sub


Public Sub crearCirculos()
    
    Dim Circulo As cDiagramaCirculo
    
    Dim o As cOperacion
    For Each o In Operaciones
        Set Circulo = New cDiagramaCirculo
    
        ' definir campos
        Circulo.Width = widthCirculo
        Circulo.Height = heightCirculo
        Circulo.x = zeroXCirculo + (o.OrdenEnElTrabajo - 1) * separacionHorizontalCirculos
        Circulo.y = zeroYCirculo + (o.Trabajo.Numero - 1) * separacionVerticalCirculos
        
        ' crear shape
        Set Circulo.miShape = New cDiagramaShape
        Set Circulo.miShape.Shape = PlanillaGraficoDisyuntivo.Shapes.AddShape(msoShapeOval, Circulo.x, Circulo.y, _
                                                                            Circulo.Width, Circulo.Height)

        With Circulo.miShape.Shape.TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .MarginLeft = 0
            .MarginRight = 0
            .TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
            .TextRange.Font.Size = 16
            .TextRange.Font.Shadow.Type = msoShadow21
            .TextRange.Characters.Text = o.indice
        End With
        
         
        Circulo.miShape.Shape.Fill.ForeColor.RGB = o.Maquina.Color
        Circulo.miShape.Shape.Line.ForeColor.RGB = o.Maquina.Color
        
        Circulo.miShape.Color = o.Maquina.Color
        
        Set Circulo.operacion = o
        Set o.Circulo = Circulo
        
        ' agregar el circulo al diagrama
        Circulos.Add Circulo
    Next
End Sub

Public Sub crearConectores()
    
    Dim Conector As cDiagramaConector
    Dim a As cArco
    
    For Each a In Arcos
    
        Set Conector = New cDiagramaConector

        Set Conector.miShape = New cDiagramaShape
        
        ' ubicacion no importa ya que reconectamos despues
        Set Conector.miShape.Shape = PlanillaGraficoDisyuntivo.Shapes.AddConnector(msoConnectorStraight, _
                                                                                            1, 1, 2, 2)
        With Conector.miShape.Shape.ConnectorFormat
             .BeginConnect ConnectedShape:=a.InicioOperacion.Circulo.miShape.Shape, ConnectionSite:=1
             .EndConnect ConnectedShape:=a.FinalOperacion.Circulo.miShape.Shape, ConnectionSite:=1
        End With
    
        With Conector.miShape.Shape.Line
            .EndArrowheadLength = msoArrowheadLong
            .EndArrowheadWidth = msoArrowheadWide
            .EndArrowheadStyle = msoArrowheadStealth
            .visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent5
            .Weight = weightLineHidden
        End With
        
        Conector.miShape.Shape.RerouteConnections
       ' Conector.miShape.invisible
        
        Set Conector.arco = a
        Set a.Conector = Conector
        
        Conectores.Add Conector
    
    Next
    
End Sub

Public Sub crearRectangulos()
    
    Set Rectangulos = New Collection
    
    Dim Rectangulo As cDiagramaRectangulo
    
    
    Dim o As cOperacion
    For Each o In Operaciones
    
        Set Rectangulo = New cDiagramaRectangulo
        
        ' definir campos
        Rectangulo.Height = heightRectangulo
        Rectangulo.Width = o.Duracion * Factor_Gantt
        Rectangulo.x = zeroXRectangulo + o.HorarioInicio * Factor_Gantt
        Rectangulo.y = distanciaTop + zeroYRectangulo + _
                    (separacionVerticalRectangulos + heightRectangulo) * (o.Maquina.Numero - 1)

        Set Rectangulo.miShape = New cDiagramaShape
        Set Rectangulo.miShape.Shape = PlanillaGraficoGantt.Shapes.AddShape(msoShapeRectangle, Rectangulo.x, _
                                                        Rectangulo.y, Rectangulo.Width, Rectangulo.Height)

        With Rectangulo.miShape.Shape.TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .MarginLeft = 0
            .MarginRight = 0
            .TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
            .TextRange.Font.Size = 16
            .TextRange.Font.Shadow.Type = msoShadow21
            .TextRange.Characters.Text = o.indice
        End With


        Rectangulo.miShape.Shape.Fill.ForeColor.RGB = o.Trabajo.Color
        Rectangulo.miShape.Shape.Line.ForeColor.RGB = o.Trabajo.Color
        Rectangulo.miShape.Color = o.Trabajo.Color
        
        Set o.Rectangulo = Rectangulo
        Set Rectangulo.operacion = o
        
        Rectangulos.Add Rectangulo
        
    Next

End Sub

Public Sub crearDiagramaDisyuntivo()
    
    crearCirculos
    crearConectores
    mandarAlFrenteCirculos
    
End Sub

Public Sub crearDiagramaGantt()
    
    crearRectangulos
    
End Sub

Public Sub actualizarRecorridoCritico()

    ' tengo que borrar las marcas del recorrido critico anterior
    Dim o As cOperacion
    For Each o In Operaciones
        o.Circulo.miShape.marcarRemover
        o.Rectangulo.miShape.marcarRemover
    Next

    Dim a As cArco
    For Each a In Arcos
        a.Conector.marcarLineaRemover
    Next

    marcarRecorrido Modelo.FabricaDeRecorridos.RecorridoCritico
    
End Sub

Public Sub actualizarPlanilla()

    Dim i As Integer
    For i = 1 To Modelo.Operaciones.count
        Modelo.PlanillaDatos.Range("DATOS").Cells(i, 5).Value = Modelo.Operaciones(i).HorarioInicio
        Modelo.PlanillaDatos.Range("DATOS").Cells(i, 6).Value = Modelo.Operaciones(i).HorarioFin
    Next i
    
    Modelo.PlanillaDatos.Range("FUNCIONAL").Cells(1, 1).Value = _
                    Modelo.FabricaDeRecorridos.RecorridoCritico.Suma
End Sub

Public Sub marcarRecorrido(rec As cRecorrido)

    Dim o As cOperacion
    For Each o In rec.Recorrido
        o.Circulo.miShape.marcar
        o.Circulo.miShape.mandarAlFrente
        o.Rectangulo.miShape.marcar
        o.Rectangulo.miShape.mandarAlFrente
    Next
    
    Dim a As cArco
    For Each a In rec.Arcos
        a.Conector.marcarLinea
    Next
    
End Sub

Public Sub marcarRemoverRecorridos(rec As cRecorrido)
    
    Dim o As cOperacion
    For Each o In rec.Recorrido
        o.Circulo.miShape.marcarRemover
        o.Rectangulo.miShape.marcarRemover
    Next
    
    Dim a As cArco
    For Each a In rec.Arcos
        a.Conector.marcarLineaRemover
    Next
    
End Sub

Public Sub actualizarDiagramaDisyuntivo()

    Dim c As cDiagramaConector
    For Each c In Conectores
        c.actualizarConector
    Next

    
End Sub

Public Sub actualizarDiagramaGantt()
    Dim r As cDiagramaRectangulo
    For Each r In Rectangulos
        r.miShape.Shape.Left = zeroXRectangulo + r.operacion.HorarioInicio * Factor_Gantt
    Next
End Sub

Public Sub crearLineasSeparacion()
        
    ' buscar el rectangulo que es mas a la derecha
    Dim rder As cDiagramaRectangulo
    Dim rbot As cDiagramaRectangulo
    Set rder = Rectangulos(1) ' eligo uno para empezar
    Set rbot = Rectangulos(1)
    
    Dim r As cDiagramaRectangulo
    For Each r In Rectangulos
        If r.miShape.Shape.Left > rder.miShape.Shape.Left Then
            Set rder = r
        End If
        
        If r.miShape.Shape.Top > rbot.miShape.Shape.Top Then
            Set rbot = r
        End If
    Next
    
    Dim maxdist As Integer
    ' este es el punto mas lejano del inicio
    maxdist = rder.miShape.Shape.Left
    
    Dim spacing As Double
    ' spacing es la distancia entre cada linea de separacion
    spacing = maxdist / numLineasSepaGantt
    
    Dim i As Integer
    For i = 1 To numLineasSepaGantt
        Dim lineaSep As New cDiagramaLineaSep
        lineaSep.StartX = zeroXRectangulo + (i - 1) * spacing
        lineaSep.StartY = zeroYRectangulo
        lineaSep.EndX = zeroXRectangulo + (i - 1) * spacing
        lineaSep.EndY = rbot.miShape.Shape.Top + heightRectangulo
        lineaSep.Tiempo = (i - 1) * (spacing / Factor_Gantt)
        lineaSep.crear
        pcolLineasSeparacion.Add lineaSep
    Next i
    
End Sub

Public Sub mostrarArcosDeOperacion(o As cOperacion)
    
    Dim a As cArco
    For Each a In o.Arcos
        a.Conector.marcarLinea
    Next
    
End Sub

Public Sub mostrarArcosDeMaquina(m As cMaquina)

    ocultarArcos

    Dim a As cArco
    For Each a In Arcos
        If a.Maquina Is m Then a.Conector.marcarLinea
    Next
    
End Sub

Public Sub mostrarRecorrido(r As cRecorrido)
    
    ocultarArcos

    Dim a As cArco
    For Each a In r.Arcos
        a.Conector.marcarLinea
    Next
    
End Sub

Public Sub ocultarArcos()
    
    Dim a As cArco
    For Each a In Arcos
        a.Conector.marcarLineaRemover
    Next
    
End Sub

