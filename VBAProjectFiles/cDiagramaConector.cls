VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDiagramaConector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private pmiShape As cDiagramaShape
Private pInicioX As Integer, pInicioY As Integer, pFinalX As Integer, pFinalY As Integer
Private pInicioOperacion As cOperacion, pFinalOperacion As cOperacion
Private pArco As cArco

'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************

Property Get arco() As cArco
    Set arco = pArco
End Property

Property Set arco(a As cArco)
    Set pArco = a
End Property

Property Get miShape() As cDiagramaShape
    Set miShape = pmiShape
End Property

Property Set miShape(s As cDiagramaShape)
    Set pmiShape = s
End Property

Property Get InicioX() As Integer
    InicioX = pInicioX
End Property

Property Let InicioX(i As Integer)
    pInicioX = i
End Property

Property Get InicioY() As Integer
    InicioY = pInicioY
End Property

Property Let InicioY(i As Integer)
    pInicioY = i
End Property

Property Get FinalX() As Integer
    FinalX = pFinalX
End Property

Property Let FinalX(i As Integer)
    pFinalX = i
End Property

Property Get FinalY() As Integer
    FinalY = pFinalY
End Property

Property Let FinalY(i As Integer)
    pFinalY = i
End Property

Property Get InicioOperacion() As cOperacion
    Set InicioOperacion = pInicioOperacion
End Property

Property Set InicioOperacion(o As cOperacion)
    Set pInicioOperacion = o
End Property

Property Get FinalOperacion() As cOperacion
    Set FinalOperacion = pFinalOperacion
End Property

Property Set FinalOperacion(o As cOperacion)
    Set pFinalOperacion = o
End Property

'********************************************************************************
'FUNCIONES
'********************************************************************************


Public Sub actualizarConector()

    With miShape.Shape.ConnectorFormat
         .BeginConnect ConnectedShape:=arco.InicioOperacion.Circulo.miShape.Shape, ConnectionSite:=1
         .EndConnect ConnectedShape:=arco.FinalOperacion.Circulo.miShape.Shape, ConnectionSite:=1
    End With
    
    miShape.Shape.RerouteConnections
    
End Sub

Public Sub marcarLinea()
    With miShape.Shape.Line
        .visible = msoTrue
        .Weight = 3
        .DashStyle = msoLineSysDash
        .BeginArrowheadStyle = msoArrowheadNone
        .EndArrowheadLength = msoArrowheadLong
        .EndArrowheadStyle = msoArrowheadStealth
        .EndArrowheadWidth = msoArrowheadWide
        .DashStyle = msoLineSysDash
        .Weight = weightLineNormal + 2
        .ForeColor.RGB = glowColor
    End With
End Sub

Public Sub marcarLineaRemover()
    With miShape.Shape.Line
        .visible = msoFalse
    End With
End Sub

Public Sub makeActive()
    With miShape.Shape.Line
        .visible = msoCTrue
        .ForeColor.RGB = activeColor
    End With
End Sub

Public Sub makeInactive()
    With miShape.Shape.Line
        .visible = msoFalse
        .ForeColor.RGB = glowColor
    End With
End Sub

