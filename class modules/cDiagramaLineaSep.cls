VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDiagramaLineaSep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pStartX As Integer
Private pStartY As Integer
Private pEndX As Integer
Private pEndY As Integer
Private pTiempo As Integer
Private pobjShape As Shape
Private pobjText As Shape

Private Sub Class_Terminate()
    Set pobjShape = Nothing
End Sub


'********************************************************************************
'FUNCIONES
'********************************************************************************

Public Sub crear()
    Set Shape = Worksheets("DEMO").Shapes.AddConnector(msoConnectorStraight, StartX, StartY, EndX, _
        EndY)
    Shape.Reflection.Type = msoReflectionType1
    With Shape
        .Line.visible = msoTrue
        .Line.DashStyle = msoLineSolid
        .Line.Weight = 1.25
        .ZOrder msoSendToBack
        .Line.ForeColor.TintAndShade = 0
        .Line.ForeColor.Brightness = -0.5
        .Line.Transparency = 0
        .Line.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Line.Transparency = 0.6
    End With
    
    Set Text = Worksheets("DEMO").Shapes.AddTextbox(msoTextOrientationHorizontal, StartX, 0, _
        50, distanciaTop * 2)
    
    With Text
        .Fill.visible = msoFalse
        .Line.visible = msoFalse
        .TextFrame2.TextRange.Characters.Font.Size = 16
        .TextFrame2.TextRange.Characters.Text = Tiempo & "'"
    End With
    
End Sub

'********************************************************************************
'GETTERS/SETTERS
'********************************************************************************

Public Property Get Text() As Shape
    Set Text = pobjText
End Property

Public Property Set Text(objText As Shape)
    Set pobjText = objText
End Property

Public Property Get Shape() As Shape
    Set Shape = pobjShape
End Property

Public Property Set Shape(objShape As Shape)
    Set pobjShape = objShape
End Property

Public Property Get StartX() As Integer
    StartX = pStartX
End Property

Public Property Let StartX(ByVal iStartX As Integer)
    pStartX = iStartX
End Property

Public Property Get StartY() As Integer
    StartY = pStartY
End Property

Public Property Let StartY(ByVal iStartY As Integer)
    pStartY = iStartY
End Property

Public Property Get EndX() As Integer
    EndX = pEndX
End Property

Public Property Let EndX(ByVal iEndX As Integer)
    pEndX = iEndX
End Property

Public Property Get EndY() As Integer
    EndY = pEndY
End Property

Public Property Let EndY(ByVal iEndY As Integer)
    pEndY = iEndY
End Property

Public Property Get Tiempo() As Integer
    Tiempo = pTiempo
End Property

Public Property Let Tiempo(iTiempo As Integer)
    pTiempo = iTiempo
End Property
